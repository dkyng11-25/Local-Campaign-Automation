from __future__ import annotations

import ast
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

import pandas as pd
import requests


# =========================================================
# 1. Configuration
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
LOCAL_MEDIA_ROOT = BASE_DIR / "media"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOCAL_MEDIA_ROOT.mkdir(parents=True, exist_ok=True)

RAW_SHEET_NAMES = ["Raw Data_원문", "Raw Data_전략법인"]
CAMPAIGN_INPUT_SHEET_NAME = "campaign_input"
MEDIA_MANIFEST_SHEET_NAME = "media_manifest"

# Excel 기준 컬럼명 행. 예: 컬럼명이 Excel 2행이면 2
EXCEL_HEADER_ROW = 2

REQUEST_TIMEOUT = (10, 60)
PROBE_BYTES = 64 * 1024
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/142.0 Safari/537.36"
)


# =========================================================
# 2. Enum / Data Models
# =========================================================

class Platform(StrEnum):
    YOUTUBE = "youtube"
    INSTAGRAM = "instagram"
    FACEBOOK = "facebook"
    TWITTER = "twitter"
    UNKNOWN = "unknown"


class MediaType(StrEnum):
    VIDEO = "video"
    IMAGE = "image"
    LINK = "link"
    CAROUSEL = "carousel"
    UNKNOWN = "unknown"


@dataclass(frozen=True, slots=True)
class AccountInfo:
    account_id: str | None = None
    account_name: str | None = None
    account_handle: str | None = None
    profile_url: str | None = None


@dataclass(frozen=True, slots=True)
class SourceMedia:
    """Raw Data의 같은 순번에 있는 URL과 media type을 하나로 묶은 객체."""

    source_url: str
    media_type: MediaType


@dataclass(frozen=True, slots=True)
class StructuredMediaInput:
    """Excel의 게시글 한 행을 표준화한 게시글 단위 입력 객체."""

    campaign_id: str
    platform: Platform
    media_type: MediaType
    original_post_url: str | None
    conversation_stream: str | None
    account: AccountInfo
    source_medias: tuple[SourceMedia, ...] = ()
    source_sheet_name: str | None = None
    raw_row_number: int | None = None


@dataclass(slots=True)
class MediaAsset:
    """실제로 검사하고 다운로드할 개별 이미지 또는 영상."""

    campaign_id: str
    asset_id: str
    asset_index: int
    platform: Platform
    media_type: MediaType
    original_post_url: str | None
    source_url: str | None

    source_sheet_name: str | None = None
    raw_row_number: int | None = None

    final_url: str | None = None
    http_status: int | None = None
    content_type: str | None = None
    file_size_bytes: int | None = None

    local_path: Path | None = None
    extraction_method: str | None = None
    status: str = "pending"
    error_message: str | None = None


@dataclass(frozen=True, slots=True)
class UrlFeasibilityResult:
    feasible: bool
    final_url: str | None
    http_status: int | None
    content_type: str | None
    detected_media_type: MediaType
    extension: str | None
    error_message: str | None = None


# 표준 컬럼명 -> 원본 Excel 컬럼명
COLUMN_MAPPINGS: dict[str, str] = {
    "conversation_stream": "Conversation Stream",
    "campaign_id": "Campaign ID",
    "profile_url": "Profile URL",
    "user_name": "User Name",
    "permalink": "Permalink",
    "platform": "snType column",
    "media_type": "Media Type",
    "source_url": "Media URL",
}

# 이 컬럼들은 원본 시트에 반드시 존재해야 함
REQUIRED_STANDARD_COLUMNS = {
    "campaign_id",
    "platform",
    "source_url",
}


# =========================================================
# 3. 날짜 기반 경로 생성
# =========================================================

def build_excel_paths(input_date: str) -> tuple[Path, Path]:
    input_date = input_date.strip()

    try:
        datetime.strptime(input_date, "%y%m%d")
    except ValueError as exc:
        raise ValueError(
            "날짜는 YYMMDD 형식으로 입력해야 합니다. 예시) 260714"
        ) from exc

    input_excel_path = (
        OUTPUT_DIR
        / f"{input_date}_SLCC_SOV_Local Campaign Tracking_7월_v01.xlsx"
    )
    output_excel_path = OUTPUT_DIR / f"{input_date}_campaign_media_result.xlsx"

    if not input_excel_path.exists():
        raise FileNotFoundError(
            f"입력 Excel 파일을 찾을 수 없습니다: {input_excel_path}"
        )

    return input_excel_path, output_excel_path


# =========================================================
# 4. Excel Input / Sheet Functions
# =========================================================

def load_raw_sheets(
    excel_path: Path,
    sheet_names: list[str],
    header_row: int = 1,
) -> dict[str, pd.DataFrame]:
    """header_row는 Excel 기준 1부터 시작한다."""

    if header_row < 1:
        raise ValueError("header_row는 1 이상의 Excel 행 번호여야 합니다.")

    sheet_data = pd.read_excel(
        excel_path,
        sheet_name=sheet_names,
        header=header_row - 1,
    )

    return {
        sheet_name: dataframe.copy()
        for sheet_name, dataframe in sheet_data.items()
    }


def prepare_input_df(
    raw_df: pd.DataFrame,
    sheet_name: str,
    column_mapping: dict[str, str],
    header_row: int = 1,
) -> pd.DataFrame:
    """필요 컬럼을 표준명으로 변경한다. 선택 컬럼이 없으면 pd.NA를 넣는다."""

    missing_required_columns = [
        column_mapping[standard_name]
        for standard_name in REQUIRED_STANDARD_COLUMNS
        if (
            standard_name not in column_mapping
            or column_mapping[standard_name] not in raw_df.columns
        )
    ]

    if missing_required_columns:
        raise ValueError(
            f"[{sheet_name}] 필수 컬럼이 없습니다: {missing_required_columns}"
        )

    # Mapping 가능한 컬럼 dict 생성
    available_mapping = {
        standard_name: original_name
        for standard_name, original_name in column_mapping.items()
        if original_name in raw_df.columns
    }

    # 중복 원본 컬럼명이 매핑될 경우 한 번만 선택
    selected_original_columns = list(
        dict.fromkeys(available_mapping.values())
    )
    prepared_df = raw_df[selected_original_columns].copy()

    rename_mapping = {
        original_name: standard_name
        for standard_name, original_name in available_mapping.items()
    }
    prepared_df = prepared_df.rename(columns=rename_mapping)

    # 시트에 없던 선택 컬럼도 최종 스키마에는 포함
    for standard_name in column_mapping:
        if standard_name not in prepared_df.columns:
            prepared_df[standard_name] = pd.NA

    prepared_df = prepared_df[list(column_mapping.keys())]
    prepared_df["source_sheet"] = sheet_name

    # DataFrame index 0은 컬럼명 행의 다음 Excel 행
    prepared_df["raw_row_number"] = raw_df.index + header_row + 1

    return prepared_df


def build_media_input_dataframe(
    raw_sheets: dict[str, pd.DataFrame],
    column_mapping: dict[str, str],
    header_row: int = 1,
) -> pd.DataFrame:
    prepared_dataframes: list[pd.DataFrame] = []

    for sheet_name, raw_df in raw_sheets.items():
        prepared_df = prepare_input_df(
            raw_df=raw_df,
            sheet_name=sheet_name,
            column_mapping=column_mapping,
            header_row=header_row,
        )
        prepared_dataframes.append(prepared_df)

    if not prepared_dataframes:
        return pd.DataFrame()

    return pd.concat(
        prepared_dataframes,
        ignore_index=True,
        sort=False,
    )


def clean_input_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned_df = df.copy()

    # media_type/source_url은 list 또는 줄바꿈 문자열일 수 있으므로
    # 여기서 무조건 astype("string") 하지 않는다.
    scalar_string_columns = [
        "conversation_stream",
        "campaign_id",
        "profile_url",
        "user_name",
        "permalink",
        "platform",
        "source_sheet",
    ]

    for column in scalar_string_columns:
        if column not in cleaned_df.columns:
            continue

        cleaned_df[column] = (
            cleaned_df[column]
            .astype("string")
            .str.strip()
        )

    cleaned_df = cleaned_df.replace(
        {
            "": pd.NA,
            "nan": pd.NA,
            "None": pd.NA,
        }
    )

    cleaned_df = cleaned_df.dropna(
        subset=["campaign_id"]
    ).copy()

    return cleaned_df


# =========================================================
# 5. Input Mapper Helpers
# =========================================================

def optional_text(value: Any) -> str | None:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    return text or None


def identify_platform(platform_value: Any) -> Platform:
    text = optional_text(platform_value)

    if text is None:
        return Platform.UNKNOWN

    platform_mapping = {
        "youtube": Platform.YOUTUBE,
        "instagram": Platform.INSTAGRAM,
        "facebook": Platform.FACEBOOK,
        "twitter": Platform.TWITTER,
        "x": Platform.TWITTER,
    }

    return platform_mapping.get(
        text.lower(),
        Platform.UNKNOWN,
    )


def identify_media_type(media_type_value: Any) -> MediaType:
    text = optional_text(media_type_value)

    if text is None:
        return MediaType.UNKNOWN

    media_type_mapping = {
        "video": MediaType.VIDEO,
        "photo": MediaType.IMAGE,
        "image": MediaType.IMAGE,
        "picture": MediaType.IMAGE,
        "link": MediaType.LINK,
        "carousel": MediaType.CAROUSEL,
    }

    return media_type_mapping.get(
        text.lower(),
        MediaType.UNKNOWN,
    )


def normalize_cell_values(value: Any) -> tuple[str, ...]:
    """단일값, 리스트, 리스트 문자열, 줄바꿈 문자열을 문자열 tuple로 통일."""

    if value is None:
        return ()

    if isinstance(value, (list, tuple, set)):
        normalized_values: list[str] = []

        for item in value:
            text = optional_text(item)
            if not text:
                continue

            normalized_values.extend(
                line.strip()
                for line in text.splitlines()
                if line.strip()
            )

        return tuple(normalized_values)

    try:
        if pd.isna(value):
            return ()
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    if not text:
        return ()

    # Excel에 "['video', 'photo']" 같은 문자열로 저장된 경우도 지원
    if text.startswith("[") and text.endswith("]"):
        try:
            parsed_value = ast.literal_eval(text)
        except (ValueError, SyntaxError):
            parsed_value = None

        if isinstance(parsed_value, (list, tuple, set)):
            return normalize_cell_values(parsed_value)

    return tuple(
        line.strip()
        for line in text.splitlines()
        if line.strip()
    )


def normalize_source_medias(
    source_url_value: Any,
    media_type_value: Any,
) -> tuple[SourceMedia, ...]:
    """같은 순번의 URL과 media type을 SourceMedia로 묶는다."""

    source_urls = normalize_cell_values(source_url_value)
    raw_media_types = normalize_cell_values(media_type_value)

    if not source_urls:
        return ()

    media_types = tuple(
        identify_media_type(value)
        for value in raw_media_types
    )

    if not media_types:
        media_types = (
            MediaType.UNKNOWN,
        ) * len(source_urls)

    elif len(media_types) == 1 and len(source_urls) > 1:
        single_type = media_types[0]

        # carousel은 게시글 전체 타입이므로 개별 asset 타입으로 쓰지 않음
        if single_type == MediaType.CAROUSEL:
            media_types = (
                MediaType.UNKNOWN,
            ) * len(source_urls)
        else:
            media_types = media_types * len(source_urls)

    elif len(media_types) != len(source_urls):
        raise ValueError(
            "Media URL 개수와 Media Type 개수가 일치하지 않습니다. "
            f"URL={len(source_urls)}, Type={len(media_types)}"
        )

    return tuple(
        SourceMedia(
            source_url=source_url,
            media_type=media_type,
        )
        for source_url, media_type in zip(
            source_urls,
            media_types,
            strict=True,
        )
    )


def identify_post_media_type(
    source_medias: tuple[SourceMedia, ...],
    raw_media_type_value: Any,
) -> MediaType:
    if len(source_medias) > 1:
        return MediaType.CAROUSEL

    if len(source_medias) == 1:
        return source_medias[0].media_type

    raw_types = normalize_cell_values(raw_media_type_value)
    if len(raw_types) == 1:
        return identify_media_type(raw_types[0])

    return MediaType.UNKNOWN


def extract_rows_to_inputs(
    df: pd.DataFrame,
) -> list[StructuredMediaInput]:
    media_inputs: list[StructuredMediaInput] = []

    for row in df.to_dict(orient="records"):
        campaign_id = optional_text(row.get("campaign_id"))

        if campaign_id is None:
            raise ValueError("campaign_id가 없는 행이 존재합니다.")

        raw_row_number_value = row.get("raw_row_number")
        if (
            raw_row_number_value is not None
            and not pd.isna(raw_row_number_value)
        ):
            raw_row_number = int(raw_row_number_value)
        else:
            raw_row_number = None

        try:
            source_medias = normalize_source_medias(
                source_url_value=row.get("source_url"),
                media_type_value=row.get("media_type"),
            )
        except ValueError as exc:
            raise ValueError(
                "Source media 매핑 실패: "
                f"sheet={row.get('source_sheet')}, "
                f"row={raw_row_number}, "
                f"campaign_id={campaign_id}. {exc}"
            ) from exc

        media_input = StructuredMediaInput(
            campaign_id=campaign_id,
            platform=identify_platform(row.get("platform")),
            media_type=identify_post_media_type(
                source_medias=source_medias,
                raw_media_type_value=row.get("media_type"),
            ),
            original_post_url=optional_text(row.get("permalink")),
            conversation_stream=optional_text(
                row.get("conversation_stream")
            ),
            account=AccountInfo(
                account_name=optional_text(row.get("user_name")),
                profile_url=optional_text(row.get("profile_url")),
            ),
            source_medias=source_medias,
            source_sheet_name=optional_text(row.get("source_sheet")),
            raw_row_number=raw_row_number,
        )
        media_inputs.append(media_input)

    return media_inputs


# =========================================================
# 6. Structured Input -> MediaAsset
# =========================================================

def build_media_assets(
    media_inputs: list[StructuredMediaInput],
) -> list[MediaAsset]:
    media_assets: list[MediaAsset] = []

    for media_input in media_inputs:
        if not media_input.source_medias:
            media_assets.append(
                MediaAsset(
                    campaign_id=media_input.campaign_id,
                    asset_id=f"{media_input.campaign_id}_01",
                    asset_index=1,
                    platform=media_input.platform,
                    media_type=media_input.media_type,
                    original_post_url=media_input.original_post_url,
                    source_url=None,
                    source_sheet_name=media_input.source_sheet_name,
                    raw_row_number=media_input.raw_row_number,
                    status="source_url_missing",
                    error_message="source URL이 존재하지 않습니다.",
                )
            )
            continue

        for asset_index, source_media in enumerate(
            media_input.source_medias,
            start=1,
        ):
            media_assets.append(
                MediaAsset(
                    campaign_id=media_input.campaign_id,
                    asset_id=(
                        f"{media_input.campaign_id}_"
                        f"{asset_index:02d}"
                    ),
                    asset_index=asset_index,
                    platform=media_input.platform,
                    media_type=source_media.media_type,
                    original_post_url=media_input.original_post_url,
                    source_url=source_media.source_url,
                    source_sheet_name=media_input.source_sheet_name,
                    raw_row_number=media_input.raw_row_number,
                )
            )

    return media_assets


# =========================================================
# 6-1. Twitter LINK -> gallery-dl Media URL Resolution
# =========================================================

def is_twitter_status_url(
    source_url: str | None,
) -> bool:
    """Twitter/X의 개별 status URL인지 확인한다."""

    if not source_url:
        return False

    try:
        parsed_url = urlparse(source_url.strip())
    except (TypeError, ValueError):
        return False

    hostname = (parsed_url.hostname or "").lower()

    if hostname not in {
        "x.com",
        "www.x.com",
        "twitter.com",
        "www.twitter.com",
        "mobile.twitter.com",
    }:
        return False

    return bool(
        re.search(
            r"/status/\d+",
            parsed_url.path,
        )
    )


def extract_twitter_media_urls_with_gallery_dl(
    source_url: str,
    timeout_seconds: int = 180,
) -> tuple[str, ...]:
    """
    Twitter/X status URL을 gallery-dl에 전달해 실제 미디어 URL을 반환한다.

    -G를 사용하므로 gallery-dl이 파일을 직접 저장하지 않는다.
    반환된 URL은 기존 feasibility/download/naming 로직에서 처리한다.
    """

    if not is_twitter_status_url(source_url):
        raise ValueError(
            "gallery-dl에는 Twitter/X status URL만 전달할 수 있습니다. "
            f"source_url={source_url}"
        )

    command = [
        "gallery-dl",
        "-G",
        "-v",
        "--no-input",
        "--no-colors",
        source_url,
    ]

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            "gallery-dl 실행 시간이 초과되었습니다. "
            f"source_url={source_url}"
        ) from exc
    except OSError as exc:
        raise RuntimeError(
            "gallery-dl 명령을 실행하지 못했습니다. "
            "현재 터미널에서 gallery-dl --version이 실행되는지 확인하세요."
        ) from exc

    extracted_urls = tuple(
        dict.fromkeys(
            line.strip()
            for line in result.stdout.splitlines()
            if line.strip().startswith(("http://", "https://"))
        )
    )

    # 일부 URL이 정상 추출된 경우에는 경고 로그가 있어도 추출 결과를 사용한다.
    if extracted_urls:
        return extracted_urls

    if result.returncode != 0:
        raise RuntimeError(
            "gallery-dl 미디어 URL 추출에 실패했습니다. "
            f"return_code={result.returncode}, "
            f"stderr={result.stderr.strip()}"
        )

    raise RuntimeError(
        "gallery-dl 실행은 완료되었지만 미디어 URL을 찾지 못했습니다. "
        f"source_url={source_url}"
    )


def resolve_twitter_link_media_assets(
    media_assets: list[MediaAsset],
) -> list[MediaAsset]:
    """
    Platform.TWITTER + MediaType.LINK + Twitter/X status URL인 asset만
    gallery-dl로 실제 이미지/영상 URL asset들로 확장한다.

    다른 플랫폼 URL과 Twitter가 아닌 LINK는 수정하지 않는다.
    확장 후 같은 Excel 행 안에서 asset_index를 다시 순서대로 부여한다.
    """

    resolved_media_assets: list[MediaAsset] = []
    next_asset_index_by_row: dict[tuple[str | None, int | None, str], int] = {}

    for asset in media_assets:
        row_key = (
            asset.source_sheet_name,
            asset.raw_row_number,
            asset.campaign_id,
        )
        next_asset_index = next_asset_index_by_row.get(
            row_key,
            1,
        )

        should_use_gallery_dl = (
            asset.platform == Platform.TWITTER
            and asset.media_type == MediaType.LINK
            and is_twitter_status_url(asset.source_url)
        )

        if not should_use_gallery_dl:
            asset.asset_index = next_asset_index
            asset.asset_id = (
                f"{asset.campaign_id}_"
                f"{next_asset_index:02d}"
            )
            resolved_media_assets.append(asset)
            next_asset_index_by_row[row_key] = next_asset_index + 1
            continue

        try:
            extracted_urls = extract_twitter_media_urls_with_gallery_dl(
                source_url=asset.source_url,
            )

            print(
                "  gallery-dl 추출 완료: "
                f"{asset.asset_id} -> {len(extracted_urls)}개"
            )

            for extracted_url in extracted_urls:
                resolved_media_assets.append(
                    MediaAsset(
                        campaign_id=asset.campaign_id,
                        asset_id=(
                            f"{asset.campaign_id}_"
                            f"{next_asset_index:02d}"
                        ),
                        asset_index=next_asset_index,
                        platform=asset.platform,
                        # 실제 IMAGE/VIDEO 판별은 기존 check_media_url에서 수행
                        media_type=MediaType.UNKNOWN,
                        original_post_url=(
                            asset.original_post_url
                            or asset.source_url
                        ),
                        source_url=extracted_url,
                        source_sheet_name=asset.source_sheet_name,
                        raw_row_number=asset.raw_row_number,
                        extraction_method="gallery_dl",
                    )
                )
                next_asset_index += 1

            next_asset_index_by_row[row_key] = next_asset_index

        except Exception as exc:
            asset.asset_index = next_asset_index
            asset.asset_id = (
                f"{asset.campaign_id}_"
                f"{next_asset_index:02d}"
            )
            asset.extraction_method = "gallery_dl"
            asset.status = "gallery_dl_failed"
            asset.error_message = f"{type(exc).__name__}: {exc}"
            resolved_media_assets.append(asset)
            next_asset_index_by_row[row_key] = next_asset_index + 1

            print(
                "  gallery-dl 추출 실패: "
                f"{asset.asset_id} - {asset.error_message}"
            )

    return resolved_media_assets


# =========================================================
# 7. URL Feasibility
# =========================================================

def detect_media_from_response(
    content_type: str | None,
    sample: bytes,
) -> tuple[MediaType, str | None]:
    normalized_content_type = (
        content_type or ""
    ).split(";", 1)[0].strip().lower()

    content_type_mapping = {
        "video/mp4": (MediaType.VIDEO, ".mp4"),
        "video/webm": (MediaType.VIDEO, ".webm"),
        "video/quicktime": (MediaType.VIDEO, ".mov"),
        "image/jpeg": (MediaType.IMAGE, ".jpg"),
        "image/png": (MediaType.IMAGE, ".png"),
        "image/webp": (MediaType.IMAGE, ".webp"),
        "image/gif": (MediaType.IMAGE, ".gif"),
    }

    if normalized_content_type in content_type_mapping:
        return content_type_mapping[normalized_content_type]

    # 서버 Content-Type이 부정확할 수 있으므로 magic bytes도 확인
    if sample.startswith(b"\xff\xd8\xff"):
        return MediaType.IMAGE, ".jpg"

    if sample.startswith(b"\x89PNG\r\n\x1a\n"):
        return MediaType.IMAGE, ".png"

    if sample.startswith((b"GIF87a", b"GIF89a")):
        return MediaType.IMAGE, ".gif"

    if (
        len(sample) >= 12
        and sample[:4] == b"RIFF"
        and sample[8:12] == b"WEBP"
    ):
        return MediaType.IMAGE, ".webp"

    # MP4/M4V/MOV 계열의 대표 signature
    if len(sample) >= 12 and sample[4:8] == b"ftyp":
        return MediaType.VIDEO, ".mp4"

    if normalized_content_type.startswith("video/"):
        return MediaType.VIDEO, ".mp4"

    if normalized_content_type.startswith("image/"):
        return MediaType.IMAGE, ".jpg"

    if normalized_content_type in {
        "text/html",
        "application/xhtml+xml",
    }:
        return MediaType.LINK, None

    return MediaType.UNKNOWN, None


def read_probe_sample(
    response: requests.Response,
    max_bytes: int = PROBE_BYTES,
) -> bytes:
    chunks: list[bytes] = []
    total_size = 0

    for chunk in response.iter_content(chunk_size=8192):
        if not chunk:
            continue

        remaining = max_bytes - total_size
        chunks.append(chunk[:remaining])
        total_size += min(len(chunk), remaining)

        if total_size >= max_bytes:
            break

    return b"".join(chunks)


def check_media_url(
    source_url: str,
    session: requests.Session,
) -> UrlFeasibilityResult:
    headers = {
        "Range": f"bytes=0-{PROBE_BYTES - 1}",
    }

    try:
        with session.get(
            source_url,
            headers=headers,
            stream=True,
            allow_redirects=True,
            timeout=REQUEST_TIMEOUT,
        ) as response:
            sample = read_probe_sample(response)
            content_type = response.headers.get("Content-Type")
            detected_media_type, extension = detect_media_from_response(
                content_type=content_type,
                sample=sample,
            )

            if response.status_code not in {200, 206}:
                return UrlFeasibilityResult(
                    feasible=False,
                    final_url=response.url,
                    http_status=response.status_code,
                    content_type=content_type,
                    detected_media_type=detected_media_type,
                    extension=None,
                    error_message=(
                        "정상적인 미디어 응답이 아닙니다: "
                        f"HTTP {response.status_code}"
                    ),
                )

            if detected_media_type == MediaType.LINK:
                return UrlFeasibilityResult(
                    feasible=False,
                    final_url=response.url,
                    http_status=response.status_code,
                    content_type=content_type,
                    detected_media_type=MediaType.LINK,
                    extension=None,
                    error_message=(
                        "HTML 페이지 URL입니다. 직접 파일 다운로드가 아니라 "
                        "플랫폼 extractor가 필요합니다."
                    ),
                )

            if detected_media_type not in {
                MediaType.VIDEO,
                MediaType.IMAGE,
            }:
                return UrlFeasibilityResult(
                    feasible=False,
                    final_url=response.url,
                    http_status=response.status_code,
                    content_type=content_type,
                    detected_media_type=MediaType.UNKNOWN,
                    extension=None,
                    error_message=(
                        "URL에는 접근했지만 이미지 또는 영상 파일로 "
                        "판별하지 못했습니다."
                    ),
                )

            return UrlFeasibilityResult(
                feasible=True,
                final_url=response.url,
                http_status=response.status_code,
                content_type=content_type,
                detected_media_type=detected_media_type,
                extension=extension,
            )

    except requests.exceptions.Timeout:
        return UrlFeasibilityResult(
            feasible=False,
            final_url=None,
            http_status=None,
            content_type=None,
            detected_media_type=MediaType.UNKNOWN,
            extension=None,
            error_message="서버 응답 시간이 초과되었습니다.",
        )

    except requests.exceptions.RequestException as exc:
        return UrlFeasibilityResult(
            feasible=False,
            final_url=None,
            http_status=None,
            content_type=None,
            detected_media_type=MediaType.UNKNOWN,
            extension=None,
            error_message=f"URL 요청 실패: {type(exc).__name__}: {exc}",
        )


# =========================================================
# 8. Direct Media Download
# =========================================================

def sanitize_filename(value: str) -> str:
    sanitized = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "_",
        value,
    )
    sanitized = sanitized.strip(" .")
    return sanitized or "unknown"


def build_media_filename_stem(
    asset: MediaAsset,
) -> str:
    """
    저장 파일명 규칙:
        첫 번째 asset: {sheet_name}_{excel_row}
        두 번째 이후: {sheet_name}_{excel_row}_{asset_index:02d}

    한 Excel 행에 미디어가 여러 개 있을 때 파일명 충돌을 방지한다.
    """

    safe_sheet_name = sanitize_filename(
        asset.source_sheet_name or "unknown_sheet"
    )

    if asset.raw_row_number is None:
        row_text = "unknown_row"
    else:
        row_text = str(asset.raw_row_number)

    if asset.asset_index == 1:
        return f"{safe_sheet_name}_{row_text}"

    return (
        f"{safe_sheet_name}_{row_text}_"
        f"{asset.asset_index:02d}"
    )


def download_media_asset(
    asset: MediaAsset,
    feasibility: UrlFeasibilityResult,
    local_media_root: Path,
    session: requests.Session,
) -> MediaAsset:
    if not feasibility.feasible:
        asset.status = "feasibility_failed"
        asset.error_message = feasibility.error_message
        return asset

    if asset.source_url is None:
        asset.status = "source_url_missing"
        asset.error_message = "다운로드할 source URL이 없습니다."
        return asset

    safe_sheet_name = sanitize_filename(
        asset.source_sheet_name or "unknown_sheet"
    )

    sheet_directory = (
        local_media_root
        / safe_sheet_name
    )
    sheet_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        with session.get(
            asset.source_url,
            stream=True,
            allow_redirects=True,
            timeout=REQUEST_TIMEOUT,
        ) as response:
            response.raise_for_status()

            chunk_iterator = response.iter_content(
                chunk_size=1024 * 1024
            )
            first_chunk = next(
                (chunk for chunk in chunk_iterator if chunk),
                b"",
            )

            actual_media_type, actual_extension = detect_media_from_response(
                content_type=response.headers.get("Content-Type"),
                sample=first_chunk,
            )

            if actual_media_type not in {
                MediaType.VIDEO,
                MediaType.IMAGE,
            }:
                raise ValueError(
                    "다운로드 응답이 이미지 또는 영상 파일이 아닙니다. "
                    f"Content-Type={response.headers.get('Content-Type')}"
                )

            extension = actual_extension or feasibility.extension
            if extension is None:
                raise ValueError("파일 확장자를 판별하지 못했습니다.")

            file_stem = build_media_filename_stem(
                asset
            )

            output_path = (
                sheet_directory
                / f"{file_stem}{extension}"
            )
            temporary_path = output_path.with_suffix(
                output_path.suffix + ".part"
            )

            if temporary_path.exists():
                temporary_path.unlink()

            with temporary_path.open("wb") as file:
                if first_chunk:
                    file.write(first_chunk)

                for chunk in chunk_iterator:
                    if chunk:
                        file.write(chunk)

            temporary_path.replace(output_path)

            asset.final_url = response.url
            asset.http_status = response.status_code
            asset.content_type = response.headers.get("Content-Type")
            asset.media_type = actual_media_type
            asset.file_size_bytes = output_path.stat().st_size
            asset.local_path = output_path
            if asset.extraction_method == "gallery_dl":
                asset.extraction_method = (
                    "gallery_dl + direct_http_download"
                )
            else:
                asset.extraction_method = "direct_http_download"

            asset.status = "downloaded"
            asset.error_message = None

            return asset

    except Exception as exc:
        asset.status = "download_failed"
        asset.error_message = f"{type(exc).__name__}: {exc}"
        return asset


def process_media_assets(
    media_assets: list[MediaAsset],
    local_media_root: Path,
) -> list[MediaAsset]:
    processed_assets: list[MediaAsset] = []

    with requests.Session() as session:
        session.headers.update(
            {
                "User-Agent": USER_AGENT,
                "Accept": "*/*",
            }
        )

        total_assets = len(media_assets)

        for current_index, asset in enumerate(
            media_assets,
            start=1,
        ):
            print(
                f"[{current_index}/{total_assets}] 검사: "
                f"{asset.asset_id}"
            )

            if asset.status == "gallery_dl_failed":
                processed_assets.append(asset)
                print(
                    f"  실패: {asset.status} - "
                    f"{asset.error_message}"
                )
                continue

            if not asset.source_url:
                asset.status = "source_url_missing"
                asset.error_message = "source URL이 존재하지 않습니다."
                processed_assets.append(asset)
                continue

            feasibility = check_media_url(
                source_url=asset.source_url,
                session=session,
            )

            asset.final_url = feasibility.final_url
            asset.http_status = feasibility.http_status
            asset.content_type = feasibility.content_type

            if not feasibility.feasible:
                if feasibility.detected_media_type == MediaType.LINK:
                    asset.status = "platform_extractor_required"
                else:
                    asset.status = "feasibility_failed"

                asset.error_message = feasibility.error_message
                processed_assets.append(asset)
                print(
                    f"  실패: {asset.status} - "
                    f"{asset.error_message}"
                )
                continue

            # Raw Data의 media type보다 실제 응답 Content-Type을 최종 기준으로 사용
            asset.media_type = feasibility.detected_media_type

            asset = download_media_asset(
                asset=asset,
                feasibility=feasibility,
                local_media_root=local_media_root,
                session=session,
            )
            processed_assets.append(asset)

            if asset.status == "downloaded":
                print(f"  저장 완료: {asset.local_path}")
            else:
                print(f"  다운로드 실패: {asset.error_message}")

    return processed_assets


# =========================================================
# 9. Manifest / Excel Output
# =========================================================

def media_assets_to_dataframe(
    media_assets: list[MediaAsset],
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []

    for asset in media_assets:
        records.append(
            {
                "campaign_id": asset.campaign_id,
                "asset_id": asset.asset_id,
                "asset_index": asset.asset_index,
                "platform": asset.platform.value,
                "media_type": asset.media_type.value,
                "source_sheet_name": asset.source_sheet_name,
                "raw_row_number": asset.raw_row_number,
                "original_post_url": asset.original_post_url,
                "source_url": asset.source_url,
                "final_url": asset.final_url,
                "http_status": asset.http_status,
                "content_type": asset.content_type,
                "file_size_bytes": asset.file_size_bytes,
                "local_path": (
                    str(asset.local_path)
                    if asset.local_path is not None
                    else None
                ),
                "extraction_method": asset.extraction_method,
                "status": asset.status,
                "error_message": asset.error_message,
            }
        )

    return pd.DataFrame(records)


def save_result_excel(
    output_excel_path: Path,
    campaign_input_df: pd.DataFrame,
    manifest_df: pd.DataFrame,
) -> None:
    with pd.ExcelWriter(
        output_excel_path,
        engine="openpyxl",
    ) as writer:
        campaign_input_df.to_excel(
            writer,
            sheet_name=CAMPAIGN_INPUT_SHEET_NAME,
            index=False,
        )
        manifest_df.to_excel(
            writer,
            sheet_name=MEDIA_MANIFEST_SHEET_NAME,
            index=False,
        )


# =========================================================
# 10. Main
# =========================================================

def main() -> None:
    input_date = input(
        "조회 날짜를 입력하세요 (YYMMDD): "
    ).strip()

    input_excel_path, output_excel_path = build_excel_paths(
        input_date
    )

    # 입력 날짜별 최상위 미디어 폴더
    date_media_root = (
        LOCAL_MEDIA_ROOT
        / sanitize_filename(input_date)
    )
    date_media_root.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(f"입력 파일: {input_excel_path}")
    print(f"출력 파일: {output_excel_path}")
    print(f"미디어 저장 폴더: {date_media_root}")

    # 1. Raw sheet 읽기
    raw_sheets = load_raw_sheets(
        excel_path=input_excel_path,
        sheet_names=RAW_SHEET_NAMES,
        header_row=EXCEL_HEADER_ROW,
    )
    print(f"Raw sheet 로드 완료: {list(raw_sheets.keys())}")

    # 2. 필요한 컬럼 추출 및 표준화
    media_input_df = build_media_input_dataframe(
        raw_sheets=raw_sheets,
        column_mapping=COLUMN_MAPPINGS,
        header_row=EXCEL_HEADER_ROW,
    )
    print(f"컬럼 표준화 완료: {media_input_df.shape}")

    # 3. 값 정리
    media_input_df = clean_input_dataframe(media_input_df)
    print(f"입력 데이터 정리 완료: {media_input_df.shape}")

    # 4. DataFrame 행 -> StructuredMediaInput
    structured_inputs = extract_rows_to_inputs(media_input_df)
    print(
        "StructuredMediaInput 생성 완료: "
        f"{len(structured_inputs)}개"
    )

    # 5. 게시글 단위 -> 개별 MediaAsset
    media_assets = build_media_assets(structured_inputs)
    print(f"MediaAsset 생성 완료: {len(media_assets)}개")

    # 5-1. Twitter LINK만 gallery-dl로 실제 미디어 URL로 확장
    media_assets = resolve_twitter_link_media_assets(
        media_assets=media_assets,
    )
    print(
        "Twitter LINK gallery-dl 해석 완료: "
        f"{len(media_assets)}개"
    )

    # 6. URL feasibility 검사 및 직접 미디어 로컬 다운로드
    processed_assets = process_media_assets(
        media_assets=media_assets,

        # media/{input_date}를 기준 폴더로 전달
        local_media_root=date_media_root,
    )

    # 7. Manifest 생성 및 결과 Excel 저장
    manifest_df = media_assets_to_dataframe(processed_assets)
    save_result_excel(
        output_excel_path=output_excel_path,
        campaign_input_df=media_input_df,
        manifest_df=manifest_df,
    )

    status_counts = manifest_df["status"].value_counts(
        dropna=False
    ).to_dict()

    print(f"결과 Excel 저장 완료: {output_excel_path}")
    print(f"처리 상태 요약: {status_counts}")


if __name__ == "__main__":
    main()
