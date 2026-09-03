from __future__ import annotations

import argparse
import importlib.util
import sys
from datetime import datetime
from pathlib import Path
from typing import Any


# =========================================================
# 1. Missing-pipeline path configuration
# =========================================================

# 이 파일의 배치 위치:
#   Local_Campaign_Automation_version7/누락/media_extractor.py
BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent

# 프로젝트 루트의 media_extractor.py를 기준 구현으로 그대로 사용한다.
ROOT_MEDIA_EXTRACTOR_PATH = PROJECT_ROOT / "media_extractor.py"

# 누락 전용 산출물 경로
OUTPUT_DIR = BASE_DIR / "output_누락"
LOCAL_MEDIA_ROOT = BASE_DIR / "media_누락"


# =========================================================
# 2. Load root media_extractor implementation
# =========================================================

def load_root_media_extractor() -> Any:
    """
    프로젝트 루트의 media_extractor.py를 별도 모듈명으로 로드한다.

    누락 버전에서 플랫폼별 미디어 추출, TikTok yt-dlp,
    Twitter gallery-dl, URL feasibility, LLM input 집계 및 Excel 저장
    로직을 복제하지 않고 루트 버전과 정확히 동일하게 사용한다.
    """

    if not ROOT_MEDIA_EXTRACTOR_PATH.is_file():
        raise FileNotFoundError(
            "프로젝트 루트의 media_extractor.py를 찾을 수 없습니다.\n"
            f"경로: {ROOT_MEDIA_EXTRACTOR_PATH}"
        )

    module_name = "_local_campaign_root_media_extractor"
    spec = importlib.util.spec_from_file_location(
        module_name,
        ROOT_MEDIA_EXTRACTOR_PATH,
    )

    if spec is None or spec.loader is None:
        raise ImportError(
            "프로젝트 루트 media_extractor.py의 모듈 spec을 "
            f"생성하지 못했습니다: {ROOT_MEDIA_EXTRACTOR_PATH}"
        )

    module = importlib.util.module_from_spec(spec)

    # dataclass/enum 등에서 __module__ 조회가 정상 동작하도록
    # exec_module 전에 sys.modules에 등록한다.
    sys.modules[module_name] = module
    spec.loader.exec_module(module)

    return module


# =========================================================
# 3. CLI / Missing directory resolution
# =========================================================

def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "프로젝트 루트 media_extractor.py와 동일한 미디어 추출 로직을 "
            "사용하되 누락/output_누락 및 누락/media_누락에 결과를 저장합니다."
        )
    )

    parser.add_argument(
        "--output-dir",
        type=Path,
        help=(
            "누락 output 폴더 override. 미지정 시 "
            "누락/output_누락/{YYMMDD}_누락 사용"
        ),
    )
    parser.add_argument(
        "--media-dir",
        type=Path,
        help=(
            "누락 media 폴더 override. 미지정 시 "
            "누락/media_누락/{YYMMDD}_누락 사용"
        ),
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help=(
            "기존 campaign_media_result.xlsx 또는 미디어 결과가 있을 때 "
            "새 결과가 모두 완성된 후 기존 결과를 교체합니다."
        ),
    )

    return parser.parse_args()


def resolve_missing_path(path: Path) -> Path:
    """누락 폴더 기준 상대경로를 절대경로로 변환한다."""

    expanded = path.expanduser()
    if not expanded.is_absolute():
        expanded = BASE_DIR / expanded
    return expanded.resolve()


def validate_input_date(input_date: str) -> str:
    normalized = input_date.strip()

    try:
        datetime.strptime(normalized, "%y%m%d")
    except ValueError as exc:
        raise ValueError(
            "날짜는 YYMMDD 형식으로 입력해야 합니다. 예시) 260714"
        ) from exc

    return normalized


def resolve_missing_execution_directories(
    input_date: str,
    cli_output_dir: Path | None,
    cli_media_dir: Path | None,
) -> tuple[Path, Path]:
    """
    누락 서브 파이프라인의 날짜별 output/media 폴더를 확정한다.

    기본 경로:
        누락/output_누락/{YYMMDD}_누락
        누락/media_누락/{YYMMDD}_누락

    --output-dir / --media-dir가 주어지면 해당 경로만 override한다.
    """

    output_dir = (
        resolve_missing_path(cli_output_dir)
        if cli_output_dir is not None
        else (OUTPUT_DIR / f"{input_date}_누락").resolve()
    )

    media_dir = (
        resolve_missing_path(cli_media_dir)
        if cli_media_dir is not None
        else (LOCAL_MEDIA_ROOT / f"{input_date}_누락").resolve()
    )

    # raw_to_processed 단계에서 입력 Excel을 넣어둔 output 폴더는
    # 이미 존재해야 한다. 잘못된 날짜/경로를 조용히 새로 만들지 않는다.
    if not output_dir.exists():
        raise FileNotFoundError(
            "누락 데이터 output 폴더를 찾을 수 없습니다.\n"
            f"경로: {output_dir}"
        )

    if not output_dir.is_dir():
        raise NotADirectoryError(
            "누락 데이터 output 경로가 폴더가 아닙니다.\n"
            f"경로: {output_dir}"
        )

    expected_output_name = f"{input_date}_누락"
    if output_dir.name != expected_output_name:
        raise ValueError(
            "누락 output 폴더명이 작업 날짜와 일치하지 않습니다.\n"
            f"기대 폴더명: {expected_output_name}\n"
            f"실제 폴더: {output_dir}"
        )

    # media 폴더는 media_extractor가 안전하게 생성한다.
    media_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    expected_media_name = f"{input_date}_누락"
    if media_dir.name != expected_media_name:
        raise ValueError(
            "누락 media 폴더명이 작업 날짜와 일치하지 않습니다.\n"
            f"기대 폴더명: {expected_media_name}\n"
            f"실제 폴더: {media_dir}"
        )

    print(f"[INFO] 누락 데이터 output 폴더: {output_dir}")
    print(f"[INFO] 누락 데이터 media 폴더: {media_dir}")

    return output_dir, media_dir


# =========================================================
# 4. Main
# =========================================================

def main() -> None:
    args = parse_arguments()
    root_media = load_root_media_extractor()

    # 현재 누락 run_pipeline.py가 날짜를 stdin으로 전달하는 방식과 호환한다.
    input_date = validate_input_date(
        input("조회 날짜를 입력하세요 (YYMMDD): ").strip()
    )

    output_dir, media_dir = resolve_missing_execution_directories(
        input_date=input_date,
        cli_output_dir=args.output_dir,
        cli_media_dir=args.media_dir,
    )

    # 루트 media_extractor와 동일한 파일 계약:
    # 입력  : {date}_SLCC_SOV_Local Campaign Tracking_{month}월_v01.xlsx
    # 출력  : {date}_campaign_media_result.xlsx
    input_excel_path, output_excel_path = root_media.build_excel_paths(
        input_date=input_date,
        output_dir=output_dir,
    )

    (
        temporary_excel_path,
        temporary_media_dir,
        backup_media_dir,
    ) = root_media.prepare_temporary_artifacts(
        output_excel_path=output_excel_path,
        media_dir=media_dir,
        overwrite=args.overwrite,
    )

    print(f"입력 파일: {input_excel_path}")
    print(f"Media extraction 결과 파일: {output_excel_path}")
    print(f"미디어 저장 폴더: {media_dir}")

    try:
        # 이하 처리 순서는 프로젝트 루트 media_extractor.py와 동일하다.
        raw_sheets = root_media.load_raw_sheets(
            excel_path=input_excel_path,
            sheet_names=root_media.RAW_SHEET_NAMES,
        )
        print(
            "Raw sheet 로드 완료: "
            f"{list(raw_sheets.keys())}"
        )

        media_input_df = root_media.build_media_input_dataframe(
            raw_sheets=raw_sheets,
            column_mapping=root_media.COLUMN_MAPPINGS,
        )
        print(
            "컬럼 표준화 완료: "
            f"{media_input_df.shape}"
        )

        media_input_df = root_media.clean_input_dataframe(
            media_input_df
        )
        media_input_df = root_media.prepare_campaign_input_for_processing(
            media_input_df
        )
        print(
            "입력 데이터 정리 완료: "
            f"{media_input_df.shape}"
        )

        structured_inputs = root_media.extract_rows_to_inputs(
            media_input_df
        )
        print(
            "StructuredMediaInput 생성 완료: "
            f"{len(structured_inputs)}개"
        )

        media_assets = root_media.build_media_assets(
            structured_inputs
        )
        print(
            "MediaAsset 생성 완료: "
            f"{len(media_assets)}개"
        )

        media_assets = root_media.process_tiktok_media_assets_with_yt_dlp(
            media_assets=media_assets,
            local_media_root=temporary_media_dir,
        )

        media_assets = root_media.resolve_page_media_assets_with_gallery_dl(
            media_assets=media_assets,
            only_failed_unknown=False,
        )

        processed_assets = root_media.process_media_assets(
            media_assets=media_assets,
            local_media_root=temporary_media_dir,
        )

        processed_assets = root_media.resolve_page_media_assets_with_gallery_dl(
            media_assets=processed_assets,
            only_failed_unknown=True,
        )

        has_pending_assets = any(
            asset.status == "pending"
            for asset in processed_assets
        )

        if has_pending_assets:
            processed_assets = root_media.process_media_assets(
                media_assets=processed_assets,
                local_media_root=temporary_media_dir,
            )

        processed_assets = root_media.mark_remaining_failures_for_manual_action(
            media_assets=processed_assets,
        )

        processed_assets = root_media.rebase_media_asset_local_paths(
            media_assets=processed_assets,
            temporary_media_dir=temporary_media_dir,
            final_media_dir=media_dir,
        )

        manifest_df = root_media.media_assets_to_dataframe(
            processed_assets
        )
        llm_input_df = root_media.build_llm_input_dataframe(
            campaign_input_df=media_input_df,
            manifest_df=manifest_df,
        )

        root_media.save_result_excel(
            output_excel_path=temporary_excel_path,
            llm_input_df=llm_input_df,
        )

        root_media.commit_output_artifacts(
            temporary_excel_path=temporary_excel_path,
            output_excel_path=output_excel_path,
            temporary_media_dir=temporary_media_dir,
            media_dir=media_dir,
            backup_media_dir=backup_media_dir,
        )

    except Exception:
        root_media.cleanup_temporary_artifacts(
            temporary_excel_path=temporary_excel_path,
            temporary_media_dir=temporary_media_dir,
        )
        raise

    if llm_input_df.empty:
        status_counts: dict[Any, int] = {}
        print(
            "처리할 게시물이 없어 헤더만 있는 "
            "llm_input 시트를 생성했습니다."
        )
    else:
        status_counts = llm_input_df[
            "status"
        ].value_counts(
            dropna=False
        ).to_dict()

    follower_count_rows = (
        int(
            llm_input_df[
                "sender_follower_count"
            ].notna().sum()
        )
        if (
            not llm_input_df.empty
            and "sender_follower_count" in llm_input_df.columns
        )
        else 0
    )

    print(
        "Media extraction 결과 Excel 저장 완료: "
        f"{output_excel_path}"
    )
    print(
        "게시물 단위 처리 상태 요약: "
        f"{status_counts}"
    )
    print(
        "Sender Follower Count 전달 완료 행 수: "
        f"{follower_count_rows}/{len(llm_input_df)}"
    )


if __name__ == "__main__":
    main()
