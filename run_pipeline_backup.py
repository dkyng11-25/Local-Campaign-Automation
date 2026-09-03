from __future__ import annotations

from datetime import datetime
import os
from pathlib import Path
import shutil
import subprocess
import sys

from pipeline_run_paths import (
    PipelineRunPaths,
    prepare_pipeline_run_paths,
)


PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = PROJECT_ROOT / "output"
MEDIA_ROOT = PROJECT_ROOT / "media"

# Buzz Volume은 날짜별/차수별 실행 폴더와 분리된 공용 작업 폴더다.
BUZZ_VOLUME_ROOT = OUTPUT_ROOT / "Buzz_Volume"
BUZZ_VOLUME_COMPLETED_DIR = (
    BUZZ_VOLUME_ROOT / "completed"
)

SPRINKLR_MODULE = "sprinklr_export_excel.py"

FOLLOW_UP_MODULES = (
    "raw_to_processed.py",
    "media_extractor.py",
    "llm_analysis_pipeline.py",
)

# 하위 모듈에 공통 실행 경로를 전달하기 위한 환경변수명
ENV_INPUT_DATE = "LOCAL_CAMPAIGN_INPUT_DATE"
ENV_RUN_NUMBER = "LOCAL_CAMPAIGN_RUN_NUMBER"
ENV_OUTPUT_DIR = "LOCAL_CAMPAIGN_OUTPUT_DIR"
ENV_MEDIA_DIR = "LOCAL_CAMPAIGN_MEDIA_DIR"


def ensure_buzz_volume_directories() -> tuple[Path, Path]:
    """
    Buzz Volume 공용 입력/완료 폴더가 존재하도록 보장한다.

    구조:
        output/
        └─ Buzz_Volume/
           ├─ 사용자가 최종 통합·정제 Excel을 넣는 위치
           └─ completed/
              └─ Buzz Volume 적재 완료 결과 저장 위치

    이 폴더들은 날짜별 실행 차수와 독립적이므로,
    이미 존재하면 그대로 유지한다.
    """

    BUZZ_VOLUME_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    BUZZ_VOLUME_COMPLETED_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    return (
        BUZZ_VOLUME_ROOT,
        BUZZ_VOLUME_COMPLETED_DIR,
    )


def _build_gcloud_command(
    gcloud_path: str,
    *arguments: str,
) -> list[str]:
    """
    현재 OS에서 gcloud 명령을 안정적으로 실행할 command list를 만든다.

    Windows에서는 gcloud가 gcloud.cmd / gcloud.bat 형태일 수 있으므로
    cmd.exe를 통해 실행한다.
    """

    suffix = Path(gcloud_path).suffix.casefold()

    if os.name == "nt" and suffix in {
        ".cmd",
        ".bat",
    }:
        return [
            "cmd.exe",
            "/d",
            "/c",
            gcloud_path,
            *arguments,
        ]

    return [
        gcloud_path,
        *arguments,
    ]


def ensure_google_adc() -> None:
    """
    Google Application Default Credentials(ADC)를 확인한다.

    처리 순서:
    1. 현재 ADC로 access token 발급 가능 여부 확인
    2. 정상이라면 그대로 pipeline 진행
    3. ADC가 없거나 만료되었다면
       `gcloud auth application-default login` 자동 실행
    4. 로그인 완료 후 ADC를 다시 검증
    """

    gcloud_path = shutil.which(
        "gcloud"
    )

    if not gcloud_path:
        raise RuntimeError(
            "gcloud CLI를 찾을 수 없습니다.\n"
            "Google Cloud SDK가 설치되어 있고 "
            "PATH에 등록되어 있는지 확인하세요."
        )

    print()
    print("=" * 70)
    print("Google Cloud 인증 확인")
    print("=" * 70)

    # =========================================================
    # 1. 현재 ADC 상태 확인
    # =========================================================
    check_result = subprocess.run(
        _build_gcloud_command(
            gcloud_path,
            "auth",
            "application-default",
            "print-access-token",
        ),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
        cwd=PROJECT_ROOT,
    )

    if check_result.returncode == 0:
        print(
            "✅ Google Cloud "
            "Application Default Credentials 정상"
        )
        return

    # =========================================================
    # 2. ADC가 없거나 만료된 경우 자동 로그인
    # =========================================================
    print(
        "Google Cloud 인증이 없거나 만료되었습니다."
    )
    print(
        "gcloud auth application-default login을 "
        "실행합니다."
    )
    print()

    login_result = subprocess.run(
        _build_gcloud_command(
            gcloud_path,
            "auth",
            "application-default",
            "login",
        ),
        check=False,
        cwd=PROJECT_ROOT,
    )

    if login_result.returncode != 0:
        raise RuntimeError(
            "Google Cloud Application Default Credentials "
            "로그인에 실패했습니다.\n"
            f"gcloud return code: "
            f"{login_result.returncode}"
        )

    # =========================================================
    # 3. 로그인 완료 후 ADC 재검증
    # =========================================================
    verify_result = subprocess.run(
        _build_gcloud_command(
            gcloud_path,
            "auth",
            "application-default",
            "print-access-token",
        ),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
        cwd=PROJECT_ROOT,
    )

    if verify_result.returncode != 0:
        raise RuntimeError(
            "Google Cloud 로그인은 완료되었지만 "
            "Application Default Credentials 검증에 "
            "실패했습니다."
        )

    print()
    print(
        "✅ Google Cloud "
        "Application Default Credentials 인증 완료"
    )


def parse_datetime(
    datetime_text: str,
    input_name: str,
) -> datetime:
    """
    YYYY-MM-DD HH:MM:SS 형식의 문자열을 datetime으로 변환한다.
    """

    datetime_text = datetime_text.strip()

    try:
        return datetime.strptime(
            datetime_text,
            "%Y-%m-%d %H:%M:%S",
        )

    except ValueError as exc:
        raise ValueError(
            f"{input_name}은 YYYY-MM-DD HH:MM:SS "
            "형식이어야 합니다.\n"
            f"입력값: {datetime_text}\n"
            "예시: 2026-07-27 19:00:00"
        ) from exc


def build_module_environment(
    run_paths: PipelineRunPaths,
) -> dict[str, str]:
    """
    하위 모듈에 전달할 환경변수를 생성한다.

    기존 모듈의 input() 입력 순서는 그대로 유지하고,
    실행 차수별 output/media 경로는 환경변수로 별도 전달한다.
    """

    module_environment = os.environ.copy()

    module_environment.update(
        {
            ENV_INPUT_DATE: run_paths.input_date,
            ENV_RUN_NUMBER: str(
                run_paths.run_number
            ),
            ENV_OUTPUT_DIR: str(
                run_paths.output_dir
            ),
            ENV_MEDIA_DIR: str(
                run_paths.media_dir
            ),
            # 하위 Python 프로세스 로그가 즉시 출력되도록 설정
            "PYTHONUNBUFFERED": "1",
        }
    )

    return module_environment


def run_module(
    module_name: str,
    module_inputs: list[str],
    run_paths: PipelineRunPaths,
) -> None:
    """
    Python 모듈을 실행하고 해당 모듈의 input()에 값을 순서대로 전달한다.

    실행 경로는 다음 환경변수로 모든 하위 모듈에 동일하게 전달한다.

        LOCAL_CAMPAIGN_INPUT_DATE
        LOCAL_CAMPAIGN_RUN_NUMBER
        LOCAL_CAMPAIGN_OUTPUT_DIR
        LOCAL_CAMPAIGN_MEDIA_DIR
    """

    module_path = (
        PROJECT_ROOT / module_name
    )

    if not module_path.exists():
        raise FileNotFoundError(
            "모듈 파일을 찾을 수 없습니다: "
            f"{module_path}"
        )

    if not module_path.is_file():
        raise FileNotFoundError(
            "모듈 경로가 파일이 아닙니다: "
            f"{module_path}"
        )

    # 각 하위 모듈의 기존 input() 호출 순서에 맞게 줄바꿈으로 연결
    stdin_text = (
        "\n".join(module_inputs)
        + "\n"
    )

    module_environment = (
        build_module_environment(
            run_paths=run_paths
        )
    )

    print()
    print("=" * 70)
    print(
        f"실행 시작: {module_name}"
    )

    for index, input_value in enumerate(
        module_inputs,
        start=1,
    ):
        print(
            f"전달 입력값 {index}: "
            f"{input_value}"
        )

    print(
        f"실행 차수: "
        f"{run_paths.run_label}"
    )
    print(
        f"Output 폴더: "
        f"{run_paths.output_dir}"
    )
    print(
        f"Media 폴더: "
        f"{run_paths.media_dir}"
    )
    print("=" * 70)

    result = subprocess.run(
        [
            sys.executable,
            str(module_path),
        ],
        input=stdin_text,
        text=True,
        check=False,
        cwd=PROJECT_ROOT,
        env=module_environment,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"{module_name} 실행 실패 "
            f"(return code: "
            f"{result.returncode})\n"
            f"실행 차수: "
            f"{run_paths.run_label}\n"
            f"Output 폴더: "
            f"{run_paths.output_dir}\n"
            f"Media 폴더: "
            f"{run_paths.media_dir}"
        )

    print(
        f"✅ 실행 완료: {module_name}"
    )


def main() -> None:
    print("=" * 70)
    print(
        "Local Campaign 전체 자동화 파이프라인"
    )
    print("=" * 70)

    # =========================================================
    # Google Cloud Application Default Credentials 확인
    #
    # ADC가 정상적이면 바로 진행하고,
    # 없거나 만료되었을 때만
    # gcloud auth application-default login을 자동 실행한다.
    # =========================================================
    ensure_google_adc()

    # Buzz Volume은 1~4단계와 별도로 실행하지만,
    # 사용자가 최종 통합·정제 파일을 넣을 공용 폴더는
    # run_pipeline.py 실행 시 자동으로 준비한다.
    (
        buzz_volume_root,
        buzz_volume_completed_dir,
    ) = ensure_buzz_volume_directories()

    print()
    print(
        "Buzz Volume 공용 폴더 확인"
    )
    print(
        "- 입력 파일 위치: "
        f"{buzz_volume_root}"
    )
    print(
        "- 완료 결과 위치: "
        f"{buzz_volume_completed_dir}"
    )

    start_datetime_text = input(
        "\nSprinklr 조회 시작 날짜와 시간을 입력하세요.\n"
        "형식: YYYY-MM-DD HH:MM:SS\n"
        "예시: 2026-07-26 19:00:00\n"
        "입력: "
    ).strip()

    end_datetime_text = input(
        "\nSprinklr 조회 종료 날짜와 시간을 입력하세요.\n"
        "형식: YYYY-MM-DD HH:MM:SS\n"
        "예시: 2026-07-27 19:00:00\n"
        "입력: "
    ).strip()

    start_datetime = parse_datetime(
        datetime_text=(
            start_datetime_text
        ),
        input_name="시작 날짜와 시간",
    )

    end_datetime = parse_datetime(
        datetime_text=(
            end_datetime_text
        ),
        input_name="종료 날짜와 시간",
    )

    if end_datetime <= start_datetime:
        raise ValueError(
            "종료 날짜와 시간은 시작 날짜와 시간보다 "
            "늦어야 합니다.\n"
            f"시작: {start_datetime}\n"
            f"종료: {end_datetime}"
        )

    normalized_start_datetime = (
        start_datetime.strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    normalized_end_datetime = (
        end_datetime.strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    # 종료 시각의 날짜를 작업 기준 날짜로 사용
    input_date = (
        end_datetime.strftime(
            "%y%m%d"
        )
    )

    print()
    print("입력값 확인")
    print(
        "- 조회 시작 시각: "
        f"{normalized_start_datetime}"
    )
    print(
        "- 조회 종료 시각: "
        f"{normalized_end_datetime}"
    )
    print(
        "- 후속 모듈 작업 날짜: "
        f"{input_date}"
    )

    # 파이프라인 전체에서 딱 한 번만
    # 실행 차수와 공통 경로를 확정한다.
    run_paths = prepare_pipeline_run_paths(
        input_date=input_date,
        output_root=OUTPUT_ROOT,
        media_root=MEDIA_ROOT,
    )

    print()
    print("실행 폴더 확인")
    print(
        f"- 실행 차수: "
        f"{run_paths.run_label}"
    )
    print(
        f"- Output 폴더: "
        f"{run_paths.output_dir}"
    )
    print(
        f"- Media 폴더: "
        f"{run_paths.media_dir}"
    )

    # =========================================================
    # 1단계
    # sprinklr_export_excel.py의 input() 호출 순서:
    #   1. start_datetime
    #   2. end_datetime
    #
    # output/media 경로는 input()이 아니라 환경변수로 전달한다.
    # =========================================================
    run_module(
        module_name=SPRINKLR_MODULE,
        module_inputs=[
            normalized_start_datetime,
            normalized_end_datetime,
        ],
        run_paths=run_paths,
    )

    # =========================================================
    # 2~4단계
    # 각 모듈의 기존 input()에는 YYMMDD 한 번 전달한다.
    # 동일 실행 경로는 환경변수로 함께 전달한다.
    # =========================================================
    for module_name in FOLLOW_UP_MODULES:
        run_module(
            module_name=module_name,
            module_inputs=[
                input_date
            ],
            run_paths=run_paths,
        )

    print()
    print("=" * 70)
    print(
        "✅ 전체 Local Campaign "
        "자동화 파이프라인 실행 완료"
    )
    print(
        f"실행 차수: "
        f"{run_paths.run_label}"
    )
    print(
        f"Output 폴더: "
        f"{run_paths.output_dir}"
    )
    print(
        f"Media 폴더: "
        f"{run_paths.media_dir}"
    )
    print(
        "Buzz Volume 입력 폴더: "
        f"{buzz_volume_root}"
    )
    print(
        "Buzz Volume 완료 폴더: "
        f"{buzz_volume_completed_dir}"
    )
    print()
    print(
        "1차·2차 결과를 통합하고 최종 정제한 Excel을 "
        "Buzz Volume 입력 폴더에 넣은 뒤 "
        "buzz_volume_adaptor.py를 별도로 실행하세요."
    )
    print("=" * 70)


if __name__ == "__main__":
    main()
