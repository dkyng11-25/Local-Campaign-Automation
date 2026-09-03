from pathlib import Path
import json
import datetime

import pandas as pd
from google.cloud import bigquery
from google.oauth2 import service_account


# ============================================================
# 1. 사용자 설정값
# ============================================================

# 서비스 계정 JSON key 경로
KEY_PATH = r"C:\Users\KEARNEY\Desktop\gcp-key\slcc-buzz-agent-dev-449cfae180df.json"

# BigQuery 정보
PROJECT_ID = "slcc-buzz-agent-dev"
DATASET_ID = "agent_data"
TABLE_NAME = "quant_dashboard"

# 최종 append 대상 테이블
TARGET_TABLE_ID = f"{PROJECT_ID}.{DATASET_ID}.{TABLE_NAME}"

# 검수 및 스키마 정리가 끝난 Input Excel 파일
EXCEL_PATH = r"C:\Users\KEARNEY\Downloads\SEJ - Product.xlsx"

# Excel sheet 이름 또는 번호
# 첫 번째 sheet면 0
SHEET_NAME = 0

# 로그 저장 경로
LOG_PATH = Path("logs/upload_log.jsonl")


# ============================================================
# 2. 로그 기록 함수
# ============================================================

def write_log(status: str, message: str, extra: dict | None = None):
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)

    log_record = {
        "timestamp": datetime.datetime.now().isoformat(),
        "status": status,
        "message": message,
        "extra": extra or {},
    }

    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(log_record, ensure_ascii=False) + "\n")


# ============================================================
# 3. BigQuery client 생성
# ============================================================

def create_bigquery_client() -> bigquery.Client:
    credentials = service_account.Credentials.from_service_account_file(
        KEY_PATH
    )

    client = bigquery.Client(
        credentials=credentials,
        project=PROJECT_ID
    )

    return client


# ============================================================
# 4. Excel 파일 읽기
# ============================================================

def read_cleaned_excel(excel_path: str, sheet_name=0) -> pd.DataFrame:
    df = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
        engine="openpyxl"
    )

    # 컬럼명 앞뒤 공백 제거
    df.columns = [str(col).strip() for col in df.columns]

    # 완전히 빈 행 제거
    df = df.dropna(how="all")

    return df

# ============================================================
# 5. BigQuery table row count 계산
# ============================================================

def get_bigquery_row_count(client: bigquery.Client, table_id: str) -> int:
    query = f"""
    SELECT COUNT(*) AS row_count
    FROM `{table_id}`
    """

    query_job = client.query(query)
    result = query_job.result()

    for row in result:
        return row.row_count

    return 0


# ============================================================
# 6. DataFrame을 BigQuery target table에 append
# ============================================================

def append_dataframe_to_bigquery(
    client: bigquery.Client,
    df: pd.DataFrame,
    table_id: str,
):
    job_config = bigquery.LoadJobConfig(
        write_disposition=bigquery.WriteDisposition.WRITE_APPEND,
        autodetect=False,
    )

    load_job = client.load_table_from_dataframe(
        df,
        table_id,
        job_config=job_config
    )

    # BigQuery load job이 끝날 때까지 대기
    load_job.result()

    if load_job.errors:
        raise RuntimeError(f"BigQuery load job failed: {load_job.errors}")

    return load_job

# ============================================================
# 7. Local path 확인 
# ============================================================

def validate_local_paths():
    key_path = Path(KEY_PATH)
    excel_path = Path(EXCEL_PATH)

    if not key_path.exists():
        raise FileNotFoundError(f"Key file not found: {key_path}")

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

# ============================================================
# 8. 전체 실행 로직
# ============================================================

def main():
    try:
        write_log(
            status="STARTED",
            message="BigQuery append pipeline started",
            extra={
                "excel_path": EXCEL_PATH,
                "target_table": TARGET_TABLE_ID,
            }
        )


        # 1. Local path 확인
        validate_local_paths()

        # 2. BigQuery client 생성
        client = create_bigquery_client()

        # 3. Excel 파일 읽기
        df = read_cleaned_excel(EXCEL_PATH, SHEET_NAME)

        # 4. Input row count 확인
        input_row_count = len(df)

        if input_row_count == 0:
            raise ValueError("Excel 파일에 업로드할 데이터가 없습니다.")

        # 5. append 전 target table row count
        before_row_count = get_bigquery_row_count(
            client=client,
            table_id=TARGET_TABLE_ID
        )

        # 6. BigQuery target table에 append
        load_job = append_dataframe_to_bigquery(
            client=client,
            df=df,
            table_id=TARGET_TABLE_ID,
        )

        # 7. append 후 target table row count
        after_row_count = get_bigquery_row_count(
            client=client,
            table_id=TARGET_TABLE_ID
        )

        # 8. append 후 row count 검증
        expected_after_count = before_row_count + input_row_count

        if after_row_count != expected_after_count:
            raise ValueError(
                f"Row count mismatch after append. "
                f"Expected: {expected_after_count}, Actual: {after_row_count}"
            )
        
        # 9. 성공 로그
        write_log(
            status="SUCCESS",
            message="Data appended to BigQuery successfully",
            extra={
                "target_table": TARGET_TABLE_ID,
                "input_row_count": input_row_count,
                "before_row_count": before_row_count,
                "after_row_count": after_row_count,
                "bigquery_job_id": load_job.job_id,
            }
        )

        print("Upload completed successfully.")
        print(f"Input rows: {input_row_count}")
        print(f"Before rows: {before_row_count}")
        print(f"After rows: {after_row_count}")
        print(f"BigQuery job ID: {load_job.job_id}")

    except Exception as e:
        write_log(
            status="FAILED",
            message=str(e),
            extra={
                "target_table": TARGET_TABLE_ID,
                "excel_path": EXCEL_PATH,
            }
        )

        print("Upload failed.")
        print(str(e))
        raise

if __name__ == "__main__":
    main()