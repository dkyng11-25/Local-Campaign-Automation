import re
import json
from pathlib import Path
from openpyxl import load_workbook
from google.cloud import storage
import os

import pandas as pd

pd.set_option("display.max_columns", None)

INPUT_PATH = Path("input/raw_campaign.xlsx")
OUTPUT_JSONL_PATH = Path("output/clean_campaign.jsonl")
OUTPUT_CSV_PATH = Path("output/clean_campaign_preview.csv")

GCS_KEY_PATH = r"C:\Users\KEARNEY\Desktop\gcp-key\slcc-buzz-agent-dev-449cfae180df.json"
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = GCS_KEY_PATH

BUCKET_NAME = "slcc-local-campaign-images"
GCS_FOLDER = "campaign_images"

HEADER_ROW_INDEX = 0
"""
Temporary set as 2026 for testing purposes. 
"""
DEFAULT_YEAR = 2026

FINAL_COLUMNS = [
    "row_order",
    "campaign_date",
    "image_url",
    "subsidiary",
    "country",
    "campaign_name",
    "product_name",
    "feature_lv1",
    "feature_lv2",
    "campaign_description",
    "mentions",
    "channel",
    "giveaway",
    "influencer",
    "htr_de",
    "conv_card",
    "hashtag",
    "url_post",
    "url_reaction",
    "remarks",
    "source_query",
]

"""앞뒤 공백 제거"""
def clean_text(value):
    if pd.isna(value):
        return None
    text = str(value).strip() 
    return text if text else None

def extract_images_from_exel(excel_path, output_dir, image_column_letter="D", header_row_index=1):    
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(excel_path)
    ws = wb.active

    """이미지 위치 및 저장 경로 mapping"""
    image_map = {}

    for idx, img in enumerate(ws._images, start=1):
        """이미지 시작 셀 위치 파악"""
        anchor = img.anchor._from 
        print(f"Image {idx}: anchor at row {anchor.row}, col {anchor.col}")

        """openpyxl row/col은 0-based -> 엑셀 기준으로 변환"""
        excel_row = anchor.row + 1
        excel_col = anchor.col + 1

        expected_col = ord(image_column_letter.upper()) - ord("A") + 1
        if excel_col != expected_col:
            continue

        ext = "png"
        file_name = f"campaign_row_{excel_row:04d}.{ext}"
        file_path = output_dir / file_name

        """이미지 바이너리 데이터 수집 및 파일로 저장"""
        image_bytes = img._data()

        with open(file_path, "wb") as f:
            f.write(image_bytes)

        image_map[excel_row] = str(file_path).replace("\\", "/")

    """row별 image path mapping 반환"""
    return image_map 

"""로컬 파일을 GCS에 업로드"""
def upload_to_gcs(client, local_path, bucket_name, destination_blob_name):
    bucket = client.bucket(bucket_name)
    blob = bucket.blob(destination_blob_name)

    blob.upload_from_filename(local_path)

    auth_url = f"https://storage.cloud.google.com/{bucket_name}/{destination_blob_name}"

    print(f"File {local_path} uploaded to {auth_url}")

    return auth_url

def upload_local_image_and_return_url(client,local_path):
    if pd.isna(local_path):
        return None

    local_path = Path(local_path)

    destination_blob_name = f"{GCS_FOLDER}/{local_path.name}"

    return upload_to_gcs(
        client,
        local_path=str(local_path),
        bucket_name=BUCKET_NAME,
        destination_blob_name=destination_blob_name
    ) 

"""Subsidiary와 Country를 분리"""
def parse_subsidiary_country(value):
    """
    예: 'SEA (Thailand)' -> subsidiary='SEA', country='Thailand'
    """
    text = clean_text(value)
    if not text:
        return None, None

    match = re.match(r"^(.*?)\s*\((.*?)\)\s*$", text)
    if match:
        subsidiary = match.group(1).strip()
        country = match.group(2).strip()
        return subsidiary, country

    return text, None

def parse_product(value):
    text = clean_text(value)
    if not text or text.upper() in ["N/A", "NA", "NONE"]:
        return []

    parts = re.split(r"[\n\r]+", text)
    products = [] 

    for part in parts:
        product = part.strip()
        if product:
            products.append(product)
    
    return list(dict.fromkeys(products))     

"""Feature Level 1과 Level 2를 분리"""
def parse_feature(value):
    """
    예: 'Design_General' -> feature_lv1='Design', feature_lv2='General'
    """
    text = clean_text(value)
    if not text or text.upper() in ["N/A", "NA", "NONE"]: 
        return None, None
    
    extract = text.split("_")

    if len(extract) == 2:
        return extract[0].strip(), extract[1].strip()
    
    return None, None

def parse_channel(value):
    text = clean_text(value)
    if not text or text.upper() in ["N/A", "NA", "NONE"]:
        return []

    parts = re.split(r"[\n\r]+", text)
    channels = [] 

    for part in parts:
        channel = part.strip()
        if channel:
            channels.append(channel)
    
    return list(dict.fromkeys(channels)) 

""" Yes/No -> True/False """
def parse_bool(value):
    text = clean_text(value)
    if not text:
        return None

    text = text.lower()

    if text in ["yes", "y", "true", "1"]:
        return True
    if text in ["no", "n", "false", "0"]:
        return False

    return None

"""줄바꿈으로 구분된 hashtag를 ARRAY<STRING> 형태로 변환"""
def parse_hashtags(value):
    """
    예:
    '#A\n#B\n#C' -> ['#A', '#B', '#C']
    """
    text = clean_text(value)
    if not text or text.upper() in ["N/A", "NA", "NONE"]:
        return []

    parts = re.split(r"[\n\r]+", text)
    hashtags = []

    for part in parts:
        tag = part.strip()
        if tag:
            hashtags.append(tag)

    # 중복 제거, 순서 유지
    return list(dict.fromkeys(hashtags))

"""URL 셀 구조 파싱 및 분리"""
def parse_urls(value):
    """
    URL 셀 구조:
    [인플루언서 or 당사 게시글]
    url1
    url2
    (줄바꿈)
    [소비자 반응]
    url3

    결과:
    url_post = [인플루언서 게시글 + 당사 게시글 아래 URL]
    url_reaction = [소비자 반응 아래 URL]
    """
    text = clean_text(value)
    if not text:
        return [], []

    lines = [line.strip() for line in re.split(r"[\n\r]+", text) if line.strip()]

    current_section = None
    post_urls = []
    reaction_urls = []

    for line in lines:
        if "인플루언서" in line:
            current_section = "post"
            continue
        if "당사" in line:
            current_section = "post"
            continue
        if "소비자" in line or "반응" in line:
            current_section = "reaction"
            continue

        if line.startswith("http://") or line.startswith("https://"):
            if current_section == "reaction":
                reaction_urls.append(line)
            else:
                post_urls.append(line)

    return post_urls, reaction_urls


def main():
    client = storage.Client()

    df = pd.read_excel(INPUT_PATH, header=HEADER_ROW_INDEX)

    df = df.iloc[:, 1:].copy()  # 첫 번째 컬럼 제거

    df = df.head(3)

    image_map = extract_images_from_exel(
        excel_path=INPUT_PATH, 
        output_dir="output/images", 
        image_column_letter="D",   
        header_row_index=1)

    print("Raw columns:")
    """Raw columns 출력"""
    print(list(df.columns)) 
    print(f"Raw shape: {df.shape}")

    column_mapping = {
        df.columns[0]: "row_order",
        df.columns[1]: "campaign_date",
        df.columns[2]: "image_url",
        df.columns[3]: "subsidiary_country",
        df.columns[4]: "campaign_name",
        df.columns[5]: "product_name",
        df.columns[6]: "features",
        df.columns[7]: "campaign_description",
        df.columns[8]: "mentions",
        df.columns[9]: "channel",
        df.columns[10]: "giveaway",
        df.columns[11]: "influencer",
        df.columns[12]: "htr_de",
        df.columns[13]: "conv_card",
        df.columns[14]: "hashtag",
        df.columns[15]: "url_raw",
        df.columns[16]: "remarks",
        df.columns[17]: "source_query"
    }

    """컬럼명 변경"""
    df = df.rename(columns=column_mapping)

    # row_order
    df["row_order"] = pd.to_numeric(df["row_order"], errors="coerce").astype("Int64")

    # image_url mapping
    df["image_url"] = df["row_order"].apply(
    lambda x: image_map.get(int(x) + 1) if pd.notna(x) else None)

    # local image path -> GCS upload -> GCS URL mapping
    df["image_url"] = df["image_url"].apply(lambda x: upload_local_image_and_return_url(client, x))

    # subsidiary / country 분리
    split_result = df["subsidiary_country"].apply(parse_subsidiary_country)
    df["subsidiary"] = split_result.apply(lambda x: x[0])
    df["country"] = split_result.apply(lambda x: x[1])

    # product_name 분리
    df["product_name"] = df["product_name"].apply(parse_product)

    # feature_lv1 / feature_lv2 분리
    split_result = df["features"].apply(parse_feature)
    df["feature_lv1"] = split_result.apply(lambda x: x[0])
    df["feature_lv2"] = split_result.apply(lambda x: x[1])

    # mentions 변환
    df["mentions"] = pd.to_numeric(df["mentions"], errors="coerce").astype("Int64")

    # channel 변환
    df["channel"] = df["channel"].apply(parse_channel)

    # bool 변환
    for col in ["giveaway", "influencer", "htr_de", "conv_card"]:
        df[col] = df[col].apply(parse_bool)

    # hashtags ARRAY<STRING>
    df["hashtag"] = df["hashtag"].apply(parse_hashtags)

    # URL 분리
    url_split = df["url_raw"].apply(parse_urls)
    df["url_post"] = url_split.apply(lambda x: x[0])
    df["url_reaction"] = url_split.apply(lambda x: x[1])

    # 문자열 컬럼 정리
    for col in [
        "image_url",
        "campaign_name",
        "campaign_description",
        "remarks",
        "source_query",
    ]:
        if col in df.columns:
            df[col] = df[col].apply(clean_text)

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = None

    clean_df = df[FINAL_COLUMNS].copy()

    print("\nCleaned preview:")
    print(clean_df.head())

    print("\nCleaned dtypes:")
    print(clean_df.dtypes)

    OUTPUT_JSONL_PATH.parent.mkdir(parents=True, exist_ok=True)

    # BigQuery ARRAY<STRING> 업로드용 JSONL
    with open(OUTPUT_JSONL_PATH, "w", encoding="utf-8") as f:
        for record in clean_df.to_dict(orient="records"):
            clean_record = {}
            for key, value in record.items():
                if pd.isna(value) if not isinstance(value, list) else False:
                    clean_record[key] = None
                else:
                    clean_record[key] = value
            f.write(json.dumps(clean_record, ensure_ascii=False, default=str) + "\n")

    # 사람이 확인하기 위한 preview CSV
    preview_df = clean_df.copy()
    preview_df["hashtag"] = preview_df["hashtag"].apply(lambda x: "|".join(x) if isinstance(x, list) else "")
    preview_df["url_post"] = preview_df["url_post"].apply(lambda x: "|".join(x) if isinstance(x, list) else "")
    preview_df["url_reaction"] = preview_df["url_reaction"].apply(lambda x: "|".join(x) if isinstance(x, list) else "")

    preview_df.to_csv(OUTPUT_CSV_PATH, index=False, encoding="utf-8-sig")

    print(f"\nSaved BigQuery JSONL to: {OUTPUT_JSONL_PATH}")
    print(f"Saved preview CSV to: {OUTPUT_CSV_PATH}")


if __name__ == "__main__":
    main()