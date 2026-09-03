from pathlib import Path

import requests


url = """https://video.twimg.com/amplify_video/2076437706637516800/vid/avc1/1080x1920/7GRDsY6bvseUNmlk.mp4?tag=16"""
output_dir = Path("downloaded_media")
output_dir.mkdir(parents=True, exist_ok=True)

output_path = output_dir / "instagram_video.mp4"

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "Chrome/142.0 Safari/537.36"
    )
}

try:
    print("영상 다운로드를 시작합니다.")

    with requests.get(
        url,
        headers=headers,
        stream=True,
        allow_redirects=True,
        timeout=(10, 60),
    ) as response:

        print("HTTP 상태 코드:", response.status_code)
        print("Content-Type:", response.headers.get("Content-Type"))
        print("최종 URL:", response.url)

        response.raise_for_status()

        content_type = (
            response.headers.get("Content-Type") or ""
        ).lower()

        if not content_type.startswith("video/"):
            raise ValueError(
                f"영상 응답이 아닙니다. Content-Type: {content_type}"
            )

        with output_path.open("wb") as file:
            for chunk in response.iter_content(
                chunk_size=1024 * 1024
            ):
                if chunk:
                    file.write(chunk)

    file_size_mb = output_path.stat().st_size / (1024 * 1024)

    print("\n다운로드 성공")
    print("저장 경로:", output_path.resolve())
    print(f"파일 크기: {file_size_mb:.2f} MB")

except requests.exceptions.Timeout:
    print("다운로드 실패: 서버 응답 시간이 초과되었습니다.")

except requests.exceptions.HTTPError as error:
    print("다운로드 실패: HTTP 오류")
    print(error)

except requests.exceptions.RequestException as error:
    print("다운로드 실패: 네트워크 요청 오류")
    print(type(error).__name__, error)

except Exception as error:
    print("다운로드 실패")
    print(type(error).__name__, error)