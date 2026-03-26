"""
downloader.py - Zoom録画ファイルのダウンロード（プログレス表示・リトライ付き）
"""
import time as time_module
from pathlib import Path

import requests
from tqdm import tqdm

from config import Config


def download_recording(
    download_url: str,
    access_token: str,
    dest_path: Path,
    expected_size: int = 0,
) -> bool:
    if dest_path.exists() and expected_size > 0:
        existing_size = dest_path.stat().st_size
        if existing_size == expected_size:
            print(f"    [SKIP] 既存ファイル（サイズ一致）: {dest_path.name}")
            return True

    headers = {
        "Authorization": f"Bearer {access_token}",
    }

    tmp_path = dest_path.with_suffix(".tmp")

    for attempt in range(1, Config.MAX_RETRIES + 1):
        try:
            print(f"    [DOWNLOAD] Attempt {attempt}/{Config.MAX_RETRIES}: {dest_path.name}")

            resp = requests.get(
                download_url,
                headers=headers,
                stream=True,
                timeout=300,
                allow_redirects=True,
            )
            resp.raise_for_status()

            total_size = int(resp.headers.get("content-length", expected_size or 0))

            with open(tmp_path, "wb") as f:
                with tqdm(
                    total=total_size if total_size > 0 else None,
                    unit="B",
                    unit_scale=True,
                    unit_divisor=1024,
                    desc=f"    {dest_path.name[:50]}",
                    ncols=100,
                ) as pbar:
                    for chunk in resp.iter_content(chunk_size=Config.DOWNLOAD_CHUNK_SIZE):
                        if chunk:
                            f.write(chunk)
                            pbar.update(len(chunk))

            tmp_path.rename(dest_path)
            size_mb = dest_path.stat().st_size / 1024 / 1024
            print(f"    [OK] Downloaded: {dest_path.name} ({size_mb:.1f} MB)")
            return True

        except requests.exceptions.HTTPError as e:
            status_code = e.response.status_code if e.response else "unknown"
            print(f"    [ERROR] HTTP {status_code}: {e}")
            if status_code == 401:
                print("    [ERROR] 401 Unauthorized - トークンが期限切れの可能性があります")
                return False
            if status_code == 404:
                print("    [ERROR] 404 Not Found - 録画が存在しないか削除済みです")
                return False

        except requests.exceptions.RequestException as e:
            print(f"    [ERROR] Request failed: {e}")

        except Exception as e:
            print(f"    [ERROR] Unexpected error: {e}")

        if attempt < Config.MAX_RETRIES:
            wait = Config.RETRY_DELAY * attempt
            print(f"    [RETRY] {wait}秒後にリトライ...")
            time_module.sleep(wait)

    if tmp_path.exists():
        tmp_path.unlink()

    print(f"    [FAIL] {Config.MAX_RETRIES}回リトライ後も失敗: {dest_path.name}")
    return False
