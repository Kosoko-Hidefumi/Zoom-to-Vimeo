"""
Vimeo へ MP4 をアップロード（Zoom ダウンロード後の別工程用）。

- CSV の「配信用動画タイトル」とローカル mp4（ファイル名 stem）を突合
- パスコードで Vimeo のパスワード制限を設定
- 既に Vimeo 上に同名（突合キー一致）がある場合はスキップ（参考実装と同様）
- --resume-from は Vimeo 専用の結果 CSV を指定（Zoom の result CSV とは別）

使い方（zoom_download ディレクトリで）:
  python vimeo_upload.py --csv consultant_vimeo.csv --out-csv vimeo_results.csv
  python vimeo_upload.py --root ../downloads --csv ../consultant_vimeo.csv --out-csv vimeo_results.csv --dry-run
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path
from typing import Set

import requests

from config import Config
from csv_parser import sanitize_filename

VIMEO_API = "https://api.vimeo.com"


def norm_loose(s: str) -> str:
    if s is None:
        return ""
    s = s.strip()
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def file_match_key(s: str) -> str:
    """Zoom 保存名と同じ sanitize_filename を通し、CSV 生タイトルと mp4 stem を揃える。"""
    return norm_loose(sanitize_filename(s))


def vimeo_request(method: str, url: str, token: str, **kwargs):
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"bearer {token}"
    headers["Accept"] = "application/vnd.vimeo.*+json;version=3.4"
    return requests.request(method, url, headers=headers, **kwargs)


def vimeo_upload(
    file_path: Path,
    title: str,
    password: str,
    token: str,
    dry_run: bool = False,
):
    if dry_run:
        return {"uri": "dry-run", "link": ""}

    size = file_path.stat().st_size

    r = vimeo_request(
        "POST",
        f"{VIMEO_API}/me/videos",
        token,
        json={
            "upload": {"approach": "tus", "size": size},
            "name": title,
        },
    )
    r.raise_for_status()
    data = r.json()
    upload_link = data["upload"]["upload_link"]
    video_uri = data["uri"]

    with open(file_path, "rb") as f:
        patch_headers = {
            "Tus-Resumable": "1.0.0",
            "Upload-Offset": "0",
            "Content-Type": "application/offset+octet-stream",
        }
        rr = requests.patch(upload_link, headers=patch_headers, data=f)
        rr.raise_for_status()

    vimeo_request(
        "PATCH",
        f"{VIMEO_API}{video_uri}",
        token,
        json={
            "privacy": {"view": "password"},
            "password": password,
        },
    ).raise_for_status()

    info = vimeo_request("GET", f"{VIMEO_API}{video_uri}", token).json()
    return {"uri": video_uri, "link": info.get("link", "")}


def vimeo_existing_title_keys(token: str) -> Set[str]:
    keys: Set[str] = set()
    url = f"{VIMEO_API}/me/videos?per_page=100"
    while url:
        r = vimeo_request("GET", url, token)
        r.raise_for_status()
        payload = r.json()
        for v in payload.get("data") or []:
            name = (v.get("name") or "").strip()
            if name:
                keys.add(file_match_key(name))
        url = (payload.get("paging") or {}).get("next") or None
    return keys


def parse_args():
    p = argparse.ArgumentParser(
        description="Vimeo へローカル MP4 をアップロード（CSV の配信用動画タイトルで突合）",
    )
    p.add_argument(
        "--root",
        type=Path,
        default=None,
        help="mp4 を再帰探索するルート（省略時: .env の DOWNLOAD_DIR）",
    )
    p.add_argument(
        "--csv",
        type=Path,
        default=Path("consultant_vimeo.csv"),
        help="入力 CSV（既定: consultant_vimeo.csv）",
    )
    p.add_argument(
        "--col-title",
        default="配信用動画タイトル",
        help="タイトル列名",
    )
    p.add_argument(
        "--col-pass",
        default="パスコード",
        help="Vimeo パスワードに使う列名（CSV のパスコード）",
    )
    p.add_argument(
        "--out-csv",
        type=Path,
        required=True,
        help="Vimeo 工程用の結果 CSV 出力先（専用形式・resume に使用）",
    )
    p.add_argument(
        "--resume-from",
        type=Path,
        default=None,
        help="前回の Vimeo 結果 CSV（status=uploaded 済みをスキップ）",
    )
    p.add_argument(
        "--no-vimeo-check",
        action="store_true",
        help="GET /me/videos による重複チェックをしない",
    )
    p.add_argument("--dry-run", action="store_true")
    return p.parse_args()


def main():
    args = parse_args()
    token = Config.VIMEO_TOKEN
    if not token:
        print("[ERROR] .env に VIMEO_TOKEN が設定されていません。", file=sys.stderr)
        sys.exit(1)

    root = args.root.resolve() if args.root else Config.DOWNLOAD_DIR.resolve()
    csv_path = args.csv

    if not csv_path.is_file():
        alt = Path(__file__).resolve().parent / csv_path.name
        if alt.is_file():
            csv_path = alt
    if not csv_path.is_file():
        print(f"[ERROR] CSV が見つかりません: {args.csv}", file=sys.stderr)
        sys.exit(1)

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    if not rows:
        print("[ERROR] CSV にデータ行がありません。", file=sys.stderr)
        sys.exit(1)

    if args.col_title not in rows[0]:
        print(f"[ERROR] 列がありません: {args.col_title}", file=sys.stderr)
        sys.exit(1)

    print(f"CSV行数: {len(rows)}")
    if not root.is_dir():
        print(f"[WARNING] ルートが存在しません（mp4 は 0 件）: {root}")

    files: dict[str, Path] = {}
    if root.is_dir():
        for p in root.rglob("*.mp4"):
            if p.is_file():
                files[file_match_key(p.stem)] = p

    print(f"mp4検出件数: {len(files)}")
    print(f"探索ルート: {root}")

    already_uploaded: Set[str] = set()
    if args.resume_from and args.resume_from.is_file():
        with open(args.resume_from, "r", encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                if (r.get("status") or "").startswith("uploaded"):
                    t = r.get(args.col_title, "").strip()
                    if t:
                        already_uploaded.add(file_match_key(t))
        print(f"resume: アップロード済み（結果 CSV）= {len(already_uploaded)} 件")

    on_vimeo: Set[str] = set()
    if not args.no_vimeo_check:
        print("Vimeo 既存動画を取得中...")
        on_vimeo = vimeo_existing_title_keys(token)
        print(f"Vimeo 既存タイトルキー数: {len(on_vimeo)}")

    results = []
    matched = uploaded = skipped_resume = skipped_vimeo = missing = errors = 0

    for row in rows:
        title = (row.get(args.col_title) or "").strip()
        if not title:
            continue

        key = file_match_key(title)
        password = (row.get(args.col_pass) or "").strip()

        if key in already_uploaded:
            skipped_resume += 1
            results.append(
                {
                    **row,
                    "status": "skipped(already_uploaded)",
                    "video_uri": "",
                    "link": "",
                }
            )
            continue

        if key in on_vimeo:
            skipped_vimeo += 1
            results.append(
                {
                    **row,
                    "status": "skipped(already_on_vimeo)",
                    "video_uri": "",
                    "link": "",
                }
            )
            continue

        file_path = files.get(key)
        if not file_path:
            missing += 1
            results.append({**row, "status": "missing", "video_uri": "", "link": ""})
            continue

        matched += 1
        print(f"\n=== {title}")
        print(f"FILE: {file_path}")
        print(f"PASS: {password}")

        try:
            res = vimeo_upload(file_path, title, password, token, args.dry_run)
            status = "uploaded" if not args.dry_run else "dry-run"
            if not args.dry_run:
                uploaded += 1
            results.append(
                {**row, "status": status, "video_uri": res["uri"], "link": res["link"]}
            )
        except Exception as e:
            errors += 1
            results.append(
                {**row, "status": f"error:{e}", "video_uri": "", "link": ""}
            )

    out_path = args.out_csv
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = list(rows[0].keys()) + ["status", "video_uri", "link"]
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        for r in results:
            w.writerow(r)

    print("\n---- SUMMARY ----")
    print(f"matched : {matched} (ファイルあり)")
    print(f"uploaded: {uploaded}")
    print(
        f"skipped : {skipped_resume + skipped_vimeo} "
        f"(resume:{skipped_resume} vimeo:{skipped_vimeo})"
    )
    print(f"missing : {missing} (CSV にあるがファイルなし)")
    print(f"errors  : {errors}")
    print(f"out_csv : {out_path.resolve()}")


if __name__ == "__main__":
    main()
