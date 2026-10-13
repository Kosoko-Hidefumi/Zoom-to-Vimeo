"""
Vimeo へ MP4 をアップロード（Zoom ダウンロード後の別工程用）。

- CSV の「配信用動画タイトル」とローカル mp4（ファイル名 stem）を突合
- パスコードで Vimeo のパスワード制限を設定
- 重複アップロード防止の一次情報は **入力 CSV 自体に書き戻す「Vimeo URI」「Vimeoリンク」列**。
  この列に値が入っている行は、Vimeo に問い合わせることなく確実にスキップする。
  （タイトル文字列で Vimeo 側と突合する方式は、128文字切り詰めや同名重複動画があると
  誤判定・再アップロードを招くため、確認専用のフォールバックに格下げしている）
- アップロード/既存確認で判明したリンクは毎回 CSV に書き戻すので、2回目以降の実行では
  Vimeo API の一覧取得（--no-vimeo-check で無効化可）すら不要になっていく。
- --resume-from は従来の Vimeo 専用結果 CSV（互換のため維持、--csv への書き戻しの方が優先）

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
    kwargs.setdefault("verify", False)
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
            "name": title[:128],
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
        rr = requests.patch(upload_link, headers=patch_headers, data=f, verify=False)
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


def vimeo_existing_videos(token: str) -> dict:
    """タイトルキー -> {uri, link, duration} の辞書を返す（重複チェックとリンク復元の両方に使用）。"""
    videos: dict = {}
    url = f"{VIMEO_API}/me/videos?per_page=100"
    while url:
        r = vimeo_request("GET", url, token)
        r.raise_for_status()
        payload = r.json()
        for v in payload.get("data") or []:
            name = (v.get("name") or "").strip()
            if name:
                videos[file_match_key(name)] = {
                    "uri": v.get("uri") or "",
                    "link": v.get("link") or "",
                    "duration": v.get("duration") or 0,
                }
        next_url = (payload.get("paging") or {}).get("next") or None
        if next_url and next_url.startswith("/"):
            next_url = f"{VIMEO_API}{next_url}"
        url = next_url
    return videos


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
        "--col-vimeo-uri",
        default="Vimeo URI",
        help="アップロード済みの Vimeo URI を保存/参照する列名（この列に値があればスキップ）",
    )
    p.add_argument(
        "--col-vimeo-link",
        default="Vimeoリンク",
        help="アップロード済みの Vimeo リンクを保存/参照する列名",
    )
    p.add_argument(
        "--out-csv",
        type=Path,
        required=True,
        help="Vimeo 工程用の結果 CSV 出力先（専用形式・--resume-from 互換用）",
    )
    p.add_argument(
        "--resume-from",
        type=Path,
        default=None,
        help="（互換用）前回の Vimeo 結果 CSV（status=uploaded 済みをスキップ）。"
        "通常は --csv 自体の Vimeo URI/リンク列があれば不要。",
    )
    p.add_argument(
        "--no-vimeo-check",
        action="store_true",
        help="GET /me/videos によるタイトル突合フォールバックをしない"
        "（--csv に Vimeo URI/リンク列がまだ無い古い行の重複検出に使われる）",
    )
    p.add_argument(
        "--no-write-back",
        action="store_true",
        help="アップロード結果を --csv 自体には書き戻さない（out-csv のみに記録）",
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
        reader = csv.DictReader(f)
        orig_fieldnames = list(reader.fieldnames or [])
        rows = list(reader)

    if not rows:
        print("[ERROR] CSV にデータ行がありません。", file=sys.stderr)
        sys.exit(1)

    if args.col_title not in rows[0]:
        print(f"[ERROR] 列がありません: {args.col_title}", file=sys.stderr)
        sys.exit(1)

    write_fieldnames = list(orig_fieldnames)
    for col in (args.col_vimeo_uri, args.col_vimeo_link):
        if col not in write_fieldnames:
            write_fieldnames.append(col)

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

    already_uploaded: dict[str, dict] = {}
    if args.resume_from and args.resume_from.is_file():
        with open(args.resume_from, "r", encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                if (r.get("status") or "").startswith("uploaded"):
                    t = r.get(args.col_title, "").strip()
                    if t:
                        already_uploaded[file_match_key(t)] = {
                            "uri": r.get("video_uri", "") or "",
                            "link": r.get("link", "") or "",
                        }
        print(f"resume: アップロード済み（結果 CSV）= {len(already_uploaded)} 件")

    on_vimeo: dict[str, dict] = {}
    on_vimeo_fetched = False

    results = []
    matched = uploaded = skipped_csv = skipped_resume = skipped_vimeo = missing = errors = 0
    csv_dirty = False

    for row in rows:
        title = (row.get(args.col_title) or "").strip()
        if not title:
            continue

        key = file_match_key(title)
        password = (row.get(args.col_pass) or "").strip()

        # 1次情報: 入力 CSV 自体に記録済みの Vimeo URI/リンク。
        # タイトル突合に頼らず確実にスキップできるため、ここが真っ先にチェックされる。
        csv_uri = (row.get(args.col_vimeo_uri) or "").strip()
        csv_link = (row.get(args.col_vimeo_link) or "").strip()
        if csv_uri or csv_link:
            skipped_csv += 1
            results.append(
                {**row, "status": "skipped(already_in_csv)", "video_uri": csv_uri, "link": csv_link}
            )
            continue

        # 互換用: 旧方式の --resume-from（前回の Vimeo 結果 CSV）
        if key in already_uploaded:
            skipped_resume += 1
            prev = already_uploaded[key]
            row[args.col_vimeo_uri] = prev["uri"]
            row[args.col_vimeo_link] = prev["link"]
            csv_dirty = True
            results.append(
                {
                    **row,
                    "status": "skipped(already_uploaded)",
                    "video_uri": prev["uri"],
                    "link": prev["link"],
                }
            )
            continue

        # フォールバック: CSV にまだ記録がない古い行だけ、Vimeo 側をタイトル突合で確認する。
        # 見つかった場合は即座に CSV へ書き戻し、以後はこの突合が不要になるようにする。
        if not args.no_vimeo_check and not on_vimeo_fetched:
            print("Vimeo 既存動画を取得中（CSV未記録の行のタイトル突合フォールバック用）...")
            on_vimeo = vimeo_existing_videos(token)
            on_vimeo_fetched = True
            print(f"Vimeo 既存タイトルキー数: {len(on_vimeo)}")

        # Vimeo はアップロード時にタイトルを128文字で切り詰める（vimeo_upload() 参照）ため、
        # 128文字超のタイトルは素の key では一致しない。切り詰め後の key でも照合する。
        vimeo_hit = on_vimeo.get(key) or on_vimeo.get(file_match_key(title[:128]))
        if vimeo_hit and int(vimeo_hit.get("duration") or 0) > 0:
            skipped_vimeo += 1
            row[args.col_vimeo_uri] = vimeo_hit["uri"]
            row[args.col_vimeo_link] = vimeo_hit["link"]
            csv_dirty = True
            results.append(
                {
                    **row,
                    "status": "skipped(already_on_vimeo:backfilled)",
                    "video_uri": vimeo_hit["uri"],
                    "link": vimeo_hit["link"],
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
                row[args.col_vimeo_uri] = res["uri"]
                row[args.col_vimeo_link] = res["link"]
                csv_dirty = True
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
    out_fieldnames = write_fieldnames + ["status", "video_uri", "link"]
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=out_fieldnames, extrasaction="ignore")
        w.writeheader()
        for r in results:
            w.writerow(r)

    if csv_dirty and not args.dry_run and not args.no_write_back:
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=write_fieldnames, extrasaction="ignore")
            w.writeheader()
            for r in rows:
                w.writerow(r)
        print(f"write-back: {csv_path.resolve()} に Vimeo URI/リンクを書き戻しました")

    print("\n---- SUMMARY ----")
    print(f"matched : {matched} (ファイルあり)")
    print(f"uploaded: {uploaded}")
    print(
        f"skipped : {skipped_csv + skipped_resume + skipped_vimeo} "
        f"(csv:{skipped_csv} resume:{skipped_resume} vimeo:{skipped_vimeo})"
    )
    print(f"missing : {missing} (CSV にあるがファイルなし)")
    print(f"errors  : {errors}")
    print(f"out_csv : {out_path.resolve()}")


if __name__ == "__main__":
    main()
