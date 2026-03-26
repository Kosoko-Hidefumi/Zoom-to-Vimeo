#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Vimeo の /me/videos とソース CSV を突き合わせ、指定列の CSV を出力する。

- 講師名・専門科などはソース CSV 由来（Vimeo API には無い）。
- リンクは Vimeo の動画オブジェクトの link。
- パスコードも API では取得できないためソース CSV の列をそのまま使う。

認証: zoom_download/.env の VIMEO_TOKEN（カレントに .env があればそちらも参照）。
突合キー・API ヘルパーは zoom_download.vimeo_upload と共通。

ソース CSV: --source-csv 省略時は本スクリプトと同じフォルダ、またはカレントの
consultant_vimeo.csv を自動使用し、全行に講師名等を載せて Vimeo の link を突合する。
Vimeo 一覧だけ欲しい場合は --vimeo-only。
"""

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

_ZOOM_DL = Path(__file__).resolve().parent / "zoom_download"
if _ZOOM_DL.is_dir() and str(_ZOOM_DL) not in sys.path:
    sys.path.insert(0, str(_ZOOM_DL))

from config import Config  # noqa: E402
from vimeo_upload import VIMEO_API, file_match_key, vimeo_request  # noqa: E402

_ROOT = Path(__file__).resolve().parent

OUT_FIELDS = [
    "講師名（英語）",
    "専門科",
    "所属",
    "レクチャータイトル（英語）",
    "配信用動画タイトル",
    "パスコード",
    "Vimeo動画のlink",
]


def fetch_vimeo_videos(token: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    url = f"{VIMEO_API}/me/videos?per_page=100"
    while url:
        r = vimeo_request("GET", url, token)
        r.raise_for_status()
        payload = r.json()
        for v in payload.get("data") or []:
            name = (v.get("name") or "").strip()
            link = (v.get("link") or "").strip()
            rows.append(
                {
                    "name": name,
                    "link": link,
                    "uri": (v.get("uri") or "").strip(),
                }
            )
        url = (payload.get("paging") or {}).get("next") or None
    return rows


def build_vimeo_link_by_key(videos: List[Dict[str, Any]]) -> Tuple[Dict[str, str], int]:
    """file_match_key -> link（同一キーが複数あるときは先頭を採用し、重複を数える）。"""
    by_key: Dict[str, str] = {}
    dup = 0
    for v in videos:
        name = v["name"]
        if not name:
            continue
        key = file_match_key(name)
        if not key:
            continue
        if key in by_key:
            dup += 1
            continue
        by_key[key] = v["link"]
    return by_key, dup


def resolve_source_csv_path(arg: str | None) -> Path | None:
    """
    --source-csv 省略または空のときは _ROOT またはカレントの consultant_vimeo.csv。
    パス指定時はファイルの存在を確認して解決。
    """
    if arg and str(arg).strip():
        p = Path(arg)
        if p.is_file():
            return p.resolve()
        alt = _ROOT / arg
        if alt.is_file():
            return alt.resolve()
        cwd_try = Path.cwd() / arg
        if cwd_try.is_file():
            return cwd_try.resolve()
        return None

    for cand in (_ROOT / "consultant_vimeo.csv", Path.cwd() / "consultant_vimeo.csv"):
        if cand.is_file():
            return cand.resolve()
    return None


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Vimeo とソース CSV をマージしてメタデータ＋リンクの CSV を出力する。"
    )
    ap.add_argument(
        "--source-csv",
        default=None,
        metavar="PATH",
        help="講師名などを含む CSV。省略時は consultant_vimeo.csv（スクリプト直下またはカレント）を自動使用。",
    )
    ap.add_argument(
        "--vimeo-only",
        action="store_true",
        help="ソース CSV を使わず Vimeo 一覧のみ（講師名等は空）。",
    )
    ap.add_argument("--out-csv", default="vimeo_metadata_export.csv", help="出力 CSV")
    ap.add_argument(
        "--col-title",
        default="配信用動画タイトル",
        help="ソース CSV 上の配信タイトル列名（マッチキー用）",
    )
    ap.add_argument(
        "--col-pass",
        default="パスコード",
        help="ソース CSV 上のパスコード列名",
    )
    args = ap.parse_args()

    token = Config.VIMEO_TOKEN
    if not token:
        print(
            "VIMEO_TOKEN が未設定です。zoom_download/.env に VIMEO_TOKEN=... を記入してください。",
            file=sys.stderr,
        )
        sys.exit(1)

    print("Vimeo 動画一覧を取得中...", file=sys.stderr)
    videos = fetch_vimeo_videos(token)
    vimeo_link_by_key, dup_keys = build_vimeo_link_by_key(videos)
    print(f"vimeo: {len(videos)} 本, ユニークキー {len(vimeo_link_by_key)}", file=sys.stderr)
    if dup_keys:
        print(f"warn: 同じタイトルキーの動画が {dup_keys} 本余分（先頭の link のみ使用）", file=sys.stderr)

    out_rows: List[Dict[str, str]] = []

    if args.vimeo_only:
        for v in videos:
            name = v["name"]
            out_rows.append(
                {
                    "講師名（英語）": "",
                    "専門科": "",
                    "所属": "",
                    "レクチャータイトル（英語）": "",
                    "配信用動画タイトル": name,
                    "パスコード": "",
                    "Vimeo動画のlink": v["link"],
                }
            )
        print(f"Vimeo のみ出力: {len(out_rows)} 行", file=sys.stderr)
    else:
        src_path = resolve_source_csv_path(args.source_csv)
        if src_path is None:
            print(
                "ソース CSV が見つかりません。"
                " consultant_vimeo.csv をこのディレクトリに置くか、"
                " --source-csv でパスを指定してください。"
                "（Vimeo だけ欲しい場合は --vimeo-only）",
                file=sys.stderr,
            )
            sys.exit(1)
        print(f"ソース CSV: {src_path}", file=sys.stderr)
        with open(src_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames or []
            required = [
                "講師名（英語）",
                "専門科",
                "所属",
                "レクチャータイトル（英語）",
                args.col_title,
                args.col_pass,
            ]
            missing = [c for c in required if c not in fieldnames]
            if missing:
                print(f"ソース CSV に次の列がありません: {missing}", file=sys.stderr)
                print(f"実際の列: {fieldnames}", file=sys.stderr)
                sys.exit(1)
            for row in reader:
                title = (row.get(args.col_title) or "").strip()
                key = file_match_key(title) if title else ""
                link = vimeo_link_by_key.get(key, "") if key else ""
                out_rows.append(
                    {
                        "講師名（英語）": (row.get("講師名（英語）") or "").strip(),
                        "専門科": (row.get("専門科") or "").strip(),
                        "所属": (row.get("所属") or "").strip(),
                        "レクチャータイトル（英語）": (row.get("レクチャータイトル（英語）") or "").strip(),
                        "配信用動画タイトル": title,
                        "パスコード": (row.get(args.col_pass) or "").strip(),
                        "Vimeo動画のlink": link,
                    }
                )
        matched = sum(1 for r in out_rows if r["Vimeo動画のlink"])
        print(f"ソース行: {len(out_rows)}, Vimeo と一致: {matched}", file=sys.stderr)

    out_path = Path(args.out_csv).resolve()
    _write_result_csv(out_path, out_rows)


def _write_result_csv(out_path: Path, out_rows: List[Dict[str, str]]) -> None:
    """CSV 書き込み。ロック中（Excel 起動中等）は別名でリトライし、手順を stderr に出す。"""
    try:
        _write_result_csv_once(out_path, out_rows)
        print(f"wrote -> {out_path}", file=sys.stderr)
        return
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = out_path.parent / f"{out_path.stem}_{ts}{out_path.suffix}"
        try:
            _write_result_csv_once(alt, out_rows)
            print(
                f"wrote -> {alt}",
                file=sys.stderr,
            )
            print(
                "注意: 元のパスへは書けませんでした（ファイルが他アプリで開かれている可能性）。"
                " Excel 等で出力 CSV を閉じてから再実行するか、この別名ファイルを使用してください。",
                file=sys.stderr,
            )
            return
        except PermissionError:
            pass
        print(
            f"エラー: 出力できません（アクセス拒否）: {out_path}",
            file=sys.stderr,
        )
        print(
            "Excel・プレビュー・別のエディタでこの CSV を開いていないか確認し、閉じてから再実行してください。",
            file=sys.stderr,
        )
        print(
            "または別パスを指定: python export_vimeo_metadata_csv.py --out-csv vimeo_export_new.csv",
            file=sys.stderr,
        )
        sys.exit(1)


def _write_result_csv_once(path: Path, out_rows: List[Dict[str, str]]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=OUT_FIELDS, extrasaction="ignore")
        w.writeheader()
        for r in out_rows:
            w.writerow(r)


if __name__ == "__main__":
    main()
