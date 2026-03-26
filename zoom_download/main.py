"""
main.py - Zoom録画ダウンロードシステム エントリーポイント

CSVスプレッドシートを「正」として、Zoom APIから録画を取得し、
講師名フォルダ/日付フォルダに整理してダウンロードする。

Usage:
    python main.py --from 2025-07-14 --to 2025-07-18
    python main.py --from 2025-07-14 --to 2025-07-18 --dry-run
    python main.py --csv other.csv --speaker "Pangilinan"
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

from config import Config
from csv_parser import parse_csv, filter_records_by_date
from zoom_client import ZoomClient
from matcher import match_recordings, MatchResult
from folder_manager import (
    create_download_path,
    build_staging_download_path,
    finalize_download_to_csv_title,
)
from downloader import download_recording
from result_csv import load_resume_keys, write_result_csv, build_result_row


def parse_date(date_str: str) -> datetime:
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(f"日付形式が不正です: {date_str} (YYYY-MM-DD)")


def main():
    parser = argparse.ArgumentParser(
        description="Zoom録画ダウンロードシステム - CSVベース突合",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
例:
  python main.py --from 2025-07-14 --to 2025-07-18
  python main.py --from 2025-07-14 --to 2025-07-18 --dry-run
  python main.py --from 2025-07-14 --to 2025-07-18 --resume-from result.csv
  python main.py --speaker "Pangilinan"
  python main.py --csv other.csv --from 2025-09-01 --to 2025-09-05 --output-csv result_sep.csv
        """,
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=Path("consultant_vimeo.csv"),
        help="入力CSV（省略時: consultant_vimeo.csv。カレントに無い場合は zoom_download 内も検索）",
    )
    parser.add_argument("--from", dest="date_from", type=parse_date, help="開始日 (YYYY-MM-DD)")
    parser.add_argument("--to", dest="date_to", type=parse_date, help="終了日 (YYYY-MM-DD)")
    parser.add_argument("--speaker", type=str, help="講師名フィルタ（部分一致）")
    parser.add_argument("--dry-run", action="store_true", help="ダウンロードせず突合結果のみ表示")
    parser.add_argument("--resume-from", type=Path, help="前回結果CSVパス（ダウンロード済みをスキップ）")
    parser.add_argument("--output-csv", type=Path, help="結果CSV出力先（省略時: result_YYYYMMDD_HHMMSS.csv）")
    parser.add_argument("--download-dir", type=Path, help="ダウンロード先ディレクトリ（.envの値を上書き）")

    args = parser.parse_args()

    if not args.dry_run:
        errors = Config.validate()
        if errors:
            print("[ERROR] 設定エラー:")
            for e in errors:
                print(f"  - {e}")
            print("\n.envファイルを確認してください。")
            sys.exit(1)

    if args.download_dir:
        Config.DOWNLOAD_DIR = args.download_dir

    print(f"\n{'='*70}")
    print("  Zoom録画ダウンロードシステム")
    print(f"{'='*70}")

    csv_path = args.csv
    if not csv_path.is_file():
        fallback = Path(__file__).resolve().parent / csv_path.name
        if fallback.is_file():
            csv_path = fallback
    if not csv_path.is_file():
        print(f"[ERROR] CSVファイルが見つかりません: {args.csv}")
        sys.exit(1)

    print(f"\n[Step 1] CSV読み込み: {csv_path}")
    all_records = parse_csv(csv_path)
    print(f"  → 全{len(all_records)}件のレクチャーを読み込み")

    print("\n[Step 2] フィルタリング")
    records = filter_records_by_date(
        all_records,
        date_from=args.date_from,
        date_to=args.date_to,
        speaker=args.speaker,
    )

    if args.date_from:
        print(f"  開始日: {args.date_from.strftime('%Y-%m-%d')}")
    if args.date_to:
        print(f"  終了日: {args.date_to.strftime('%Y-%m-%d')}")
    if args.speaker:
        print(f"  講師名: {args.speaker}")
    print(f"  → フィルタ後: {len(records)}件")

    if not records:
        print("[WARNING] 対象レクチャーが0件です。フィルタ条件を確認してください。")
        sys.exit(0)

    with_zoom = [r for r in records if r.has_zoom]
    without_zoom = [r for r in records if not r.has_zoom]
    print(f"  → Zoom URL/Meeting IDあり: {len(with_zoom)}件")
    print(f"  → Zoom URL/Meeting IDなし: {len(without_zoom)}件（スキップ）")

    resume_keys: set[str] = set()
    if args.resume_from:
        print(f"\n[Step 3] Resume判定: {args.resume_from}")
        resume_keys = load_resume_keys(args.resume_from)
    else:
        print("\n[Step 3] Resume: なし（フル実行）")

    print("\n[Step 4] Zoom APIから録画一覧取得")

    if args.dry_run:
        print("  [DRY-RUN] Zoom APIへのアクセスをスキップ")
        print("\n[Step 5] 突合結果（dry-run）")
        for r in records:
            zoom_flag = "○" if r.has_zoom else "×"
            skip_flag = " [SKIP:resume]" if r.lecture_key in resume_keys else ""
            title = (r.title_en or r.title_ja or "")[:40]
            print(
                f"  {r.date.strftime('%Y-%m-%d')} {r.time_slot:6s} "
                f"{(r.instructor_name_en or ''):30s} "
                f"{title:40s} "
                f"Zoom:{zoom_flag} MID:{r.meeting_id_raw:18s}"
                f"{skip_flag}"
            )
        print(f"\n  合計: {len(records)}件（Zoomあり: {len(with_zoom)}件）")
        print("  [DRY-RUN] 実際のダウンロードは行いません。")
        sys.exit(0)

    dates = [r.date for r in with_zoom]
    if not dates:
        print("  [WARNING] Zoom録画のあるレクチャーが0件です。")
        sys.exit(0)

    api_from = min(dates)
    api_to = max(dates)

    zoom_client = ZoomClient()
    zoom_recordings = zoom_client.list_recordings(api_from, api_to)

    print("\n[Step 5] CSVとZoom録画の突合")
    match_results = match_recordings(with_zoom, zoom_recordings)

    status_counts: dict[str, int] = {}
    for m in match_results:
        status_counts[m.status] = status_counts.get(m.status, 0) + 1
    print("  突合結果サマリー:")
    for status, count in sorted(status_counts.items()):
        print(f"    {status}: {count}件")

    print("\n[Step 6] ダウンロード実行")
    result_rows: list[dict] = []

    for r in without_zoom:
        result_rows.append(
            build_result_row(
                MatchResult(
                    lecture=r,
                    recording=None,
                    status="no_zoom_url",
                    message="Zoom URL/Meeting IDなし",
                ),
            )
        )

    download_count = 0
    skip_count = 0
    error_count = 0

    for match in match_results:
        lecture = match.lecture

        if lecture.lecture_key in resume_keys:
            result_rows.append(build_result_row(match, final_status="skipped_resume"))
            skip_count += 1
            print(f"  [SKIP] resume済み: {lecture.video_title or lecture.title_en}")
            continue

        if match.status != "matched":
            result_rows.append(build_result_row(match))
            if match.status in ("not_found", "not_ready"):
                error_count += 1
            print(
                f"  [{match.status.upper()}] {lecture.video_title or lecture.title_en}: {match.message}"
            )
            continue

        recording = match.recording
        final_path = create_download_path(lecture)
        staging_path = build_staging_download_path(lecture, final_path)
        print(f"\n  --- ダウンロード [{download_count + 1}] ---")
        print(f"  講師: {lecture.instructor_name_en}")
        print(f"  日付: {lecture.date.strftime('%Y-%m-%d')} {lecture.time_slot}")
        print(f"  タイトル: {lecture.video_title or lecture.title_en}")
        print(f"  録画タイプ: {recording.recording_type}")
        print(f"  サイズ: {recording.file_size / 1024 / 1024:.1f} MB")
        print(f"  一時保存: {staging_path.name}")
        print(f"  確定予定（配信用動画タイトル）: {final_path.name}")

        if (
            final_path.is_file()
            and recording.file_size > 0
            and final_path.stat().st_size == recording.file_size
        ):
            print(f"    [SKIP] 既存の確定ファイル（サイズ一致）: {final_path.name}")
            download_count += 1
            result_rows.append(
                build_result_row(
                    match,
                    local_path=str(final_path),
                    final_status="downloaded",
                )
            )
            continue

        token = zoom_client.get_download_url_with_token(recording.download_url)
        success = download_recording(
            download_url=recording.download_url,
            access_token=token,
            dest_path=staging_path,
            expected_size=recording.file_size,
        )

        if success:
            resolved = finalize_download_to_csv_title(staging_path, lecture)
            if resolved is not None:
                download_count += 1
                result_rows.append(
                    build_result_row(
                        match,
                        local_path=str(resolved),
                        final_status="downloaded",
                    )
                )
            else:
                error_count += 1
                result_rows.append(
                    build_result_row(
                        match,
                        local_path=str(staging_path),
                        final_status="rename_error",
                    )
                )
        else:
            error_count += 1
            result_rows.append(
                build_result_row(
                    match,
                    final_status="download_error",
                )
            )

    print("\n[Step 7] 結果CSV出力")

    if args.output_csv:
        output_csv_path = args.output_csv
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_csv_path = Path(f"result_{timestamp}.csv")

    write_result_csv(result_rows, output_csv_path)

    print(f"\n{'='*70}")
    print("  処理完了サマリー")
    print(f"{'='*70}")
    print(f"  CSV全レコード数:        {len(all_records)}")
    print(f"  フィルタ後:              {len(records)}")
    print(f"  Zoom URLあり:            {len(with_zoom)}")
    print(f"  Zoom URLなし（スキップ）: {len(without_zoom)}")
    print(f"  ダウンロード成功:        {download_count}")
    print(f"  スキップ（resume）:      {skip_count}")
    print(f"  エラー/未取得:           {error_count}")
    print(f"  結果CSV:                 {output_csv_path}")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
