"""
folder_manager.py - フォルダ構造の作成
"""
from pathlib import Path

from csv_parser import LectureRecord
from config import Config


def create_download_path(lecture: LectureRecord) -> Path:
    """
    CSV（配信用動画タイトル等）に基づく最終保存パス。
    講師名フォルダ/日付フォルダ/ファイル名
    """
    base_dir = Config.DOWNLOAD_DIR
    speaker_dir = base_dir / lecture.speaker_folder_name
    date_dir = speaker_dir / lecture.date_folder_name
    file_path = date_dir / lecture.download_filename

    date_dir.mkdir(parents=True, exist_ok=True)

    return file_path


def build_staging_download_path(lecture: LectureRecord, final_path: Path) -> Path:
    """ダウンロード中の一時ファイルパス（同一フォルダ内・確定名とは別名）。"""
    final_path.parent.mkdir(parents=True, exist_ok=True)
    safe_mid = lecture.meeting_id or "noid"
    return final_path.parent / f".__zoom_dl_r{lecture.row_number}_{safe_mid}.mp4"


def finalize_download_to_csv_title(staging_path: Path, lecture: LectureRecord) -> Path | None:
    """
    ダウンロード完了後、CSVの配信用動画タイトル（lecture.download_filename）へリネームする。
    成功時は最終パス、失敗時は None。
    """
    if not staging_path.is_file():
        print(f"    [ERROR] 一時ファイルが見つかりません: {staging_path}")
        return None

    final_path = create_download_path(lecture)

    try:
        if staging_path.resolve() == final_path.resolve():
            print(f"    [OK] 保存完了: {final_path.name}")
            return final_path
        if final_path.exists():
            final_path.unlink()
        staging_path.rename(final_path)
        print(f"    [RENAME] CSVの配信用動画タイトルに合わせて確定: {final_path.name}")
        return final_path
    except OSError as e:
        print(f"    [ERROR] リネーム失敗（一時ファイルは残しています）: {e}")
        return None
