"""
result_csv.py - 結果CSV出力・resume判定
"""
import csv
from datetime import datetime
from pathlib import Path

from matcher import MatchResult


RESULT_COLUMNS = [
    "lecture_key",
    "row_number",
    "date",
    "time_slot",
    "start_time",
    "instructor_name_en",
    "specialty",
    "video_title",
    "meeting_id",
    "status",
    "recording_type",
    "file_size_mb",
    "local_file_path",
    "message",
    "processed_at",
]


def load_resume_keys(resume_csv_path: Path) -> set[str]:
    completed_keys: set[str] = set()
    if not resume_csv_path.exists():
        return completed_keys

    with open(resume_csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("status") == "downloaded":
                key = row.get("lecture_key", "")
                if key:
                    completed_keys.add(key)

    print(f"  [RESUME] {len(completed_keys)}件のダウンロード済みレコードをスキップ")
    return completed_keys


def write_result_csv(
    results: list[dict],
    output_path: Path,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=RESULT_COLUMNS)
        writer.writeheader()
        for row in results:
            writer.writerow(row)

    print(f"  [CSV] 結果CSV出力: {output_path} ({len(results)}件)")


def build_result_row(
    match: MatchResult,
    local_path: str = "",
    final_status: str = "",
) -> dict:
    lecture = match.lecture
    recording = match.recording

    status = final_status or match.status

    return {
        "lecture_key": lecture.lecture_key,
        "row_number": lecture.row_number,
        "date": lecture.date.strftime("%Y-%m-%d"),
        "time_slot": lecture.time_slot,
        "start_time": lecture.start_time.strftime("%H:%M") if lecture.start_time else "",
        "instructor_name_en": lecture.instructor_name_en,
        "specialty": lecture.specialty,
        "video_title": lecture.video_title,
        "meeting_id": lecture.meeting_id,
        "status": status,
        "recording_type": recording.recording_type if recording else "",
        "file_size_mb": f"{recording.file_size / 1024 / 1024:.1f}" if recording else "",
        "local_file_path": local_path,
        "message": match.message,
        "processed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
