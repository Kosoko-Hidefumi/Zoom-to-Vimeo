"""
matcher.py - CSV ↔ Zoom録画の突合ロジック
"""
from dataclasses import dataclass
from datetime import timedelta

from csv_parser import LectureRecord
from zoom_client import ZoomRecordingFile
from config import Config


@dataclass
class MatchResult:
    """突合結果"""
    lecture: LectureRecord
    recording: ZoomRecordingFile | None
    status: str
    message: str = ""


def select_best_recording(
    candidates: list[ZoomRecordingFile],
) -> ZoomRecordingFile | None:
    mp4_files = [c for c in candidates if c.file_type.upper() == Config.TARGET_FILE_TYPE]
    if not mp4_files:
        return None

    completed = [c for c in mp4_files if c.status == "completed"]
    if not completed:
        processing = [c for c in mp4_files if c.status == "processing"]
        if processing:
            return None
        return None

    for preferred_type in Config.RECORDING_TYPE_PRIORITY:
        type_matches = [
            c for c in completed if c.recording_type.lower() == preferred_type.lower()
        ]
        if type_matches:
            return max(type_matches, key=lambda x: x.file_size)

    return max(completed, key=lambda x: x.file_size)


def match_recordings(
    lectures: list[LectureRecord],
    recordings: list[ZoomRecordingFile],
) -> list[MatchResult]:
    results: list[MatchResult] = []

    rec_index: dict[str, dict[str, list[ZoomRecordingFile]]] = {}
    for rec in recordings:
        mid = rec.meeting_id
        date_key = rec.recording_date_jst
        rec_index.setdefault(mid, {}).setdefault(date_key, []).append(rec)

    for lecture in lectures:
        if not lecture.has_zoom:
            results.append(
                MatchResult(
                    lecture=lecture,
                    recording=None,
                    status="no_zoom_url",
                    message="Zoom URL/Meeting IDが空のためスキップ",
                )
            )
            continue

        mid = lecture.meeting_id
        date_key = lecture.date.strftime("%Y-%m-%d")

        candidates = rec_index.get(mid, {}).get(date_key, [])

        if not candidates:
            prev_day = (lecture.date - timedelta(days=1)).strftime("%Y-%m-%d")
            next_day = (lecture.date + timedelta(days=1)).strftime("%Y-%m-%d")
            candidates = (
                rec_index.get(mid, {}).get(prev_day, [])
                + rec_index.get(mid, {}).get(next_day, [])
            )

        if not candidates:
            results.append(
                MatchResult(
                    lecture=lecture,
                    recording=None,
                    status="not_found",
                    message=f"Meeting ID {mid} の {date_key} に該当録画なし",
                )
            )
            continue

        if lecture.start_time:
            time_filtered = []
            for c in candidates:
                rec_time = c.recording_time_jst
                lecture_minutes = lecture.start_time.hour * 60 + lecture.start_time.minute
                rec_minutes = rec_time.hour * 60 + rec_time.minute
                diff = abs(lecture_minutes - rec_minutes)
                if diff <= 90:
                    time_filtered.append(c)

            if time_filtered:
                candidates = time_filtered

        best = select_best_recording(candidates)

        if best is None:
            processing_any = any(
                c.status == "processing" and c.file_type.upper() == "MP4"
                for c in candidates
            )
            if processing_any:
                results.append(
                    MatchResult(
                        lecture=lecture,
                        recording=None,
                        status="not_ready",
                        message=f"Meeting ID {mid} の {date_key} の録画は処理中",
                    )
                )
            else:
                results.append(
                    MatchResult(
                        lecture=lecture,
                        recording=None,
                        status="not_found",
                        message=f"Meeting ID {mid} の {date_key} に適切なMP4録画なし",
                    )
                )
            continue

        results.append(
            MatchResult(
                lecture=lecture,
                recording=best,
                status="matched",
                message=f"突合成功: {best.recording_type} ({best.file_size / 1024 / 1024:.1f} MB)",
            )
        )

    return results
