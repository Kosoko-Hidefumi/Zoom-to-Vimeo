"""
csv_parser.py - CSVの読み込み・パース・バリデーション
"""
import csv
import re
from dataclasses import dataclass, field
from datetime import datetime, time
from pathlib import Path


def sanitize_filename(name: str) -> str:
    """ファイルシステムで使えない文字を置換"""
    invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
    sanitized = re.sub(invalid_chars, "_", name)
    sanitized = sanitized.strip(" .")
    sanitized = re.sub(r"_+", "_", sanitized)
    if len(sanitized) > 200:
        sanitized = sanitized[:200]
    return sanitized


@dataclass
class LectureRecord:
    """CSVの1行に対応するレクチャー情報"""
    row_number: int
    category: str
    instructor_id: str
    instructor_name_ja: str
    instructor_name_en: str
    specialty: str
    affiliation: str
    date: datetime
    time_slot: str
    start_time: time | None
    end_time: time | None
    location: str
    title_ja: str
    title_en: str
    video_title: str
    zoom_url: str
    meeting_id: str
    meeting_id_raw: str
    passcode: str
    remarks: str

    lecture_key: str = field(init=False)

    def __post_init__(self):
        date_str = self.date.strftime("%Y-%m-%d")
        time_str = self.start_time.strftime("%H:%M") if self.start_time else "unknown"
        self.lecture_key = f"{self.meeting_id}_{date_str}_{time_str}"

    @property
    def has_zoom(self) -> bool:
        return bool(self.meeting_id)

    @property
    def speaker_folder_name(self) -> str:
        name = self.instructor_name_en or self.instructor_name_ja
        return sanitize_filename(name) if name else "Unknown_Speaker"

    @property
    def date_folder_name(self) -> str:
        return self.date.strftime("%Y-%m-%d")

    @property
    def download_filename(self) -> str:
        if self.video_title:
            return sanitize_filename(self.video_title) + ".mp4"
        name = self.instructor_name_en or self.instructor_name_ja or "Unknown"
        title = self.title_en or self.title_ja or "Untitled"
        date_str = self.date.strftime("%Y.%m.%d")
        return sanitize_filename(f"[{self.specialty}] {name} - {title} ({date_str})") + ".mp4"


def parse_time(time_str: str) -> time | None:
    """時刻文字列をtimeオブジェクトに変換"""
    if not time_str or not time_str.strip():
        return None
    time_str = time_str.strip()
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(time_str, fmt).time()
        except ValueError:
            continue
    return None


def normalize_meeting_id(mid: str) -> str:
    """Meeting IDから空白を除去して数字のみにする"""
    if not mid:
        return ""
    return re.sub(r"\s+", "", mid.strip())


def parse_csv(csv_path: Path) -> list[LectureRecord]:
    """
    CSVファイルを読み込み、LectureRecordのリストを返す。
    """
    records: list[LectureRecord] = []

    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)

        for row_num, row in enumerate(reader, start=2):
            try:
                date_str = row.get("日付", "").strip()
                if not date_str:
                    continue

                lecture_date = None
                for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d"):
                    try:
                        lecture_date = datetime.strptime(date_str, fmt)
                        break
                    except ValueError:
                        continue
                if lecture_date is None:
                    print(f"  [WARNING] Row {row_num}: 日付パース失敗 '{date_str}'")
                    continue

                raw_mid = row.get("ミーティングID", "").strip()
                meeting_id = normalize_meeting_id(raw_mid)

                record = LectureRecord(
                    row_number=row_num,
                    category=row.get("区分", "").strip(),
                    instructor_id=row.get("講師No.", "").strip(),
                    instructor_name_ja=row.get("講師名（日本語）", "").strip(),
                    instructor_name_en=row.get("講師名（英語）", "").strip(),
                    specialty=row.get("専門科", "").strip(),
                    affiliation=row.get("所属", "").strip(),
                    date=lecture_date,
                    time_slot=row.get("時間帯", "").strip(),
                    start_time=parse_time(row.get("開始時刻", "")),
                    end_time=parse_time(row.get("終了時刻", "")),
                    location=row.get("場所", "").strip(),
                    title_ja=row.get("レクチャータイトル（日本語）", "").strip(),
                    title_en=row.get("レクチャータイトル（英語）", "").strip(),
                    video_title=row.get("配信用動画タイトル", "").strip(),
                    zoom_url=row.get("Zoom URL", "").strip(),
                    meeting_id=meeting_id,
                    meeting_id_raw=raw_mid,
                    passcode=row.get("パスコード", "").strip(),
                    remarks=row.get("備考", "").strip(),
                )
                records.append(record)

            except Exception as e:
                print(f"  [ERROR] Row {row_num}: パース例外 {e}")
                continue

    return records


def filter_records_by_date(
    records: list[LectureRecord],
    date_from: datetime | None = None,
    date_to: datetime | None = None,
    speaker: str | None = None,
) -> list[LectureRecord]:
    """日付範囲・講師名でフィルタリング"""
    filtered = records

    if date_from:
        filtered = [r for r in filtered if r.date >= date_from]
    if date_to:
        filtered = [r for r in filtered if r.date <= date_to]
    if speaker:
        speaker_lower = speaker.lower()
        filtered = [
            r
            for r in filtered
            if speaker_lower in (r.instructor_name_en or "").lower()
            or speaker_lower in (r.instructor_name_ja or "").lower()
        ]

    return filtered
