"""
zoom_client.py - Zoom API認証・録画情報取得
"""
import time as time_module
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone, time
import requests

from config import Config

JST = timezone(timedelta(hours=9))


@dataclass
class ZoomRecordingFile:
    """Zoom録画ファイル1つの情報"""
    id: str
    meeting_id: str
    meeting_uuid: str
    topic: str
    recording_start: datetime
    recording_start_jst: datetime
    recording_end: datetime | None
    file_type: str
    file_size: int
    recording_type: str
    download_url: str
    status: str

    @property
    def recording_date_jst(self) -> str:
        return self.recording_start_jst.strftime("%Y-%m-%d")

    @property
    def recording_time_jst(self) -> time:
        return self.recording_start_jst.time()


class ZoomClient:
    """Zoom Server-to-Server OAuth Client"""

    def __init__(self):
        self.account_id = Config.ZOOM_ACCOUNT_ID
        self.client_id = Config.ZOOM_CLIENT_ID
        self.client_secret = Config.ZOOM_CLIENT_SECRET
        self.user_id = Config.ZOOM_USER_ID
        self.base_url = Config.ZOOM_API_BASE
        self._access_token: str | None = None
        self._token_expires_at: float = 0

    def _get_access_token(self) -> str:
        now = time_module.time()
        if self._access_token and now < self._token_expires_at - 60:
            return self._access_token

        print("  [ZOOM] Requesting new access token...")
        resp = requests.post(
            Config.ZOOM_OAUTH_URL,
            data={
                "grant_type": "account_credentials",
                "account_id": self.account_id,
            },
            auth=(self.client_id, self.client_secret),
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=30,
        )
        if not resp.ok:
            detail = resp.text
            try:
                err = resp.json()
                reason = err.get("reason")
                err_code = err.get("error")
                if reason or err_code:
                    detail = f"{err_code or ''} {reason or ''}".strip() or detail
            except Exception:
                pass
            print(f"  [ZOOM] HTTP {resp.status_code}: トークン取得に失敗しました。")
            print(f"  [ZOOM] Zoomからの応答: {detail}")
            print(
                "  [ZOOM] 確認: (1) Marketplace の Server-to-Server OAuth アプリが Activate 済み"
                " (2) App Credentials の Account ID / Client ID / Client Secret が .env と一致"
                " (3) 値に余計なスペースや引用符がない (4) EU アカウントなら .env に"
                " ZOOM_OAUTH_URL=https://zoom.eu/oauth/token を試す"
            )
            resp.raise_for_status()
        data = resp.json()

        self._access_token = data["access_token"]
        self._token_expires_at = now + data.get("expires_in", 3600)
        print("  [ZOOM] Access token obtained successfully.")
        return self._access_token

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._get_access_token()}",
            "Content-Type": "application/json",
        }

    def list_recordings(
        self,
        date_from: datetime,
        date_to: datetime,
    ) -> list[ZoomRecordingFile]:
        """
        ホストユーザーの録画一覧を期間指定で取得。
        """
        all_files: list[ZoomRecordingFile] = []
        current_from = date_from

        while current_from <= date_to:
            current_to = min(
                current_from + timedelta(days=Config.MAX_DATE_RANGE_DAYS - 1),
                date_to,
            )

            from_str = current_from.strftime("%Y-%m-%d")
            to_str = current_to.strftime("%Y-%m-%d")
            print(f"  [ZOOM] Fetching recordings: {from_str} ~ {to_str}")

            page_token = ""
            while True:
                params = {
                    "from": from_str,
                    "to": to_str,
                    "page_size": 300,
                }
                if page_token:
                    params["next_page_token"] = page_token

                resp = requests.get(
                    f"{self.base_url}/users/{self.user_id}/recordings",
                    headers=self._headers(),
                    params=params,
                    timeout=60,
                )
                resp.raise_for_status()
                data = resp.json()

                meetings = data.get("meetings", [])
                for meeting in meetings:
                    meeting_id_num = str(meeting.get("id", ""))
                    meeting_uuid = meeting.get("uuid", "")
                    topic = meeting.get("topic", "")

                    for rf in meeting.get("recording_files", []):
                        file_type = rf.get("file_type", "")
                        status = rf.get("status", "")

                        rec_start_str = rf.get("recording_start", "")
                        rec_end_str = rf.get("recording_end", "")
                        rec_start_utc = self._parse_zoom_datetime(rec_start_str)
                        rec_end_utc = (
                            self._parse_zoom_datetime(rec_end_str) if rec_end_str else None
                        )

                        if rec_start_utc is None:
                            continue

                        rec_start_jst = rec_start_utc.astimezone(JST)

                        recording_file = ZoomRecordingFile(
                            id=rf.get("id", ""),
                            meeting_id=meeting_id_num,
                            meeting_uuid=meeting_uuid,
                            topic=topic,
                            recording_start=rec_start_utc,
                            recording_start_jst=rec_start_jst,
                            recording_end=rec_end_utc,
                            file_type=file_type,
                            file_size=rf.get("file_size", 0),
                            recording_type=rf.get("recording_type", ""),
                            download_url=rf.get("download_url", ""),
                            status=status,
                        )
                        all_files.append(recording_file)

                page_token = data.get("next_page_token", "")
                if not page_token:
                    break

            current_from = current_to + timedelta(days=1)

        print(f"  [ZOOM] Total recording files fetched: {len(all_files)}")
        return all_files

    def get_download_url_with_token(self, download_url: str) -> str:
        return self._get_access_token()

    @staticmethod
    def _parse_zoom_datetime(dt_str: str) -> datetime | None:
        if not dt_str:
            return None
        for fmt in (
            "%Y-%m-%dT%H:%M:%SZ",
            "%Y-%m-%dT%H:%M:%S.%fZ",
            "%Y-%m-%dT%H:%M:%S",
        ):
            try:
                dt = datetime.strptime(dt_str, fmt)
                return dt.replace(tzinfo=timezone.utc)
            except ValueError:
                continue
        return None
