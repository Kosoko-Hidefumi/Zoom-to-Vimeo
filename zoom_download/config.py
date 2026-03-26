"""
config.py - 設定管理
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# どのカレントディレクトリから起動しても zoom_download 内の .env を読む
_APP_DIR = Path(__file__).resolve().parent
load_dotenv(_APP_DIR / ".env")
load_dotenv()  # カレントディレクトリの .env（あれば上書き）


def _env(key: str, default: str = "") -> str:
    """環境変数の前後空白・改行を除去（コピペミス対策）。"""
    raw = os.getenv(key, default)
    if raw is None:
        return default
    return raw.strip()


class Config:
    # Zoom API
    ZOOM_ACCOUNT_ID: str = _env("ZOOM_ACCOUNT_ID")
    ZOOM_CLIENT_ID: str = _env("ZOOM_CLIENT_ID")
    ZOOM_CLIENT_SECRET: str = _env("ZOOM_CLIENT_SECRET")
    ZOOM_USER_ID: str = _env("ZOOM_USER_ID", "me")

    # Vimeo（別コマンド vimeo_upload.py 用）
    VIMEO_TOKEN: str = _env("VIMEO_TOKEN")

    # Paths
    DOWNLOAD_DIR: Path = Path(_env("DOWNLOAD_DIR", "./downloads"))

    # Timezone
    TIMEZONE: str = _env("TIMEZONE", "Asia/Tokyo")

    # Recording type priority (descending)
    RECORDING_TYPE_PRIORITY: list[str] = [
        "shared_screen_with_speaker_view",
        "shared_screen_with_speaker_view(CC)",
        "shared_screen_with_gallery_view",
        "active_speaker",
        "gallery_view",
        "shared_screen",
    ]

    # File type filter
    TARGET_FILE_TYPE: str = "MP4"

    # API（EU データ所在地などで zoom.eu が必要な場合は .env で ZOOM_OAUTH_URL を上書き）
    ZOOM_OAUTH_URL: str = _env("ZOOM_OAUTH_URL", "https://zoom.us/oauth/token")
    ZOOM_API_BASE: str = _env("ZOOM_API_BASE", "https://api.zoom.us/v2")

    # Retry
    MAX_RETRIES: int = 3
    RETRY_DELAY: float = 5.0  # seconds
    DOWNLOAD_CHUNK_SIZE: int = 8192  # bytes

    # List recordings API: max date range is 1 month
    MAX_DATE_RANGE_DAYS: int = 30

    @classmethod
    def validate(cls) -> list[str]:
        """必須設定の検証"""
        errors = []
        if not cls.ZOOM_ACCOUNT_ID:
            errors.append("ZOOM_ACCOUNT_ID is not set")
        if not cls.ZOOM_CLIENT_ID:
            errors.append("ZOOM_CLIENT_ID is not set")
        if not cls.ZOOM_CLIENT_SECRET:
            errors.append("ZOOM_CLIENT_SECRET is not set")
        for label, val in (
            ("ZOOM_ACCOUNT_ID", cls.ZOOM_ACCOUNT_ID),
            ("ZOOM_CLIENT_ID", cls.ZOOM_CLIENT_ID),
            ("ZOOM_CLIENT_SECRET", cls.ZOOM_CLIENT_SECRET),
        ):
            if val and ("ここに" in val or "貼り付け" in val or "your_" in val.lower()):
                errors.append(
                    f"{label} がプレースホルダのままです。"
                    " Zoom Marketplace の Server-to-Server OAuth アプリの実値を入れてください。"
                )
        return errors
