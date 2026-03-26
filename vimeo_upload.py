"""
ZOOM フォルダ直下から実行するランチャー（zoom_download/vimeo_upload.py を起動）。
"""
import importlib.util
import sys
from pathlib import Path

_APP = Path(__file__).resolve().parent / "zoom_download"
_SCRIPT = _APP / "vimeo_upload.py"

if not _SCRIPT.is_file():
    sys.exit(f"見つかりません: {_SCRIPT}")

sys.path.insert(0, str(_APP))

spec = importlib.util.spec_from_file_location("zoom_vimeo_upload", _SCRIPT)
if spec is None or spec.loader is None:
    sys.exit(f"読み込めません: {_SCRIPT}")

_mod = importlib.util.module_from_spec(spec)
sys.modules["zoom_vimeo_upload"] = _mod
spec.loader.exec_module(_mod)
_mod.main()
