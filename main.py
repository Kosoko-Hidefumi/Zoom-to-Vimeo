"""
ZOOM フォルダ直下から実行するランチャー。
実体は zoom_download/main.py（既定CSV: consultant_vimeo.csv）です。

例:
  python main.py --from 2025-07-14 --to 2025-07-18 --dry-run
  python main.py --csv other.csv --from 2025-07-14 --to 2025-07-18
"""
import importlib.util
import sys
from pathlib import Path

_APP = Path(__file__).resolve().parent / "zoom_download"
_MAIN = _APP / "main.py"

if not _MAIN.is_file():
    sys.exit(f"見つかりません: {_MAIN}\nzoom_download フォルダを確認してください。")

sys.path.insert(0, str(_APP))

spec = importlib.util.spec_from_file_location("zoom_download_main", _MAIN)
if spec is None or spec.loader is None:
    sys.exit(f"読み込めません: {_MAIN}")

_mod = importlib.util.module_from_spec(spec)
sys.modules["zoom_download_main"] = _mod
spec.loader.exec_module(_mod)
_mod.main()
