# Zoom録画ダウンロード（CSV突合）

CSVを正として Zoom クラウド録画を突合・ダウンロードします。Server-to-Server OAuth を使用します。

## ディレクトリ

- **`main.py`**（このフォルダ直下）: ランチャー。内部で `zoom_download/main.py` を実行します。
- **`zoom_download/`**: アプリ本体（設定・モジュール・`.env` はここを基準にします）。

## セットアップ

1. Zoom Marketplace で **Server-to-Server OAuth** アプリを作成する。
2. **Scopes** に以下を追加する。
   - `cloud_recording:read:list_user_recordings:admin`
   - `cloud_recording:read:list_recording_files:admin`
3. アプリをアクティベートする。
4. 依存関係を入れる。

```bash
cd zoom_download
pip install -r requirements.txt
copy .env.example .env
```

5. `zoom_download/.env` に Account ID / Client ID / Client Secret を記入する。
6. Vimeo にアップロードする場合は同じ `.env` に **`VIMEO_TOKEN`** を記入する。

## 使い方

プロジェクト直下（`ZOOM`）から:

```bash
# 既定CSV: consultant_vimeo.csv（カレントまたは zoom_download 内）

python main.py --from 2025-07-14 --to 2025-07-18 --dry-run
python main.py --from 2025-07-14 --to 2025-07-18
python main.py --speaker "Pangilinan"
python main.py
python main.py --csv other.csv --from 2025-07-14 --to 2025-07-18
python main.py --from 2025-07-14 --to 2025-07-18 --resume-from result_20250720_100000.csv
```

`zoom_download` に入って直接実行する場合も同様です（`python main.py ...`）。

### Vimeo アップロード（別コマンド）

Zoom ダウンロード後、**Vimeo 専用の結果 CSV**（例: `vimeo_results.csv`）を `--out-csv` / `--resume-from` に使います（Zoom の `result_*.csv` とは別）。

- 突合: CSV の **配信用動画タイトル** と、`DOWNLOAD_DIR` 以下の **mp4 のファイル名（拡張子除く）** — Zoom 側と同じ `sanitize_filename` でキー化。
- 既に Vimeo に同じキーのタイトルがある場合は **スキップ**（`skipped(already_on_vimeo)`）。
- API は `https://api.vimeo.com` 固定。

プロジェクト直下から:

```bash
python vimeo_upload.py --csv consultant_vimeo.csv --out-csv vimeo_results.csv --dry-run
python vimeo_upload.py --csv consultant_vimeo.csv --out-csv vimeo_results.csv
python vimeo_upload.py --root downloads --csv consultant_vimeo.csv --out-csv vimeo_results.csv --resume-from vimeo_results.csv
```

`--no-vimeo-check` で Vimeo 上の一覧取得を省略（オフライン確認用）。

## CSV 形式

`consultant_vimeo.csv` と同じ日本語ヘッダー（区分、講師No.、日付、ミーティングID など）に対応しています。

## 設計メモ

- 突合: ミーティングID（空白除去）＋日付（JST）＋開始時刻（±90分）。
- 録画種別は `shared_screen_with_speaker_view` を最優先し、MP4 のみ、同一種別はサイズ最大を採用。
