# 手順書：Zoom クラウド録画の取得から Vimeo アップロードまで

このドキュメントは、**`consultant_vimeo.csv` を正**として Zoom から MP4 を落とし、続けて Vimeo にアップロードするまでの流れです。作業ディレクトリの例は **`D:\code4biz\ZOOM`** です。

---

## 1. 前提

- **Python 3.10+**（3.12 利用可）がインストールされていること
- **Zoom**：Marketplace で **Server-to-Server OAuth** アプリが作成済みで Activate 済みであること
- **Vimeo**：アップロード用の **アクセストークン**（`VIMEO_TOKEN`）を取得済みであること
- 講師・日付・ミーティング ID・配信タイトルなどが入った **`consultant_vimeo.csv`**（または同等の CSV）があること

---

## 2. 初回セットアップ


### 2.1 依存パッケージ

PowerShell 例：

```powershell
cd D:\code4biz\ZOOM\zoom_download
pip install -r requirements.txt
```

### 2.2 環境変数ファイル

```powershell
cd D:\code4biz\ZOOM\zoom_download
copy .env.example .env
notepad .env
```

`zoom_download\.env` に **少なくとも** 次を設定します。

| 変数 | 用途 |
|------|------|
| `ZOOM_ACCOUNT_ID` | Zoom S2S OAuth：Account ID |
| `ZOOM_CLIENT_ID` | Zoom S2S OAuth：Client ID |
| `ZOOM_CLIENT_SECRET` | Zoom S2S OAuth：Client Secret |
| `ZOOM_USER_ID` | 通常は `me`（特定ユーザーならメール等） |
| `DOWNLOAD_DIR` | 保存先（例：`./downloads` ※実行時のカレントからの相対） |
| `VIMEO_TOKEN` | Vimeo API 用トークン（アップロード工程で使用） |

Zoom OAuth が 400 になる場合は、[README.md](README.md) の Zoom EU 向け URL 記載や、Marketplace の Credentials を再確認してください。

### 2.3 CSV の配置

- **`D:\code4biz\ZOOM\consultant_vimeo.csv`** に置く（推奨）  
  または `zoom_download\consultant_vimeo.csv`  
- Zoom の `main.py` は **既定で `consultant_vimeo.csv`** を探します（カレント → `zoom_download` 内）。

---

## 3. Zoom からのダウンロード

プロジェクト直下でランチャーを使います。

```powershell
cd D:\code4biz\ZOOM
```

### 3.1 動作確認（API・ダウンロードなし）

```powershell
python main.py --from 2025-07-14 --to 2025-07-18 --dry-run
```

### 3.2 期間を指定してダウンロード

```powershell
python main.py --from 2025-07-14 --to 2025-07-18
```

### 3.3 CSV 全行を対象にする（日付フィルタなし）

```powershell
python main.py
```

### 3.4 別の CSV を使う

```powershell
python main.py --csv .\別名.csv --from 2025-09-01 --to 2025-09-10
```

### 3.5 途中から再開（Zoom 結果 CSV）

前回出力された **`result_YYYYMMDD_HHMMSS.csv`** を `--resume-from` に指定します（**Zoom 工程用**。Vimeo 用 CSV とは別です）。

```powershell
python main.py --from 2025-07-14 --to 2025-07-18 --resume-from .\result_20250326_120000.csv
```

### 3.6 保存先について

- `DOWNLOAD_DIR` 既定 `./downloads` のとき、`D:\code4biz\ZOOM` から実行すると **`D:\code4biz\ZOOM\downloads\`** 以下に  
  **`講師名フォルダ\日付\配信用動画タイトル由来のファイル名.mp4`** で保存されます。  
- 一時ファイルはダウンロード後に **CSV の「配信用動画タイトル」に合わせて確定名**へリネームされます。

### 3.7 Zoom 結果 CSV

実行ごとに **`result_*.csv`** がカレント（例：`D:\code4biz\ZOOM`）に出力されます。ダウンロード成功・失敗の記録用です。

---

## 4. Vimeo へのアップロード

**Zoom の結果 CSV は使いません。** Vimeo 工程では **`--out-csv` で出力する専用 CSV** を `--resume-from` に渡して再開します。

```powershell
cd D:\code4biz\ZOOM
```

### 4.1 事前確認（アップロードは実行しない）

```powershell
python vimeo_upload.py --csv consultant_vimeo.csv --out-csv vimeo_results.csv --dry-run
```

### 4.2 本番アッpload

- **`--root`** を省略すると `.env` の **`DOWNLOAD_DIR`** から MP4 を再帰検索します（通常は `downloads` で問題ありません）。

```powershell
python vimeo_upload.py --csv consultant_vimeo.csv --out-csv vimeo_results.csv
```

明示的にルートを指定する場合：

```powershell
python vimeo_upload.py --root .\downloads --csv consultant_vimeo.csv --out-csv vimeo_results.csv
```

### 4.3 途中から再開（Vimeo 専用 CSV）

```powershell
python vimeo_upload.py --csv consultant_vimeo.csv --out-csv vimeo_results.csv --resume-from .\vimeo_results.csv
```

### 4.4 補足

- 突合は **CSV「配信用動画タイトル」** と **ローカル MP4 のファイル名（拡張子除く）** を同一ルールでキー化して行います（Zoom 保存時と整合）。
- **既に Vimeo 上に同じタイトルキーがある動画はスキップ**されます（`skipped(already_on_vimeo)`）。
- 一覧 API を叩きたくないとき：`--no-vimeo-check`

---

## 5. （任意）Vimeo リンク付きメタデータ CSV の出力

`consultant_vimeo.csv` の講師名・専門科などと、Vimeo の **`link`** を突合した一覧を出します。

```powershell
cd D:\code4biz\ZOOM
python export_vimeo_metadata_csv.py --out-csv vimeo_metadata_export.csv
```

- `--source-csv` を省略すると **`consultant_vimeo.csv`** を自動検索します。
- **エクスポート先の CSV を Excel で開いたまま**書き込むと `PermissionError` になることがあります。**ファイルを閉じてから**再実行するか、`--out-csv` で別名を指定してください。スクリプト側でロック時に **日時付き別名**へ逃がす処理も入っています。

Vimeo の動画一覧だけ欲しい場合：

```powershell
python export_vimeo_metadata_csv.py --vimeo-only --out-csv vimeo_only_list.csv
```

---

## 6. 処理の流れ（一覧）

1. `zoom_download\.env` を整備（Zoom + `VIMEO_TOKEN`）
2. `consultant_vimeo.csv` を配置
3. `python main.py ...` で Zoom から MP4 を `downloads`（または `DOWNLOAD_DIR`）へ取得
4. `python vimeo_upload.py ... --out-csv vimeo_results.csv` で Vimeo にアップロード
5. 必要なら `python export_vimeo_metadata_csv.py` でリンク付き一覧 CSV を出力

---

## 7. トラブルのヒント

| 現象 | 対処の例 |
|------|-----------|
| Zoom トークン 400 | `.env` の Account ID / Client ID / Secret、アプリ Activate、EU 向け URL（README 参照） |
| Vimeo 書き込み Permission denied | 出力・入力 CSV を Excel で閉じる、`--out-csv` で別ファイル |
| アップロードで `missing` が多い | `DOWNLOAD_DIR` と `--root` が実際の MP4 場所と一致しているか、ファイル名が CSV の「配信用動画タイトル」と一致しているか |

以上です。
