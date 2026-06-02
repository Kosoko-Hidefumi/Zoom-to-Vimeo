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

Zoom ダウンロード・Vimeo アップロード用（必須）:

```powershell
cd D:\code4biz\ZOOM\zoom_download
pip install -r requirements.txt
```

メール解析（`parse_lecture_email.py`）を使う場合は追加でインストール:

```powershell
# Claude API（必須）
pip install anthropic

# Outlook 連携（Windows + Outlook がインストール済みの場合のみ）
pip install pywin32
```

### 2.2 環境変数ファイル

```powershell
cd D:\code4biz\ZOOM\zoom_download
copy .env.example .env
notepad .env
```

`zoom_download\.env` に **少なくとも** 次を設定します。

| 変数 | 用途 | 既定値 |
|------|------|--------|
| `ZOOM_ACCOUNT_ID` | Zoom S2S OAuth：Account ID | （必須） |
| `ZOOM_CLIENT_ID` | Zoom S2S OAuth：Client ID | （必須） |
| `ZOOM_CLIENT_SECRET` | Zoom S2S OAuth：Client Secret | （必須） |
| `ZOOM_USER_ID` | 通常は `me`（特定ユーザーならメール等） | `me` |
| `DOWNLOAD_DIR` | 保存先（例：`./downloads` ※実行時のカレントからの相対） | `./downloads` |
| `VIMEO_TOKEN` | Vimeo API 用トークン（アップロード工程で使用） | （必須） |
| `ANTHROPIC_API_KEY` | Claude API キー（`parse_lecture_email.py` で使用） | （メール解析時のみ必須） |
| `TIMEZONE` | 日付突合に使うタイムゾーン | `Asia/Tokyo` |

> **EU アカウントの場合（オプション）**  
> Zoom が 400 エラーを返す場合は、次の変数も追記してください。  
> `ZOOM_OAUTH_URL=https://zoom.eu/oauth/token`  
> `ZOOM_API_BASE=https://api.eu.zoom.us/v2`

スクリプトは `zoom_download/.env` → カレントの `.env` の順で読み込みます。  
`ANTHROPIC_API_KEY` はプロジェクト直下の `.env`（`D:\code4biz\ZOOM\.env`）に書いても構いません。

### 2.3 CSV の配置

- **`D:\code4biz\ZOOM\consultant_vimeo.csv`** に置く（推奨）  
  または `zoom_download\consultant_vimeo.csv`  
- Zoom の `main.py` は **既定で `consultant_vimeo.csv`** を探します（カレント → `zoom_download` 内）。

---

## 3. 招聘メールから `consultant_vimeo.csv` への自動追記

コンサルタント招聘案内メールを **Claude API** で解析し、`consultant_vimeo.csv` に行を追記するツールです。  
Zoom ダウンロードの前工程として実行します。

```powershell
cd D:\code4biz\ZOOM
```

### 3.1 テキストファイルから読み込む（テスト用）

```powershell
python parse_lecture_email.py --file email_sample.txt
```

### 3.2 Outlook 受信トレイから自動取得

件名に「コンサルタント」「レクチャー」「Lecture」を含むメールを検索します（Outlook の起動が必要）。

```powershell
python parse_lecture_email.py              # 過去7日間
python parse_lecture_email.py --days 14   # 過去14日間
```

複数メールが見つかった場合は番号で選択します。

### 3.3 別の CSV を指定する

```powershell
python parse_lecture_email.py --csv D:\other\path\consultant_vimeo.csv --file email.txt
```

### 3.4 重複行を強制追記する

デフォルトは「日付 × 開始時刻 × 講師名（英語）」の重複をスキップします。

```powershell
python parse_lecture_email.py --file email.txt --force
```

### 3.5 タイトル未確定行を更新する（`--update` モード）

備考欄に `【タイトル未確定】` と入っている行のタイトルを、新しいメールの情報で上書きします。

```powershell
python parse_lecture_email.py --update --file email_with_titles.txt
python parse_lecture_email.py --update              # Outlook から取得
```

> **補足**  
> - 追記前に必ずプレビューが表示され、`y` で確定 / `Enter` でキャンセルします。  
> - 追記後、`consultant_vimeo.csv` の「配信用動画タイトル」が空の行は `--update` で補完できます。

---

## 4. Zoom からのダウンロード

プロジェクト直下でランチャーを使います。

```powershell
cd D:\code4biz\ZOOM
```

### 4.1 動作確認（API・ダウンロードなし）

```powershell
python main.py --from 2025-07-14 --to 2025-07-18 --dry-run
```

### 4.2 期間を指定してダウンロード

```powershell
python main.py --from 2025-07-14 --to 2025-07-18
```

### 4.3 CSV 全行を対象にする（日付フィルタなし）

```powershell
python main.py
```

### 4.4 別の CSV を使う

```powershell
python main.py --csv .\別名.csv --from 2025-09-01 --to 2025-09-10
```

### 4.5 途中から再開（Zoom 結果 CSV）

前回出力された **`result_YYYYMMDD_HHMMSS.csv`** を `--resume-from` に指定します（**Zoom 工程用**。Vimeo 用 CSV とは別です）。

```powershell
python main.py --from 2025-07-14 --to 2025-07-18 --resume-from .\result_20250326_120000.csv
```

### 4.6 保存先について

- `DOWNLOAD_DIR` 既定 `./downloads` のとき、`D:\code4biz\ZOOM` から実行すると **`D:\code4biz\ZOOM\downloads\`** 以下に  
  **`講師名フォルダ\日付\配信用動画タイトル由来のファイル名.mp4`** で保存されます。  
- 一時ファイルはダウンロード後に **CSV の「配信用動画タイトル」に合わせて確定名**へリネームされます。

### 4.7 Zoom 結果 CSV

実行ごとに **`result_*.csv`** がカレント（例：`D:\code4biz\ZOOM`）に出力されます。ダウンロード成功・失敗の記録用です。

---

## 5. Vimeo へのアップロード

**Zoom の結果 CSV は使いません。** Vimeo 工程では **`--out-csv` で出力する専用 CSV** を `--resume-from` に渡して再開します。

```powershell
cd D:\code4biz\ZOOM
```

### 5.1 事前確認（アップロードは実行しない）

```powershell
python vimeo_upload.py --csv consultant_vimeo.csv --out-csv vimeo_results.csv --dry-run
```

### 5.2 本番アップロード

- **`--root`** を省略すると `.env` の **`DOWNLOAD_DIR`** から MP4 を再帰検索します（通常は `downloads` で問題ありません）。

```powershell
python vimeo_upload.py --csv consultant_vimeo.csv --out-csv vimeo_results.csv
```

明示的にルートを指定する場合：

```powershell
python vimeo_upload.py --root .\downloads --csv consultant_vimeo.csv --out-csv vimeo_results.csv
```

### 5.3 途中から再開（Vimeo 専用 CSV）

```powershell
python vimeo_upload.py --csv consultant_vimeo.csv --out-csv vimeo_results.csv --resume-from .\vimeo_results.csv
```

### 5.4 補足

- 突合は **CSV「配信用動画タイトル」** と **ローカル MP4 のファイル名（拡張子除く）** を同一ルールでキー化して行います（Zoom 保存時と整合）。
- **既に Vimeo 上に同じタイトルキーがある動画はスキップ**されます（`skipped(already_on_vimeo)`）。
- **前回の結果 CSV で `status=uploaded` の行もスキップ**されます（`skipped(already_uploaded)`）。
- 一覧 API を叩きたくないとき：`--no-vimeo-check`
- アップロード時に CSV の「パスコード」列が Vimeo のパスワード制限に設定されます。

---

## 6. （任意）Vimeo リンク付きメタデータ CSV の出力

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

## 7. Vercel ギャラリーの自動デプロイ

`vimeo_metadata_export.csv` を **`main` ブランチへ push** すると、GitHub Actions が自動的に Vercel の Deploy Hook を呼び出し、ギャラリーを再デプロイします。

### 7.1 設定方法

1. Vercel でデプロイ対象プロジェクトの **Deploy Hook URL** を取得する  
   （Vercel ダッシュボード → Project Settings → Git → Deploy Hooks）
2. GitHub リポジトリの **Settings → Secrets and variables → Actions** に  
   `VERCEL_DEPLOY_HOOK_URL` という名前でシークレットを登録する

### 7.2 トリガー条件

`.github/workflows/trigger-vercel-gallery-deploy.yml` で定義されており、  
`main` ブランチへの push で **`vimeo_metadata_export.csv` が変更されたときのみ**実行されます。

---

## 8. 処理の流れ（一覧）

1. `zoom_download\.env` を整備（Zoom + `VIMEO_TOKEN` + `ANTHROPIC_API_KEY`）
2. `python parse_lecture_email.py ...` で招聘メールを解析し `consultant_vimeo.csv` に追記
3. `python main.py ...` で Zoom から MP4 を `downloads`（または `DOWNLOAD_DIR`）へ取得
4. `python vimeo_upload.py ... --out-csv vimeo_results.csv` で Vimeo にアップロード
5. 必要なら `python export_vimeo_metadata_csv.py` でリンク付き一覧 CSV を出力
6. `vimeo_metadata_export.csv` を `git push` すると Vercel ギャラリーが自動更新される

---

## 9. トラブルのヒント

| 現象 | 対処の例 |
|------|-----------|
| Zoom トークン 400 | `.env` の Account ID / Client ID / Secret、アプリ Activate、EU 向け URL（`ZOOM_OAUTH_URL` 変数） |
| Vimeo 書き込み Permission denied | 出力・入力 CSV を Excel で閉じる、`--out-csv` で別ファイル |
| アップロードで `missing` が多い | `DOWNLOAD_DIR` と `--root` が実際の MP4 場所と一致しているか、ファイル名が CSV の「配信用動画タイトル」と一致しているか |
| `parse_lecture_email.py` が「Outlook に接続できません」 | Outlook を起動してから再実行、または `--file` でテキストファイルを指定 |
| `parse_lecture_email.py` が「ANTHROPIC_API_KEY が設定されていません」 | `.env`（プロジェクト直下 or `zoom_download` 配下）に `ANTHROPIC_API_KEY=...` を追記 |
| `pywin32` が見つからない | `pip install pywin32` を実行（Outlook 連携時のみ必要） |
| Vercel が自動デプロイされない | GitHub リポジトリに `VERCEL_DEPLOY_HOOK_URL` シークレットが設定されているか確認 |

以上です。
