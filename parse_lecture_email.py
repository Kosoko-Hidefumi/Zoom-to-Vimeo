#!/usr/bin/env python3
"""
parse_lecture_email.py
コンサルタント招聘メールを解析して consultant_vimeo.csv に追記する。

使用例:
  python parse_lecture_email.py                          # Outlook受信トレイ（過去7日）
  python parse_lecture_email.py --days 14               # 過去14日間
  python parse_lecture_email.py --file email_sample.txt # テキストファイルから
  python parse_lecture_email.py --csv path/to/other.csv # CSVパス指定
  python parse_lecture_email.py --force                 # 重複行も追記
  python parse_lecture_email.py --update --file x.txt   # タイトル更新モード
"""

import argparse
import csv
import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent

# .env探索順: スクリプトと同階層 → zoom_download サブディレクトリ
for _env in [SCRIPT_DIR / ".env", SCRIPT_DIR / "zoom_download" / ".env"]:
    if _env.exists():
        load_dotenv(_env)
        break

DEFAULT_CSV = SCRIPT_DIR / "consultant_vimeo.csv"

CSV_COLUMNS = [
    "区分", "講師No.", "講師名（日本語）", "講師名（英語）",
    "専門科", "所属", "日付", "時間帯", "開始時刻", "終了時刻",
    "場所", "レクチャータイトル（日本語）", "レクチャータイトル（英語）",
    "配信用動画タイトル", "Zoom URL", "ミーティングID", "パスコード", "備考",
]

SLOT_LABELS = {
    "早朝": "早朝レクチャー",
    "コア": "コアレクチャー",
    "午後": "午後レクチャー",
    "夕方": "夕方レクチャー",
}


# ---------------------------------------------------------------------------
# 引数解析
# ---------------------------------------------------------------------------

def parse_args():
    p = argparse.ArgumentParser(description="コンサルタント招聘メールを解析してCSVに追記")
    p.add_argument("--days", type=int, default=7, help="Outlookから過去N日間を検索（デフォルト: 7）")
    p.add_argument("--file", help="テキストファイルからメール本文を読み込む（テスト用）")
    p.add_argument("--csv", default=str(DEFAULT_CSV), help="追記対象のCSVパス")
    p.add_argument("--force", action="store_true", help="重複行も強制追記")
    p.add_argument("--update", action="store_true", help="【タイトル未確定】行のタイトルを更新する")
    return p.parse_args()


# ---------------------------------------------------------------------------
# メール取得
# ---------------------------------------------------------------------------

def load_email_from_file(path: str) -> str:
    with open(path, encoding="utf-8", errors="replace") as f:
        return f.read()


def get_emails_from_outlook(days: int) -> list:
    try:
        import win32com.client
    except ImportError:
        print("エラー: pywin32 がインストールされていません。")
        print("  pip install pywin32")
        sys.exit(1)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"エラー: Outlook に接続できません。Outlook が起動しているか確認してください。\n{e}")
        sys.exit(1)

    inbox = ns.GetDefaultFolder(6)  # 6 = 受信トレイ
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    cutoff = datetime.now() - timedelta(days=days)
    keywords = ["コンサルタント", "レクチャー", "Lecture"]
    found = []

    for msg in items:
        try:
            received = msg.ReceivedTime
            # タイムゾーンを除去して比較
            if hasattr(received, "replace"):
                received_naive = received.replace(tzinfo=None)
            else:
                received_naive = received
            if received_naive < cutoff:
                break
            subject = msg.Subject or ""
            if any(kw in subject for kw in keywords):
                found.append({
                    "subject": subject,
                    "received": received_naive,
                    "body": msg.Body,
                })
        except Exception:
            continue

    return found


def select_email(emails: list) -> str:
    if len(emails) == 1:
        print(f"メールを1件取得: {emails[0]['subject']}")
        return emails[0]["body"]

    print(f"\n{len(emails)} 件のメールが見つかりました：")
    for i, m in enumerate(emails, 1):
        print(f"  [{i}] {m['received'].strftime('%Y/%m/%d')}  {m['subject']}")

    while True:
        choice = input("\n番号を選択してください（Enterでキャンセル）: ").strip()
        if choice == "":
            print("キャンセルしました。")
            sys.exit(0)
        if choice.isdigit() and 1 <= int(choice) <= len(emails):
            return emails[int(choice) - 1]["body"]
        print("無効な番号です。再入力してください。")


# ---------------------------------------------------------------------------
# Claude API 解析
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = "あなたは医療研修のコンサルタント招聘メールを解析するアシスタントです。メール本文から講義スケジュールを読み取り、指定されたJSON形式のみを出力してください。説明文は不要です。"

USER_PROMPT_TEMPLATE = """以下のメール本文を解析し、各セッションの情報をJSON配列で出力してください。

## 抽出ルール

各セッション1オブジェクトとして以下フィールドを抽出：

| フィールド | ルール |
|-----------|--------|
| 講師名_日本語 | 日本人のみ。外国人は "" |
| 講師名_英語 | 例: "Dr. Neal A. Palafox" |
| 専門科 | 例: "総合診療科" |
| 所属 | そのまま |
| 日付 | "YYYY/M/D" 形式 例: "2026/5/11" |
| 時間帯 | 下記ルール |
| 開始時刻 | "H:MM" 例: "7:30" |
| 終了時刻 | "H:MM" 例: "8:30" |
| 場所 | そのまま |
| タイトル_英語 | "〇〇" や ‟〇〇" で囲まれた部分のみ。"Lecture by Dr.〇〇" は空文字 |
| タイトル_日本語 | 日本語タイトルがあれば。なければタイトル_英語と同じ |
| zoom_url | 時間帯に対応するURLを割り当て。なければ "" |
| meeting_id | スペース区切りそのまま 例: "892 5023 5315"。なければ "" |
| passcode | 数字そのまま。なければ "" |
| タイトル未確定 | "Lecture by Dr.〇〇" など明示タイトルなしの場合 true |

## 時間帯の判定
- 7:00〜9:00開始 → "早朝"
- 13:00〜13:30開始 → "コア"
- 14:00〜16:00開始 → "午後"
- 17:00〜18:00開始 → "夕方"
- その他 → 時刻をそのまま文字列で

## Zoom URL割り当て
メール内の Zoom Information セクション：
- "Morning Lecture" / "早朝" → 早朝セッションのZoom URL
- "Core Lecture" / "コアレクチャー" / "コア" → コアセッションのZoom URL
- 午後/夕方も対応するURLを割り当て
- 対応するURLがない場合 → ""

## 出力形式（このJSON配列のみ出力）
```json
[
  {{
    "講師名_日本語": "",
    "講師名_英語": "Dr. Neal A. Palafox",
    "専門科": "総合診療科",
    "所属": "Professor, Department of ...",
    "日付": "2026/5/11",
    "時間帯": "早朝",
    "開始時刻": "7:30",
    "終了時刻": "8:30",
    "場所": "第1会議室",
    "タイトル_英語": "Atraumatic Subcutaneous Hemorrhage in a 40s Female",
    "タイトル_日本語": "Atraumatic Subcutaneous Hemorrhage in a 40s Female",
    "zoom_url": "https://...",
    "meeting_id": "892 5023 5315",
    "passcode": "579919",
    "タイトル未確定": false
  }}
]
```

## メール本文:
{email_body}
"""


def call_claude_api(email_body: str) -> list:
    try:
        import anthropic
    except ImportError:
        print("エラー: anthropic パッケージがインストールされていません。")
        print("  pip install anthropic")
        sys.exit(1)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("エラー: ANTHROPIC_API_KEY が設定されていません。.env ファイルを確認してください。")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)
    prompt = USER_PROMPT_TEMPLATE.format(email_body=email_body)
    last_error = None

    for attempt in range(2):
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4096,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = resp.content[0].text.strip()

            # コードブロックを除去
            if "```" in raw:
                parts = raw.split("```")
                for part in parts:
                    stripped = part.strip()
                    if stripped.startswith("json"):
                        stripped = stripped[4:].strip()
                    if stripped.startswith("["):
                        raw = stripped
                        break

            return json.loads(raw)

        except json.JSONDecodeError as e:
            last_error = f"JSON解析エラー: {e}"
            if attempt == 0:
                print(f"  {last_error}（リトライ中...）")
        except Exception as e:
            last_error = f"API呼び出しエラー: {e}"
            if attempt == 0:
                print(f"  {last_error}（リトライ中...）")

    print(f"\nAPI呼び出しに失敗しました: {last_error}")
    sys.exit(1)


# ---------------------------------------------------------------------------
# データ変換
# ---------------------------------------------------------------------------

def build_csv_rows(sessions: list) -> list:
    rows = []
    for s in sessions:
        specialist = s.get("専門科", "")
        name_en = s.get("講師名_英語", "")
        title_en = s.get("タイトル_英語", "")
        title_ja = s.get("タイトル_日本語", "") or title_en
        date_str = s.get("日付", "")
        time_slot = s.get("時間帯", "")
        start = s.get("開始時刻", "")
        end = s.get("終了時刻", "")
        is_undecided = bool(s.get("タイトル未確定", False))

        # 備考
        slot_label = SLOT_LABELS.get(time_slot, time_slot)
        biko = f"{slot_label}（{start}～{end}）"
        if is_undecided:
            biko += "、【タイトル未確定】"

        # 配信用動画タイトル
        if is_undecided or not title_en:
            video_title = ""
        else:
            try:
                dt = datetime.strptime(date_str, "%Y/%m/%d")
                date_fmt = dt.strftime("%Y.%m.%d")
            except ValueError:
                date_fmt = date_str.replace("/", ".")
            video_title = f"[{specialist}] {name_en} - {title_en} ({date_fmt})"

        rows.append({
            "区分": "短期",
            "講師No.": "",
            "講師名（日本語）": s.get("講師名_日本語", ""),
            "講師名（英語）": name_en,
            "専門科": specialist,
            "所属": s.get("所属", ""),
            "日付": date_str,
            "時間帯": time_slot,
            "開始時刻": start,
            "終了時刻": end,
            "場所": s.get("場所", ""),
            "レクチャータイトル（日本語）": title_ja,
            "レクチャータイトル（英語）": title_en,
            "配信用動画タイトル": video_title,
            "Zoom URL": s.get("zoom_url", ""),
            "ミーティングID": s.get("meeting_id", ""),
            "パスコード": s.get("passcode", ""),
            "備考": biko,
        })
    return rows


# ---------------------------------------------------------------------------
# CSV操作
# ---------------------------------------------------------------------------

def load_existing_keys(csv_path: str) -> set:
    keys = set()
    if not os.path.exists(csv_path):
        return keys
    with open(csv_path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            keys.add((
                row.get("日付", ""),
                row.get("開始時刻", ""),
                row.get("講師名（英語）", ""),
            ))
    return keys


def preview_rows(rows: list, csv_path: str):
    print("\n【追記予定の行】")
    print("-" * 72)
    print(f"  {'日付':<12} {'時間帯':<6} {'講師名（英語）':<26} タイトル")
    for r in rows:
        title = r["レクチャータイトル（英語）"] or "【タイトル未確定】"
        if len(title) > 28:
            title = title[:26] + "..."
        name = r["講師名（英語）"]
        if len(name) > 24:
            name = name[:22] + "..."
        print(f"  {r['日付']:<12} {r['時間帯']:<6} {name:<26} {title}")
    print("-" * 72)
    print(f"全 {len(rows)} 行を {csv_path} に追記します。続行しますか？ [y/N]: ", end="", flush=True)


def load_all_rows(csv_path: str) -> list:
    if not os.path.exists(csv_path):
        return []
    with open(csv_path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def preview_updates(updates: list):
    print("\n【更新予定の行】")
    print("-" * 72)
    print(f"  {'日付':<12} {'時間帯':<6} {'講師名（英語）':<22} 更新内容")
    for _, old_row, new_row in updates:
        name = old_row["講師名（英語）"]
        if len(name) > 20:
            name = name[:18] + "..."
        old_title = old_row["レクチャータイトル（英語）"] or "(未確定)"
        new_title = new_row["レクチャータイトル（英語）"]
        if len(old_title) > 18:
            old_title = old_title[:16] + "..."
        if len(new_title) > 18:
            new_title = new_title[:16] + "..."
        print(f"  {old_row['日付']:<12} {old_row['時間帯']:<6} {name:<22} {old_title} → {new_title}")
    print("-" * 72)
    print(f"全 {len(updates)} 行を更新します。続行しますか？ [y/N]: ", end="", flush=True)


def write_all_rows(rows: list, csv_path: str):
    try:
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(rows)
    except PermissionError:
        print(f"\nエラー: {csv_path} への書き込みに失敗しました。")
        print("Excel などでファイルを開いている場合は閉じてから再実行してください。")
        sys.exit(1)


def append_to_csv(rows: list, csv_path: str):
    file_exists = os.path.exists(csv_path)
    try:
        with open(csv_path, "a", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
            if not file_exists:
                writer.writeheader()
            writer.writerows(rows)
    except PermissionError:
        print(f"\nエラー: {csv_path} への書き込みに失敗しました。")
        print("Excel などでファイルを開いている場合は閉じてから再実行してください。")
        sys.exit(1)


# ---------------------------------------------------------------------------
# タイトル更新モード
# ---------------------------------------------------------------------------

def run_update_mode(args):
    # ① メール本文取得
    if args.file:
        print(f"ファイルからメール本文を読み込み中: {args.file}")
        email_body = load_email_from_file(args.file)
    else:
        print(f"Outlook の受信トレイから過去 {args.days} 日間を検索中...")
        emails = get_emails_from_outlook(args.days)
        if not emails:
            print(f"対象メールが見つかりませんでした（過去 {args.days} 日間）。")
            sys.exit(0)
        email_body = select_email(emails)

    # ② API解析
    print("\nClaude API でメールを解析中...")
    sessions = call_claude_api(email_body)
    if not sessions:
        print("解析結果が空でした。")
        sys.exit(1)
    print(f"{len(sessions)} 件のセッションを検出しました。")

    new_rows = build_csv_rows(sessions)

    # ③ 既存CSV全行読み込み
    existing_rows = load_all_rows(args.csv)
    if not existing_rows:
        print(f"エラー: {args.csv} が見つかりません。")
        sys.exit(1)

    # ④ 更新候補を抽出（キー一致 かつ 備考に【タイトル未確定】 かつ 新タイトルあり）
    # existing_rows のインデックスをキーにしたマップを作成
    key_to_idx = {}
    for i, row in enumerate(existing_rows):
        key = (row.get("日付", ""), row.get("開始時刻", ""), row.get("講師名（英語）", ""))
        key_to_idx[key] = i

    updates = []  # (idx, old_row, new_row)
    for new_row in new_rows:
        if not new_row["レクチャータイトル（英語）"]:
            continue  # 新タイトルがなければスキップ
        key = (new_row["日付"], new_row["開始時刻"], new_row["講師名（英語）"])
        idx = key_to_idx.get(key)
        if idx is None:
            continue  # 対応する既存行なし
        old_row = existing_rows[idx]
        if "【タイトル未確定】" not in old_row.get("備考", ""):
            continue  # タイトル未確定でない行はスキップ
        updates.append((idx, old_row, new_row))

    if not updates:
        print("\n更新対象の行がありません（【タイトル未確定】かつ新タイトルありの行が見つかりませんでした）。")
        sys.exit(0)

    # ⑤ プレビュー & 確認
    preview_updates(updates)
    answer = input().strip().lower()
    if answer != "y":
        print("キャンセルしました。")
        sys.exit(0)

    # ⑥ 対象列を上書きして全行書き直し
    UPDATE_COLS = [
        "レクチャータイトル（日本語）",
        "レクチャータイトル（英語）",
        "配信用動画タイトル",
        "備考",
    ]
    for idx, _, new_row in updates:
        for col in UPDATE_COLS:
            existing_rows[idx][col] = new_row[col]

    write_all_rows(existing_rows, args.csv)
    print(f"\n{len(updates)} 行を更新しました → {args.csv}")


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------

def main():
    args = parse_args()

    if args.update:
        run_update_mode(args)
        return

    # ① メール本文取得
    if args.file:
        print(f"ファイルからメール本文を読み込み中: {args.file}")
        email_body = load_email_from_file(args.file)
    else:
        print(f"Outlook の受信トレイから過去 {args.days} 日間を検索中...")
        emails = get_emails_from_outlook(args.days)
        if not emails:
            print(f"対象メールが見つかりませんでした（過去 {args.days} 日間）。")
            print("--file オプションでテキストファイルから読み込むことができます。")
            sys.exit(0)
        email_body = select_email(emails)

    # ② Claude API で解析
    print("\nClaude API でメールを解析中...")
    sessions = call_claude_api(email_body)
    if not sessions:
        print("解析結果が空でした。メール本文を確認してください。")
        sys.exit(1)
    print(f"{len(sessions)} 件のセッションを検出しました。")

    # ③ CSV行に変換
    rows = build_csv_rows(sessions)

    # ④ 重複チェック
    existing_keys = load_existing_keys(args.csv)
    new_rows = []
    skipped = []
    for row in rows:
        key = (row["日付"], row["開始時刻"], row["講師名（英語）"])
        if key in existing_keys and not args.force:
            skipped.append(row)
        else:
            new_rows.append(row)

    if skipped:
        print(f"\n{len(skipped)} 件は既に CSV に存在するためスキップします：")
        for r in skipped:
            print(f"  {r['日付']}  {r['開始時刻']}  {r['講師名（英語）']}")
        print("（--force オプションで強制追記できます）")

    if not new_rows:
        print("\n追記する行がありません。処理を終了します。")
        sys.exit(0)

    # ⑤ プレビュー & 確認
    preview_rows(new_rows, args.csv)
    answer = input().strip().lower()
    if answer != "y":
        print("キャンセルしました。")
        sys.exit(0)

    # ⑥ CSV追記
    append_to_csv(new_rows, args.csv)
    print(f"\n{len(new_rows)} 行を追記しました → {args.csv}")


if __name__ == "__main__":
    main()
