from fastapi import FastAPI, Request
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from gspread_formatting import *
import requests
import os
import json

app = FastAPI()

# SlackのWebhook URL
WEBHOOK_URL = "https://hooks.slack.com/services/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"

# Google認証設定
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS_JSON'])
credentials = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)

gc = gspread.authorize(credentials)



# 話題ごとの最後に作成されたスプレッドシートIDを保存する辞書
last_spreadsheet_ids = {}

# ========================== ここから関数 ==========================

def create_new_spreadsheet(title, topic_name):
    """新しいスプレッドシートを作成"""
    sh = gc.create(title)
    worksheet = sh.sheet1
    # ✅ 自分に共有
    sh.share('nattsuchanneru@gmail.com', perm_type='user', role='writer')
    # ✅ Slackに通知送信
    spreadsheet_url = sh.url
    slack_message = f"✅ 新しいスプレッドシートが作成されました！\n{spreadsheet_url}"
    send_slack_notification(slack_message, WEBHOOK_URL)

    # ✅ 作ったスプレッドシートIDを記録
    last_spreadsheet_ids[topic_name] = sh.id

    return sh, worksheet

def get_or_create_spreadsheet(topic_name, force_new=False):
    """スプレッドシートを取得 or 新規作成"""
    if force_new:
        # 強制新規作成なら新しいシートを作る
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = f"{topic_name}_{now}"
        return create_new_spreadsheet(title, topic_name)

    # 直近作成スプレッドシートIDがあるかチェック
    if topic_name in last_spreadsheet_ids:
        try:
            sh = gc.open_by_key(last_spreadsheet_ids[topic_name])
            worksheet = sh.sheet1
            return sh, worksheet
        except gspread.exceptions.SpreadsheetNotFound:
            # もし無かったら普通に探す
            pass

    # 話題名で検索して開く（バックアップパターン）
    try:
        sh = gc.open(topic_name)
        worksheet = sh.sheet1
        return sh, worksheet
    except gspread.exceptions.SpreadsheetNotFound:
        # 無かったら新規作成
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = f"{topic_name}_{now}"
        return create_new_spreadsheet(title, topic_name)

def send_slack_notification(message, webhook_url):
    payload = {"text": message}
    headers = {"Content-Type": "application/json"}
    response = requests.post(webhook_url, json=payload, headers=headers)
    return response.status_code

# ========================== ここまで関数 ==========================

@app.post("/trigger")
async def receive_data(request: Request):
    data = await request.json()

    for row in data:
        force_new = row.get("新規作成", False)
        topic = row.get("話題", "未分類")

        # シート取得
        sh, worksheet = get_or_create_spreadsheet(topic, force_new)

        # ヘッダー行チェック
        if worksheet.row_count == 0 or worksheet.acell('A1').value is None:
            headers = ["話題", "内容", "得た情報", "メモ", "参考URL", "アクション"]
            worksheet.append_row(headers)

            # ヘッダーに書式設定
            header_format = cellFormat(
                backgroundColor=color(0.9, 0.9, 0.9),
                textFormat=textFormat(bold=True),
                horizontalAlignment='CENTER'
            )
            format_cell_range(worksheet, f'A1:{chr(65+len(headers)-1)}1', header_format)

        # データを並び順に合わせて整える
        row_data = [
            row.get("話題", ""),
            row.get("内容", ""),
            row.get("得た情報", ""),
            row.get("メモ", ""),
            row.get("参考URL", ""),
            row.get("アクション", ""),
        ]
        worksheet.append_row(row_data)

        # --- ここで文字数から列幅を自動設定する ---
        def calculate_column_width(text):
            width = 0
            for ch in text:
                if ord(ch) < 128:
                    width += 1
                else:
                    width += 1.5
            return min(max(100, int(width * 7)), 400)  # ちょっとゆとり持たせる

        headers = ["話題", "内容", "得た情報", "メモ", "参考URL", "アクション"]
        for i, header in enumerate(headers):
            column_letter = chr(65 + i)  # 'A', 'B', 'C', ...
            # ヘッダーと最新のデータを比較して大きい方に
            text = header + str(row.get(header, ""))
            width = calculate_column_width(text)
            set_column_width(worksheet, column_letter, width)
        # -------------------------------------------

        spreadsheet_url = sh.url

    slack_message = f"📝 スプレッドシートにデータを追記しました！\n{spreadsheet_url}"
    send_slack_notification(slack_message, WEBHOOK_URL)

    print("受け取ったデータ:", data)
    return {"status": "success", "received": data}
