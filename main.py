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
credentials = Credentials.from_service_account_info(credentials_info)

gc = gspread.authorize(credentials)

# 話題ごとの最後に作成されたスプレッドシートIDを保存する辞書
last_spreadsheet_ids = {}

# ========================== ここから関数 ==========================

def create_new_spreadsheet(title, topic_name):
    """新しいスプレッドシートを作成"""
    sh = gc.create(title)
    worksheet = sh.sheet1
    sh.share('nattsuchanneru@gmail.com', perm_type='user', role='writer')

    # カラム幅を均等に広げる（初期化）
    default_column_width = 200
    set_column_width(worksheet, 'A', default_column_width)
    set_column_width(worksheet, 'B', default_column_width)
    set_column_width(worksheet, 'C', default_column_width)
    set_column_width(worksheet, 'D', default_column_width)
    set_column_width(worksheet, 'E', default_column_width)
    set_column_width(worksheet, 'F', default_column_width)

    # Slack通知
    spreadsheet_url = sh.url
    slack_message = f"✅ 新しいスプレッドシートが作成されました！\n{spreadsheet_url}"
    send_slack_notification(slack_message, WEBHOOK_URL)

    last_spreadsheet_ids[topic_name] = sh.id

    return sh, worksheet

def get_or_create_spreadsheet(topic_name, force_new=False):
    """スプレッドシートを取得 or 新規作成"""
    if force_new:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = f"{topic_name}_{now}"
        return create_new_spreadsheet(title, topic_name)

    if topic_name in last_spreadsheet_ids:
        try:
            sh = gc.open_by_key(last_spreadsheet_ids[topic_name])
            worksheet = sh.sheet1
            return sh, worksheet
        except gspread.exceptions.SpreadsheetNotFound:
            pass

    try:
        sh = gc.open(topic_name)
        worksheet = sh.sheet1
        return sh, worksheet
    except gspread.exceptions.SpreadsheetNotFound:
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

    if isinstance(data, dict):
        data = [data]  # 単一オブジェクトならリストに包む

    for row in data:
        if isinstance(row, str):
            row = json.loads(row)  # 文字列ならパース

        force_new = row.get("新規作成", False)
        topic = row.get("話題", "未分類")

        sh, worksheet = get_or_create_spreadsheet(topic, force_new)

        headers = ["話題", "内容", "得た情報", "メモ", "参考URL", "アクション"]

        # ヘッダー行チェック
        if worksheet.row_count == 0 or worksheet.acell('A1').value is None:
            worksheet.append_row(headers)
            header_format = cellFormat(
                backgroundColor=color(0.9, 0.9, 0.9),
                textFormat=textFormat(bold=True),
                horizontalAlignment='CENTER'
            )
            format_cell_range(worksheet, f'A1:{chr(65+len(headers)-1)}1', header_format)

        row_data = [
            row.get("話題", ""),
            row.get("内容", ""),
            row.get("得た情報", ""),
            row.get("メモ", ""),
            row.get("参考URL", ""),
            row.get("アクション", "")
        ]
        worksheet.append_row(row_data)

    spreadsheet_url = sh.url
    slack_message = f"📝 スプレッドシートにデータを追記しました！\n{spreadsheet_url}"
    send_slack_notification(slack_message, WEBHOOK_URL)

    print("受け取ったデータ:", data)
    return {"status": "success", "received": data}
