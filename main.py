from fastapi import FastAPI, Request
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from gspread_formatting import *
import requests
import os
import json
from googleapiclient.discovery import build

app = FastAPI()

# SlackのWebhook URL
WEBHOOK_URL = os.environ['SLACK_WEBHOOK_URL']

# Google認証設定
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents"
]
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS_JSON'])
credentials = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)

# gspreadの認証
gc = gspread.authorize(credentials)

# GoogleドキュメントのAPIクライアント
docs_service = build('docs', 'v1', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)

# 話題ごとの最後に作成されたスプレッドシートIDを保存する辞書
last_spreadsheet_ids = {}

# ========================== ここから関数 ==========================

def auto_resize_columns(worksheet):
    all_values = worksheet.get_all_values()
    if not all_values:
        return
    columns = list(zip(*all_values))
    for i, col in enumerate(columns):
        max_length = max(len(str(cell)) for cell in col)
        width = max(100, min(max_length * 10, 400))
        set_column_width(worksheet, chr(65+i), width)

def create_new_spreadsheet(title, topic_name):
    sh = gc.create(title)
    worksheet = sh.sheet1
    sh.share('nattsuchanneru@gmail.com', perm_type='user', role='writer')
    spreadsheet_url = sh.url
    slack_message = f"✅ 新しいスプレッドシートが作成されました！\n{spreadsheet_url}"
    send_slack_notification(slack_message, WEBHOOK_URL)
    last_spreadsheet_ids[topic_name] = sh.id
    return sh, worksheet

def create_new_document(title, content_dict):
    doc = docs_service.documents().create(body={"title": title}).execute()
    document_id = doc['documentId']

    # コンテンツの挿入
    requests_body = []

    # タイトルを見出しに
    requests_body.append({
        'insertText': {
            'location': {"index": 1},
            'text': title + "\n\n"
        }
    })
    requests_body.append({
        'updateParagraphStyle': {
            'range': {"startIndex": 1, "endIndex": len(title) + 2},
            'paragraphStyle': {"namedStyleType": "TITLE"},
            'fields': "namedStyleType"
        }
    })

    # 内容を本文に追加
    for key, value in content_dict.items():
        if value:
            text = f"{key}: {value}\n"
            requests_body.append({
                'insertText': {
                    'location': {"index": 1},
                    'text': text
                }
            })

    docs_service.documents().batchUpdate(documentId=document_id, body={"requests": requests_body}).execute()

    drive_service.permissions().create(
        fileId=document_id,
        body={
            'type': 'user',
            'role': 'writer',
            'emailAddress': 'nattsuchanneru@gmail.com'
        },
        fields='id',
        sendNotificationEmail=False
    ).execute()

    document_url = f"https://docs.google.com/document/d/{document_id}/edit"
    slack_message = f"\ud83d\udcc4 新しいGoogleドキュメントが作成されました！\n{document_url}"
    send_slack_notification(slack_message, WEBHOOK_URL)

    return document_url

def get_or_create_spreadsheet(topic_name, force_new=False):
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

    for row in data:
        force_new = row.get("新規作成", False)
        topic = row.get("話題", "未分類")

        sh, worksheet = get_or_create_spreadsheet(topic, force_new)

        if worksheet.row_count == 0 or worksheet.acell('A1').value is None:
            headers = list(row.keys())
            worksheet.append_row(headers)
            header_format = cellFormat(
                backgroundColor=color(0.9, 0.9, 0.9),
                textFormat=textFormat(bold=True),
                horizontalAlignment='CENTER'
            )
            format_cell_range(worksheet, f'A1:{chr(65+len(headers)-1)}1', header_format)

        worksheet.append_row(list(row.values()))
        auto_resize_columns(worksheet)


        create_new_document(topic, row)

        spreadsheet_url = sh.url

    slack_message = f"\ud83d\udcdc スプレッドシートにデータを追記しました！\n{spreadsheet_url}"
    send_slack_notification(slack_message, WEBHOOK_URL)

    print("受け取ったデータ:", data)
    return {"status": "success", "received": data}
