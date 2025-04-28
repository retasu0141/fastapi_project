from fastapi import FastAPI, Request
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from gspread_formatting import *
import requests
import os
import json
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

app = FastAPI()

# 環境変数から取得
WEBHOOK_URL = os.environ['SLACK_WEBHOOK_URL']
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS_JSON'])

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents"
]
credentials = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
gc = gspread.authorize(credentials)
docs_service = build('docs', 'v1', credentials=credentials)

google_account_email = credentials_info.get('client_email')

# スプレッドシートID保存用
document_tracker = {}

# ========================== 関数 ==========================

def auto_resize_columns(worksheet):
    all_values = worksheet.get_all_values()
    if not all_values:
        return
    columns = list(zip(*all_values))
    for i, col in enumerate(columns):
        max_length = max(len(str(cell)) for cell in col)
        width = max(100, min(max_length * 10, 400))
        set_column_width(worksheet, chr(65+i), width)

def send_slack_notification(message, webhook_url):
    payload = {"text": message}
    headers = {"Content-Type": "application/json"}
    response = requests.post(webhook_url, json=payload, headers=headers)
    return response.status_code

def create_spreadsheet(title):
    sh = gc.create(title)
    worksheet = sh.sheet1
    sh.share(google_account_email, perm_type='user', role='writer')
    return sh, worksheet

def create_document(title):
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc.get('documentId')
    permission = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': google_account_email
    }
    drive_service = build('drive', 'v3', credentials=credentials)
    drive_service.permissions().create(fileId=doc_id, body=permission, sendNotificationEmail=False).execute()
    return doc_id

def append_to_document(doc_id, topic, rows):
    requests_batch = []
    # 最初にタイトルとしてtopicを追加
    requests_batch.append({
        'insertText': {
            'location': {'index': 1},
            'text': f"{topic}\n"
        }
    })
    requests_batch.append({
        'updateParagraphStyle': {
            'range': {'startIndex': 1, 'endIndex': 1 + len(topic)},
            'paragraphStyle': {'namedStyleType': 'TITLE'},
            'fields': 'namedStyleType'
        }
    })

    for row in rows:
        for key, value in row.items():
            if key in ["新規作成"]:
                continue
            key_text = key.replace(':', '')
            start_idx = None
            text = f"\n{key_text}\n{value}\n"
            requests_batch.append({
                'insertText': {
                    'location': {'index': 1},
                    'text': text
                }
            })
            # 見出しに設定
            requests_batch.append({
                'updateParagraphStyle': {
                    'range': {'startIndex': 1, 'endIndex': 1 + len(key_text)},
                    'paragraphStyle': {'namedStyleType': 'HEADING_1'},
                    'fields': 'namedStyleType'
                }
            })

    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests_batch}).execute()

# ========================== APIエンドポイント ==========================

@app.post("/trigger")
async def receive_data(request: Request):
    raw_data = await request.json()

    if isinstance(raw_data, dict):
        data = [raw_data]
    elif isinstance(raw_data, list):
        if all(isinstance(row, str) for row in raw_data):
            data = [json.loads(row) for row in raw_data]
        else:
            data = raw_data
    else:
        raise ValueError("Unexpected data format")

    for row in data:
        force_new = row.get("新規作成", False)
        topic = row.get("話題", "未分類")

        if force_new or topic not in document_tracker:
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            sheet_title = f"{topic}_{now}"
            doc_title = f"{topic}議事録_{now}"
            sh, worksheet = create_spreadsheet(sheet_title)
            doc_id = create_document(doc_title)
            document_tracker[topic] = {'spreadsheet_id': sh.id, 'doc_id': doc_id}
        else:
            sh = gc.open_by_key(document_tracker[topic]['spreadsheet_id'])
            worksheet = sh.sheet1
            doc_id = document_tracker[topic]['doc_id']

        # スプレッドシートへの書き込み
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

        # ドキュメントへの書き込み
        append_to_document(doc_id, topic, [row])

        spreadsheet_url = sh.url
        doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"

    slack_message = f"✅ 新しいデータが記録されました！\nスプレッドシート: {spreadsheet_url}\nドキュメント: {doc_url}"
    send_slack_notification(slack_message, WEBHOOK_URL)

    return {"status": "success", "received": data}
