from fastapi import FastAPI, Request
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime
from gspread_formatting import *
import requests
import os
import json

app = FastAPI()

# 環境変数から設定を取得
WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL")
credentials_info = json.loads(os.getenv("GOOGLE_CREDENTIALS_JSON"))
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents"
]
credentials = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
gc = gspread.authorize(credentials)
docs_service = build('docs', 'v1', credentials=credentials)

# スプレッドシート管理用
last_spreadsheet_ids = {}

# ========= 共通関数 =========

def send_slack_notification(message):
    payload = {"text": message}
    headers = {"Content-Type": "application/json"}
    requests.post(WEBHOOK_URL, json=payload, headers=headers)

def auto_resize_columns(worksheet):
    all_values = worksheet.get_all_values()
    if not all_values:
        return
    columns = list(zip(*all_values))
    for i, col in enumerate(columns):
        max_length = max(len(str(cell)) for cell in col)
        width = max(100, min(max_length * 10, 400))
        set_column_width(worksheet, chr(65+i), width)

# ========= スプレッドシート関連 =========

def create_spreadsheet(title, headers):
    sh = gc.create(title)
    worksheet = sh.sheet1
    sh.share('nattsuchanneru@gmail.com', perm_type='user', role='writer')
    worksheet.append_row(headers)
    auto_resize_columns(worksheet)
    return sh, worksheet

def write_to_spreadsheet(headers, rows, topic):
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    title = f"{topic}_{now}"
    sh, worksheet = create_spreadsheet(title, headers)
    for row in rows:
        worksheet.append_row(row)
    auto_resize_columns(worksheet)
    send_slack_notification(f"📄 新しいスプレッドシートを作成しました！\n{sh.url}")

# ========= ドキュメント関連 =========

def create_document(title):
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc.get("documentId")
    return doc_id

def write_to_document(doc_id, contents):
    requests_list = []
    for block in contents:
        if block['type'] == 'heading':
            requests_list.append({
                "insertText": {
                    "location": {"index": 1},
                    "text": block['text'] + "\n"
                }
            })
            requests_list.append({
                "updateParagraphStyle": {
                    "range": {"startIndex": 1, "endIndex": 1 + len(block['text']) + 1},
                    "paragraphStyle": {"namedStyleType": "HEADING_1"},
                    "fields": "namedStyleType"
                }
            })
        elif block['type'] == 'paragraph':
            requests_list.append({
                "insertText": {
                    "location": {"index": 1},
                    "text": block['text'] + "\n"
                }
            })
    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": list(reversed(requests_list))}).execute()

# ========= エンドポイント =========

@app.post("/trigger")
async def trigger(request: Request):
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
        if row.get("type") == "spreadsheet":
            headers = row.get("headers")
            rows = row.get("rows")
            topic = row.get("topic", "未分類")
            if headers and rows:
                write_to_spreadsheet(headers, rows, topic)

        elif row.get("type") == "document":
            topic = row.get("topic", "未分類")
            contents = row.get("contents", [])
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            title = f"{topic}_議事録_{now}"
            doc_id = create_document(title)
            write_to_document(doc_id, contents)
            send_slack_notification(f"📝 新しいドキュメントを作成しました！\nhttps://docs.google.com/document/d/{doc_id}/edit")

    return {"status": "success", "received": data}
