from fastapi import FastAPI, Request
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *
from googleapiclient.discovery import build
import requests
import os
import json
from datetime import datetime

app = FastAPI()

# 環境変数読み込み
WEBHOOK_URL = os.environ['SLACK_WEBHOOK_URL']
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents"
]
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS_JSON'])
credentials = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)

gc = gspread.authorize(credentials)
docs_service = build('docs', 'v1', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)

# ========== 関数定義 ==========

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

def create_spreadsheet(title, headers, rows):
    sh = gc.create(title)
    worksheet = sh.sheet1
    worksheet.append_row(headers)
    for row in rows:
        worksheet.append_row(row)
    auto_resize_columns(worksheet)
    sh.share('nattsuchanneru@gmail.com', perm_type='user', role='writer')
    return sh.url

def create_document(title, contents):
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc.get('documentId')

    # ドキュメント作成後、共有権限を付与
    drive_service.permissions().create(
        fileId=doc_id,
        body={
            'type': 'user',
            'role': 'writer',
            'emailAddress': 'nattsuchanneru@gmail.com'
        },
        fields='id'
    ).execute()

    requests_body = []
    # タイトル挿入
    requests_body.append({
        "insertText": {
            "location": {"index": 1},
            "text": title + "\n"
        }
    })
    requests_body.append({
        "updateParagraphStyle": {
            "range": {"startIndex": 1, "endIndex": len(title)+1},
            "paragraphStyle": {"namedStyleType": "TITLE"},
            "fields": "namedStyleType"
        }
    })
    idx = len(title) + 1
    # セクション挿入
    for section in contents:
        heading = section['heading']
        body = section['body']
        requests_body.append({
            "insertText": {
                "location": {"index": idx},
                "text": heading + "\n"
            }
        })
        requests_body.append({
            "updateParagraphStyle": {
                "range": {"startIndex": idx, "endIndex": idx+len(heading)},
                "paragraphStyle": {"namedStyleType": "HEADING_2"},
                "fields": "namedStyleType"
            }
        })
        idx += len(heading) + 1

        requests_body.append({
            "insertText": {
                "location": {"index": idx},
                "text": body + "\n\n"
            }
        })
        idx += len(body) + 2

    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests_body}).execute()
    return f"https://docs.google.com/document/d/{doc_id}/edit"

# ========== エンドポイント ==========

@app.post("/trigger")
async def receive_data(request: Request):
    raw_data = await request.json()

    if isinstance(raw_data, dict):
        data = [raw_data]
    elif isinstance(raw_data, list):
        if all(isinstance(item, str) for item in raw_data):
            data = [json.loads(item) for item in raw_data]
        else:
            data = raw_data
    else:
        raise ValueError("Unexpected data format")

    for block in data:
        content_type = block.get("type")
        topic = block.get("topic", "未分類")
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = f"{topic}_{now}"

        if content_type == "spreadsheet":
            headers = block.get("headers", [])
            rows = block.get("rows", [])
            url = create_spreadsheet(title, headers, rows)
            send_slack_notification(f"✅ 新しいスプレッドシートが作成されました！\n{url}")

        elif content_type == "document":
            contents = block.get("contents", [])
            url = create_document(title, contents)
            send_slack_notification(f"📄 新しいGoogleドキュメントが作成されました！\n{url}")

    return {"status": "success"}
