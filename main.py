from fastapi import FastAPI, Request
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from gspread_formatting import *
import requests
import os
import json
from datetime import datetime

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

gc = gspread.authorize(credentials)

drive_service = build('drive', 'v3', credentials=credentials)
docs_service = build('docs', 'v1', credentials=credentials)

last_spreadsheet_ids = {}
last_document_ids = {}

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
    last_spreadsheet_ids[topic_name] = sh.id
    return sh, worksheet, spreadsheet_url

def create_new_document(title, topic_name):
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc['documentId']
    drive_service.permissions().create(
        fileId=doc_id,
        body={"type": "user", "role": "writer", "emailAddress": 'nattsuchanneru@gmail.com'}
    ).execute()
    document_url = f"https://docs.google.com/document/d/{doc_id}/edit"
    last_document_ids[topic_name] = doc_id
    return doc_id, document_url

def send_slack_notification(message, webhook_url):
    payload = {"text": message}
    headers = {"Content-Type": "application/json"}
    response = requests.post(webhook_url, json=payload, headers=headers)
    return response.status_code

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
        force_new = row.get("新規作成", False)
        topic = row.get("話題", "未分類")

        # スプレッドシート処理
        if force_new or topic not in last_spreadsheet_ids:
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            title = f"{topic}_{now}"
            sh, worksheet, spreadsheet_url = create_new_spreadsheet(title, topic)
            # ヘッダー追加
            worksheet.append_row([str(cell) for cell in row.keys()])
        else:
            sh = gc.open_by_key(last_spreadsheet_ids[topic])
            worksheet = sh.sheet1
            spreadsheet_url = sh.url

        worksheet.append_row([str(cell) for cell in row.values()])
        auto_resize_columns(worksheet)

        # ドキュメント処理
        if force_new or topic not in last_document_ids:
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            title = f"{topic}議事録_{now}"
            doc_id, document_url = create_new_document(title, topic)
        else:
            doc_id = last_document_ids[topic]
            document_url = f"https://docs.google.com/document/d/{doc_id}/edit"

        write_to_document(doc_id, row, topic)

        # Slack通知まとめて
        slack_message = f"✅ スプレッドシート記録: {spreadsheet_url}\n📄 ドキュメント記録: {document_url}"
        send_slack_notification(slack_message, WEBHOOK_URL)

    return {"status": "success"}

def write_to_document(doc_id, row, topic):
    requests_list = []

    # トピックを最初に挿入
    requests_list.append({
        "insertText": {
            "location": {"index": 1},
            "text": f"{topic}\n"
        }
    })
    requests_list.append({
        "updateParagraphStyle": {
            "range": {"startIndex": 1, "endIndex": len(topic)+1},
            "paragraphStyle": {"namedStyleType": "TITLE"},
            "fields": "namedStyleType"
        }
    })

    # 中身を見出し・本文で追加
    cursor = len(topic) + 2
    for key, value in row.items():
        if key == "話題" or key == "新規作成":
            continue
        # 見出し
        requests_list.append({
            "insertText": {
                "location": {"index": cursor},
                "text": f"{key}\n"
            }
        })
        requests_list.append({
            "updateParagraphStyle": {
                "range": {"startIndex": cursor, "endIndex": cursor+len(key)},
                "paragraphStyle": {"namedStyleType": "HEADING_2"},
                "fields": "namedStyleType"
            }
        })
        cursor += len(key) + 1

        # 本文
        value_text = str(value)
        requests_list.append({
            "insertText": {
                "location": {"index": cursor},
                "text": f"{value_text}\n"
            }
        })
        cursor += len(value_text) + 1

    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests_list}).execute()
