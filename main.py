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

# === 環境変数・認証設定 ===
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

# === 内部管理用 ===
last_spreadsheet_ids = {}
last_document_ids = {}

# === ユーティリティ関数 ===

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

# === スプレッドシート操作 ===

def create_new_spreadsheet(title, topic_name):
    sh = gc.create(title)
    worksheet = sh.sheet1
    sh.share('nattsuchanneru@gmail.com', perm_type='user', role='writer')
    last_spreadsheet_ids[topic_name] = sh.id
    return sh, worksheet

def get_or_create_spreadsheet(topic_name, force_new):
    if force_new or topic_name not in last_spreadsheet_ids:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = f"{topic_name}_{now}"
        return create_new_spreadsheet(title, topic_name)
    else:
        sh = gc.open_by_key(last_spreadsheet_ids[topic_name])
        return sh, sh.sheet1

# === ドキュメント操作 ===

def create_new_document(title, topic_name):
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc['documentId']
    last_document_ids[topic_name] = doc_id
    return doc_id

def get_or_create_document(topic_name, force_new):
    if force_new or topic_name not in last_document_ids:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = f"{topic_name}_議事録_{now}"
        return create_new_document(title, topic_name)
    else:
        return last_document_ids[topic_name]

# === メインエンドポイント ===

@app.post("/trigger")
async def trigger(request: Request):
    raw_data = await request.json()

    if isinstance(raw_data, dict):
        data = [raw_data]
    elif isinstance(raw_data, list):
        data = raw_data
    else:
        raise ValueError("Invalid data format")

    for entry in data:
        force_new = entry.get("新規作成", False)
        topic = entry.get("話題", "未分類")

        # === スプレッドシート処理 ===
        sh, worksheet = get_or_create_spreadsheet(topic, force_new)

        if worksheet.row_count == 0 or worksheet.acell('A1').value is None:
            worksheet.append_row(list(entry.keys()))

        worksheet.append_row(list(entry.values()))
        auto_resize_columns(worksheet)

        spreadsheet_url = sh.url

        # === ドキュメント処理 ===
        doc_id = get_or_create_document(topic, force_new)

        requests_list = []

        # 各フィールドを見出し＋本文で追加
        cursor = 1
        for key, value in entry.items():
            if key == "新規作成" or key == "話題":
                continue
            key_text = str(key)
            value_text = str(value)

            requests_list.append({
                "insertText": {
                    "location": {"index": cursor},
                    "text": key_text + "\n"
                }
            })
            requests_list.append({
                "updateParagraphStyle": {
                    "range": {"startIndex": cursor, "endIndex": cursor + len(key_text)},
                    "paragraphStyle": {"namedStyleType": "HEADING_2"},
                    "fields": "namedStyleType"
                }
            })
            cursor += len(key_text) + 1

            requests_list.append({
                "insertText": {
                    "location": {"index": cursor},
                    "text": value_text + "\n"
                }
            })
            cursor += len(value_text) + 1

        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": list(reversed(requests_list))}
        ).execute()

        doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"

        # === Slack通知 ===
        slack_message = f"✅ スプレッドシート記録: {spreadsheet_url}\n📄 ドキュメント記録: {doc_url}"
        send_slack_notification(slack_message)

    return {"status": "success", "message": "保存完了"}
