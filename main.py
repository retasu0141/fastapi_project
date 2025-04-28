from fastapi import FastAPI, Request
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from gspread_formatting import *
from googleapiclient.discovery import build
import requests
import os
import json

app = FastAPI()

# 環境変数
WEBHOOK_URL = os.environ['SLACK_WEBHOOK_URL']
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents"
]
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS_JSON'])
credentials = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)

# Google APIクライアント
gc = gspread.authorize(credentials)
docs_service = build('docs', 'v1', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)

last_spreadsheet_ids = {}
last_document_ids = {}

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

def create_new_spreadsheet(title, topic_name):
    sh = gc.create(title)
    worksheet = sh.sheet1
    sh.share('nattsuchanneru@gmail.com', perm_type='user', role='writer')
    last_spreadsheet_ids[topic_name] = sh.id
    return sh, worksheet

def create_new_document(title, topic_name):
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc['documentId']
    drive_service.permissions().create(
        fileId=doc_id,
        body={"role": "writer", "type": "user", "emailAddress": "nattsuchanneru@gmail.com"},
        fields="id"
    ).execute()
    last_document_ids[topic_name] = doc_id
    return doc_id

def write_to_document(doc_id, row, topic):
    requests_list = []

    for key, value in row.items():
        clean_key = key.replace(':', '')
        requests_list.append({
            "insertText": {
                "location": {"index": 1},
                "text": f"\n{value}\n"
            }
        })
        requests_list.append({
            "updateParagraphStyle": {
                "range": {"startIndex": 1, "endIndex": 1+len(value)},
                "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                "fields": "namedStyleType"
            }
        })
        requests_list.append({
            "insertText": {
                "location": {"index": 1},
                "text": f"{clean_key}\n"
            }
        })
        requests_list.append({
            "updateParagraphStyle": {
                "range": {"startIndex": 1, "endIndex": 1+len(clean_key)},
                "paragraphStyle": {"namedStyleType": "HEADING_2"},
                "fields": "namedStyleType"
            }
        })

    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": list(reversed(requests_list))}).execute()

def send_slack_notification(message):
    payload = {"text": message}
    headers = {"Content-Type": "application/json"}
    requests.post(WEBHOOK_URL, json=payload, headers=headers)

def get_or_create_spreadsheet(topic_name, force_new=False):
    if force_new or topic_name not in last_spreadsheet_ids:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = f"{topic_name}_{now}"
        return create_new_spreadsheet(title, topic_name)

    try:
        sh = gc.open_by_key(last_spreadsheet_ids[topic_name])
        worksheet = sh.sheet1
        return sh, worksheet
    except gspread.exceptions.SpreadsheetNotFound:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = f"{topic_name}_{now}"
        return create_new_spreadsheet(title, topic_name)

# ========================== ルート ==========================

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

        flat_row = []
        for v in row.values():
            if isinstance(v, list):
                flat_row.append(", ".join(map(str, v)))
            else:
                flat_row.append(str(v))

        worksheet.append_row(flat_row)
        auto_resize_columns(worksheet)
        spreadsheet_url = sh.url

        # ドキュメント
        if force_new or topic not in last_document_ids:
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            doc_title = f"{topic}議事録_{now}"
            doc_id = create_new_document(doc_title, topic)
        else:
            doc_id = last_document_ids[topic]

        write_to_document(doc_id, row, topic)
        document_url = f"https://docs.google.com/document/d/{doc_id}/edit"

        send_slack_notification(f"✅スプレッドシート: {spreadsheet_url}\n📝ドキュメント: {document_url}")

    return {"status": "success", "received": data}
