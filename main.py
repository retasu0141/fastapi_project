from fastapi import FastAPI, Request
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *
import requests
import os
import json
from datetime import datetime
from googleapiclient.discovery import build

app = FastAPI()

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

last_spreadsheet_ids = {}

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
    spreadsheet_url = sh.url
    #slack_message = f"✅ 新しいスプレッドシートが作成されました！\n{spreadsheet_url}"
    #send_slack_notification(slack_message)
    last_spreadsheet_ids[topic_name] = sh.id
    return sh, worksheet

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

def send_slack_notification(message):
    payload = {"text": message}
    headers = {"Content-Type": "application/json"}
    requests.post(WEBHOOK_URL, json=payload, headers=headers)

def create_google_document(title, contents_dict):
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc['documentId']

    # 作成直後に共有設定を付与
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
    requests_body.append({
        'insertText': {
            'location': {'index': 1},
            'text': title + "\n"
        }
    })
    requests_body.append({
        'updateParagraphStyle': {
            'range': {'startIndex': 1, 'endIndex': 1 + len(title)},
            'paragraphStyle': {'namedStyleType': 'TITLE'},
            'fields': 'namedStyleType'
        }
    })

    cursor = 1 + len(title) + 1

    for header, body in contents_dict.items():
        if not body:
            continue
        requests_body.append({
            'insertText': {
                'location': {'index': cursor},
                'text': header + "\n"
            }
        })
        requests_body.append({
            'updateParagraphStyle': {
                'range': {'startIndex': cursor, 'endIndex': cursor + len(header)},
                'paragraphStyle': {'namedStyleType': 'HEADING_2'},
                'fields': 'namedStyleType'
            }
        })
        cursor += len(header) + 1

        requests_body.append({
            'insertText': {
                'location': {'index': cursor},
                'text': body + "\n\n"
            }
        })
        cursor += len(body) + 2

    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests_body}).execute()
    return f"https://docs.google.com/document/d/{doc_id}/edit"

# ========================== エンドポイント ==========================

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

        worksheet.append_row(list(row.values()))
        auto_resize_columns(worksheet)

        spreadsheet_url = sh.url

        contents_dict = {k.rstrip(':'): v for k, v in row.items() if v}
        document_url = create_google_document(topic, contents_dict)

        slack_message = f"✅ 新しいスプレッドシートが作成されました！\n{spreadsheet_url}\n📄 新しいGoogleドキュメントが作成されました！\n{document_url}"
        send_slack_notification(slack_message)

    return {"status": "success", "received": data}
