from fastapi import FastAPI, Request
import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build
import json
import os
import datetime
import requests

app = FastAPI()

# Google API認証
credentials = service_account.Credentials.from_service_account_file(
    "service_account.json",
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/documents"
    ],
)

gspread_client = gspread.authorize(credentials)

docs_service = build('docs', 'v1', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)

# Slack Webhook URL
SLACK_WEBHOOK_URL = os.environ.get("SLACK_WEBHOOK_URL")

# ドキュメントに書き込む関数
def write_to_document(doc_id, row, topic):
    requests_list = []
    for key, value in row.items():
        if value:
            # まずテキストを挿入
            requests_list.append({
                "insertText": {
                    "location": {"index": 1},
                    "text": f"{key}\n"
                }
            })
            # 見出しスタイルを適用
            requests_list.append({
                "updateParagraphStyle": {
                    "range": {
                        "startIndex": 1,
                        "endIndex": 1 + len(key) + 1
                    },
                    "paragraphStyle": {
                        "namedStyleType": "HEADING_2"
                    },
                    "fields": "namedStyleType"
                }
            })
            # 本文を挿入
            requests_list.append({
                "insertText": {
                    "location": {"index": 1},
                    "text": f"{value}\n\n"
                }
            })

    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": list(reversed(requests_list))}
    ).execute()

# スプレッドシートに書き込む関数
def write_to_spreadsheet(spreadsheet_id, worksheet_title, headers, rows):
    sh = gspread_client.open_by_key(spreadsheet_id)
    try:
        worksheet = sh.worksheet(worksheet_title)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=worksheet_title, rows="100", cols="20")

    if worksheet.row_count == 0:
        worksheet.append_row(headers)

    for row in rows:
        worksheet.append_row(row)

# Slack通知関数
def send_slack_notification(message):
    if SLACK_WEBHOOK_URL:
        requests.post(SLACK_WEBHOOK_URL, json={"text": message})

@app.post("/trigger")
async def trigger(request: Request):
    data = await request.json()

    if not isinstance(data, dict):
        raise ValueError("Invalid data format")

    type_ = data.get("type")
    topic = data.get("topic", "未分類")

    if type_ == "spreadsheet":
        headers = data.get("headers", [])
        rows = data.get("rows", [])

        now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        sheet_title = f"{topic}_{now}"
        spreadsheet = gspread_client.create(sheet_title)
        spreadsheet.share(None, perm_type='anyone', role='writer')

        write_to_spreadsheet(spreadsheet.id, "Sheet1", headers, rows)

        slack_message = f"✅ スプレッドシートが作成されました！\n{spreadsheet.url}"
        send_slack_notification(slack_message)

    elif type_ == "document":
        now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        doc_title = f"{topic}_議事録_{now}"
        doc = drive_service.files().create(
            body={"name": doc_title, "mimeType": "application/vnd.google-apps.document"},
            fields="id"
        ).execute()
        doc_id = doc.get("id")

        write_to_document(doc_id, data.get("content", {}), topic)

        slack_message = f"📝 ドキュメントが作成されました！\nhttps://docs.google.com/document/d/{doc_id}/edit"
        send_slack_notification(slack_message)

    return {"status": "success"}
