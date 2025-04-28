from fastapi import FastAPI, Request
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import requests
import os
import json
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

app = FastAPI()

# 環境変数から情報取得
WEBHOOK_URL = os.environ['SLACK_WEBHOOK_URL']
GOOGLE_CREDENTIALS = json.loads(os.environ['GOOGLE_CREDENTIALS_JSON'])
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/documents"]

# Googleサービス認証
credentials = Credentials.from_service_account_info(GOOGLE_CREDENTIALS, scopes=SCOPES)
gc = gspread.authorize(credentials)
docs_service = build('docs', 'v1', credentials=credentials)

gsheet_last_ids = {}
gdoc_last_ids = {}

# スプレッドシートの列自動調整
def auto_resize_columns(worksheet):
    all_values = worksheet.get_all_values()
    if not all_values:
        return
    columns = list(zip(*all_values))
    for i, col in enumerate(columns):
        max_length = max(len(str(cell)) for cell in col)
        width = max(100, min(max_length * 10, 400))
        worksheet.format(chr(65+i)+":"+chr(65+i), {"pixelSize": width})

# スプレッドシート作成
def create_spreadsheet(title, email, topic):
    sh = gc.create(title)
    sh.share(email, perm_type='user', role='writer')
    worksheet = sh.sheet1
    gsheet_last_ids[topic] = sh.id
    return sh, worksheet

# ドキュメント作成
def create_document(title, email, topic):
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc.get('documentId')
    gdoc_last_ids[topic] = doc_id
    return doc_id

# Slack通知
def send_slack(message):
    payload = {"text": message}
    headers = {"Content-Type": "application/json"}
    requests.post(WEBHOOK_URL, json=payload, headers=headers)

# ドキュメントへ書き込み
def write_to_document(doc_id, rows, topic):
    requests_list = []
    cursor = 1  # 最初にタイトルが入っているので、1から開始

    # まず本文挿入リクエストをまとめる
    for row in rows:
        for key, value in row.items():
            text = f"{key}\n{value}\n\n"
            requests_list.append({
                "insertText": {
                    "location": {"index": cursor},
                    "text": text
                }
            })
            cursor += len(text)

    # その後、見出しスタイル適用リクエストをまとめる
    paragraph_cursor = 1
    for row in rows:
        for key, value in row.items():
            key_len = len(key)
            requests_list.append({
                "updateParagraphStyle": {
                    "range": {
                        "startIndex": paragraph_cursor,
                        "endIndex": paragraph_cursor + key_len
                    },
                    "paragraphStyle": {"namedStyleType": "HEADING_2"},
                    "fields": "namedStyleType"
                }
            })
            paragraph_cursor += len(key) + len(value) + 2  # key + 改行 + value + 改行

    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests_list}).execute()

@app.post("/trigger")
async def trigger(request: Request):
    raw_data = await request.json()

    if isinstance(raw_data, dict):
        data = [raw_data]
    elif isinstance(raw_data, list):
        if all(isinstance(d, str) for d in raw_data):
            data = [json.loads(d) for d in raw_data]
        else:
            data = raw_data
    else:
        raise ValueError("Unexpected data format")

    for row in data:
        force_new = row.get("新規作成", False)
        topic = row.get("話題", "未分類")

        # スプレッドシート
        if force_new or topic not in gsheet_last_ids:
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            title = f"{topic}_{now}"
            sh, worksheet = create_spreadsheet(title, 'nattsuchanneru@gmail.com', topic)
            send_slack(f"✅ 新しいスプレッドシートが作成されました！\n{sh.url}")
        else:
            sh = gc.open_by_key(gsheet_last_ids[topic])
            worksheet = sh.sheet1

        if worksheet.row_count == 0 or worksheet.acell('A1').value is None:
            worksheet.append_row(list(row.keys()))

        worksheet.append_row(list(row.values()))
        auto_resize_columns(worksheet)

        # ドキュメント
        if force_new or topic not in gdoc_last_ids:
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            doc_title = f"{topic}議事録_{now}"
            doc_id = create_document(doc_title, 'nattsuchanneru@gmail.com', topic)
            send_slack(f"📄 新しいGoogleドキュメントが作成されました！\nhttps://docs.google.com/document/d/{doc_id}/edit")
        else:
            doc_id = gdoc_last_ids[topic]

        write_to_document(doc_id, [row], topic)

    return {"status": "success"}
