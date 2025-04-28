from fastapi import FastAPI, Request
import os
import json
from pydantic import BaseModel
from google.oauth2 import service_account
from googleapiclient.discovery import build
import gspread
import requests

app = FastAPI()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents"
]

gcp_credentials = json.loads(os.getenv("GOOGLE_CREDENTIALS_JSON"))
credentials = service_account.Credentials.from_service_account_info(
    gcp_credentials, scopes=SCOPES
)

gspread_client = gspread.authorize(credentials)
drive_service = build("drive", "v3", credentials=credentials)
docs_service = build("docs", "v1", credentials=credentials)
slack_webhook_url = os.getenv("SLACK_WEBHOOK_URL")

class Payload(BaseModel):
    新規作成: bool
    話題: str
    内容: str
    得た情報: str
    メモ: str
    参考URL: str
    アクション: str

def create_spreadsheet(title):
    spreadsheet = gspread_client.create(title)
    spreadsheet.share(None, perm_type="anyone", role="writer")
    return spreadsheet

def create_document(title):
    body = {"title": title}
    doc = docs_service.documents().create(body=body).execute()
    doc_id = doc["documentId"]
    drive_service.permissions().create(
        fileId=doc_id,
        body={"type": "anyone", "role": "writer"},
        fields="id"
    ).execute()
    return doc_id

def write_to_spreadsheet(sheet, row):
    worksheet = sheet.sheet1
    worksheet.append_row(list(row.values()))

def write_to_document(doc_id, row, topic):
    requests_list = []
    # 見出し（話題）
    requests_list.append({
        "insertText": {
            "location": {"index": 1},
            "text": f"{topic}\n"
        }
    })
    requests_list.append({
        "updateParagraphStyle": {
            "range": {
                "startIndex": 1,
                "endIndex": 1 + len(topic) + 1
            },
            "paragraphStyle": {
                "namedStyleType": "HEADING_1"
            },
            "fields": "namedStyleType"
        }
    })
    # 本文（内容などまとめて）
    content = f"\n".join([
        f"- 内容: {row['内容']}",
        f"- 得た情報: {row['得た情報']}",
        f"- メモ: {row['メモ']}",
        f"- 参考URL: {row['参考URL']}",
        f"- アクション: {row['アクション']}"
    ])
    requests_list.append({
        "insertText": {
            "location": {"index": 1 + len(topic) + 1},
            "text": f"{content}\n"
        }
    })

    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": requests_list}
    ).execute()

def send_slack_notification(sheet_url, doc_url):
    message = {
        "text": f"✅ データ保存完了！\nスプレッドシート: {sheet_url}\nドキュメント: {doc_url}"
    }
    requests.post(slack_webhook_url, json=message)

@app.post("/trigger")
async def trigger(payload: Payload):
    data = payload.dict()
    title_base = data["話題"].replace(" ", "_").replace("-", "_")

    if data["新規作成"]:
        spreadsheet = create_spreadsheet(f"{title_base}_sheet")
        sheet_url = spreadsheet.url
        doc_id = create_document(f"{title_base}_doc")
        doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"

        properties = spreadsheet.sheet1.row_values(1)
        if not properties:
            spreadsheet.sheet1.append_row(["話題", "内容", "得た情報", "メモ", "参考URL", "アクション"])

    else:
        sheet_url = data.get("スプレッドシートURL")
        doc_url = data.get("ドキュメントURL")
        spreadsheet = gspread_client.open_by_url(sheet_url)
        doc_id = doc_url.split("/d/")[1].split("/")[0]

    row = {
        "話題": data["話題"],
        "内容": data["内容"],
        "得た情報": data["得た情報"],
        "メモ": data["メモ"],
        "参考URL": data["参考URL"],
        "アクション": data["アクション"]
    }

    write_to_spreadsheet(spreadsheet, row)
    write_to_document(doc_id, row, data["話題"])
    send_slack_notification(sheet_url, doc_url)

    return {"status": "success"}
