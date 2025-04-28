import os
import json
import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
import requests

app = FastAPI()

# スコープを設定
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents'
]

# service_account情報の読み込み
credentials_info = json.loads(os.getenv("GOOGLE_CREDENTIALS_JSON"))
credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)

gs_client = gspread.authorize(credentials)
drive_service = build('drive', 'v3', credentials=credentials)
docs_service = build('docs', 'v1', credentials=credentials)
slack_webhook_url = os.getenv("SLACK_WEBHOOK_URL")

@app.post("/trigger")
async def trigger(request: Request):
    data = await request.json()

    file_link = ""

    if data['type'] == 'spreadsheet':
        # スプレッドシート新規作成
        sheet = gs_client.create(data['topic'])
        sheet.share(None, perm_type='anyone', role='writer')
        worksheet = sheet.get_worksheet(0)

        # headers追加
        worksheet.append_row(data['headers'])
        # rows追加
        for row in data['rows']:
            worksheet.append_row(row)

        file_link = f"https://docs.google.com/spreadsheets/d/{sheet.id}"

    elif data['type'] == 'document':
        # ドキュメント新規作成
        doc = docs_service.documents().create(body={"title": data['topic']}).execute()
        doc_id = doc['documentId']

        requests_list = []
        for content in data['contents']:
            # セクションタイトル
            requests_list.append({
                "insertText": {
                    "location": {"index": 1},
                    "text": f"\n{content['heading']}\n"
                }
            })
            # 本文
            requests_list.append({
                "insertText": {
                    "location": {"index": 1},
                    "text": f"{content['body']}\n"
                }
            })

        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": list(reversed(requests_list))}
        ).execute()

        file_link = f"https://docs.google.com/document/d/{doc_id}"

    else:
        return JSONResponse(status_code=400, content={"message": "Invalid type. Must be 'spreadsheet' or 'document'."})

    # Slack通知
    if slack_webhook_url:
        message = {"text": f"保存が完了しました！\n{file_link}"}
        try:
            requests.post(slack_webhook_url, json=message)
        except Exception as e:
            print(f"Slack通知失敗: {e}")

    return {"message": "保存完了", "link": file_link}
