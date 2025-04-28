import os
import json
import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
import requests
from gspread_formatting import format_cell_range, cellFormat, color, textFormat, set_column_width

app = FastAPI()

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents'
]

credentials_info = json.loads(os.getenv("GOOGLE_CREDENTIALS_JSON"))
credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)

gs_client = gspread.authorize(credentials)
drive_service = build('drive', 'v3', credentials=credentials)
docs_service = build('docs', 'v1', credentials=credentials)
slack_webhook_url = os.getenv("SLACK_WEBHOOK_URL")

def set_document_permissions(file_id):
    drive_service.permissions().create(
        fileId=file_id,
        body={
            'role': 'writer',
            'type': 'anyone'
        }
    ).execute()

def format_headers(worksheet, header_count):
    header_format = cellFormat(
        backgroundColor=color(0.9, 0.9, 0.9),
        textFormat=textFormat(bold=True),
        horizontalAlignment='CENTER'
    )
    format_cell_range(worksheet, f"A1:{chr(64+header_count)}1", header_format)

def auto_resize_columns(worksheet):
    all_values = worksheet.get_all_values()
    if not all_values:
        return
    columns = list(zip(*all_values))
    for i, col in enumerate(columns):
        max_length = max(len(str(cell)) for cell in col)
        width = max(100, min(max_length * 10, 400))
        set_column_width(worksheet, chr(65+i), width)

@app.post("/trigger")
async def trigger(request: Request):
    data = await request.json()

    file_link = ""

    if data['type'] == 'spreadsheet':
        sheet = gs_client.create(data['topic'])
        sheet.share(None, perm_type='anyone', role='writer')
        worksheet = sheet.get_worksheet(0)

        worksheet.append_row(data['headers'])
        for row in data['rows']:
            worksheet.append_row(row)

        format_headers(worksheet, len(data['headers']))
        auto_resize_columns(worksheet)

        file_link = f"https://docs.google.com/spreadsheets/d/{sheet.id}"

    elif data['type'] == 'document':
        doc = docs_service.documents().create(body={"title": data['topic']}).execute()
        doc_id = doc['documentId']

        set_document_permissions(doc_id)

        requests_list = []
        cursor = 1

        # タイトルを本文に挿入 + 見出し1に設定
        title_text = data['topic'] + "\n"
        requests_list.append({
            "insertText": {
                "location": {"index": cursor},
                "text": title_text
            }
        })
        requests_list.append({
            "updateParagraphStyle": {
                "range": {
                    "startIndex": cursor,
                    "endIndex": cursor + len(title_text) - 1
                },
                "paragraphStyle": {
                    "namedStyleType": "HEADING_1"
                },
                "fields": "namedStyleType"
            }
        })
        cursor += len(title_text)

        for content in data['contents']:
            heading_text = content['heading'] + "\n"
            body_text = content['body'] + "\n"

            # 見出し挿入（HEADING_2）
            requests_list.append({
                "insertText": {
                    "location": {"index": cursor},
                    "text": heading_text
                }
            })
            requests_list.append({
                "updateParagraphStyle": {
                    "range": {
                        "startIndex": cursor,
                        "endIndex": cursor + len(heading_text) - 1
                    },
                    "paragraphStyle": {
                        "namedStyleType": "HEADING_2"
                    },
                    "fields": "namedStyleType"
                }
            })
            cursor += len(heading_text)

            # 本文挿入（標準テキスト）
            requests_list.append({
                "insertText": {
                    "location": {"index": cursor},
                    "text": body_text
                }
            })
            cursor += len(body_text)

        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": requests_list}
        ).execute()

        file_link = f"https://docs.google.com/document/d/{doc_id}"

    else:
        return JSONResponse(status_code=400, content={"message": "Invalid type. Must be 'spreadsheet' or 'document'."})

    if slack_webhook_url:
        message = {"text": f"保存が完了しました！\n{file_link}"}
        try:
            requests.post(slack_webhook_url, json=message)
        except Exception as e:
            print(f"Slack通知失敗: {e}")

    return {"message": "保存完了", "link": file_link}
