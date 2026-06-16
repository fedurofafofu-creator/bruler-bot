import json
import gspread
from google.oauth2.service_account import Credentials
from config import GOOGLE_CREDENTIALS_JSON, SPREADSHEET_ID

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_client = None
_spreadsheet = None


def get_client():
    global _client
    if _client is None:
        if GOOGLE_CREDENTIALS_JSON:
            creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
        else:
            # локальная разработка: файл credentials.json рядом с кодом
            with open("credentials.json") as f:
                creds_dict = json.load(f)
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        _client = gspread.authorize(creds)
    return _client


def get_spreadsheet():
    global _spreadsheet
    if _spreadsheet is None:
        _spreadsheet = get_client().open_by_key(SPREADSHEET_ID)
    return _spreadsheet


def get_sheet(name: str):
    ss = get_spreadsheet()
    try:
        return ss.worksheet(name)
    except gspread.WorksheetNotFound:
        return ss.add_worksheet(title=name, rows=1000, cols=20)


def ensure_headers(sheet, headers: list[str]):
    """Ставит заголовки если лист пустой."""
    if sheet.row_count == 0 or sheet.cell(1, 1).value != headers[0]:
        sheet.insert_row(headers, 1)
