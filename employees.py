from datetime import datetime
from sheets.client import get_sheet, ensure_headers

HEADERS = ["tg_id", "username", "full_name", "role", "registered_at"]
# role: employee | admin


def _sheet():
    s = get_sheet("employees")
    ensure_headers(s, HEADERS)
    return s


def register(tg_id: int, username: str, full_name: str, role: str = "employee"):
    sheet = _sheet()
    records = sheet.get_all_records()
    for r in records:
        if str(r["tg_id"]) == str(tg_id):
            return False  # уже зарегистрирован
    sheet.append_row([
        tg_id, username or "", full_name,
        role, datetime.now().strftime("%Y-%m-%d %H:%M")
    ])
    return True


def get_all():
    return _sheet().get_all_records()


def get_employees():
    return [r for r in get_all() if r["role"] == "employee"]


def get_admins():
    return [r for r in get_all() if r["role"] == "admin"]


def get_by_tg_id(tg_id: int):
    for r in get_all():
        if str(r["tg_id"]) == str(tg_id):
            return r
    return None


def is_registered(tg_id: int) -> bool:
    return get_by_tg_id(tg_id) is not None


def is_admin(tg_id: int) -> bool:
    r = get_by_tg_id(tg_id)
    return r is not None and r["role"] == "admin"


def set_admin(tg_id: int):
    sheet = _sheet()
    records = sheet.get_all_records()
    for i, r in enumerate(records, start=2):
        if str(r["tg_id"]) == str(tg_id):
            sheet.update_cell(i, HEADERS.index("role") + 1, "admin")
            return True
    return False
