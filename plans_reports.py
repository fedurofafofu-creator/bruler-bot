from datetime import datetime, date
from sheets.client import get_sheet, ensure_headers

P_HEADERS = ["date", "tg_id", "full_name", "plan_text", "submitted_at"]
R_HEADERS = ["date", "tg_id", "full_name", "report_text", "submitted_at"]


def _plans():
    s = get_sheet("plans")
    ensure_headers(s, P_HEADERS)
    return s


def _reports():
    s = get_sheet("reports")
    ensure_headers(s, R_HEADERS)
    return s


def save_plan(tg_id: int, full_name: str, text: str):
    today = date.today().strftime("%Y-%m-%d")
    _plans().append_row([today, tg_id, full_name, text,
                         datetime.now().strftime("%Y-%m-%d %H:%M")])


def save_report(tg_id: int, full_name: str, text: str):
    today = date.today().strftime("%Y-%m-%d")
    _reports().append_row([today, tg_id, full_name, text,
                           datetime.now().strftime("%Y-%m-%d %H:%M")])


def get_plans_today():
    today = date.today().strftime("%Y-%m-%d")
    return [r for r in _plans().get_all_records() if r["date"] == today]


def get_reports_today():
    today = date.today().strftime("%Y-%m-%d")
    return [r for r in _reports().get_all_records() if r["date"] == today]


def has_plan_today(tg_id: int) -> bool:
    return any(str(r["tg_id"]) == str(tg_id) for r in get_plans_today())


def has_report_today(tg_id: int) -> bool:
    return any(str(r["tg_id"]) == str(tg_id) for r in get_reports_today())


def get_plans_for_week(iso_week: str):
    """iso_week: '2026-W24'"""
    all_plans = _plans().get_all_records()
    result = []
    for r in all_plans:
        try:
            d = datetime.strptime(r["date"], "%Y-%m-%d")
            if d.strftime("%Y-W%W") == iso_week:
                result.append(r)
        except Exception:
            pass
    return result


def get_reports_for_week(iso_week: str):
    all_reports = _reports().get_all_records()
    result = []
    for r in all_reports:
        try:
            d = datetime.strptime(r["date"], "%Y-%m-%d")
            if d.strftime("%Y-W%W") == iso_week:
                result.append(r)
        except Exception:
            pass
    return result
