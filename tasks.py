import uuid
from datetime import datetime, date
from sheets.client import get_sheet, ensure_headers

HEADERS = [
    "task_id", "created_by_id", "created_by_name",
    "assigned_to_id", "assigned_to_name",
    "title", "deadline", "status", "created_at", "done_at", "comment"
]
# status: open | in_progress | done | overdue


def _sheet():
    s = get_sheet("tasks")
    ensure_headers(s, HEADERS)
    return s


def create_task(created_by_id: int, created_by_name: str,
                assigned_to_id: int, assigned_to_name: str,
                title: str, deadline: str) -> str:
    task_id = str(uuid.uuid4())[:8].upper()
    _sheet().append_row([
        task_id, created_by_id, created_by_name,
        assigned_to_id, assigned_to_name,
        title, deadline, "open",
        datetime.now().strftime("%Y-%m-%d %H:%M"), "", ""
    ])
    return task_id


def get_all_tasks():
    return _sheet().get_all_records()


def get_task(task_id: str):
    for r in get_all_tasks():
        if r["task_id"] == task_id.upper():
            return r
    return None


def get_tasks_for_user(tg_id: int):
    return [r for r in get_all_tasks()
            if str(r["assigned_to_id"]) == str(tg_id)
            and r["status"] not in ("done",)]


def get_open_tasks():
    return [r for r in get_all_tasks() if r["status"] == "open"]


def _find_row(task_id: str):
    records = get_all_tasks()
    for i, r in enumerate(records, start=2):
        if r["task_id"] == task_id.upper():
            return i
    return None


def mark_done(task_id: str, comment: str = "") -> bool:
    row = _find_row(task_id)
    if row is None:
        return False
    sheet = _sheet()
    sheet.update_cell(row, HEADERS.index("status") + 1, "done")
    sheet.update_cell(row, HEADERS.index("done_at") + 1,
                      datetime.now().strftime("%Y-%m-%d %H:%M"))
    if comment:
        sheet.update_cell(row, HEADERS.index("comment") + 1, comment)
    return True


def mark_in_progress(task_id: str) -> bool:
    row = _find_row(task_id)
    if row is None:
        return False
    _sheet().update_cell(row, HEADERS.index("status") + 1, "in_progress")
    return True


def get_tasks_due_tomorrow():
    tomorrow = (datetime.now().date())
    from datetime import timedelta
    tomorrow = (datetime.now().date() + timedelta(days=1)).strftime("%Y-%m-%d")
    return [r for r in get_all_tasks()
            if r["deadline"] == tomorrow and r["status"] != "done"]


def get_tasks_due_today():
    today = date.today().strftime("%Y-%m-%d")
    return [r for r in get_all_tasks()
            if r["deadline"] == today and r["status"] != "done"]


def get_overdue_tasks():
    today = date.today().strftime("%Y-%m-%d")
    result = []
    for r in get_all_tasks():
        if r["status"] != "done" and r["deadline"] and r["deadline"] < today:
            result.append(r)
    return result
