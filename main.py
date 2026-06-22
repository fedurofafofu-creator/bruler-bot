import asyncio
import logging
import json
import uuid
import re
import io
import time
import pytz
import gspread
import calendar
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials

from telegram import Update, Bot, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, ContextTypes, filters, CallbackQueryHandler
)
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
import os

logging.basicConfig(format="%(asctime)s [%(levelname)s] %(name)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ── CONFIG ────────────────────────────────────────────────────────────────────
BOT_TOKEN      = os.getenv("BOT_TOKEN", "8767537310:AAHK1-RmvgH6yF6ZShQFq3A1DeLU0uFsAMs")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1LCmoydfS73DKwjQcwMSnpOZxRHiKMAMjOZRRk0AwYnE")
ADMIN_CHAT_ID  = int(os.getenv("ADMIN_CHAT_ID", "-5486975908"))
TZ             = ZoneInfo("Europe/Moscow")
APZ            = pytz.timezone("Europe/Moscow")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

DEPT_LIST = ["Маркетинг", "Продажи", "Бухгалтерия", "Производство и дизайн", "Руководство"]

# Частые русские имена (для определения порядка Имя/Фамилия при нормализации)
COMMON_FIRST_NAMES = {
    "александр","алексей","анатолий","андрей","антон","аркадий","артём","артем",
    "борис","вадим","валентин","валерий","василий","виктор","виталий","владимир",
    "владислав","геннадий","георгий","григорий","даниил","денис","дмитрий","евгений",
    "егор","иван","игорь","илья","кирилл","константин","леонид","максим","максимилиан",
    "михаил","никита","николай","олег","павел","петр","пётр","роман","руслан",
    "семен","семён","сергей","степан","тимофей","тимур","федор","фёдор","юрий","ярослав",
    "александра","алина","алла","анастасия","анжела","анна","антонина","валентина",
    "валерия","вера","вероника","виктория","галина","дарья","диана","евгения",
    "екатерина","елена","елизавета","зоя","ирина","карина","кристина","ксения",
    "лариса","лидия","любовь","людмила","маргарита","марина","мария","надежда",
    "наталья","оксана","ольга","полина","раиса","светлана","софия","софья",
    "тамара","татьяна","юлия","яна",
}

def normalize_full_name(raw: str) -> str:
    """Приводит ввод к формату 'Имя Фамилия'. Распознаёт частые имена,
    переставляет слова если первое слово похоже на фамилию."""
    parts = raw.strip().split()
    if len(parts) < 2:
        return raw.strip().title()
    parts = [p.capitalize() for p in parts[:2]]
    first_lower = parts[0].lower()
    second_lower = parts[1].lower()
    if first_lower in COMMON_FIRST_NAMES:
        return f"{parts[0]} {parts[1]}"
    if second_lower in COMMON_FIRST_NAMES:
        return f"{parts[1]} {parts[0]}"
    # не распознали — оставляем как ввёл пользователь
    return f"{parts[0]} {parts[1]}"

# Стартовые значения — далее отдел хранится в Google Sheets (колонка department)
# и может быть переопределён через кнопки при регистрации / командой /setdept
DEPARTMENTS_SEED = {
    "7070230704": "Руководство",
    "7198542902": "Руководство",
    "8151347813": "Продажи",
    "195676845":  "Маркетинг",
    "8069881891": "Бухгалтерия",
    "458764300":  "Производство и дизайн",
    "860192861":  "Производство и дизайн",
    "89555212":   "Производство и дизайн",
    "549232571":  "Маркетинг",
}

def get_dept(tg_id) -> str:
    """Отдел сотрудника: сначала смотрим в Sheets, иначе — seed-словарь, иначе 'Без отдела'."""
    r = emp_by_id(tg_id)
    if r and r.get("department"):
        return r["department"]
    return DEPARTMENTS_SEED.get(str(tg_id), "Без отдела")

# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
_gc = None
_ss = None
_ws_cache = {}        # name -> Worksheet object, кэшируется навсегда (объект не меняется)
_data_cache = {}       # name -> (timestamp, rows), короткий TTL чтобы не дёргать API на каждый вызов
DATA_CACHE_TTL = 8     # секунд — этого достаточно чтобы пережить серию из 5-6 команд подряд

def gc():
    global _gc
    if _gc is None:
        raw = os.getenv("GOOGLE_CREDENTIALS_JSON", "")
        if raw:
            info = json.loads(raw)
        else:
            with open("credentials.json") as f:
                info = json.load(f)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        _gc = gspread.authorize(creds)
    return _gc

def ss():
    global _ss
    if _ss is None:
        _ss = gc().open_by_key(SPREADSHEET_ID)
    return _ss

def sheet(name):
    """Кэширует объект Worksheet — избегает повторного fetch_sheet_metadata
    при каждом вызове (это отдельный API-запрос, который быстро съедает квоту)."""
    if name in _ws_cache:
        return _ws_cache[name]
    try:
        ws = ss().worksheet(name)
    except gspread.WorksheetNotFound:
        ws = ss().add_worksheet(title=name, rows=2000, cols=20)
    _ws_cache[name] = ws
    return ws

def invalidate_cache(name=None):
    """Сбрасывает кэш данных после любой записи, чтобы следующее чтение было свежим."""
    if name:
        _data_cache.pop(name, None)
    else:
        _data_cache.clear()

def ensure_headers(ws, headers):
    """Гарантирует корректный, не дублирующийся заголовок в первой строке."""
    vals = ws.row_values(1)
    if vals != headers:
        # перезаписываем строку заголовка целиком (без сдвига остальных строк)
        ws.update('A1', [headers])

def safe_records(ws, headers):
    """
    Читает данные листа вручную по позиции колонки, без gspread.get_all_records().
    Устойчиво к повреждённому заголовку и расхождению числа колонок после миграции схемы.
    Кэширует результат на DATA_CACHE_TTL секунд — несколько команд подряд от одного
    или разных пользователей переиспользуют один и тот же снимок данных, что резко
    снижает число запросов к Google Sheets API и защищает от 429 Quota exceeded.
    При срабатывании квоты делает 2 повторные попытки с задержкой; если квота всё
    ещё превышена — отдаёт последний известный кэш (даже просроченный), а не падает.
    """
    cache_key = ws.title
    cached = _data_cache.get(cache_key)
    if cached and (datetime.now() - cached[0]).total_seconds() < DATA_CACHE_TTL:
        return cached[1]

    last_error = None
    for attempt in range(3):
        try:
            ensure_headers(ws, headers)
            all_values = ws.get_all_values()
            n = len(headers)
            result = []
            if len(all_values) >= 2:
                for row in all_values[1:]:
                    if len(row) < n:
                        row = row + [""] * (n - len(row))
                    elif len(row) > n:
                        row = row[:n]
                    result.append({headers[i]: row[i] for i in range(n)})
            _data_cache[cache_key] = (datetime.now(), result)
            return result
        except gspread.exceptions.APIError as e:
            last_error = e
            if "429" in str(e) or "Quota exceeded" in str(e):
                # Блокирующий sleep здесь осознанно: это редкий аварийный путь
                # (квота превышена), а не штатный режим — благодаря кэшу выше
                # обычные обращения сюда вообще не доходят.
                time.sleep(2 * (attempt + 1))
                continue
            raise

    # квота не отпустила за 3 попытки — отдаём устаревший кэш, если он есть,
    # лучше показать чуть старые данные чем уронить команду полностью
    if cached:
        logger.warning(f"Sheets quota exceeded for {cache_key}, serving stale cache")
        return cached[1]
    logger.error(f"Sheets quota exceeded for {cache_key}, no cache available")
    raise last_error

# ── EMPLOYEES ─────────────────────────────────────────────────────────────────
EMP_H = ["tg_id","username","full_name","role","registered_at","department"]

def emp_sheet():
    ws = sheet("employees"); ensure_headers(ws, EMP_H); return ws

def emp_all():
    return safe_records(emp_sheet(), EMP_H)

def emp_employees():
    """Все, кто проходит ежедневный цикл план→EOD: employee, dept_head и admin —
    руководители тоже работают и должны отчитываться о своих задачах."""
    return [r for r in emp_all() if r["role"] in ("employee", "dept_head", "admin")]

def emp_admins():
    return [r for r in emp_all() if r["role"] == "admin"]

def emp_dept_heads():
    return [r for r in emp_all() if r["role"] == "dept_head"]

def emp_by_id(tg_id):
    for r in emp_all():
        if str(r["tg_id"]) == str(tg_id):
            return r
    return None

def emp_registered(tg_id):
    return emp_by_id(tg_id) is not None

def emp_is_admin(tg_id):
    r = emp_by_id(tg_id)
    return r is not None and r["role"] == "admin"

def emp_is_dept_head(tg_id):
    r = emp_by_id(tg_id)
    return r is not None and r["role"] == "dept_head"

def emp_has_management_access(tg_id):
    """Полный доступ к панели (admin) или ограниченный (dept_head)."""
    r = emp_by_id(tg_id)
    return r is not None and r["role"] in ("admin", "dept_head")

def emp_managed_dept(tg_id) -> str:
    """Отдел, который видит dept_head. Для admin — пусто (видит всё)."""
    r = emp_by_id(tg_id)
    if not r:
        return ""
    if r["role"] == "admin":
        return ""
    return get_dept(tg_id)

def emp_register(tg_id, username, full_name, role="employee", department=""):
    if emp_registered(tg_id):
        return False
    emp_sheet().append_row([tg_id, username or "", full_name, role,
                             datetime.now().strftime("%Y-%m-%d %H:%M"), department])
    invalidate_cache("employees")
    return True

def emp_set_admin(tg_id):
    ws = emp_sheet()
    for i, r in enumerate(safe_records(ws, EMP_H), start=2):
        if str(r["tg_id"]) == str(tg_id):
            ws.update_cell(i, EMP_H.index("role") + 1, "admin")
            invalidate_cache("employees")
            return True
    return False

def emp_set_role(tg_id, role: str):
    ws = emp_sheet()
    for i, r in enumerate(safe_records(ws, EMP_H), start=2):
        if str(r["tg_id"]) == str(tg_id):
            ws.update_cell(i, EMP_H.index("role") + 1, role)
            invalidate_cache("employees")
            return True
    return False

def emp_set_department(tg_id, department: str):
    ws = emp_sheet()
    for i, r in enumerate(safe_records(ws, EMP_H), start=2):
        if str(r["tg_id"]) == str(tg_id):
            ws.update_cell(i, EMP_H.index("department") + 1, department)
            invalidate_cache("employees")
            return True
    return False

def emp_set_full_name(tg_id, full_name: str):
    ws = emp_sheet()
    for i, r in enumerate(safe_records(ws, EMP_H), start=2):
        if str(r["tg_id"]) == str(tg_id):
            ws.update_cell(i, EMP_H.index("full_name") + 1, full_name)
            invalidate_cache("employees")
            return True
    return False

# ── PLANS & REPORTS ───────────────────────────────────────────────────────────
PH = ["date","tg_id","full_name","plan_text","submitted_at"]
RH = ["date","tg_id","full_name","report_text","submitted_at"]

def plans_sheet():
    ws = sheet("plans"); ensure_headers(ws, PH); return ws
def reports_sheet():
    ws = sheet("reports"); ensure_headers(ws, RH); return ws

def today_str():
    return date.today().strftime("%Y-%m-%d")

def save_plan(tg_id, name, text):
    plans_sheet().append_row([today_str(), tg_id, name, text,
                               datetime.now().strftime("%Y-%m-%d %H:%M")])
    invalidate_cache("plans")

def save_report(tg_id, name, text):
    reports_sheet().append_row([today_str(), tg_id, name, text,
                                 datetime.now().strftime("%Y-%m-%d %H:%M")])
    invalidate_cache("reports")

def plans_today():
    return [r for r in safe_records(plans_sheet(), PH) if r["date"] == today_str()]

def reports_today():
    return [r for r in safe_records(reports_sheet(), RH) if r["date"] == today_str()]

def has_plan_today(tg_id):
    return any(str(r["tg_id"]) == str(tg_id) for r in plans_today())

def has_report_today(tg_id):
    return any(str(r["tg_id"]) == str(tg_id) for r in reports_today())

def records_for_week(ws_name, headers):
    ws = sheet(ws_name)
    iso = datetime.now().strftime("%Y-W%W")
    result = []
    for r in safe_records(ws, headers):
        try:
            d = datetime.strptime(r["date"], "%Y-%m-%d")
            if d.strftime("%Y-W%W") == iso:
                result.append(r)
        except Exception:
            pass
    return result

def records_for_month(ws_name, headers):
    ws = sheet(ws_name)
    month = datetime.now().strftime("%Y-%m")
    return [r for r in safe_records(ws, headers) if r.get("date","").startswith(month)]

def reports_for_period(date_from, date_to):
    ws = sheet("reports")
    return [r for r in safe_records(ws, RH)
            if r.get("date","") >= date_from and r.get("date","") <= date_to]

def plans_for_period(date_from, date_to):
    ws = sheet("plans")
    return [p for p in safe_records(ws, PH)
            if p.get("date","") >= date_from and p.get("date","") <= date_to]

# ── TASKS ─────────────────────────────────────────────────────────────────────
# status: open | in_progress | done | paused | overdue
TH = ["task_id","created_by_id","created_by_name","assigned_to_id",
      "assigned_to_name","title","deadline","status","created_at",
      "done_at","comment","result_link","source","channel"]

CHANNEL_LIST = ["Сайт", "Маркетплейсы", "Комиссионеры", "Опт", "Розница", "Bruler Studio", "Не применимо"]

def tasks_sheet():
    ws = sheet("tasks"); ensure_headers(ws, TH); return ws

def tasks_all():
    return safe_records(tasks_sheet(), TH)

def task_by_id(tid):
    for r in tasks_all():
        if r["task_id"] == tid.upper():
            return r
    return None

def tasks_for_user(tg_id):
    return [r for r in tasks_all()
            if str(r["assigned_to_id"]) == str(tg_id) and r["status"] not in ("done",)]

def tasks_open():
    return [r for r in tasks_all() if r["status"] in ("open","in_progress","paused")]

def tasks_overdue():
    today = today_str()
    return [r for r in tasks_all()
            if r["status"] not in ("done",) and r.get("deadline","") and r["deadline"] < today]

def tasks_due_tomorrow():
    tmr = (date.today() + timedelta(days=1)).strftime("%Y-%m-%d")
    return [r for r in tasks_all() if r.get("deadline") == tmr and r["status"] != "done"]

def tasks_due_today_list():
    return [r for r in tasks_all()
            if r.get("deadline") == today_str() and r["status"] != "done"]

def task_find_row(tid):
    for i, r in enumerate(safe_records(tasks_sheet(), TH), start=2):
        if r["task_id"] == tid.upper():
            return i
    return None

def task_create(by_id, by_name, to_id, to_name, title, deadline, source="manual", channel=""):
    tid = str(uuid.uuid4())[:8].upper()
    tasks_sheet().append_row([
        tid, by_id, by_name, to_id, to_name, title,
        deadline, "open", datetime.now().strftime("%Y-%m-%d %H:%M"),
        "", "", "", source, channel
    ])
    invalidate_cache("tasks")
    return tid

def task_update_channel(tid, channel: str):
    row = task_find_row(tid)
    if not row: return False
    tasks_sheet().update_cell(row, TH.index("channel") + 1, channel)
    invalidate_cache("tasks")
    return True

def task_update_status(tid, status):
    row = task_find_row(tid)
    if not row: return False
    ws = tasks_sheet()
    ws.update_cell(row, TH.index("status") + 1, status)
    if status == "done":
        ws.update_cell(row, TH.index("done_at") + 1, datetime.now().strftime("%Y-%m-%d %H:%M"))
    invalidate_cache("tasks")
    return True

def task_update_comment(tid, comment, link=""):
    row = task_find_row(tid)
    if not row: return False
    ws = tasks_sheet()
    ws.update_cell(row, TH.index("comment") + 1, comment)
    if link:
        ws.update_cell(row, TH.index("result_link") + 1, link)
    invalidate_cache("tasks")
    return True

def task_update_title(tid, title: str):
    row = task_find_row(tid)
    if not row: return False
    tasks_sheet().update_cell(row, TH.index("title") + 1, title)
    invalidate_cache("tasks")
    return True

def task_update_deadline(tid, deadline: str):
    row = task_find_row(tid)
    if not row: return False
    tasks_sheet().update_cell(row, TH.index("deadline") + 1, deadline)
    invalidate_cache("tasks")
    return True

def tasks_today_for_user(tg_id):
    """Все активные задачи пользователя для EOD-опроса: задачи с дедлайном сегодня
    ИЛИ просроченные ИЛИ длительные (дедлайн в будущем, но уже в работе/приостановлены).
    Это даёт промежуточный статус каждый вечер даже для многодневных задач."""
    today = today_str()
    return [r for r in tasks_all()
            if str(r["assigned_to_id"]) == str(tg_id)
            and r["status"] != "done"
            and (
                r.get("deadline","") == today           # дедлайн сегодня
                or r.get("deadline","") < today          # просрочена
                or r["status"] in ("in_progress", "paused")  # длительная, уже стартовала
            )]

# ── PARSE PLAN ────────────────────────────────────────────────────────────────
def parse_plan_items(text: str) -> list:
    """
    Разбирает план на пункты. Каждая строка может опционально содержать дату
    в конце в формате [ДД.ММ.ГГГГ] или (ДД.ММ.ГГГГ) — если её нет, дедлайн = сегодня.
    Пример: "Согласовать смету [25.06.2026]" → задача с дедлайном 25.06.2026.
    Возвращает список (title, deadline) где deadline — строка YYYY-MM-DD.
    """
    items = []
    today = today_str()
    date_pattern = re.compile(r"[\[\(]\s*(\d{2})\.(\d{2})\.(\d{4})\s*[\]\)]\s*$")
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        line = re.sub(r"^\d+[\.\)]\s*", "", line)
        line = re.sub(r"^[-–—•*]\s*", "", line)
        line = re.sub(r"^[✅🪡💘🩵🔜👌📌⚡✔▪▸►→]\s*", "", line)
        line = line.strip()

        deadline = today
        m = date_pattern.search(line)
        if m:
            d, mo, y = m.groups()
            deadline = f"{y}-{mo}-{d}"
            line = date_pattern.sub("", line).strip()

        if len(line) > 3:
            items.append((line, deadline))
    return items

# ── CONVERSATION STATES ───────────────────────────────────────────────────────
S_NAME         = 1
S_DEPT         = 2    # выбор отдела при регистрации
S_PLAN         = 10
S_REPORT       = 11
S_TASK         = 20
S_EOD_STATUS   = 30   # выбор статуса задачи конец дня
S_EOD_COMMENT  = 31   # комментарий + ссылка
S_EOD_EXTRA    = 32   # были ли другие задачи
S_EOD_EXTRA_TEXT = 33 # текст других задач

# Храним состояние сессии EOD (end-of-day) в bot_data
# bot_data["eod"][tg_id] = {"tasks": [...], "current_idx": 0, "results": [...]}

def dept_keyboard(prefix: str, tg_id_for_callback: str = "") -> InlineKeyboardMarkup:
    """Клавиатура выбора отдела. prefix: 'reg_dept_' (при регистрации) или 'admset_dept_{tg_id}_'."""
    rows = []
    row = []
    for d in DEPT_LIST:
        cb = f"{prefix}{tg_id_for_callback}_{d}" if tg_id_for_callback else f"{prefix}{d}"
        row.append(InlineKeyboardButton(d, callback_data=cb))
        if len(row) == 2:
            rows.append(row); row = []
    if row: rows.append(row)
    return InlineKeyboardMarkup(rows)

async def cmd_ping(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Мгновенный health-check, без обращения к Google Sheets — проверяет,
    что бот вообще отвечает, до проверки конкретных функций."""
    await update.effective_message.reply_text(
        f"🟢 Бот на связи.\n{datetime.now(TZ).strftime('%d.%m.%Y %H:%M:%S')} МСК"
    )

# ── /start ────────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    if emp_registered(tg_id):
        u = emp_by_id(tg_id)
        role_labels = {"admin": "руководитель", "dept_head": "руководитель отдела", "employee": "сотрудник"}
        role = role_labels.get(u["role"], "сотрудник")
        dept = get_dept(tg_id)
        await update.effective_message.reply_text(
            f"👋 Привет, {u['full_name']}! Ты зарегистрирован как {role} ({dept}).\n\n"
            "/plan — план на день\n/report — отчёт\n"
            "/task — поставить задачу\n/mytasks — мои задачи\n"
            "/done ID — отметить выполненной\n/eod — закрыть день"
        )
        return ConversationHandler.END
    await update.effective_message.reply_text(
        "👋 Привет! Это трекер команды Brûler d'Amour.\nНапиши имя и фамилию:"
    )
    return S_NAME

async def recv_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    raw_name = update.message.text.strip()
    name = normalize_full_name(raw_name)
    ctx.user_data["reg_name"] = name
    await update.message.reply_text(
        f"Приятно познакомиться, {name}!\nВыбери свой отдел:",
        reply_markup=dept_keyboard("reg_dept_")
    )
    return ConversationHandler.END

async def cb_reg_dept(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Сотрудник выбрал отдел при регистрации."""
    query = update.callback_query
    await query.answer()
    tg_id = query.from_user.id
    dept = query.data.replace("reg_dept_", "", 1)
    name = ctx.user_data.get("reg_name") or query.from_user.full_name or "Без имени"
    username = query.from_user.username or ""

    emp_register(tg_id, username, name, role="employee", department=dept)

    await query.edit_message_text(
        f"✅ Готово, {name}! Отдел: <b>{dept}</b>\n\n"
        "Каждый день в 11:00 — план, в 19:00 — закрытие дня.\n"
        "/plan — написать план сейчас",
        parse_mode="HTML"
    )

    # уведомляем всех руководителей о новом сотруднике
    notify_text = (
        f"🆕 <b>Новый сотрудник зарегистрирован</b>\n\n"
        f"Имя: {name}\n"
        f"Username: @{username}\n" if username else f"🆕 <b>Новый сотрудник зарегистрирован</b>\n\nИмя: {name}\n"
    )
    notify_text += f"Отдел (указал сам): {dept}\n\nМожешь изменить отдел или назначить руководителем отдела:"
    keyboard_rows = [
        [InlineKeyboardButton(f"📁 {d}", callback_data=f"admset_dept_{tg_id}_{d}") for d in DEPT_LIST[i:i+2]]
        for i in range(0, len(DEPT_LIST), 2)
    ]
    keyboard_rows.append([InlineKeyboardButton("⭐ Сделать руководителем отдела", callback_data=f"admset_head_{tg_id}")])
    keyboard = InlineKeyboardMarkup(keyboard_rows)

    for adm in emp_admins():
        try:
            await ctx.bot.send_message(int(adm["tg_id"]), notify_text, parse_mode="HTML", reply_markup=keyboard)
        except Exception as e:
            logger.warning(f"notify admin {adm['tg_id']}: {e}")

    return ConversationHandler.END

async def cb_admset_dept(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Руководитель меняет отдел сотрудника из уведомления о регистрации."""
    query = update.callback_query
    if not emp_is_admin(query.from_user.id):
        await query.answer("⛔ Только для администраторов.", show_alert=True); return
    await query.answer()
    # callback: admset_dept_{tg_id}_{Отдел}
    rest = query.data.replace("admset_dept_", "", 1)
    tg_id_str, dept = rest.split("_", 1)
    emp_set_department(int(tg_id_str), dept)
    emp = emp_by_id(int(tg_id_str))
    await query.edit_message_text(
        query.message.text_html + f"\n\n✅ Отдел установлен: <b>{dept}</b>",
        parse_mode="HTML"
    )

async def cb_admset_head(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Руководитель назначает сотрудника руководителем отдела из уведомления о регистрации."""
    query = update.callback_query
    if not emp_is_admin(query.from_user.id):
        await query.answer("⛔ Только для администраторов.", show_alert=True); return
    await query.answer()
    tg_id_str = query.data.replace("admset_head_", "", 1)
    emp_set_role(int(tg_id_str), "dept_head")
    emp = emp_by_id(int(tg_id_str))
    dept = get_dept(int(tg_id_str))
    await query.edit_message_text(
        query.message.text_html + f"\n\n⭐ Назначен руководителем отдела «{dept}»",
        parse_mode="HTML"
    )
    try:
        await ctx.bot.send_message(
            int(tg_id_str),
            f"⭐ Тебя назначили руководителем отдела «{dept}»!\n"
            f"Теперь доступна команда /menu — сводки и задачи по твоему отделу."
        )
    except Exception: pass

# ── /plan ─────────────────────────────────────────────────────────────────────
async def cmd_plan(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not emp_registered(update.effective_user.id):
        await update.effective_message.reply_text("Сначала /start")
        return ConversationHandler.END
    u = emp_by_id(update.effective_user.id)
    await update.effective_message.reply_text(
        f"📋 {u['full_name']}, напиши план на сегодня:\n\n"
        "Каждый пункт с новой строки — они автоматически станут задачами в трекере."
    )
    return S_PLAN

async def recv_plan(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    u = emp_by_id(tg_id)
    text = update.message.text.strip()
    today = today_str()

    save_plan(tg_id, u["full_name"], text)

    # удаляем старые задачи из плана на сегодня
    ws = tasks_sheet()
    all_rows = safe_records(ws, TH)
    to_delete = [i+2 for i, r in enumerate(all_rows)
                 if str(r["assigned_to_id"]) == str(tg_id)
                 and r.get("created_at","").startswith(today)
                 and r.get("source","") == "plan"]
    for row_i in sorted(to_delete, reverse=True):
        ws.delete_rows(row_i)
    if to_delete:
        invalidate_cache("tasks")

    # создаём задачи из плана
    items = parse_plan_items(text)
    created_ids = []
    for item, item_deadline in items:
        tid = task_create(tg_id, u["full_name"], tg_id, u["full_name"],
                          item, item_deadline, source="plan")
        created_ids.append((tid, item, item_deadline))

    if created_ids:
        task_lines = []
        for tid, title, dl in created_ids:
            dl_mark = f" 📅{fmt_dl(dl)}" if dl != today else ""
            task_lines.append(f"  • <code>{tid}</code> {title}{dl_mark}")
        task_list = "\n".join(task_lines)
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Подтвердить задачи", callback_data=f"confirm_plan_{tg_id}"),
        ]])
        await update.message.reply_text(
            f"📋 План записан! Создано <b>{len(created_ids)} задач</b>:\n\n{task_list}\n\n"
            f"Совет: укажи [ДД.ММ.ГГГГ] в конце строки, если срок не сегодня.\n"
            f"Нажми кнопку чтобы подтвердить — или добавь задачи вручную через /task",
            parse_mode="HTML", reply_markup=keyboard
        )
    else:
        await update.message.reply_text("✅ План записан! Увидимся в 19:00 🌆")
    return ConversationHandler.END

async def cb_confirm_plan(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer("✅ Задачи активированы!")
    await query.edit_message_reply_markup(reply_markup=None)
    tg_id = query.data.split("_")[-1]
    tasks = tasks_today_for_user(int(tg_id))
    await query.message.reply_text(
        f"✅ {len(tasks)} задач активны на сегодня.\n"
        f"/mytasks — посмотреть список\n"
        f"В 19:00 я попрошу статус по каждой."
    )

# ── /report ───────────────────────────────────────────────────────────────────
async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not emp_registered(update.effective_user.id):
        await update.effective_message.reply_text("Сначала /start")
        return ConversationHandler.END
    u = emp_by_id(update.effective_user.id)
    await update.effective_message.reply_text(
        f"📊 {u['full_name']}, напиши итоговый отчёт за день:\n"
        "Что сделано, что переносится, что заблокировано."
    )
    return S_REPORT

async def recv_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    u = emp_by_id(tg_id)
    save_report(tg_id, u["full_name"], update.message.text.strip())
    await update.message.reply_text("✅ Отчёт записан! Хорошего вечера 🌙")
    return ConversationHandler.END

# ── END-OF-DAY FLOW ───────────────────────────────────────────────────────────
def status_keyboard(task_id: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Выполнено",     callback_data=f"eods_done_{task_id}"),
            InlineKeyboardButton("🔄 В работе",      callback_data=f"eods_in_progress_{task_id}"),
        ],
        [
            InlineKeyboardButton("⏸ Приостановлено", callback_data=f"eods_paused_{task_id}"),
            InlineKeyboardButton("⬜ Не начато",      callback_data=f"eods_open_{task_id}"),
        ],
    ])

async def start_eod_flow(bot: Bot, tg_id: int):
    """Запускает опрос статусов задач для сотрудника."""
    tasks = tasks_today_for_user(tg_id)
    if not tasks:
        try:
            await bot.send_message(tg_id,
                "🌆 Рабочий день завершается!\n"
                "Задач на сегодня не было. Напиши отчёт: /report")
        except Exception as e:
            logger.warning(f"eod no tasks {tg_id}: {e}")
        return

    if "eod" not in bot.bot_data:
        bot.bot_data["eod"] = {}
    bot.bot_data["eod"][tg_id] = {
        "tasks": tasks,
        "current_idx": 0,
        "results": []
    }

    task = tasks[0]
    try:
        await bot.send_message(
            tg_id,
            f"🌆 Подводим итоги дня! У тебя <b>{len(tasks)} задач</b>.\n\n"
            f"Задача 1/{len(tasks)}:\n<b>{task['title']}</b>\n\n"
            "Выбери статус:",
            parse_mode="HTML",
            reply_markup=status_keyboard(task["task_id"])
        )
    except Exception as e:
        logger.warning(f"eod start {tg_id}: {e}")

async def cb_eod_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Callback: сотрудник выбрал статус задачи."""
    query = update.callback_query
    await query.answer()
    tg_id = query.from_user.id

    parts = query.data.split("_")
    # format: eods_{status}_{task_id}
    # eods can be: eods_in_progress, eods_done, eods_paused, eods_open
    task_id = parts[-1]
    status = "_".join(parts[1:-1])  # handles in_progress

    task_update_status(task_id, status)

    eod = ctx.bot_data.get("eod", {}).get(tg_id)
    if not eod:
        await query.edit_message_reply_markup(reply_markup=None)
        return

    eod["results"].append({"task_id": task_id, "status": status})

    status_labels = {
        "done": "✅ Выполнено",
        "in_progress": "🔄 В работе",
        "paused": "⏸ Приостановлено",
        "open": "⬜ Не начато"
    }
    label = status_labels.get(status, status)

    await query.edit_message_text(
        query.message.text + f"\n\n<b>Статус: {label}</b>",
        parse_mode="HTML"
    )

    # Запрашиваем комментарий
    ctx.bot_data.setdefault("eod_pending_comment", {})[tg_id] = task_id
    await query.message.reply_text(
        f"💬 Обязательный комментарий по задаче:\n<b>{eod['tasks'][eod['current_idx']]['title']}</b>\n\n"
        "Опиши результат / причину статуса.\n"
        "Если есть документ — прикрепи ссылку в конце через пробел или с новой строки.\n\n"
        "<i>Пример: Выгрузила отчёт, согласовала с Настей\nhttps://docs.google.com/...</i>",
        parse_mode="HTML"
    )

async def recv_eod_comment(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Получаем комментарий к задаче EOD."""
    tg_id = update.effective_user.id
    pending = ctx.bot_data.get("eod_pending_comment", {})
    task_id = pending.get(tg_id)
    if not task_id:
        return  # не в этом потоке

    text = update.message.text.strip()

    # ищем ссылку в тексте
    url_pattern = r'https?://\S+'
    links = re.findall(url_pattern, text)
    link = links[0] if links else ""
    comment = re.sub(url_pattern, "", text).strip()

    task_update_comment(task_id, comment, link)
    del pending[tg_id]

    eod = ctx.bot_data.get("eod", {}).get(tg_id)
    if not eod:
        return

    eod["current_idx"] += 1
    idx = eod["current_idx"]

    if idx < len(eod["tasks"]):
        # следующая задача
        task = eod["tasks"][idx]
        await update.message.reply_text(
            f"Задача {idx+1}/{len(eod['tasks'])}:\n<b>{task['title']}</b>\n\nВыбери статус:",
            parse_mode="HTML",
            reply_markup=status_keyboard(task["task_id"])
        )
    else:
        # все задачи обработаны
        await ask_extra_tasks(update, ctx, tg_id)

async def ask_extra_tasks(update, ctx, tg_id):
    """Спрашиваем были ли другие задачи."""
    keyboard = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Да, были", callback_data="eod_extra_yes"),
        InlineKeyboardButton("❌ Нет", callback_data="eod_extra_no"),
    ]])
    await update.message.reply_text(
        "🎉 Все задачи из плана отмечены!\n\n"
        "Были ли сегодня <b>другие задачи</b>, не из плана?",
        parse_mode="HTML",
        reply_markup=keyboard
    )

async def cb_eod_extra_yes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_reply_markup(reply_markup=None)
    ctx.bot_data.setdefault("eod_extra_pending", set()).add(query.from_user.id)
    await query.message.reply_text(
        "📝 Напиши что ещё сделал сегодня (каждое дело с новой строки):"
    )

async def cb_eod_extra_no(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_reply_markup(reply_markup=None)
    tg_id = query.from_user.id
    await finish_eod(query.message, ctx, tg_id)

async def recv_eod_extra(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Получаем текст доп. задач."""
    tg_id = update.effective_user.id
    extra_set = ctx.bot_data.get("eod_extra_pending", set())
    if tg_id not in extra_set:
        return

    extra_set.discard(tg_id)
    text = update.message.text.strip()
    items = parse_plan_items(text)
    u = emp_by_id(tg_id)

    for item, item_deadline in items:
        tid = task_create(tg_id, u["full_name"], tg_id, u["full_name"],
                          item, item_deadline, source="extra")
        task_update_status(tid, "done")
        task_update_comment(tid, "Выполнено вне плана")

    await update.message.reply_text(f"✅ Добавлено ещё {len(items)} выполненных задач.")
    await finish_eod(update.message, ctx, tg_id)

async def finish_eod(message, ctx, tg_id):
    """Завершаем EOD — просим написать отчёт."""
    eod = ctx.bot_data.get("eod", {}).pop(tg_id, {})
    results = eod.get("results", [])
    done_c = sum(1 for r in results if r["status"] == "done")
    total_c = len(results)

    streak_line = ""
    if total_c > 0 and done_c == total_c:
        streak = calc_streak(tg_id)
        if streak >= 2:
            streak_line = f"\n🔥 {streak} дней подряд все задачи закрыты!"
        elif streak == 1:
            streak_line = "\n🌱 Отличное начало серии — продолжай завтра!"

    await message.reply_text(
        f"✅ День закрыт! Выполнено {done_c} из {total_c} задач.{streak_line}\n\n"
        "Теперь напиши итоговый отчёт за день — /report\n"
        "Хорошего вечера! 🌙"
    )

# ── /eod — запуск вручную ────────────────────────────────────────────────────
async def cmd_eod(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not emp_registered(update.effective_user.id):
        await update.effective_message.reply_text("Сначала /start"); return
    await start_eod_flow(ctx.bot, update.effective_user.id)

# ── /task ─────────────────────────────────────────────────────────────────────
def parse_task(text):
    m = re.match(r"@(\S+)\s+(.+?)\s*\|\s*(\d{2}\.\d{2}\.\d{4})", text)
    if not m: return None
    username, title, dl = m.groups()
    p = dl.split(".")
    return username, title.strip(), f"{p[2]}-{p[1]}-{p[0]}"

async def cmd_task(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    Новый флоу: /task без аргументов → выбор отдела кнопками → выбор сотрудника
    по имени кнопками → текстовый ввод названия задачи и срока.
    Старый текстовый формат @username Название | Дата всё ещё поддерживается
    для тех, кто уже привык.
    """
    if not emp_registered(update.effective_user.id):
        await update.effective_message.reply_text("Сначала /start")
        return ConversationHandler.END
    args = " ".join(ctx.args) if ctx.args else ""
    if args and parse_task(args):
        return await do_create_task(update, ctx, parse_task(args))

    # диалоговый режим — выбор отдела
    depts_present = sorted(set(get_dept(e["tg_id"]) for e in emp_employees()))
    buttons = []
    row = []
    for d in depts_present:
        row.append(InlineKeyboardButton(d, callback_data=f"tasknew_dept_{d}"))
        if len(row) == 2:
            buttons.append(row); row = []
    if row: buttons.append(row)
    await update.effective_message.reply_text(
        "📌 Кому поставить задачу? Выбери отдел:",
        reply_markup=InlineKeyboardMarkup(buttons)
    )
    return ConversationHandler.END

async def cb_tasknew_dept(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Выбран отдел — показываем сотрудников этого отдела по именам."""
    query = update.callback_query; await query.answer()
    dept = query.data.replace("tasknew_dept_", "", 1)
    emps = [e for e in emp_employees() if get_dept(e["tg_id"]) == dept]
    if not emps:
        await query.message.reply_text("В этом отделе нет сотрудников."); return
    buttons = []
    row = []
    for e in emps:
        row.append(InlineKeyboardButton(e["full_name"], callback_data=f"tasknew_emp_{e['tg_id']}"))
        if len(row) == 2:
            buttons.append(row); row = []
    if row: buttons.append(row)
    await query.message.reply_text(f"Сотрудники отдела «{dept}»:", reply_markup=InlineKeyboardMarkup(buttons))

async def cb_tasknew_emp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Выбран сотрудник — просим текст задачи и срок."""
    query = update.callback_query; await query.answer()
    assignee_tg_id = query.data.replace("tasknew_emp_", "", 1)
    assignee = emp_by_id(int(assignee_tg_id))
    if not assignee:
        await query.message.reply_text("Сотрудник не найден."); return
    ctx.user_data["tasknew_assignee"] = assignee_tg_id
    await query.message.reply_text(
        f"Задача для <b>{assignee['full_name']}</b>.\n\n"
        f"Напиши название и срок в формате:\n"
        f"<code>Название задачи | ДД.ММ.ГГГГ</code>\n\n"
        f"Если срок не важен — просто название, поставим на сегодня.",
        parse_mode="HTML"
    )
    ctx.bot_data.setdefault("tasknew_pending", {})[query.from_user.id] = assignee_tg_id

async def recv_tasknew_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Получаем текст задачи после выбора сотрудника кнопками."""
    tg_id = update.effective_user.id
    pending = ctx.bot_data.get("tasknew_pending", {})
    assignee_tg_id = pending.get(tg_id)
    if not assignee_tg_id:
        return  # не в этом потоке
    del pending[tg_id]

    text = update.message.text.strip()
    if "|" in text:
        title, date_part = text.split("|", 1)
        title = title.strip()
        m = re.match(r"^(\d{2})\.(\d{2})\.(\d{4})$", date_part.strip())
        if m:
            d, mo, y = m.groups()
            deadline = f"{y}-{mo}-{d}"
        else:
            deadline = today_str()
    else:
        title = text
        deadline = today_str()

    creator = emp_by_id(tg_id)
    assignee = emp_by_id(int(assignee_tg_id))
    tid = task_create(tg_id, creator["full_name"],
                      int(assignee_tg_id), assignee["full_name"], title, deadline)
    dl_fmt = fmt_dl(deadline)
    try:
        await ctx.bot.send_message(
            int(assignee_tg_id),
            f"📌 Тебе поставлена задача!\n\n<b>{title}</b>\n"
            f"От: {creator['full_name']}\nСрок: {dl_fmt}\n"
            f"ID: <code>{tid}</code>\n\n/done {tid} — отметить выполненной",
            parse_mode="HTML"
        )
    except Exception: pass
    await update.message.reply_text(
        f"✅ Задача создана!\n<b>{title}</b>\nИсполнитель: {assignee['full_name']}\n"
        f"Срок: {dl_fmt}\nID: <code>{tid}</code>", parse_mode="HTML"
    )

async def recv_task(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    parsed = parse_task(update.message.text.strip())
    if not parsed:
        await update.message.reply_text(
            "Не понял формат:\n<code>@username Название | ДД.ММ.ГГГГ</code>",
            parse_mode="HTML"
        )
        return S_TASK
    return await do_create_task(update, ctx, parsed)

async def do_create_task(update, ctx, parsed):
    username, title, deadline = parsed
    creator = emp_by_id(update.effective_user.id)
    assignee = next((e for e in emp_all()
                     if e["username"].lstrip("@").lower() == username.lower()), None)
    if not assignee:
        await update.effective_message.reply_text(f"❌ @{username} не найден.")
        return ConversationHandler.END
    tid = task_create(update.effective_user.id, creator["full_name"],
                      assignee["tg_id"], assignee["full_name"], title, deadline)
    dl_fmt = deadline[8:]+"."+deadline[5:7]+"."+deadline[:4]
    try:
        await ctx.bot.send_message(
            int(assignee["tg_id"]),
            f"📌 Тебе поставлена задача!\n\n<b>{title}</b>\n"
            f"От: {creator['full_name']}\nСрок: {dl_fmt}\n"
            f"ID: <code>{tid}</code>\n\n/done {tid} — отметить выполненной",
            parse_mode="HTML"
        )
    except Exception: pass
    await update.effective_message.reply_text(
        f"✅ Задача создана!\n<b>{title}</b>\nИсполнитель: {assignee['full_name']}\n"
        f"Срок: {dl_fmt}\nID: <code>{tid}</code>", parse_mode="HTML"
    )
    return ConversationHandler.END

# ── /done /status /mytasks ────────────────────────────────────────────────────
async def cmd_done(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ctx.args:
        await update.effective_message.reply_text("Укажи ID: /done ABCD1234"); return
    tid = ctx.args[0].upper()
    comment = " ".join(ctx.args[1:])
    task = task_by_id(tid)
    if not task:
        await update.effective_message.reply_text(f"❌ Задача {tid} не найдена."); return
    tg_id = update.effective_user.id
    if str(task["assigned_to_id"]) != str(tg_id) and not emp_is_admin(tg_id):
        await update.effective_message.reply_text("❌ Это не твоя задача."); return
    task_update_status(tid, "done")
    if comment:
        task_update_comment(tid, comment)
    try:
        await ctx.bot.send_message(
            int(task["created_by_id"]),
            f"✅ Задача выполнена!\n<b>{task['title']}</b>\n"
            f"Выполнил: {task['assigned_to_name']}\nID: <code>{tid}</code>"
            + (f"\n{comment}" if comment else ""),
            parse_mode="HTML"
        )
    except Exception: pass
    await update.effective_message.reply_text(
        f"✅ Задача <code>{tid}</code> выполнена!", parse_mode="HTML"
    )

async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ctx.args:
        await update.effective_message.reply_text("Укажи ID: /status ABCD1234"); return
    task = task_by_id(ctx.args[0].upper())
    if not task:
        await update.effective_message.reply_text("❌ Не найдена."); return
    sl = {"open":"⬜ Не начато","in_progress":"🔄 В работе",
          "done":"✅ Выполнена","overdue":"🔴 Просрочена","paused":"⏸ Приостановлено"}
    dl = task["deadline"]
    dl_fmt = dl[8:]+"."+dl[5:7]+"."+dl[:4] if dl else "—"
    link_line = f"\n🔗 {task['result_link']}" if task.get("result_link") else ""
    comment_line = f"\n💬 {task['comment']}" if task.get("comment") else ""
    tag_line = f"\n🏷 {task['channel']}" if task.get("channel") else ""
    await update.effective_message.reply_text(
        f"📌 <b>{task['title']}</b>\n{sl.get(task['status'],task['status'])}\n"
        f"Исполнитель: {task['assigned_to_name']}\nПостановщик: {task['created_by_name']}\n"
        f"Срок: {dl_fmt}{tag_line}{comment_line}{link_line}\nID: <code>{task['task_id']}</code>",
        parse_mode="HTML"
    )

async def cmd_edit(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    /edit ID title Новое название
    /edit ID deadline ДД.ММ.ГГГГ
    Изменить может постановщик, исполнитель или администратор.
    """
    if len(ctx.args) < 3:
        await update.effective_message.reply_text(
            "Формат:\n"
            "<code>/edit ID title Новое название</code>\n"
            "<code>/edit ID deadline ДД.ММ.ГГГГ</code>",
            parse_mode="HTML"
        )
        return
    tid = ctx.args[0].upper()
    field = ctx.args[1].lower()
    value = " ".join(ctx.args[2:])

    task = task_by_id(tid)
    if not task:
        await update.effective_message.reply_text(f"❌ Задача {tid} не найдена."); return

    tg_id = update.effective_user.id
    allowed = (str(task["assigned_to_id"]) == str(tg_id)
               or str(task["created_by_id"]) == str(tg_id)
               or emp_is_admin(tg_id))
    if not allowed:
        await update.effective_message.reply_text("❌ Можно редактировать только свои задачи."); return

    if field == "title":
        task_update_title(tid, value)
        await update.effective_message.reply_text(f"✅ Название обновлено:\n<b>{value}</b>", parse_mode="HTML")
    elif field == "deadline":
        m = re.match(r"^(\d{2})\.(\d{2})\.(\d{4})$", value.strip())
        if not m:
            await update.effective_message.reply_text("❌ Формат даты: ДД.ММ.ГГГГ"); return
        d, mo, y = m.groups()
        new_deadline = f"{y}-{mo}-{d}"
        task_update_deadline(tid, new_deadline)
        await update.effective_message.reply_text(f"✅ Срок обновлён: {value}")
    else:
        await update.effective_message.reply_text("Поле должно быть title или deadline.")

async def cmd_tag(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    /tag — диалоговый выбор: своя активная задача → канал, кнопками.
    /tag ID Канал — прямой текстовый вариант для опытных пользователей.
    """
    tg_id = update.effective_user.id
    if not emp_registered(tg_id):
        await update.effective_message.reply_text("Сначала /start"); return

    if ctx.args and len(ctx.args) >= 2:
        tid = ctx.args[0].upper()
        channel = " ".join(ctx.args[1:])
        if channel not in CHANNEL_LIST:
            await update.effective_message.reply_text(
                "Канал должен быть одним из: " + ", ".join(CHANNEL_LIST)
            )
            return
        task = task_by_id(tid)
        if not task:
            await update.effective_message.reply_text(f"❌ Задача {tid} не найдена."); return
        task_update_channel(tid, channel)
        await update.effective_message.reply_text(f"✅ Тег «{channel}» проставлен для задачи {tid}.")
        return

    my_tasks = tasks_for_user(tg_id)
    if not my_tasks:
        await update.effective_message.reply_text("У тебя нет активных задач для тегирования."); return

    buttons = []
    for t in my_tasks[:20]:
        label = t["title"][:35] + ("…" if len(t["title"]) > 35 else "")
        buttons.append([InlineKeyboardButton(label, callback_data=f"tagtask_{t['task_id']}")])
    await update.effective_message.reply_text(
        "Выбери задачу для тегирования каналом:",
        reply_markup=InlineKeyboardMarkup(buttons)
    )

async def cb_tagtask_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Выбрана задача — показываем список каналов."""
    query = update.callback_query; await query.answer()
    tid = query.data.replace("tagtask_", "", 1)
    buttons = []
    row = []
    for ch in CHANNEL_LIST:
        row.append(InlineKeyboardButton(ch, callback_data=f"tagchannel_{tid}_{ch}"))
        if len(row) == 2:
            buttons.append(row); row = []
    if row: buttons.append(row)
    await query.message.reply_text("Выбери канал:", reply_markup=InlineKeyboardMarkup(buttons))

async def cb_tagchannel_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Выбран канал — сохраняем тег."""
    query = update.callback_query; await query.answer()
    rest = query.data.replace("tagchannel_", "", 1)
    tid, channel = rest.split("_", 1)
    task_update_channel(tid, channel)
    task = task_by_id(tid)
    title = task["title"] if task else tid
    await query.edit_message_text(f"✅ Тег «{channel}» проставлен:\n{title}")

def calc_streak(tg_id: int) -> int:
    """Считает, сколько последних дней подряд у сотрудника ВСЕ задачи плана
    были закрыты статусом done (без 'paused'/'open' в конце дня)."""
    ws = tasks_sheet()
    all_t = safe_records(ws, TH)
    my_tasks = [t for t in all_t if str(t["assigned_to_id"]) == str(tg_id) and t.get("source") == "plan"]
    if not my_tasks:
        return 0
    by_date: dict = {}
    for t in my_tasks:
        d = t.get("deadline", "")
        if d:
            by_date.setdefault(d, []).append(t)

    streak = 0
    day = date.today()
    # сегодняшний день не считаем, если он ещё не закрыт — начинаем со вчера
    while True:
        d_str = day.strftime("%Y-%m-%d")
        day_tasks = by_date.get(d_str)
        if day_tasks is None:
            # нет задач в этот день — пропускаем день, не прерывая стрик (выходной/отсутствие плана)
            day -= timedelta(days=1)
            if (date.today() - day).days > 60:
                break
            continue
        if all(t["status"] == "done" for t in day_tasks):
            streak += 1
            day -= timedelta(days=1)
        else:
            break
    return streak

async def cmd_mytasks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    tasks = tasks_for_user(tg_id)
    streak = calc_streak(tg_id)
    streak_line = f"🔥 Подряд закрытых дней: {streak}\n\n" if streak >= 2 else ""
    if not tasks:
        await update.effective_message.reply_text(f"{streak_line}✅ Нет активных задач!", parse_mode="HTML")
        return
    sl = {"open":"⬜","in_progress":"🔄","paused":"⏸","overdue":"🔴"}
    lines = [f"{streak_line}📋 <b>Твои задачи:</b>\n"] if streak_line else ["📋 <b>Твои задачи:</b>\n"]
    for t in tasks:
        dl = t["deadline"]
        dl_fmt = dl[8:]+"."+dl[5:7] if dl else "—"
        src = " <i>(план)</i>" if t.get("source") == "plan" else ""
        tag = f" 🏷{t['channel']}" if t.get("channel") else ""
        lines.append(f"{sl.get(t['status'],'⚪')} <code>{t['task_id']}</code> {t['title']} — до {dl_fmt}{src}{tag}")
    await update.effective_message.reply_text("\n".join(lines), parse_mode="HTML")

# ── ADMIN KEYBOARD ────────────────────────────────────────────────────────────
def build_admin_keyboard():
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📋 Задачи сегодня",   callback_data="tasks_today"),
            InlineKeyboardButton("📋 Все активные",     callback_data="show_all_tasks"),
        ],
        [
            InlineKeyboardButton("✅ Закрытые задачи",  callback_data="closed_period_pick"),
            InlineKeyboardButton("📊 По отделам",       callback_data="summary_depts"),
        ],
        [
            InlineKeyboardButton("👤 По сотруднику",    callback_data="summary_person_list"),
        ],
        [
            InlineKeyboardButton("📅 За неделю",        callback_data="period_week"),
            InlineKeyboardButton("📅 За месяц",         callback_data="period_month"),
        ],
        [
            InlineKeyboardButton("📥 Экспорт отдела",   callback_data="export_dept_pick"),
            InlineKeyboardButton("📈 Динамика",         callback_data="dynamics_dept"),
        ],
        [
            InlineKeyboardButton("🔧 Восстановить задачи", callback_data="recover_period_pick"),
            InlineKeyboardButton("📨 Запросить статусы",   callback_data="checkstatuses_now"),
        ],
    ])

def fmt_dl(deadline: str) -> str:
    if not deadline: return "—"
    return deadline[8:]+"."+deadline[5:7]+"."+deadline[:4]

def is_admin_check(query) -> bool:
    return emp_has_management_access(query.from_user.id)

def dept_filter_for(query) -> str:
    """Пустая строка для admin (видит всё), иначе название отдела."""
    return emp_managed_dept(query.from_user.id)

def tasks_for_date(target_date: str) -> list:
    return [t for t in tasks_all() if t.get("deadline","") == target_date and t["status"] != "done"]

# ── ADMIN COMMANDS ────────────────────────────────────────────────────────────
async def cmd_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    if not emp_has_management_access(tg_id):
        await update.effective_message.reply_text("⛔ Только для руководителей.")
        return
    dept_filter = emp_managed_dept(tg_id)  # "" для admin, название отдела для dept_head
    today = date.today().strftime("%d.%m.%Y")
    employees = emp_employees()
    if dept_filter:
        employees = [e for e in employees if get_dept(e["tg_id"]) == dept_filter]

    plan_ids   = {str(p["tg_id"]) for p in plans_today()}
    report_ids = {str(r["tg_id"]) for r in reports_today()}
    over       = tasks_overdue()
    open_t     = tasks_open()
    if dept_filter:
        over   = [t for t in over   if get_dept(t["assigned_to_id"]) == dept_filter]
        open_t = [t for t in open_t if get_dept(t["assigned_to_id"]) == dept_filter]
        my_plan_ids   = {str(e["tg_id"]) for e in employees} & plan_ids
        my_report_ids = {str(e["tg_id"]) for e in employees} & report_ids
    else:
        my_plan_ids, my_report_ids = plan_ids, report_ids

    dept_label = f" — {dept_filter}" if dept_filter else ""
    text = (
        f"🎛 <b>Панель управления{dept_label} — {today}</b>\n\n"
        f"📋 Планов: {len(my_plan_ids)}/{len(employees)}\n"
        f"📝 Отчётов: {len(my_report_ids)}/{len(employees)}\n"
        f"✅ Задач активных: {len(open_t)}\n"
        + (f"⚠️ Просроченных: {len(over)}\n" if over else "")
    )
    await update.effective_message.reply_text(
        text, parse_mode="HTML", reply_markup=build_admin_keyboard()
    )

async def cmd_fixsheets(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Аварийная команда: чинит заголовки во всех листах."""
    if not emp_is_admin(update.effective_user.id):
        await update.effective_message.reply_text("⛔ Только для администраторов."); return
    try:
        ensure_headers(emp_sheet(), EMP_H)
        ensure_headers(plans_sheet(), PH)
        ensure_headers(reports_sheet(), RH)
        ensure_headers(tasks_sheet(), TH)
        await update.effective_message.reply_text("✅ Заголовки во всех листах исправлены.")
    except Exception as e:
        await update.effective_message.reply_text(f"❌ Ошибка: {e}")


async def cmd_makeadmin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    if not emp_registered(tg_id):
        await update.effective_message.reply_text("Сначала /start"); return
    admins = emp_admins()
    if admins and not emp_is_admin(tg_id):
        await update.effective_message.reply_text("⛔ Уже есть администраторы."); return
    emp_set_admin(tg_id)
    await update.effective_message.reply_text("✅ Ты теперь администратор!")

async def cmd_setadmin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not emp_is_admin(update.effective_user.id):
        await update.effective_message.reply_text("⛔ Только для руководителей."); return
    if not ctx.args:
        await update.effective_message.reply_text("Укажи: /setadmin @username"); return
    username = ctx.args[0].lstrip("@")
    for e in emp_all():
        if e["username"].lstrip("@").lower() == username.lower():
            emp_set_admin(int(e["tg_id"]))
            await update.effective_message.reply_text(f"✅ {e['full_name']} теперь администратор.")
            return
    await update.effective_message.reply_text(f"❌ @{username} не найден.")

async def cmd_setdepthead(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    /setdepthead — диалоговый выбор: отдел → сотрудник из активных, кнопками.
    /setdepthead @username — старый текстовый вариант для тех, кто уже привык.
    """
    if not emp_is_admin(update.effective_user.id):
        await update.effective_message.reply_text("⛔ Только для администраторов."); return

    if ctx.args:
        username = ctx.args[0].lstrip("@")
        for e in emp_all():
            if e["username"].lstrip("@").lower() == username.lower():
                emp_set_role(int(e["tg_id"]), "dept_head")
                dept = get_dept(e["tg_id"])
                await update.effective_message.reply_text(
                    f"✅ {e['full_name']} теперь руководитель подразделения «{dept}».\n"
                    f"Видит планы/отчёты/задачи только своего отдела через /menu"
                )
                return
        await update.effective_message.reply_text(f"❌ @{username} не найден.")
        return

    # диалоговый режим — сначала отдел
    depts_present = sorted(set(get_dept(e["tg_id"]) for e in emp_employees()))
    if not depts_present:
        await update.effective_message.reply_text("Нет зарегистрированных сотрудников."); return
    buttons = []
    row = []
    for d in depts_present:
        row.append(InlineKeyboardButton(d, callback_data=f"sethead_dept_{d}"))
        if len(row) == 2:
            buttons.append(row); row = []
    if row: buttons.append(row)
    await update.effective_message.reply_text(
        "⭐ Назначение руководителя отдела.\nВыбери отдел:",
        reply_markup=InlineKeyboardMarkup(buttons)
    )

async def cb_sethead_dept(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Выбран отдел — показываем активных сотрудников этого отдела."""
    query = update.callback_query; await query.answer()
    if not emp_is_admin(query.from_user.id):
        await query.answer("⛔ Только для администраторов.", show_alert=True); return
    dept = query.data.replace("sethead_dept_", "", 1)
    emps = [e for e in emp_employees() if get_dept(e["tg_id"]) == dept]
    if not emps:
        await query.message.reply_text(f"В отделе «{dept}» нет сотрудников."); return
    buttons = []
    row = []
    for e in emps:
        role_mark = " ⭐" if e["role"] == "dept_head" else ""
        row.append(InlineKeyboardButton(e["full_name"]+role_mark, callback_data=f"sethead_emp_{e['tg_id']}"))
        if len(row) == 2:
            buttons.append(row); row = []
    if row: buttons.append(row)
    await query.message.reply_text(
        f"Сотрудники отдела «{dept}»:\n(⭐ — уже руководитель)",
        reply_markup=InlineKeyboardMarkup(buttons)
    )

async def cb_sethead_emp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Выбран сотрудник — назначаем руководителем отдела."""
    query = update.callback_query; await query.answer()
    if not emp_is_admin(query.from_user.id):
        await query.answer("⛔ Только для администраторов.", show_alert=True); return
    target_tg_id = query.data.replace("sethead_emp_", "", 1)
    emp_set_role(int(target_tg_id), "dept_head")
    emp = emp_by_id(int(target_tg_id))
    dept = get_dept(int(target_tg_id))
    await query.edit_message_text(
        f"✅ {emp['full_name']} теперь руководитель подразделения «{dept}».\n"
        f"Видит планы/отчёты/задачи только своего отдела через /menu"
    )
    try:
        await ctx.bot.send_message(
            int(target_tg_id),
            f"⭐ Тебя назначили руководителем отдела «{dept}»!\n"
            f"Теперь доступна команда /menu — сводки и задачи по твоему отделу."
        )
    except Exception: pass

async def cmd_setdept(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/setdept @username — открыть кнопки выбора отдела для сотрудника."""
    if not emp_is_admin(update.effective_user.id):
        await update.effective_message.reply_text("⛔ Только для администраторов."); return
    if not ctx.args:
        await update.effective_message.reply_text("Укажи: /setdept @username"); return
    username = ctx.args[0].lstrip("@")
    for e in emp_all():
        if e["username"].lstrip("@").lower() == username.lower():
            await update.effective_message.reply_text(
                f"Выбери отдел для {e['full_name']}:",
                reply_markup=dept_keyboard("admset_dept_", str(e["tg_id"]))
            )
            return
    await update.effective_message.reply_text(f"❌ @{username} не найден.")

async def cmd_fixname(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/fixname @username Имя Фамилия — вручную исправить порядок имени/фамилии."""
    if not emp_is_admin(update.effective_user.id):
        await update.effective_message.reply_text("⛔ Только для администраторов."); return
    if len(ctx.args) < 2:
        await update.effective_message.reply_text("Укажи: /fixname @username Имя Фамилия"); return
    username = ctx.args[0].lstrip("@")
    new_name = " ".join(ctx.args[1:])
    for e in emp_all():
        if e["username"].lstrip("@").lower() == username.lower():
            emp_set_full_name(int(e["tg_id"]), new_name)
            await update.effective_message.reply_text(f"✅ Имя обновлено: {new_name}")
            return
    await update.effective_message.reply_text(f"❌ @{username} не найден.")

async def cmd_fixallnames(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/fixallnames — прогоняет нормализацию имён по всем сотрудникам сразу."""
    if not emp_is_admin(update.effective_user.id):
        await update.effective_message.reply_text("⛔ Только для администраторов."); return
    changed = []
    for e in emp_all():
        old_name = e["full_name"]
        new_name = normalize_full_name(old_name)
        if new_name != old_name:
            emp_set_full_name(int(e["tg_id"]), new_name)
            changed.append(f"{old_name} → {new_name}")
    if changed:
        await update.effective_message.reply_text(
            "✅ Имена нормализованы:\n" + "\n".join(changed)
        )
    else:
        await update.effective_message.reply_text("Все имена уже в формате «Имя Фамилия».")

async def cmd_team(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    if not emp_has_management_access(tg_id):
        await update.effective_message.reply_text("⛔ Только для руководителей."); return
    dept_filter = emp_managed_dept(tg_id)
    employees = emp_employees()
    if dept_filter:
        employees = [e for e in employees if get_dept(e["tg_id"]) == dept_filter]
    plan_ids  = {str(p["tg_id"]) for p in plans_today()}
    report_ids = {str(r["tg_id"]) for r in reports_today()}
    dept_label = f" — {dept_filter}" if dept_filter else ""
    lines = [f"📊 <b>Команда{dept_label} {datetime.now().strftime('%d.%m.%Y')}:</b>\n"]
    for e in employees:
        tid = str(e["tg_id"])
        lines.append(
            f"{'✅' if tid in report_ids else '❌'} <b>{e['full_name']}</b>  "
            f"план {'✅' if tid in plan_ids else '❌'}  отчёт {'✅' if tid in report_ids else '—'}"
        )
    await update.effective_message.reply_text("\n".join(lines), parse_mode="HTML")

async def cmd_tasks_all(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    if not emp_has_management_access(tg_id):
        await update.effective_message.reply_text("⛔ Только для руководителей."); return
    dept_filter = emp_managed_dept(tg_id)
    open_t = tasks_open()
    over_t = tasks_overdue()
    if dept_filter:
        open_t = [t for t in open_t if get_dept(t["assigned_to_id"]) == dept_filter]
        over_t = [t for t in over_t if get_dept(t["assigned_to_id"]) == dept_filter]
    dept_label = f" — {dept_filter}" if dept_filter else ""
    lines = [f"📋 <b>Активные задачи{dept_label}:</b>\n"]
    if not open_t:
        lines.append("Нет активных задач.")
    by_dept: dict = {}
    for t in open_t:
        dept = get_dept(t["assigned_to_id"])
        by_dept.setdefault(dept, []).append(t)
    for dept, dt_list in sorted(by_dept.items()):
        lines.append(f"<b>— {dept} —</b>")
        for t in dt_list:
            icon = "🔴" if t in over_t else "🔵"
            src = " <i>(план)</i>" if t.get("source") == "plan" else ""
            lines.append(f"  {icon} <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']}{src}")
        lines.append("")
    await update.effective_message.reply_text("\n".join(lines), parse_mode="HTML")

async def cmd_checkstatuses(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Запросить статусы по текущим задачам у всей команды (или своего отдела) прямо сейчас."""
    tg_id = update.effective_user.id
    if not emp_has_management_access(tg_id):
        await update.effective_message.reply_text("⛔ Только для руководителей."); return
    dept_filter = emp_managed_dept(tg_id)
    employees = emp_employees()
    if dept_filter:
        employees = [e for e in employees if get_dept(e["tg_id"]) == dept_filter]

    if not employees:
        await update.effective_message.reply_text("Нет сотрудников для запроса."); return

    sent = 0
    for e in employees:
        try:
            await start_eod_flow(ctx.bot, int(e["tg_id"]))
            sent += 1
        except Exception as ex:
            logger.warning(f"checkstatuses {e['tg_id']}: {ex}")

    dept_label = f" отдела «{dept_filter}»" if dept_filter else " команды"
    await update.effective_message.reply_text(
        f"📨 Запрос статусов отправлен {sent} сотрудникам{dept_label}."
    )

def recover_tasks_from_plans(date_from: str, date_to: str, dept_filter: str = ""):
    """
    Сравнивает планы за период с уже существующими задачами (source=plan)
    и досоздаёт недостающие. Возвращает (recovered_count, by_person, affected_dates).
    """
    ws_plans = sheet("plans"); ensure_headers(ws_plans, PH)
    all_plans = [p for p in safe_records(ws_plans, PH)
                 if date_from <= p.get("date","") <= date_to]
    if dept_filter:
        all_plans = [p for p in all_plans if get_dept(p["tg_id"]) == dept_filter]

    all_t = tasks_all()
    # ключ существующей задачи из плана: (tg_id, date, title) — чтобы не дублировать
    existing_keys = set()
    for t in all_t:
        if t.get("source") == "plan":
            key = (str(t["assigned_to_id"]), t.get("deadline",""), t.get("title","").strip())
            existing_keys.add(key)

    recovered = 0
    by_person: dict = {}
    affected_dates = set()

    # берём последнюю запись плана на каждую (сотрудник, дата) — план может быть переписан несколько раз за день
    latest_plan: dict = {}
    for p in all_plans:
        key = (str(p["tg_id"]), p["date"])
        # submitted_at растёт по времени добавления строк, последняя запись с этим ключом — самая свежая
        latest_plan[key] = p

    for (tg_id, plan_date), p in latest_plan.items():
        items = parse_plan_items(p["plan_text"])
        for item, item_deadline in items:
            key = (tg_id, item_deadline, item.strip())
            if key in existing_keys:
                continue
            task_create(int(tg_id), p["full_name"], int(tg_id), p["full_name"],
                       item, item_deadline, source="plan")
            existing_keys.add(key)
            recovered += 1
            by_person[p["full_name"]] = by_person.get(p["full_name"], 0) + 1
            affected_dates.add(item_deadline)

    return recovered, by_person, sorted(affected_dates)

async def cmd_recovertasks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/recovertasks [N] — восстановить задачи из планов за последние N дней (по умолчанию 7)."""
    tg_id = update.effective_user.id
    if not emp_has_management_access(tg_id):
        await update.effective_message.reply_text("⛔ Только для руководителей."); return
    dept_filter = emp_managed_dept(tg_id)

    days_back = 7
    if ctx.args and ctx.args[0].isdigit():
        days_back = int(ctx.args[0])

    date_to = date.today().strftime("%Y-%m-%d")
    date_from = (date.today() - timedelta(days=days_back)).strftime("%Y-%m-%d")

    await update.effective_message.reply_text(
        f"🔧 Сканирую планы с {date_from} по {date_to}..."
    )

    recovered, by_person, affected_dates = recover_tasks_from_plans(date_from, date_to, dept_filter)

    if recovered == 0:
        await update.effective_message.reply_text("✅ Расхождений не найдено — все задачи из планов на месте.")
        return

    lines = [f"🔧 <b>Восстановлено задач: {recovered}</b>\n"]
    for person, cnt in sorted(by_person.items()):
        lines.append(f"  {person}: +{cnt}")
    lines.append(f"\nЗатронутые даты: {', '.join(affected_dates)}")
    await update.effective_message.reply_text("\n".join(lines), parse_mode="HTML")

    # отправляем уведомление всем администраторам (в т.ч. инициатору, если он не получил выше)
    notify = (
        f"🔧 <b>Восстановление задач выполнено</b>\n"
        f"Инициатор: {emp_by_id(tg_id)['full_name']}\n"
        f"Период: {date_from} — {date_to}\n"
        f"Восстановлено задач: {recovered}\n"
    )
    for adm in emp_admins():
        if int(adm["tg_id"]) == tg_id:
            continue
        try:
            await ctx.bot.send_message(int(adm["tg_id"]), notify, parse_mode="HTML")
        except Exception as ex:
            logger.warning(f"recover notify admin {adm['tg_id']}: {ex}")

    # сразу предлагаем актуализировать статусы у всех затронутых
    keyboard = InlineKeyboardMarkup([[
        InlineKeyboardButton("📨 Запросить статусы сейчас", callback_data="checkstatuses_now")
    ]])
    await update.effective_message.reply_text(
        "Хочешь сразу запросить актуальные статусы по восстановленным задачам?",
        reply_markup=keyboard
    )

async def cb_checkstatuses_now(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Кнопка-дублёр /checkstatuses, вызываемая после восстановления задач."""
    query = update.callback_query
    if not emp_has_management_access(query.from_user.id):
        await query.answer("⛔ Только для руководителей.", show_alert=True); return
    await query.answer("Запускаю опрос...")
    dept_filter = emp_managed_dept(query.from_user.id)
    employees = emp_employees()
    if dept_filter:
        employees = [e for e in employees if get_dept(e["tg_id"]) == dept_filter]

    sent = 0
    for e in employees:
        try:
            await start_eod_flow(ctx.bot, int(e["tg_id"]))
            sent += 1
        except Exception as ex:
            logger.warning(f"checkstatuses_now {e['tg_id']}: {ex}")

    dept_label = f" отдела «{dept_filter}»" if dept_filter else " команды"
    await query.message.reply_text(f"📨 Запрос статусов отправлен {sent} сотрудникам{dept_label}.")

# ── ADMIN CALLBACKS ───────────────────────────────────────────────────────────
async def cb_back_main(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query; await query.answer()
    await query.message.reply_text("Выбери действие:", reply_markup=build_admin_keyboard())

async def cb_recover_period_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Выбор периода для восстановления задач из планов."""
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Сегодня",      callback_data="recover_run_1")],
        [InlineKeyboardButton("3 дня",        callback_data="recover_run_3")],
        [InlineKeyboardButton("Неделя",       callback_data="recover_run_7")],
        [InlineKeyboardButton("Месяц",        callback_data="recover_run_30")],
        [InlineKeyboardButton("◀️ Назад",     callback_data="back_main")],
    ])
    await query.message.reply_text(
        "За какой период проверить планы и досоздать пропущенные задачи?",
        reply_markup=keyboard
    )

async def cb_recover_run(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Запускает восстановление задач за выбранный период."""
    query = update.callback_query; await query.answer("Сканирую планы...")
    if not is_admin_check(query): return
    tg_id = query.from_user.id
    dept_filter = dept_filter_for(query)

    days_back = int(query.data.replace("recover_run_", "", 1))
    date_to = date.today().strftime("%Y-%m-%d")
    date_from = (date.today() - timedelta(days=days_back)).strftime("%Y-%m-%d")

    recovered, by_person, affected_dates = recover_tasks_from_plans(date_from, date_to, dept_filter)

    if recovered == 0:
        await query.message.reply_text("✅ Расхождений не найдено — все задачи из планов на месте.")
        return

    lines = [f"🔧 <b>Восстановлено задач: {recovered}</b>\n"]
    for person, cnt in sorted(by_person.items()):
        lines.append(f"  {person}: +{cnt}")
    lines.append(f"\nЗатронутые даты: {', '.join(affected_dates)}")
    keyboard = InlineKeyboardMarkup([[
        InlineKeyboardButton("📨 Запросить статусы сейчас", callback_data="checkstatuses_now")
    ]])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)

    # уведомляем остальных администраторов
    notify = (
        f"🔧 <b>Восстановление задач выполнено</b>\n"
        f"Инициатор: {emp_by_id(tg_id)['full_name']}\n"
        f"Период: {date_from} — {date_to}\n"
        f"Восстановлено задач: {recovered}\n"
    )
    for adm in emp_admins():
        if int(adm["tg_id"]) == tg_id:
            continue
        try:
            await ctx.bot.send_message(int(adm["tg_id"]), notify, parse_mode="HTML")
        except Exception as ex:
            logger.warning(f"recover notify admin {adm['tg_id']}: {ex}")

async def cb_tasks_today(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    today = date.today().strftime("%Y-%m-%d")
    today_fmt = date.today().strftime("%d.%m.%Y")
    tasks = tasks_for_date(today)
    over = tasks_overdue()
    if dept_filter:
        tasks = [t for t in tasks if get_dept(t["assigned_to_id"]) == dept_filter]
        over  = [t for t in over  if get_dept(t["assigned_to_id"]) == dept_filter]
    dept_label = f" — {dept_filter}" if dept_filter else ""
    lines = [f"📋 <b>Задачи на сегодня{dept_label} ({today_fmt}):</b>\n"]
    if not tasks and not over:
        lines.append("✅ Нет задач на сегодня!")
    if over:
        lines.append(f"🔴 <b>Просроченные ({len(over)}):</b>")
        for t in over:
            lines.append(f"  🔴 <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']} (срок {fmt_dl(t['deadline'])})")
        lines.append("")
    if tasks:
        lines.append(f"🔵 <b>На сегодня ({len(tasks)}):</b>")
        for t in tasks:
            src = " <i>(план)</i>" if t.get("source") == "plan" else ""
            sl = {"open":"⬜","in_progress":"🔄","paused":"⏸","done":"✅"}
            icon = sl.get(t["status"],"🔵")
            lines.append(f"  {icon} <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']}{src}")
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔄 Обновить", callback_data="tasks_today")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)

async def cb_show_all_tasks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    open_t = tasks_open(); over_t = tasks_overdue()
    if dept_filter:
        open_t = [t for t in open_t if get_dept(t["assigned_to_id"]) == dept_filter]
        over_t = [t for t in over_t if get_dept(t["assigned_to_id"]) == dept_filter]
    dept_label = f" — {dept_filter}" if dept_filter else ""
    lines = [f"📋 <b>Активные задачи{dept_label} ({len(open_t)}):</b>\n"]
    if not open_t:
        lines.append("✅ Нет активных задач!")
    else:
        by_dept: dict = {}
        for t in open_t:
            dept = get_dept(t["assigned_to_id"])
            by_dept.setdefault(dept, []).append(t)
        for dept, dt_list in sorted(by_dept.items()):
            lines.append(f"<b>— {dept} —</b>")
            for t in dt_list:
                icon = "🔴" if t in over_t else "🔵"
                src = " <i>(план)</i>" if t.get("source") == "plan" else ""
                lines.append(f"  {icon} <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']} — до {fmt_dl(t['deadline'])}{src}")
            lines.append("")
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔄 Обновить", callback_data="show_all_tasks")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)

async def cb_summary_depts(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    employees = emp_employees()
    if dept_filter:
        employees = [e for e in employees if get_dept(e["tg_id"]) == dept_filter]
    plan_ids   = {str(p["tg_id"]) for p in plans_today()}
    report_ids = {str(r["tg_id"]) for r in reports_today()}
    all_t = tasks_all()
    today = today_str(); today_fmt = date.today().strftime("%d.%m.%Y")
    dept_employees: dict = {}
    for e in employees:
        dept = get_dept(e["tg_id"])
        dept_employees.setdefault(dept, []).append(e)
    dept_label = f" — {dept_filter}" if dept_filter else ""
    lines = [f"📊 <b>Сводка по отделам{dept_label} — {today_fmt}</b>\n"]
    for dept, emps in sorted(dept_employees.items()):
        submitted = sum(1 for e in emps if str(e["tg_id"]) in report_ids)
        planned   = sum(1 for e in emps if str(e["tg_id"]) in plan_ids)
        overdue_c = sum(1 for t in all_t
                        if get_dept(t["assigned_to_id"]) == dept
                        and t["status"] != "done" and t.get("deadline","") < today)
        icon = "✅" if submitted == len(emps) else ("🟡" if submitted > 0 else "🔴")
        lines.append(f"{icon} <b>{dept}</b> ({len(emps)} чел.)")
        lines.append(f"  планов {planned}/{len(emps)}  отчётов {submitted}/{len(emps)}"
                     + (f"  ⚠️{overdue_c}" if overdue_c else ""))
        for e in emps:
            tid = str(e["tg_id"])
            lines.append(f"    {e['full_name']}  {'✅' if tid in plan_ids else '❌'}план  {'✅' if tid in report_ids else '—'}отчёт")
        lines.append("")
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("👤 По сотруднику", callback_data="summary_person_list")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)

async def cb_summary_person_list(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    employees = emp_employees()
    if dept_filter:
        employees = [e for e in employees if get_dept(e["tg_id"]) == dept_filter]
    buttons = []
    row = []
    for e in employees:
        row.append(InlineKeyboardButton(e["full_name"].split()[0], callback_data=f"person_{e['tg_id']}"))
        if len(row) == 2:
            buttons.append(row); row = []
    if row: buttons.append(row)
    buttons.append([InlineKeyboardButton("◀️ Назад", callback_data="back_main")])
    await query.message.reply_text("Выбери сотрудника:", reply_markup=InlineKeyboardMarkup(buttons))

async def cb_summary_person(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    tg_id = query.data.split("_", 1)[1]
    emp = emp_by_id(int(tg_id))
    if not emp:
        await query.message.reply_text("Не найден."); return
    if dept_filter and get_dept(tg_id) != dept_filter:
        await query.answer("⛔ Сотрудник не из твоего отдела.", show_alert=True); return
    today = today_str(); today_fmt = date.today().strftime("%d.%m.%Y")
    plan_ids   = {str(p["tg_id"]) for p in plans_today()}
    report_ids = {str(r["tg_id"]) for r in reports_today()}
    all_t = tasks_all()
    emp_tasks = [t for t in all_t if str(t["assigned_to_id"]) == tg_id and t["status"] != "done"]
    done_today = [t for t in all_t if str(t["assigned_to_id"]) == tg_id
                  and t["status"] == "done" and t.get("done_at","").startswith(today)]
    overdue = [t for t in emp_tasks if t.get("deadline","") < today]
    sl = {"open":"⬜","in_progress":"🔄","paused":"⏸","done":"✅"}
    lines = [
        f"👤 <b>{emp['full_name']}</b>  ({get_dept(tg_id)})",
        f"Сегодня {today_fmt}:  план {'✅' if tg_id in plan_ids else '❌'}  отчёт {'✅' if tg_id in report_ids else '❌'}\n",
    ]
    if emp_tasks:
        lines.append(f"<b>Активные ({len(emp_tasks)}):</b>")
        for t in emp_tasks:
            icon = "🔴" if t in overdue else sl.get(t["status"],"🔵")
            src = " <i>(план)</i>" if t.get("source") == "plan" else ""
            comment = f"\n    💬 {t['comment']}" if t.get("comment") else ""
            link = f"\n    🔗 {t['result_link']}" if t.get("result_link") else ""
            lines.append(f"  {icon} {t['title']} — до {fmt_dl(t['deadline'])}{src}{comment}{link}")
    if done_today:
        lines.append(f"\n<b>Выполнено сегодня ({len(done_today)}):</b>")
        for t in done_today:
            lines.append(f"  ✅ {t['title']}")
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("👥 Другой", callback_data="summary_person_list")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)

async def cb_period_week(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    today = date.today()
    week_start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
    week_end   = today.strftime("%Y-%m-%d")
    week_start_fmt = (today - timedelta(days=today.weekday())).strftime("%d.%m")
    week_end_fmt   = today.strftime("%d.%m.%Y")
    employees  = emp_employees()
    if dept_filter:
        employees = [e for e in employees if get_dept(e["tg_id"]) == dept_filter]
    reps       = reports_for_period(week_start, week_end)
    all_t      = tasks_all()
    done_w     = [t for t in all_t if t["status"] == "done"
                  and t.get("done_at","")[:10] >= week_start]
    over       = tasks_overdue()
    if dept_filter:
        done_w = [t for t in done_w if get_dept(t["assigned_to_id"]) == dept_filter]
        over   = [t for t in over   if get_dept(t["assigned_to_id"]) == dept_filter]
    dept_label = f" — {dept_filter}" if dept_filter else ""
    lines = [
        f"📅 <b>Неделя{dept_label} ({week_start_fmt} — {week_end_fmt})</b>\n",
        f"Уникальных отчётов: {len(set(r['tg_id'] for r in reps))} / {len(employees)}",
        f"Задач выполнено: {len(done_w)}",
        f"Просроченных: {len(over)}\n",
        "<b>По сотрудникам:</b>",
    ]
    for e in employees:
        tid = str(e["tg_id"])
        r_c = sum(1 for r in reps if str(r["tg_id"]) == tid)
        d_c = sum(1 for t in done_w if str(t["assigned_to_id"]) == tid)
        o_c = sum(1 for t in over if str(t["assigned_to_id"]) == tid)
        icon = "✅" if r_c >= 4 else ("🟡" if r_c >= 1 else "🔴")
        lines.append(f"{icon} <b>{e['full_name']}</b> ({get_dept(e['tg_id'])}): "
                     f"отчётов {r_c}, задач ✅{d_c}" + (f" ⚠️{o_c}" if o_c else ""))
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("📅 За месяц", callback_data="period_month")],
        [InlineKeyboardButton("◀️ Назад",    callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)

async def cb_period_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    today = date.today()
    month_start = today.replace(day=1).strftime("%Y-%m-%d")
    month_end   = today.strftime("%Y-%m-%d")
    month_label = today.strftime("%B %Y")
    _, last = calendar.monthrange(today.year, today.month)
    wd = sum(1 for d in range(1, today.day+1) if date(today.year, today.month, d).weekday() < 5)
    employees = emp_employees()
    if dept_filter:
        employees = [e for e in employees if get_dept(e["tg_id"]) == dept_filter]
    reps      = reports_for_period(month_start, month_end)
    all_t     = tasks_all()
    done_m    = [t for t in all_t if t["status"] == "done"
                 and t.get("done_at","")[:10] >= month_start]
    over      = tasks_overdue()
    if dept_filter:
        done_m = [t for t in done_m if get_dept(t["assigned_to_id"]) == dept_filter]
        over   = [t for t in over   if get_dept(t["assigned_to_id"]) == dept_filter]
    dept_label = f" — {dept_filter}" if dept_filter else ""
    lines = [
        f"📆 <b>{month_label}{dept_label}</b>\n",
        f"Рабочих дней прошло: {wd}",
        f"Задач выполнено: {len(done_m)}",
        f"Просроченных: {len(over)}\n",
        "<b>По сотрудникам:</b>",
    ]
    for e in employees:
        tid = str(e["tg_id"])
        r_c = sum(1 for r in reps if str(r["tg_id"]) == tid)
        d_c = sum(1 for t in done_m if str(t["assigned_to_id"]) == tid)
        pct = round(r_c / max(wd, 1) * 100)
        icon = "✅" if pct >= 80 else ("🟡" if pct >= 40 else "🔴")
        lines.append(f"{icon} <b>{e['full_name']}</b> ({get_dept(e['tg_id'])}): "
                     f"отчётов {r_c}/{wd} ({pct}%), задач ✅{d_c}")
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("📅 За неделю", callback_data="period_week")],
        [InlineKeyboardButton("◀️ Назад",     callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)

async def cb_closed_period_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Выбор периода для просмотра закрытых (выполненных) задач."""
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Сегодня",   callback_data="closed_today")],
        [InlineKeyboardButton("Неделя",    callback_data="closed_week")],
        [InlineKeyboardButton("Месяц",     callback_data="closed_month")],
        [InlineKeyboardButton("◀️ Назад",  callback_data="back_main")],
    ])
    await query.message.reply_text("За какой период показать закрытые задачи?", reply_markup=keyboard)

async def cb_closed_tasks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Список выполненных задач за выбранный период, с фильтром по отделу."""
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    period = query.data.replace("closed_", "", 1)  # today|week|month
    today = date.today()
    if period == "today":
        start = today.strftime("%Y-%m-%d")
        label = f"сегодня ({today.strftime('%d.%m.%Y')})"
    elif period == "week":
        start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
        label = f"за неделю (с {(today - timedelta(days=today.weekday())).strftime('%d.%m')})"
    else:
        start = today.replace(day=1).strftime("%Y-%m-%d")
        label = f"за {today.strftime('%B %Y')}"

    all_t = tasks_all()
    done = [t for t in all_t if t["status"] == "done" and t.get("done_at","")[:10] >= start]
    if dept_filter:
        done = [t for t in done if get_dept(t["assigned_to_id"]) == dept_filter]
    dept_label = f" — {dept_filter}" if dept_filter else ""

    lines = [f"✅ <b>Закрытые задачи{dept_label}, {label}</b>\n", f"Всего: {len(done)}\n"]
    by_person: dict = {}
    for t in done:
        by_person.setdefault(t["assigned_to_name"], []).append(t)
    for person, tlist in sorted(by_person.items()):
        lines.append(f"<b>{person}</b> ({len(tlist)}):")
        for t in tlist:
            src = " <i>(план)</i>" if t.get("source") == "plan" else ""
            comment = f" — {t['comment']}" if t.get("comment") else ""
            lines.append(f"  ✅ {t['title']}{src}{comment}")
        lines.append("")
    if not done:
        lines.append("Нет закрытых задач за этот период.")

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Сегодня", callback_data="closed_today"),
         InlineKeyboardButton("Неделя",  callback_data="closed_week"),
         InlineKeyboardButton("Месяц",   callback_data="closed_month")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)

async def cb_export_dept_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Выбор периода для экспорта в Excel."""
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Неделя", callback_data="export_week")],
        [InlineKeyboardButton("Месяц",  callback_data="export_month")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("За какой период сформировать Excel-отчёт?", reply_markup=keyboard)

def build_export_workbook(dept_filter: str, period: str):
    """Строит .xlsx с задачами за период, фильтр по отделу опционален."""
    today = date.today()
    if period == "week":
        start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
        label = f"неделя с {(today - timedelta(days=today.weekday())).strftime('%d.%m.%Y')}"
    else:
        start = today.replace(day=1).strftime("%Y-%m-%d")
        label = today.strftime("%B %Y")

    all_t = tasks_all()
    rows = [t for t in all_t if t.get("created_at","")[:10] >= start
            or t.get("done_at","")[:10] >= start]
    if dept_filter:
        rows = [t for t in rows if get_dept(t["assigned_to_id"]) == dept_filter]

    wb = Workbook()
    ws = wb.active
    ws.title = "Задачи"

    headers = ["ID", "Сотрудник", "Отдел", "Название", "Срок", "Статус", "Комментарий", "Ссылка", "Источник"]
    ws.append(headers)
    header_fill = PatternFill(start_color="185FA5", end_color="185FA5", fill_type="solid")
    for col_i, _ in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col_i)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center")

    status_labels = {"open":"Не начато","in_progress":"В работе","paused":"Приостановлено","done":"Выполнено"}
    for t in rows:
        ws.append([
            t["task_id"], t["assigned_to_name"], get_dept(t["assigned_to_id"]),
            t["title"], t.get("deadline",""), status_labels.get(t["status"], t["status"]),
            t.get("comment",""), t.get("result_link",""), t.get("source","manual"),
        ])

    widths = [10, 20, 18, 40, 12, 14, 30, 30, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, label

async def cb_export_period(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Формирует и отправляет xlsx-файл."""
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    period = query.data.replace("export_", "", 1)  # week|month

    buf, label = build_export_workbook(dept_filter, period)
    dept_part = dept_filter.replace(" ", "_") if dept_filter else "all"
    filename = f"report_{dept_part}_{date.today().strftime('%Y%m%d')}.xlsx"

    await query.message.reply_document(
        document=InputFile(buf, filename=filename),
        caption=f"📥 Отчёт за {label}" + (f", отдел «{dept_filter}»" if dept_filter else ", все отделы")
    )

async def cb_dynamics_dept(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Динамика выполнения задач по неделям — текстовый спарклайн за последние 4 недели."""
    query = update.callback_query; await query.answer()
    if not is_admin_check(query): return
    dept_filter = dept_filter_for(query)
    all_t = tasks_all()
    if dept_filter:
        all_t = [t for t in all_t if get_dept(t["assigned_to_id"]) == dept_filter]

    today = date.today()
    weeks = []
    for i in range(3, -1, -1):
        w_start = today - timedelta(days=today.weekday() + i*7)
        w_end = w_start + timedelta(days=6)
        weeks.append((w_start, w_end))

    dept_label = f" — {dept_filter}" if dept_filter else ""
    lines = [f"📈 <b>Динамика выполнения{dept_label}</b>\n"]
    counts = []
    for w_start, w_end in weeks:
        s = w_start.strftime("%Y-%m-%d")
        e = w_end.strftime("%Y-%m-%d")
        done_count = sum(1 for t in all_t if t["status"] == "done"
                         and s <= t.get("done_at","")[:10] <= e)
        counts.append(done_count)
        bar = "█" * min(done_count, 20) + ("" if done_count <= 20 else f" (+{done_count-20})")
        lines.append(f"{w_start.strftime('%d.%m')}–{w_end.strftime('%d.%m')}: {bar} {done_count}")

    if len(counts) >= 2 and counts[-2] > 0:
        delta = counts[-1] - counts[-2]
        trend = "📈 рост" if delta > 0 else ("📉 спад" if delta < 0 else "➡️ без изменений")
        lines.append(f"\n{trend} к прошлой неделе ({'+' if delta>=0 else ''}{delta})")

    keyboard = InlineKeyboardMarkup([[InlineKeyboardButton("◀️ Назад", callback_data="back_main")]])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)

# ── NOTIFICATIONS ─────────────────────────────────────────────────────────────
async def job_request_plans(bot: Bot):
    for e in emp_employees():
        if has_plan_today(int(e["tg_id"])): continue
        try:
            await bot.send_message(int(e["tg_id"]),
                f"☀️ Доброе утро, {e['full_name']}!\nНапиши план на день: /plan\n"
                "Каждый пункт — отдельная задача в трекере.")
        except Exception as ex: logger.warning(f"request_plans: {ex}")

async def job_remind_plans(bot: Bot):
    for e in emp_employees():
        if not has_plan_today(int(e["tg_id"])):
            try:
                await bot.send_message(int(e["tg_id"]),
                    "🔔 Напоминание: напиши план на день /plan")
            except Exception as ex: logger.warning(f"remind_plans: {ex}")

async def job_request_eod(bot: Bot):
    """19:00 — запускаем EOD-опрос для всех сотрудников."""
    for e in emp_employees():
        await start_eod_flow(bot, int(e["tg_id"]))

async def job_remind_reports(bot: Bot):
    for e in emp_employees():
        if not has_report_today(int(e["tg_id"])):
            try:
                await bot.send_message(int(e["tg_id"]),
                    "🔔 Напоминание: напиши итоговый отчёт /report")
            except Exception as ex: logger.warning(f"remind_reports: {ex}")

async def job_ping_deadlines(bot: Bot):
    for t in tasks_due_tomorrow():
        try:
            await bot.send_message(int(t["assigned_to_id"]),
                f"⏰ Завтра срок задачи!\n<b>{t['title']}</b>\n"
                f"ID: <code>{t['task_id']}</code>\n/done {t['task_id']}",
                parse_mode="HTML")
        except Exception as ex: logger.warning(f"ping_tomorrow: {ex}")
    for t in tasks_due_today_list():
        try:
            await bot.send_message(int(t["assigned_to_id"]),
                f"🚨 Срок задачи сегодня!\n<b>{t['title']}</b>\n"
                f"ID: <code>{t['task_id']}</code>",
                parse_mode="HTML")
        except Exception as ex: logger.warning(f"ping_today: {ex}")

    # уведомление руководителю отдела о новых просрочках его сотрудников
    overdue = tasks_overdue()
    if not overdue:
        return
    by_dept: dict = {}
    for t in overdue:
        by_dept.setdefault(get_dept(t["assigned_to_id"]), []).append(t)

    for head in emp_dept_heads():
        head_dept = get_dept(head["tg_id"])
        dept_overdue = by_dept.get(head_dept, [])
        if not dept_overdue:
            continue
        lines = [f"⚠️ <b>Просроченные задачи в отделе «{head_dept}»</b>\n"]
        for t in dept_overdue:
            lines.append(f"🔴 {t['assigned_to_name']}: {t['title']} (срок {fmt_dl(t['deadline'])})")
        try:
            await bot.send_message(int(head["tg_id"]), "\n".join(lines), parse_mode="HTML")
        except Exception as ex:
            logger.warning(f"overdue notify dept_head {head['tg_id']}: {ex}")

async def job_daily_digest(bot: Bot):
    today = datetime.now().strftime("%d.%m.%Y")
    employees = emp_employees()
    p_list = plans_today(); r_list = reports_today()
    plan_ids   = {str(p["tg_id"]) for p in p_list}
    report_ids = {str(r["tg_id"]) for r in r_list}
    all_t = tasks_all()
    over = tasks_overdue()
    done_today = [t for t in all_t if t["status"] == "done"
                  and t.get("done_at","").startswith(datetime.now().strftime("%Y-%m-%d"))]

    # группируем сотрудников по отделам
    dept_employees: dict = {}
    for e in employees:
        dept_employees.setdefault(get_dept(e["tg_id"]), []).append(e)

    # ── Полная сводка в общий чат руководителей ──
    lines = [
        f"📊 <b>Сводка команды — {today}</b>",
        f"Планов: {len(plan_ids)}/{len(employees)}   Отчётов: {len(report_ids)}/{len(employees)}   "
        f"Закрыто задач: {len(done_today)}\n",
    ]
    for dept, emps in sorted(dept_employees.items()):
        submitted = sum(1 for e in emps if str(e["tg_id"]) in report_ids)
        icon = "✅" if submitted == len(emps) else ("🟡" if submitted > 0 else "🔴")
        dept_done = sum(1 for t in done_today if get_dept(t["assigned_to_id"]) == dept)
        dept_over = sum(1 for t in over if get_dept(t["assigned_to_id"]) == dept)
        lines.append(f"{icon} <b>{dept}</b>: отчётов {submitted}/{len(emps)}, закрыто {dept_done}"
                     + (f", ⚠️ просрочено {dept_over}" if dept_over else ""))
    if over:
        lines.append(f"\n⚠️ <b>Всего просроченных: {len(over)}</b>")
        for t in over[:3]:
            lines.append(f"  🔴 {t['assigned_to_name']}: {t['title']}")
        if len(over) > 3: lines.append(f"  ...и ещё {len(over)-3}")
    try:
        await bot.send_message(ADMIN_CHAT_ID, "\n".join(lines),
                               parse_mode="HTML", reply_markup=build_admin_keyboard())
    except Exception as ex: logger.warning(f"daily_digest: {ex}")

    # ── Персональная сводка по отделу для каждого dept_head ──
    for head in emp_dept_heads():
        head_dept = get_dept(head["tg_id"])
        dept_emps = dept_employees.get(head_dept, [])
        if not dept_emps:
            continue
        no_plan   = [e["full_name"] for e in dept_emps if str(e["tg_id"]) not in plan_ids]
        no_report = [e["full_name"] for e in dept_emps if str(e["tg_id"]) not in report_ids]
        dept_done = [t for t in done_today if get_dept(t["assigned_to_id"]) == head_dept]
        dept_over = [t for t in over if get_dept(t["assigned_to_id"]) == head_dept]
        h_lines = [
            f"📊 <b>Сводка отдела «{head_dept}» — {today}</b>\n",
            f"Закрыто задач: {len(dept_done)}",
        ]
        if no_plan:   h_lines.append(f"❌ <b>Нет плана:</b> {', '.join(no_plan)}")
        if no_report: h_lines.append(f"❌ <b>Нет отчёта:</b> {', '.join(no_report)}")
        if dept_over:
            h_lines.append(f"\n⚠️ <b>Просрочено в отделе: {len(dept_over)}</b>")
            for t in dept_over[:5]:
                h_lines.append(f"  🔴 {t['assigned_to_name']}: {t['title']}")
        try:
            await bot.send_message(int(head["tg_id"]), "\n".join(h_lines),
                                   parse_mode="HTML", reply_markup=build_admin_keyboard())
        except Exception as ex:
            logger.warning(f"daily_digest dept_head {head['tg_id']}: {ex}")

async def job_weekly_audit(bot: Bot):
    now = datetime.now()
    ws = now.strftime("%Y-W%W")
    w0 = (now - timedelta(days=now.weekday())).strftime("%d.%m")
    w1 = now.strftime("%d.%m.%Y")
    employees = emp_employees()
    plans_w   = records_for_week("plans", PH)
    reports_w = records_for_week("reports", RH)
    all_t     = tasks_all()
    done_w    = [t for t in all_t if t["status"]=="done" and _is_this_week(t.get("done_at",""), ws)]
    over      = tasks_overdue()
    lines = [f"📅 <b>Еженедельный аудит ({w0} — {w1})</b>\n",
             f"Отчётов: {len(set(r['tg_id'] for r in reports_w))}/{len(employees)}",
             f"Задач выполнено: {len(done_w)}", f"Просроченных: {len(over)}\n",
             "<b>По сотрудникам:</b>"]
    for e in employees:
        tid = str(e["tg_id"])
        r_c = sum(1 for r in reports_w if str(r["tg_id"]) == tid)
        d_c = sum(1 for t in done_w if str(t["assigned_to_id"]) == tid)
        o_c = sum(1 for t in over if str(t["assigned_to_id"]) == tid)
        icon = "✅" if r_c >= 4 else ("🟡" if r_c >= 1 else "🔴")
        lines.append(f"{icon} <b>{e['full_name']}</b> ({get_dept(e['tg_id'])}): "
                     f"отчётов {r_c}, задач ✅{d_c}" + (f" ⚠️{o_c}" if o_c else ""))
    try:
        await bot.send_message(ADMIN_CHAT_ID, "\n".join(lines), parse_mode="HTML")
    except Exception as ex: logger.warning(f"weekly_audit: {ex}")

async def job_monthly_audit(bot: Bot):
    now = datetime.now()
    month = now.strftime("%Y-%m"); month_label = now.strftime("%B %Y")
    _, last = calendar.monthrange(now.year, now.month)
    wd = sum(1 for d in range(1, now.day+1) if datetime(now.year, now.month, d).weekday() < 5)
    employees = emp_employees()
    reps      = records_for_month("reports", RH)
    all_t     = tasks_all()
    done_m    = [t for t in all_t if t["status"]=="done" and t.get("done_at","").startswith(month)]
    over      = tasks_overdue()
    lines = [f"📆 <b>Ежемесячный аудит — {month_label}</b>\n",
             f"Рабочих дней: {wd}", f"Задач выполнено: {len(done_m)}", f"Просроченных: {len(over)}\n",
             "<b>По сотрудникам:</b>"]
    for e in employees:
        tid = str(e["tg_id"])
        r_c = sum(1 for r in reps if str(r["tg_id"]) == tid)
        d_c = sum(1 for t in done_m if str(t["assigned_to_id"]) == tid)
        pct = round(r_c / max(wd, 1) * 100)
        icon = "✅" if pct >= 80 else ("🟡" if pct >= 40 else "🔴")
        lines.append(f"{icon} <b>{e['full_name']}</b> ({get_dept(e['tg_id'])}): "
                     f"отчётов {r_c}/{wd} ({pct}%), задач ✅{d_c}")
    try:
        await bot.send_message(ADMIN_CHAT_ID, "\n".join(lines), parse_mode="HTML")
    except Exception as ex: logger.warning(f"monthly_audit: {ex}")

def _is_this_week(dt_str, iso_week):
    try:
        return datetime.strptime(dt_str[:10], "%Y-%m-%d").strftime("%Y-W%W") == iso_week
    except Exception: return False

# ── SCHEDULER ─────────────────────────────────────────────────────────────────
def build_scheduler(bot: Bot):
    s = AsyncIOScheduler(timezone=APZ)
    def j(fn, **kw):
        s.add_job(fn, CronTrigger(timezone=APZ, **kw), args=[bot])
    j(job_request_plans,   day_of_week="mon-fri", hour=11, minute=0)
    j(job_remind_plans,    day_of_week="mon-fri", hour=11, minute=30)
    j(job_ping_deadlines,  day_of_week="mon-fri", hour=10, minute=0)
    j(job_request_eod,     day_of_week="mon-fri", hour=19, minute=0)
    j(job_remind_reports,  day_of_week="mon-fri", hour=19, minute=30)
    j(job_daily_digest,    day_of_week="mon-fri", hour=19, minute=35)
    j(job_weekly_audit,    day_of_week="fri",     hour=18, minute=0)
    j(job_monthly_audit,   day=28,                hour=17, minute=0)
    return s

# ── MAIN ──────────────────────────────────────────────────────────────────────
async def post_init(app):
    scheduler = build_scheduler(app.bot)
    scheduler.start()
    logger.info("Scheduler started")
    await app.bot.set_my_commands([
        ("start",     "Регистрация / меню"),
        ("ping",      "Проверка что бот на связи"),
        ("plan",      "План на день — поддерживает [ДД.ММ.ГГГГ] для другого срока"),
        ("report",    "Итоговый отчёт за день"),
        ("eod",       "Закрыть день — статусы задач"),
        ("task",      "Поставить задачу — выбор отдела и сотрудника кнопками"),
        ("tag",       "Тег канала для задачи: /tag — выбор кнопками"),
        ("mytasks",   "Мои активные задачи"),
        ("done",      "Выполнено: /done ID"),
        ("status",    "Статус: /status ID"),
        ("edit",      "Изменить задачу: /edit ID title|deadline значение"),
        ("menu",      "Панель руководителя с кнопками"),
        ("team",      "Сводка команды"),
        ("tasks_all", "Все задачи"),
        ("checkstatuses", "Запросить статусы у команды/отдела прямо сейчас"),
        ("recovertasks", "Восстановить пропавшие задачи из планов: /recovertasks [дней]"),
        ("makeadmin", "Стать администратором"),
        ("setdepthead", "Назначить руководителя отдела: /setdepthead @username"),
        ("setdept", "Сменить отдел сотрудника: /setdept @username"),
        ("fixname", "Исправить имя: /fixname @username Имя Фамилия"),
        ("fixallnames", "Нормализовать имена всех сотрудников"),
    ])

async def global_error_handler(update: object, ctx: ContextTypes.DEFAULT_TYPE):
    """Логирует все необработанные исключения, не даёт им тихо остановить polling."""
    logger.error(f"Unhandled exception: {ctx.error}", exc_info=ctx.error)
    try:
        if isinstance(update, Update) and update.effective_message:
            await update.effective_message.reply_text(
                "⚠️ Произошла ошибка при обработке запроса. Попробуй ещё раз или напиши /menu."
            )
        elif isinstance(update, Update) and update.callback_query:
            await update.callback_query.answer("⚠️ Ошибка, попробуй снова.", show_alert=True)
    except Exception:
        pass

def main():
    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()
    app.add_error_handler(global_error_handler)
    cancel = CommandHandler("cancel", lambda u,c: ConversationHandler.END)

    app.add_handler(ConversationHandler(
        entry_points=[CommandHandler("start", cmd_start)],
        states={S_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_name)]},
        fallbacks=[cancel],
    ))
    app.add_handler(ConversationHandler(
        entry_points=[CommandHandler("plan", cmd_plan)],
        states={S_PLAN: [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_plan)]},
        fallbacks=[cancel],
    ))
    app.add_handler(ConversationHandler(
        entry_points=[CommandHandler("report", cmd_report)],
        states={S_REPORT: [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_report)]},
        fallbacks=[cancel],
    ))
    app.add_handler(ConversationHandler(
        entry_points=[CommandHandler("task", cmd_task)],
        states={S_TASK: [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_task)]},
        fallbacks=[cancel],
    ))

    # EOD text handlers (не ConversationHandler — работают через bot_data флаги)
    app.add_handler(MessageHandler(
        filters.TEXT & ~filters.COMMAND & filters.ChatType.PRIVATE,
        eod_text_router
    ), group=1)

    app.add_handler(CommandHandler("eod",       cmd_eod))
    app.add_handler(CommandHandler("done",      cmd_done))
    app.add_handler(CommandHandler("status",    cmd_status))
    app.add_handler(CommandHandler("mytasks",   cmd_mytasks))
    app.add_handler(CommandHandler("menu",      cmd_menu))
    app.add_handler(CommandHandler("makeadmin", cmd_makeadmin))
    app.add_handler(CommandHandler("setadmin",  cmd_setadmin))
    app.add_handler(CommandHandler("setdepthead", cmd_setdepthead))
    app.add_handler(CommandHandler("setdept", cmd_setdept))
    app.add_handler(CommandHandler("fixsheets", cmd_fixsheets))
    app.add_handler(CommandHandler("fixname", cmd_fixname))
    app.add_handler(CommandHandler("fixallnames", cmd_fixallnames))
    app.add_handler(CommandHandler("edit", cmd_edit))
    app.add_handler(CommandHandler("tag", cmd_tag))
    app.add_handler(CommandHandler("ping", cmd_ping))
    app.add_handler(CommandHandler("checkstatuses", cmd_checkstatuses))
    app.add_handler(CommandHandler("recovertasks", cmd_recovertasks))
    app.add_handler(CommandHandler("team",      cmd_team))
    app.add_handler(CommandHandler("tasks_all", cmd_tasks_all))

    # Callbacks
    app.add_handler(CallbackQueryHandler(cb_confirm_plan,        pattern="^confirm_plan_"))
    app.add_handler(CallbackQueryHandler(cb_tasknew_dept,        pattern="^tasknew_dept_"))
    app.add_handler(CallbackQueryHandler(cb_tasknew_emp,         pattern="^tasknew_emp_"))
    app.add_handler(CallbackQueryHandler(cb_tagtask_pick,        pattern="^tagtask_"))
    app.add_handler(CallbackQueryHandler(cb_tagchannel_pick,     pattern="^tagchannel_"))
    app.add_handler(CallbackQueryHandler(cb_sethead_dept,        pattern="^sethead_dept_"))
    app.add_handler(CallbackQueryHandler(cb_sethead_emp,         pattern="^sethead_emp_"))
    app.add_handler(CallbackQueryHandler(cb_reg_dept,            pattern="^reg_dept_"))
    app.add_handler(CallbackQueryHandler(cb_admset_head,         pattern="^admset_head_"))
    app.add_handler(CallbackQueryHandler(cb_admset_dept,         pattern="^admset_dept_"))
    app.add_handler(CallbackQueryHandler(cb_eod_status,          pattern="^eods_"))
    app.add_handler(CallbackQueryHandler(cb_eod_extra_yes,       pattern="^eod_extra_yes$"))
    app.add_handler(CallbackQueryHandler(cb_eod_extra_no,        pattern="^eod_extra_no$"))
    app.add_handler(CallbackQueryHandler(cb_back_main,           pattern="^back_main$"))
    app.add_handler(CallbackQueryHandler(cb_tasks_today,         pattern="^tasks_today$"))
    app.add_handler(CallbackQueryHandler(cb_show_all_tasks,      pattern="^show_all_tasks$"))
    app.add_handler(CallbackQueryHandler(cb_summary_depts,       pattern="^summary_depts$"))
    app.add_handler(CallbackQueryHandler(cb_summary_person_list, pattern="^summary_person_list$"))
    app.add_handler(CallbackQueryHandler(cb_summary_person,      pattern="^person_"))
    app.add_handler(CallbackQueryHandler(cb_period_week,         pattern="^period_week$"))
    app.add_handler(CallbackQueryHandler(cb_period_month,        pattern="^period_month$"))
    app.add_handler(CallbackQueryHandler(cb_closed_period_pick,  pattern="^closed_period_pick$"))
    app.add_handler(CallbackQueryHandler(cb_closed_tasks,        pattern="^closed_(today|week|month)$"))
    app.add_handler(CallbackQueryHandler(cb_export_dept_pick,    pattern="^export_dept_pick$"))
    app.add_handler(CallbackQueryHandler(cb_export_period,       pattern="^export_(week|month)$"))
    app.add_handler(CallbackQueryHandler(cb_dynamics_dept,       pattern="^dynamics_dept$"))
    app.add_handler(CallbackQueryHandler(cb_recover_period_pick, pattern="^recover_period_pick$"))
    app.add_handler(CallbackQueryHandler(cb_recover_run,         pattern=r"^recover_run_\d+$"))
    app.add_handler(CallbackQueryHandler(cb_checkstatuses_now,   pattern="^checkstatuses_now$"))

    logger.info("Bot starting...")
    app.run_polling(drop_pending_updates=True)

# ── EOD TEXT ROUTER ───────────────────────────────────────────────────────────
async def eod_text_router(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Роутер для текстовых ответов в EOD-потоке и в новом диалоговом /task."""
    tg_id = update.effective_user.id
    # Комментарий к задаче
    if tg_id in ctx.bot_data.get("eod_pending_comment", {}):
        await recv_eod_comment(update, ctx)
        return
    # Доп. задачи
    if tg_id in ctx.bot_data.get("eod_extra_pending", set()):
        await recv_eod_extra(update, ctx)
        return
    # Новая задача после выбора сотрудника кнопками в /task
    if tg_id in ctx.bot_data.get("tasknew_pending", {}):
        await recv_tasknew_text(update, ctx)
        return

if __name__ == "__main__":
    main()
