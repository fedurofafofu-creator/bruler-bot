import asyncio
import logging
import json
import uuid
import re
import pytz
import gspread
import calendar

from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials

from telegram import Update, Bot, InlineKeyboardButton, InlineKeyboardMarkup
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

DEPARTMENTS = {
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
    return DEPARTMENTS.get(str(tg_id), "Без отдела")

# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
_gc = None
_ss = None

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
    try:
        return ss().worksheet(name)
    except gspread.WorksheetNotFound:
        return ss().add_worksheet(title=name, rows=2000, cols=20)

def ensure_headers(ws, headers):
    """Гарантирует корректный, не дублирующийся заголовок в первой строке."""
    vals = ws.row_values(1)
    if vals != headers:
        # перезаписываем строку заголовка целиком (без сдвига остальных строк)
        ws.update('A1', [headers])

def safe_records(ws, headers):
    """get_all_records, устойчивый к повреждённому заголовку."""
    ensure_headers(ws, headers)
    try:
        return ws.get_all_records(expected_headers=headers)
    except Exception:
        # fallback: читаем вручную по индексам колонок
        all_values = ws.get_all_values()
        if len(all_values) < 2:
            return []
        rows = all_values[1:]
        result = []
        for row in rows:
            row = row + [""] * (len(headers) - len(row))
            result.append({headers[i]: row[i] for i in range(len(headers))})
        return result

# ── EMPLOYEES ─────────────────────────────────────────────────────────────────
EMP_H = ["tg_id","username","full_name","role","registered_at"]

def emp_sheet():
    ws = sheet("employees"); ensure_headers(ws, EMP_H); return ws

def emp_all():
    return safe_records(emp_sheet(), EMP_H)

def emp_employees():
    return [r for r in emp_all() if r["role"] in ("employee", "dept_head")]

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

def emp_register(tg_id, username, full_name, role="employee"):
    if emp_registered(tg_id):
        return False
    emp_sheet().append_row([tg_id, username or "", full_name, role,
                             datetime.now().strftime("%Y-%m-%d %H:%M")])
    return True

def emp_set_admin(tg_id):
    ws = emp_sheet()
    for i, r in enumerate(safe_records(ws, EMP_H), start=2):
        if str(r["tg_id"]) == str(tg_id):
            ws.update_cell(i, EMP_H.index("role") + 1, "admin")
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

def save_report(tg_id, name, text):
    reports_sheet().append_row([today_str(), tg_id, name, text,
                                 datetime.now().strftime("%Y-%m-%d %H:%M")])

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
      "done_at","comment","result_link","source"]

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

def task_create(by_id, by_name, to_id, to_name, title, deadline, source="manual"):
    tid = str(uuid.uuid4())[:8].upper()
    tasks_sheet().append_row([
        tid, by_id, by_name, to_id, to_name, title,
        deadline, "open", datetime.now().strftime("%Y-%m-%d %H:%M"),
        "", "", "", source
    ])
    return tid

def task_update_status(tid, status):
    row = task_find_row(tid)
    if not row: return False
    ws = tasks_sheet()
    ws.update_cell(row, TH.index("status") + 1, status)
    if status == "done":
        ws.update_cell(row, TH.index("done_at") + 1, datetime.now().strftime("%Y-%m-%d %H:%M"))
    return True

def task_update_comment(tid, comment, link=""):
    row = task_find_row(tid)
    if not row: return False
    ws = tasks_sheet()
    ws.update_cell(row, TH.index("comment") + 1, comment)
    if link:
        ws.update_cell(row, TH.index("result_link") + 1, link)
    return True

def tasks_today_for_user(tg_id):
    """Задачи пользователя с дедлайном сегодня (из плана и вручную)."""
    return [r for r in tasks_all()
            if str(r["assigned_to_id"]) == str(tg_id)
            and r.get("deadline","") == today_str()]

# ── PARSE PLAN ────────────────────────────────────────────────────────────────
def parse_plan_items(text: str) -> list:
    items = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        line = re.sub(r"^\d+[\.\)]\s*", "", line)
        line = re.sub(r"^[-–—•*]\s*", "", line)
        line = re.sub(r"^[✅🪡💘🩵🔜👌📌⚡✔▪▸►→]\s*", "", line)
        line = line.strip()
        if len(line) > 3:
            items.append(line)
    return items

# ── CONVERSATION STATES ───────────────────────────────────────────────────────
S_NAME         = 1
S_PLAN         = 10
S_REPORT       = 11
S_TASK         = 20
S_EOD_STATUS   = 30   # выбор статуса задачи конец дня
S_EOD_COMMENT  = 31   # комментарий + ссылка
S_EOD_EXTRA    = 32   # были ли другие задачи
S_EOD_EXTRA_TEXT = 33 # текст других задач

# Храним состояние сессии EOD (end-of-day) в bot_data
# bot_data["eod"][tg_id] = {"tasks": [...], "current_idx": 0, "results": [...]}

# ── /start ────────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    if emp_registered(tg_id):
        u = emp_by_id(tg_id)
        role = "руководитель" if u["role"] == "admin" else "сотрудник"
        await update.effective_message.reply_text(
            f"👋 Привет, {u['full_name']}! Ты зарегистрирован как {role}.\n\n"
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
    tg_id = update.effective_user.id
    name = update.message.text.strip()
    emp_register(tg_id, update.effective_user.username or "", name)
    await update.message.reply_text(
        f"✅ Готово, {name}!\nКаждый день в 11:00 — план, в 19:00 — отчёт.\n"
        "/plan — написать план сейчас"
    )
    return ConversationHandler.END

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

    # создаём задачи из плана
    items = parse_plan_items(text)
    created_ids = []
    for item in items:
        tid = task_create(tg_id, u["full_name"], tg_id, u["full_name"],
                          item, today, source="plan")
        created_ids.append((tid, item))

    if created_ids:
        task_list = "\n".join(f"  • <code>{tid}</code> {title}" for tid, title in created_ids)
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Подтвердить задачи", callback_data=f"confirm_plan_{tg_id}"),
        ]])
        await update.message.reply_text(
            f"📋 План записан! Создано <b>{len(created_ids)} задач</b>:\n\n{task_list}\n\n"
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
    today = today_str()
    u = emp_by_id(tg_id)

    for item in items:
        tid = task_create(tg_id, u["full_name"], tg_id, u["full_name"],
                          item, today, source="extra")
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

    await message.reply_text(
        f"✅ День закрыт! Выполнено {done_c} из {total_c} задач.\n\n"
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
    if not emp_registered(update.effective_user.id):
        await update.effective_message.reply_text("Сначала /start")
        return ConversationHandler.END
    args = " ".join(ctx.args) if ctx.args else ""
    if args and parse_task(args):
        return await do_create_task(update, ctx, parse_task(args))
    await update.effective_message.reply_text(
        "📌 Формат:\n<code>@username Название | ДД.ММ.ГГГГ</code>\n\n"
        "Пример:\n<code>@tanya Контент-план | 20.06.2026</code>",
        parse_mode="HTML"
    )
    return S_TASK

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
    await update.effective_message.reply_text(
        f"📌 <b>{task['title']}</b>\n{sl.get(task['status'],task['status'])}\n"
        f"Исполнитель: {task['assigned_to_name']}\nПостановщик: {task['created_by_name']}\n"
        f"Срок: {dl_fmt}{comment_line}{link_line}\nID: <code>{task['task_id']}</code>",
        parse_mode="HTML"
    )

async def cmd_mytasks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tasks = tasks_for_user(update.effective_user.id)
    if not tasks:
        await update.effective_message.reply_text("✅ Нет активных задач!"); return
    sl = {"open":"⬜","in_progress":"🔄","paused":"⏸","overdue":"🔴"}
    lines = ["📋 <b>Твои задачи:</b>\n"]
    for t in tasks:
        dl = t["deadline"]
        dl_fmt = dl[8:]+"."+dl[5:7] if dl else "—"
        src = " <i>(план)</i>" if t.get("source") == "plan" else ""
        lines.append(f"{sl.get(t['status'],'⚪')} <code>{t['task_id']}</code> {t['title']} — до {dl_fmt}{src}")
    await update.effective_message.reply_text("\n".join(lines), parse_mode="HTML")

# ── ADMIN KEYBOARD ────────────────────────────────────────────────────────────
def build_admin_keyboard():
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📋 Задачи сегодня",   callback_data="tasks_today"),
            InlineKeyboardButton("📋 Все активные",     callback_data="show_all_tasks"),
        ],
        [
            InlineKeyboardButton("📊 По отделам",       callback_data="summary_depts"),
            InlineKeyboardButton("👤 По сотруднику",    callback_data="summary_person_list"),
        ],
        [
            InlineKeyboardButton("📅 За неделю",        callback_data="period_week"),
            InlineKeyboardButton("📅 За месяц",         callback_data="period_month"),
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
    """/setdepthead @username — назначить руководителем подразделения (видит только свой отдел)."""
    if not emp_is_admin(update.effective_user.id):
        await update.effective_message.reply_text("⛔ Только для администраторов."); return
    if not ctx.args:
        await update.effective_message.reply_text("Укажи: /setdepthead @username"); return
    username = ctx.args[0].lstrip("@")
    for e in emp_all():
        if e["username"].lstrip("@").lower() == username.lower():
            row = None
            ws = emp_sheet()
            for i, r in enumerate(safe_records(ws, EMP_H), start=2):
                if str(r["tg_id"]) == str(e["tg_id"]):
                    row = i; break
            if row:
                ws.update_cell(row, EMP_H.index("role") + 1, "dept_head")
                dept = get_dept(e["tg_id"])
                await update.effective_message.reply_text(
                    f"✅ {e['full_name']} теперь руководитель подразделения «{dept}».\n"
                    f"Видит планы/отчёты/задачи только своего отдела через /menu"
                )
            return
    await update.effective_message.reply_text(f"❌ @{username} не найден.")

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

# ── ADMIN CALLBACKS ───────────────────────────────────────────────────────────
async def cb_back_main(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query; await query.answer()
    await query.message.reply_text("Выбери действие:", reply_markup=build_admin_keyboard())

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

async def job_daily_digest(bot: Bot):
    today = datetime.now().strftime("%d.%m.%Y")
    employees = emp_employees()
    p_list = plans_today(); r_list = reports_today()
    plan_ids   = {str(p["tg_id"]) for p in p_list}
    report_ids = {str(r["tg_id"]) for r in r_list}
    no_plan   = [e["full_name"] for e in employees if str(e["tg_id"]) not in plan_ids]
    no_report = [e["full_name"] for e in employees if str(e["tg_id"]) not in report_ids]
    over = tasks_overdue()
    lines = [
        f"📊 <b>Сводка команды — {today}</b>",
        f"Планов: {len(plan_ids)}/{len(employees)}   Отчётов: {len(report_ids)}/{len(employees)}\n",
    ]
    if no_plan:   lines.append(f"❌ <b>Нет плана:</b> {', '.join(no_plan)}")
    if no_report: lines.append(f"❌ <b>Нет отчёта:</b> {', '.join(no_report)}")
    if over:
        lines.append(f"\n⚠️ <b>Просроченных задач: {len(over)}</b>")
        for t in over[:3]:
            lines.append(f"  🔴 {t['assigned_to_name']}: {t['title']}")
        if len(over) > 3: lines.append(f"  ...и ещё {len(over)-3}")
    try:
        await bot.send_message(ADMIN_CHAT_ID, "\n".join(lines),
                               parse_mode="HTML", reply_markup=build_admin_keyboard())
    except Exception as ex: logger.warning(f"daily_digest: {ex}")

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
        ("plan",      "Plan на день"),
        ("report",    "Итоговый отчёт за день"),
        ("eod",       "Закрыть день — статусы задач"),
        ("task",      "Поставить задачу @username Название | ДД.ММ.ГГГГ"),
        ("mytasks",   "Мои активные задачи"),
        ("done",      "Выполнено: /done ID"),
        ("status",    "Статус: /status ID"),
        ("menu",      "Панель руководителя с кнопками"),
        ("team",      "Сводка команды"),
        ("tasks_all", "Все задачи"),
        ("makeadmin", "Стать администратором"),
        ("setdepthead", "Назначить руководителя отдела: /setdepthead @username"),
    ])

def main():
    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()
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
    app.add_handler(CommandHandler("fixsheets", cmd_fixsheets))
    app.add_handler(CommandHandler("team",      cmd_team))
    app.add_handler(CommandHandler("tasks_all", cmd_tasks_all))

    # Callbacks
    app.add_handler(CallbackQueryHandler(cb_confirm_plan,        pattern="^confirm_plan_"))
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

    logger.info("Bot starting...")
    app.run_polling(drop_pending_updates=True)

# ── EOD TEXT ROUTER ───────────────────────────────────────────────────────────
async def eod_text_router(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Роутер для текстовых ответов в EOD-потоке."""
    tg_id = update.effective_user.id
    # Комментарий к задаче
    if tg_id in ctx.bot_data.get("eod_pending_comment", {}):
        await recv_eod_comment(update, ctx)
        return
    # Доп. задачи
    if tg_id in ctx.bot_data.get("eod_extra_pending", set()):
        await recv_eod_extra(update, ctx)
        return

if __name__ == "__main__":
    main()
