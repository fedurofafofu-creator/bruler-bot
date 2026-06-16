import asyncio
import logging
import json
import uuid
import re
import pytz
import gspread

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
        return ss().add_worksheet(title=name, rows=1000, cols=20)

def ensure_headers(ws, headers):
    vals = ws.row_values(1)
    if not vals or vals[0] != headers[0]:
        ws.insert_row(headers, 1)

# ── EMPLOYEES ─────────────────────────────────────────────────────────────────
EMP_H = ["tg_id","username","full_name","role","registered_at"]

def emp_sheet():
    ws = sheet("employees")
    ensure_headers(ws, EMP_H)
    return ws

def emp_all():
    return emp_sheet().get_all_records()

def emp_employees():
    return [r for r in emp_all() if r["role"] == "employee"]

def emp_admins():
    return [r for r in emp_all() if r["role"] == "admin"]

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

def emp_register(tg_id, username, full_name, role="employee"):
    if emp_registered(tg_id):
        return False
    emp_sheet().append_row([tg_id, username or "", full_name, role,
                             datetime.now().strftime("%Y-%m-%d %H:%M")])
    return True

def emp_set_admin(tg_id):
    ws = emp_sheet()
    for i, r in enumerate(ws.get_all_records(), start=2):
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
    return [r for r in plans_sheet().get_all_records() if r["date"] == today_str()]

def reports_today():
    return [r for r in reports_sheet().get_all_records() if r["date"] == today_str()]

def has_plan_today(tg_id):
    return any(str(r["tg_id"]) == str(tg_id) for r in plans_today())

def has_report_today(tg_id):
    return any(str(r["tg_id"]) == str(tg_id) for r in reports_today())

def records_for_week(ws_name, headers):
    ws = sheet(ws_name); ensure_headers(ws, headers)
    iso = datetime.now().strftime("%Y-W%W")
    result = []
    for r in ws.get_all_records():
        try:
            d = datetime.strptime(r["date"], "%Y-%m-%d")
            if d.strftime("%Y-W%W") == iso:
                result.append(r)
        except Exception:
            pass
    return result

def records_for_month(ws_name, headers):
    ws = sheet(ws_name); ensure_headers(ws, headers)
    month = datetime.now().strftime("%Y-%m")
    return [r for r in ws.get_all_records() if r.get("date","").startswith(month)]

# ── TASKS ─────────────────────────────────────────────────────────────────────
TH = ["task_id","created_by_id","created_by_name","assigned_to_id",
      "assigned_to_name","title","deadline","status","created_at","done_at","comment"]

def tasks_sheet():
    ws = sheet("tasks"); ensure_headers(ws, TH); return ws

def tasks_all():
    return tasks_sheet().get_all_records()

def task_by_id(tid):
    for r in tasks_all():
        if r["task_id"] == tid.upper():
            return r
    return None

def tasks_for_user(tg_id):
    return [r for r in tasks_all()
            if str(r["assigned_to_id"]) == str(tg_id) and r["status"] != "done"]

def tasks_open():
    return [r for r in tasks_all() if r["status"] in ("open","in_progress")]

def tasks_overdue():
    today = today_str()
    return [r for r in tasks_all()
            if r["status"] != "done" and r.get("deadline","") and r["deadline"] < today]

def tasks_due_tomorrow():
    tmr = (date.today() + timedelta(days=1)).strftime("%Y-%m-%d")
    return [r for r in tasks_all() if r.get("deadline") == tmr and r["status"] != "done"]

def tasks_due_today():
    return [r for r in tasks_all()
            if r.get("deadline") == today_str() and r["status"] != "done"]

def task_find_row(tid):
    for i, r in enumerate(tasks_sheet().get_all_records(), start=2):
        if r["task_id"] == tid.upper():
            return i
    return None

def task_create(by_id, by_name, to_id, to_name, title, deadline):
    tid = str(uuid.uuid4())[:8].upper()
    tasks_sheet().append_row([tid, by_id, by_name, to_id, to_name, title,
                               deadline, "open",
                               datetime.now().strftime("%Y-%m-%d %H:%M"), "", ""])
    return tid

def task_done(tid, comment=""):
    row = task_find_row(tid)
    if not row: return False
    ws = tasks_sheet()
    ws.update_cell(row, TH.index("status") + 1, "done")
    ws.update_cell(row, TH.index("done_at") + 1, datetime.now().strftime("%Y-%m-%d %H:%M"))
    if comment:
        ws.update_cell(row, TH.index("comment") + 1, comment)
    return True

# ── CONVERSATION STATES ───────────────────────────────────────────────────────
S_NAME = 1
S_PLAN = 10
S_REPORT = 11
S_TASK = 20

# ── /start ────────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    if emp_registered(tg_id):
        u = emp_by_id(tg_id)
        role = "руководитель" if u["role"] == "admin" else "сотрудник"
        await update.message.reply_text(
            f"👋 Привет, {u['full_name']}! Ты зарегистрирован как {role}.\n\n"
            "/plan — план на день\n/report — отчёт\n"
            "/task — поставить задачу\n/mytasks — мои задачи\n"
            "/done ID — отметить выполненной\n/status ID — статус задачи"
        )
        return ConversationHandler.END
    await update.message.reply_text(
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
        await update.message.reply_text("Сначала /start"); return ConversationHandler.END
    u = emp_by_id(update.effective_user.id)
    await update.message.reply_text(f"📋 {u['full_name']}, напиши план на сегодня:")
    return S_PLAN

def parse_plan_items(text: str) -> list[str]:
    """Разбирает текст плана на отдельные задачи."""
    items = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        # убираем маркеры списков: 1. 2. / - / • / ✅ / 🪡 / 💘 / 🩵 и т.п.
        line = re.sub(r"^[\d]+[\.\)]\s*", "", line)          # 1. 1)
        line = re.sub(r"^[-–—•*]\s*", "", line)               # - • *
        line = re.sub(r"^[✅🪡💘🩵🔜👌📌⚡️✔️▪️▸►→]\s*", "", line)  # эмодзи-маркеры
        line = line.strip()
        if len(line) > 3:   # игнорируем слишком короткие строки
            items.append(line)
    return items


async def recv_plan(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    u = emp_by_id(tg_id)
    text = update.message.text.strip()
    today = date.today().strftime("%Y-%m-%d")

    # сохраняем полный текст плана
    save_plan(tg_id, u["full_name"], text)

    # удаляем старые задачи из плана на сегодня (если сотрудник переписал план)
    ws = tasks_sheet()
    all_rows = ws.get_all_records()
    rows_to_delete = []
    for i, r in enumerate(all_rows, start=2):
        if (str(r["assigned_to_id"]) == str(tg_id)
                and r.get("created_at", "").startswith(today)
                and r.get("comment", "") == "из плана"):
            rows_to_delete.append(i)
    # удаляем снизу вверх чтобы не сбивать индексы
    for row_i in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row_i)

    # создаём задачи из пунктов плана
    items = parse_plan_items(text)
    created = []
    for item in items:
        tid = str(uuid.uuid4())[:8].upper()
        ws.append_row([
            tid, tg_id, u["full_name"],
            tg_id, u["full_name"],   # назначена себе
            item, today, "open",
            datetime.now().strftime("%Y-%m-%d %H:%M"), "", "из плана"
        ])
        created.append(item)

    if created:
        task_list = "\n".join(f"  • {t}" for t in created)
        await update.message.reply_text(
            f"✅ План записан! Создано {len(created)} задач на сегодня:\n{task_list}\n\n"
            f"Увидимся в 19:00 🌆\n/mytasks — посмотреть свои задачи"
        )
    else:
        await update.message.reply_text("✅ План записан! Увидимся в 19:00 🌆")
    return ConversationHandler.END

# ── /report ───────────────────────────────────────────────────────────────────
async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not emp_registered(update.effective_user.id):
        await update.message.reply_text("Сначала /start"); return ConversationHandler.END
    u = emp_by_id(update.effective_user.id)
    await update.message.reply_text(f"📊 {u['full_name']}, напиши отчёт за день:")
    return S_REPORT

async def recv_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    u = emp_by_id(tg_id)
    save_report(tg_id, u["full_name"], update.message.text.strip())
    await update.message.reply_text("✅ Отчёт записан! Хорошего вечера 🌙")
    return ConversationHandler.END

# ── /task ─────────────────────────────────────────────────────────────────────
def parse_task(text):
    m = re.match(r"@(\S+)\s+(.+?)\s*\|\s*(\d{2}\.\d{2}\.\d{4})", text)
    if not m: return None
    username, title, dl = m.groups()
    p = dl.split(".")
    return username, title.strip(), f"{p[2]}-{p[1]}-{p[0]}"

async def cmd_task(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not emp_registered(update.effective_user.id):
        await update.message.reply_text("Сначала /start"); return ConversationHandler.END
    args = " ".join(ctx.args) if ctx.args else ""
    if args and parse_task(args):
        return await do_create_task(update, ctx, parse_task(args))
    await update.message.reply_text(
        "📌 Формат:\n<code>@username Название | ДД.ММ.ГГГГ</code>\n\n"
        "Пример:\n<code>@tanya Контент-план | 20.06.2026</code>",
        parse_mode="HTML"
    )
    return S_TASK

async def recv_task(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    parsed = parse_task(update.message.text.strip())
    if not parsed:
        await update.message.reply_text(
            "Не понял формат. Попробуй:\n<code>@username Название | ДД.ММ.ГГГГ</code>",
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
        await update.message.reply_text(f"❌ @{username} не найден. Пусть сделает /start.")
        return ConversationHandler.END
    tid = task_create(update.effective_user.id, creator["full_name"],
                      assignee["tg_id"], assignee["full_name"], title, deadline)
    dl_fmt = deadline[8:] + "." + deadline[5:7] + "." + deadline[:4]
    try:
        await ctx.bot.send_message(
            int(assignee["tg_id"]),
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
    return ConversationHandler.END

# ── /done /status /mytasks ────────────────────────────────────────────────────
async def cmd_done(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ctx.args:
        await update.message.reply_text("Укажи ID: /done ABCD1234"); return
    tid = ctx.args[0].upper()
    comment = " ".join(ctx.args[1:])
    task = task_by_id(tid)
    if not task:
        await update.message.reply_text(f"❌ Задача {tid} не найдена."); return
    tg_id = update.effective_user.id
    if str(task["assigned_to_id"]) != str(tg_id) and not emp_is_admin(tg_id):
        await update.message.reply_text("❌ Это не твоя задача."); return
    task_done(tid, comment)
    try:
        await ctx.bot.send_message(
            int(task["created_by_id"]),
            f"✅ Задача выполнена!\n<b>{task['title']}</b>\n"
            f"Выполнил: {task['assigned_to_name']}\nID: <code>{tid}</code>"
            + (f"\nКомментарий: {comment}" if comment else ""),
            parse_mode="HTML"
        )
    except Exception: pass
    await update.message.reply_text(f"✅ Задача <code>{tid}</code> выполнена!", parse_mode="HTML")

async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ctx.args:
        await update.message.reply_text("Укажи ID: /status ABCD1234"); return
    task = task_by_id(ctx.args[0].upper())
    if not task:
        await update.message.reply_text("❌ Не найдена."); return
    sl = {"open":"🔵 Открыта","in_progress":"🟡 В работе","done":"✅ Выполнена","overdue":"🔴 Просрочена"}
    dl = task["deadline"]
    dl_fmt = dl[8:]+"."+dl[5:7]+"."+dl[:4] if dl else "—"
    await update.message.reply_text(
        f"📌 <b>{task['title']}</b>\n{sl.get(task['status'], task['status'])}\n"
        f"Исполнитель: {task['assigned_to_name']}\nПостановщик: {task['created_by_name']}\n"
        f"Срок: {dl_fmt}\nID: <code>{task['task_id']}</code>", parse_mode="HTML"
    )

async def cmd_mytasks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tasks = tasks_for_user(update.effective_user.id)
    if not tasks:
        await update.message.reply_text("✅ Нет активных задач!"); return
    sl = {"open":"🔵","in_progress":"🟡","overdue":"🔴"}
    lines = ["📋 <b>Твои задачи:</b>\n"]
    for t in tasks:
        dl = t["deadline"]
        dl_fmt = dl[8:]+"."+dl[5:7] if dl else "—"
        lines.append(f"{sl.get(t['status'],'⚪')} <code>{t['task_id']}</code> {t['title']} — до {dl_fmt}")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")

# ── ADMIN ─────────────────────────────────────────────────────────────────────
async def cmd_makeadmin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    if not emp_registered(tg_id):
        await update.message.reply_text("Сначала /start"); return
    admins = emp_admins()
    if admins and not emp_is_admin(tg_id):
        await update.message.reply_text("⛔ Уже есть администраторы."); return
    emp_set_admin(tg_id)
    await update.message.reply_text("✅ Ты теперь администратор!")

async def cmd_setadmin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not emp_is_admin(update.effective_user.id):
        await update.message.reply_text("⛔ Только для руководителей."); return
    if not ctx.args:
        await update.message.reply_text("Укажи: /setadmin @username"); return
    username = ctx.args[0].lstrip("@")
    for e in emp_all():
        if e["username"].lstrip("@").lower() == username.lower():
            emp_set_admin(int(e["tg_id"]))
            await update.message.reply_text(f"✅ {e['full_name']} теперь администратор.")
            return
    await update.message.reply_text(f"❌ @{username} не найден.")

async def cmd_team(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not emp_is_admin(update.effective_user.id):
        await update.message.reply_text("⛔ Только для руководителей."); return
    employees = emp_employees()
    plan_ids  = {str(p["tg_id"]) for p in plans_today()}
    report_ids = {str(r["tg_id"]) for r in reports_today()}
    lines = [f"📊 <b>Команда {datetime.now().strftime('%d.%m.%Y')}:</b>\n"]
    for e in employees:
        tid = str(e["tg_id"])
        lines.append(
            f"{'✅' if tid in report_ids else '❌'} <b>{e['full_name']}</b>  "
            f"план {'✅' if tid in plan_ids else '❌'}  отчёт {'✅' if tid in report_ids else '—'}"
        )
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")

async def cmd_tasks_all(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not emp_is_admin(update.effective_user.id):
        await update.message.reply_text("⛔ Только для руководителей."); return
    open_t = tasks_open()
    over_t = tasks_overdue()
    lines = ["📋 <b>Все активные задачи:</b>\n"]
    if not open_t and not over_t:
        lines.append("Нет активных задач.")
    for t in over_t:
        dl = t["deadline"]; dl_fmt = dl[8:]+"."+dl[5:7] if dl else "—"
        lines.append(f"🔴 <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']} (просрочено {dl_fmt})")
    for t in open_t:
        if t in over_t: continue
        dl = t["deadline"]; dl_fmt = dl[8:]+"."+dl[5:7] if dl else "—"
        lines.append(f"🔵 <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']} — до {dl_fmt}")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")

# ── NOTIFICATIONS ─────────────────────────────────────────────────────────────
async def job_request_plans(bot: Bot):
    for e in emp_employees():
        if has_plan_today(int(e["tg_id"])): continue
        try:
            await bot.send_message(int(e["tg_id"]),
                f"☀️ Доброе утро, {e['full_name']}!\nНапиши план на день: /plan")
        except Exception as ex: logger.warning(f"request_plans: {ex}")

async def job_remind_plans(bot: Bot):
    for e in emp_employees():
        if not has_plan_today(int(e["tg_id"])):
            try:
                await bot.send_message(int(e["tg_id"]), "🔔 Напоминание: напиши план /plan")
            except Exception as ex: logger.warning(f"remind_plans: {ex}")

async def job_request_reports(bot: Bot):
    for e in emp_employees():
        try:
            await bot.send_message(int(e["tg_id"]),
                f"🌆 {e['full_name']}, время отчёта!\nЧто сделал сегодня? /report")
        except Exception as ex: logger.warning(f"request_reports: {ex}")

async def job_remind_reports(bot: Bot):
    for e in emp_employees():
        if not has_report_today(int(e["tg_id"])):
            try:
                await bot.send_message(int(e["tg_id"]), "🔔 Напоминание: напиши отчёт /report")
            except Exception as ex: logger.warning(f"remind_reports: {ex}")

async def job_ping_deadlines(bot: Bot):
    for t in tasks_due_tomorrow():
        try:
            dl = t["deadline"]; dl_fmt = dl[8:]+"."+dl[5:7]+"."+dl[:4]
            await bot.send_message(int(t["assigned_to_id"]),
                f"⏰ Завтра срок задачи!\n<b>{t['title']}</b>\nID: <code>{t['task_id']}</code>\n"
                f"/done {t['task_id']} — отметить выполненной", parse_mode="HTML")
        except Exception as ex: logger.warning(f"ping_tomorrow: {ex}")
    for t in tasks_due_today():
        try:
            await bot.send_message(int(t["assigned_to_id"]),
                f"🚨 Срок задачи истекает сегодня!\n<b>{t['title']}</b>\nID: <code>{t['task_id']}</code>",
                parse_mode="HTML")
        except Exception as ex: logger.warning(f"ping_today: {ex}")

# ── CALLBACK BUTTONS ─────────────────────────────────────────────────────────
async def cb_show_all_tasks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not emp_is_admin(query.from_user.id):
        await query.answer("⛔ Только для руководителей.", show_alert=True)
        return
    open_t = tasks_open()
    over_t = tasks_overdue()
    if not open_t and not over_t:
        await query.message.reply_text("✅ Нет активных задач!")
        return
    lines = [f"📋 <b>Все задачи команды ({len(open_t)} активных):</b>\n"]
    if over_t:
        lines.append("<b>🔴 Просроченные:</b>")
        for t in over_t:
            dl = t["deadline"]; dl_fmt = dl[8:]+"."+dl[5:7]+"."+dl[:4] if dl else "—"
            lines.append(f"  🔴 <code>{t['task_id']}</code> <b>{t['assigned_to_name']}</b>: {t['title']} (срок {dl_fmt})")
        lines.append("")
    active = [t for t in open_t if t not in over_t]
    if active:
        lines.append("<b>🔵 Активные:</b>")
        for t in active:
            dl = t["deadline"]; dl_fmt = dl[8:]+"."+dl[5:7]+"."+dl[:4] if dl else "—"
            lines.append(f"  🔵 <code>{t['task_id']}</code> <b>{t['assigned_to_name']}</b>: {t['title']} — до {dl_fmt}")
    keyboard = InlineKeyboardMarkup([[
        InlineKeyboardButton("🔄 Обновить", callback_data="show_all_tasks")
    ]])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)


async def cb_show_team_summary(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not emp_is_admin(query.from_user.id):
        await query.answer("⛔ Только для руководителей.", show_alert=True)
        return
    employees = emp_employees()
    plan_ids   = {str(p["tg_id"]) for p in plans_today()}
    report_ids = {str(r["tg_id"]) for r in reports_today()}
    all_t = tasks_all()
    lines = [f"📊 <b>Сводка команды на {datetime.now().strftime('%d.%m.%Y')}:</b>\n"]
    for e in employees:
        tid = str(e["tg_id"])
        active = sum(1 for t in all_t if str(t["assigned_to_id"])==tid and t["status"]!="done")
        over   = sum(1 for t in all_t if str(t["assigned_to_id"])==tid and t["status"]!="done"
                     and t.get("deadline","") and t["deadline"] < datetime.now().strftime("%Y-%m-%d"))
        p_icon = "✅" if tid in plan_ids else "❌"
        r_icon = "✅" if tid in report_ids else "❌"
        over_str = f"  ⚠️просрочено: {over}" if over else ""
        lines.append(
            f"<b>{e['full_name']}</b>\n"
            f"  план {p_icon}  отчёт {r_icon}  "
            f"задач активных: {active}{over_str}"
        )
    await query.message.reply_text("\n".join(lines), parse_mode="HTML")


async def job_daily_digest(bot: Bot):
    today = datetime.now().strftime("%d.%m.%Y")
    employees = emp_employees()
    p_list = plans_today(); r_list = reports_today()
    plan_map   = {str(p["tg_id"]): p["plan_text"] for p in p_list}
    report_map = {str(r["tg_id"]): r["report_text"] for r in r_list}
    lines = [f"📊 <b>Сводка за {today}</b>",
             f"Отчётов: {len(r_list)}/{len(employees)}\n"]
    for e in employees:
        tid = str(e["tg_id"]); name = e["full_name"]
        if tid in report_map:
            lines.append(f"✅ <b>{name}</b>")
            if tid in plan_map: lines.append(f"  📋 {plan_map[tid][:200]}")
            lines.append(f"  📝 {report_map[tid][:300]}")
        elif tid in plan_map:
            lines.append(f"🟡 <b>{name}</b> — план есть, отчёт не сдан\n  📋 {plan_map[tid][:200]}")
        else:
            lines.append(f"❌ <b>{name}</b> — ни плана, ни отчёта")
        lines.append("")
    over = tasks_overdue()
    if over:
        lines.append(f"⚠️ <b>Просроченные задачи ({len(over)}):</b>")
        for t in over[:5]:
            lines.append(f"  🔴 {t['assigned_to_name']}: {t['title']} (срок {t['deadline']})")
    keyboard = InlineKeyboardMarkup([[
        InlineKeyboardButton("📋 Все задачи команды", callback_data="show_all_tasks"),
        InlineKeyboardButton("📊 По сотрудникам", callback_data="show_team_summary"),
    ]])
    try:
        await bot.send_message(ADMIN_CHAT_ID, "\n".join(lines),
                               parse_mode="HTML", reply_markup=keyboard)
    except Exception as ex: logger.warning(f"daily_digest: {ex}")

async def job_weekly_audit(bot: Bot):
    now = datetime.now()
    ws = now.strftime("%Y-W%W")
    w0 = (now - timedelta(days=now.weekday())).strftime("%d.%m")
    w1 = now.strftime("%d.%m.%Y")
    employees = emp_employees()
    plans_w   = records_for_week("plans", PH)
    reports_w = records_for_week("reports", RH)
    done_w    = [t for t in tasks_all() if t["status"]=="done" and _is_this_week(t.get("done_at",""), ws)]
    over      = tasks_overdue()
    lines = [f"📅 <b>Еженедельный аудит ({w0} — {w1})</b>\n",
             f"Планов: {len(set(p['tg_id'] for p in plans_w))}/{len(employees)}",
             f"Отчётов: {len(set(r['tg_id'] for r in reports_w))}/{len(employees)}",
             f"Задач выполнено: {len(done_w)}",
             f"Просроченных: {len(over)}\n",
             "<b>По сотрудникам:</b>"]
    for e in employees:
        tid = str(e["tg_id"])
        p_c = sum(1 for p in plans_w if str(p["tg_id"])==tid)
        r_c = sum(1 for r in reports_w if str(r["tg_id"])==tid)
        d_c = sum(1 for t in done_w if str(t["assigned_to_id"])==tid)
        o_c = sum(1 for t in over if str(t["assigned_to_id"])==tid)
        icon = "✅" if r_c >= 4 else ("🟡" if r_c >= 2 else "🔴")
        lines.append(f"{icon} <b>{e['full_name']}</b>: планов {p_c}, отчётов {r_c}, выполнено задач {d_c}"
                     + (f", просрочено {o_c}" if o_c else ""))
    try:
        await bot.send_message(ADMIN_CHAT_ID, "\n".join(lines), parse_mode="HTML")
    except Exception as ex: logger.warning(f"weekly_audit: {ex}")

async def job_monthly_audit(bot: Bot):
    now = datetime.now()
    month = now.strftime("%Y-%m")
    month_label = now.strftime("%B %Y")
    employees = emp_employees()
    plans_m   = records_for_month("plans", PH)
    reports_m = records_for_month("reports", RH)
    done_m    = [t for t in tasks_all() if t["status"]=="done" and t.get("done_at","").startswith(month)]
    over      = tasks_overdue()
    import calendar
    _, last = calendar.monthrange(now.year, now.month)
    wd = sum(1 for d in range(1, last+1) if datetime(now.year, now.month, d).weekday() < 5)
    lines = [f"📆 <b>Ежемесячный аудит — {month_label}</b>\n",
             f"Рабочих дней: {wd}",
             f"Задач выполнено: {len(done_m)}",
             f"Просроченных: {len(over)}\n",
             "<b>По сотрудникам:</b>"]
    for e in employees:
        tid = str(e["tg_id"])
        r_c = sum(1 for r in reports_m if str(r["tg_id"])==tid)
        d_c = sum(1 for t in done_m if str(t["assigned_to_id"])==tid)
        pct = round(r_c / max(wd, 1) * 100)
        icon = "✅" if pct >= 80 else ("🟡" if pct >= 50 else "🔴")
        lines.append(f"{icon} <b>{e['full_name']}</b>: отчётов {r_c}/{wd} ({pct}%), задач {d_c}")
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
    j(job_request_reports, day_of_week="mon-fri", hour=19, minute=0)
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
        ("plan",      "План на день"),
        ("report",    "Отчёт за день"),
        ("task",      "Поставить задачу @username Название | ДД.ММ.ГГГГ"),
        ("mytasks",   "Мои активные задачи"),
        ("done",      "Выполнено: /done ID"),
        ("status",    "Статус: /status ID"),
        ("team",      "Сводка команды (руководители)"),
        ("tasks_all", "Все задачи (руководители)"),
        ("makeadmin", "Стать администратором (первый запуск)"),
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
    app.add_handler(CommandHandler("done",      cmd_done))
    app.add_handler(CommandHandler("status",    cmd_status))
    app.add_handler(CommandHandler("mytasks",   cmd_mytasks))
    app.add_handler(CommandHandler("makeadmin", cmd_makeadmin))
    app.add_handler(CommandHandler("setadmin",  cmd_setadmin))
    app.add_handler(CommandHandler("team",      cmd_team))
    app.add_handler(CommandHandler("tasks_all", cmd_tasks_all))
    app.add_handler(CallbackQueryHandler(cb_show_all_tasks,    pattern="^show_all_tasks$"))
    app.add_handler(CallbackQueryHandler(cb_show_team_summary, pattern="^show_team_summary$"))

    logger.info("Bot starting...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
