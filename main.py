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

# Отделы сотрудников (tg_id -> отдел)
DEPARTMENTS = {
    "7070230704":  "Руководство",
    "7198542902":  "Руководство",
    "8151347813":  "Продажи",
    "195676845":   "Маркетинг",
    "8069881891":  "Бухгалтерия",
    "458764300":   "Производство и дизайн",
    "860192861":   "Производство и дизайн",
    "89555212":    "Производство и дизайн",
    "549232571":   "Маркетинг",
}

def get_dept(tg_id) -> str:
    return DEPARTMENTS.get(str(tg_id), "Без отдела")
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

# ── HELPER FUNCTIONS ─────────────────────────────────────────────────────────
def fmt_dl(deadline: str) -> str:
    if not deadline: return "—"
    return deadline[8:]+"."+deadline[5:7]+"."+deadline[:4]

def is_admin_check(query) -> bool:
    if not emp_is_admin(query.from_user.id):
        return False
    return True

def tasks_for_date(target_date: str) -> list:
    """Задачи с дедлайном на конкретную дату."""
    return [t for t in tasks_all() if t.get("deadline","") == target_date and t["status"] != "done"]

def tasks_for_period(date_from: str, date_to: str) -> list:
    """Задачи за период."""
    return [t for t in tasks_all()
            if t.get("deadline","") >= date_from and t.get("deadline","") <= date_to]

def reports_for_period(date_from: str, date_to: str) -> list:
    ws = sheet("reports"); ensure_headers(ws, RH)
    return [r for r in ws.get_all_records()
            if r.get("date","") >= date_from and r.get("date","") <= date_to]

def plans_for_period(date_from: str, date_to: str) -> list:
    ws = sheet("plans"); ensure_headers(ws, PH)
    return [p for p in ws.get_all_records()
            if p.get("date","") >= date_from and p.get("date","") <= date_to]

def back_keyboard():
    return InlineKeyboardMarkup([[InlineKeyboardButton("◀️ Назад", callback_data="back_main")]])


# ── CALLBACK HANDLERS ─────────────────────────────────────────────────────────
async def cb_back_main(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.message.reply_text("Выбери действие:", reply_markup=build_admin_keyboard())


async def cb_tasks_today(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Все активные задачи с дедлайном сегодня."""
    query = update.callback_query
    await query.answer()
    if not is_admin_check(query):
        await query.answer("⛔ Только для руководителей.", show_alert=True); return
    today = date.today().strftime("%Y-%m-%d")
    today_fmt = date.today().strftime("%d.%m.%Y")
    tasks = tasks_for_date(today)
    over = [t for t in tasks_overdue() if t not in tasks]

    lines = [f"📋 <b>Задачи на сегодня ({today_fmt}):</b>\n"]
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
            src = " <i>(из плана)</i>" if t.get("comment") == "из плана" else ""
            lines.append(f"  🔵 <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']}{src}")

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔄 Обновить", callback_data="tasks_today")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)


async def cb_show_all_tasks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Все активные задачи команды."""
    query = update.callback_query
    await query.answer()
    if not is_admin_check(query):
        await query.answer("⛔ Только для руководителей.", show_alert=True); return
    open_t = tasks_open()
    over_t = tasks_overdue()
    lines = [f"📋 <b>Все активные задачи ({len(open_t)}):</b>\n"]
    if not open_t:
        lines.append("✅ Нет активных задач!")
    else:
        # группируем по отделу
        by_dept: dict[str, list] = {}
        for t in open_t:
            dept = get_dept(t["assigned_to_id"])
            by_dept.setdefault(dept, []).append(t)
        for dept, dt_list in sorted(by_dept.items()):
            lines.append(f"<b>— {dept} —</b>")
            for t in dt_list:
                icon = "🔴" if t in over_t else "🔵"
                src = " <i>(план)</i>" if t.get("comment") == "из плана" else ""
                lines.append(f"  {icon} <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']} — до {fmt_dl(t['deadline'])}{src}")
            lines.append("")
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔄 Обновить", callback_data="show_all_tasks")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)


async def cb_summary_depts(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Сводка по отделам за сегодня."""
    query = update.callback_query
    await query.answer()
    if not is_admin_check(query):
        await query.answer("⛔ Только для руководителей.", show_alert=True); return
    employees = emp_employees()
    plan_ids   = {str(p["tg_id"]) for p in plans_today()}
    report_ids = {str(r["tg_id"]) for r in reports_today()}
    all_t = tasks_all()
    today = date.today().strftime("%Y-%m-%d")
    today_fmt = date.today().strftime("%d.%m.%Y")

    # группируем по отделам
    dept_employees: dict[str, list] = {}
    for e in employees:
        dept = get_dept(e["tg_id"])
        dept_employees.setdefault(dept, []).append(e)

    lines = [f"📊 <b>Сводка по отделам — {today_fmt}</b>\n"]
    for dept, emps in sorted(dept_employees.items()):
        submitted = sum(1 for e in emps if str(e["tg_id"]) in report_ids)
        planned   = sum(1 for e in emps if str(e["tg_id"]) in plan_ids)
        overdue_c = sum(1 for t in all_t
                        if get_dept(t["assigned_to_id"]) == dept
                        and t["status"] != "done" and t.get("deadline","") < today)
        icon = "✅" if submitted == len(emps) else ("🟡" if submitted > 0 else "🔴")
        lines.append(f"{icon} <b>{dept}</b> ({len(emps)} чел.)")
        lines.append(f"  планов {planned}/{len(emps)}  отчётов {submitted}/{len(emps)}"
                     + (f"  ⚠️просрочено {overdue_c}" if overdue_c else ""))
        for e in emps:
            tid = str(e["tg_id"])
            p = "✅" if tid in plan_ids else "❌"
            r = "✅" if tid in report_ids else "—"
            lines.append(f"    {e['full_name']}  план {p}  отчёт {r}")
        lines.append("")

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("👤 По сотруднику", callback_data="summary_person_list")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)


async def cb_summary_person_list(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Список сотрудников для выбора."""
    query = update.callback_query
    await query.answer()
    if not is_admin_check(query):
        await query.answer("⛔ Только для руководителей.", show_alert=True); return
    employees = emp_employees()
    buttons = []
    row = []
    for i, e in enumerate(employees):
        row.append(InlineKeyboardButton(
            e["full_name"].split()[0],
            callback_data=f"person_{e['tg_id']}"
        ))
        if len(row) == 2:
            buttons.append(row); row = []
    if row: buttons.append(row)
    buttons.append([InlineKeyboardButton("◀️ Назад", callback_data="back_main")])
    await query.message.reply_text("Выбери сотрудника:", reply_markup=InlineKeyboardMarkup(buttons))


async def cb_summary_person(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Детальная сводка по конкретному сотруднику."""
    query = update.callback_query
    await query.answer()
    if not is_admin_check(query):
        await query.answer("⛔ Только для руководителей.", show_alert=True); return
    tg_id = query.data.split("_", 1)[1]
    emp = emp_by_id(int(tg_id))
    if not emp:
        await query.message.reply_text("Сотрудник не найден."); return

    today = date.today().strftime("%Y-%m-%d")
    today_fmt = date.today().strftime("%d.%m.%Y")
    plan_ids   = {str(p["tg_id"]) for p in plans_today()}
    report_ids = {str(r["tg_id"]) for r in reports_today()}
    all_t = tasks_all()

    emp_tasks = [t for t in all_t if str(t["assigned_to_id"]) == tg_id and t["status"] != "done"]
    done_today = [t for t in all_t if str(t["assigned_to_id"]) == tg_id
                  and t["status"] == "done" and t.get("done_at","").startswith(today)]
    overdue = [t for t in emp_tasks if t.get("deadline","") < today]

    has_plan   = tg_id in plan_ids
    has_report = tg_id in report_ids

    lines = [
        f"👤 <b>{emp['full_name']}</b>",
        f"Отдел: {get_dept(tg_id)}",
        f"Сегодня {today_fmt}:",
        f"  план {'✅' if has_plan else '❌'}  отчёт {'✅' if has_report else '❌'}\n",
    ]
    if emp_tasks:
        lines.append(f"<b>Активные задачи ({len(emp_tasks)}):</b>")
        for t in emp_tasks:
            icon = "🔴" if t in overdue else "🔵"
            src = " <i>(план)</i>" if t.get("comment") == "из плана" else ""
            lines.append(f"  {icon} {t['title']} — до {fmt_dl(t['deadline'])}{src}")
    if done_today:
        lines.append(f"\n<b>Выполнено сегодня ({len(done_today)}):</b>")
        for t in done_today:
            lines.append(f"  ✅ {t['title']}")
    if not emp_tasks and not done_today:
        lines.append("Нет задач.")

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("👥 Другой сотрудник", callback_data="summary_person_list")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)


async def cb_period_week(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Сводка за текущую неделю."""
    query = update.callback_query
    await query.answer()
    if not is_admin_check(query):
        await query.answer("⛔ Только для руководителей.", show_alert=True); return
    today = date.today()
    week_start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
    week_end   = today.strftime("%Y-%m-%d")
    week_start_fmt = (today - timedelta(days=today.weekday())).strftime("%d.%m")
    week_end_fmt   = today.strftime("%d.%m.%Y")

    employees  = emp_employees()
    reps       = reports_for_period(week_start, week_end)
    plans_w    = plans_for_period(week_start, week_end)
    all_t      = tasks_all()
    done_w     = [t for t in all_t if t["status"] == "done"
                  and t.get("done_at","")[:10] >= week_start]
    over       = tasks_overdue()

    lines = [
        f"📅 <b>Сводка за неделю ({week_start_fmt} — {week_end_fmt})</b>\n",
        f"Сотрудников: {len(employees)}",
        f"Уникальных отчётов: {len(set(r['tg_id'] for r in reps))} / {len(employees)}",
        f"Задач выполнено: {len(done_w)}",
        f"Просроченных сейчас: {len(over)}\n",
        "<b>По сотрудникам:</b>",
    ]
    for e in employees:
        tid = str(e["tg_id"])
        r_c = sum(1 for r in reps if str(r["tg_id"]) == tid)
        d_c = sum(1 for t in done_w if str(t["assigned_to_id"]) == tid)
        o_c = sum(1 for t in over if str(t["assigned_to_id"]) == tid)
        icon = "✅" if r_c >= 4 else ("🟡" if r_c >= 1 else "🔴")
        lines.append(f"{icon} <b>{e['full_name']}</b> ({get_dept(e['tg_id'])}): "
                     f"отчётов {r_c}, задач ✅ {d_c}" + (f", ⚠️ {o_c}" if o_c else ""))

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("📅 За месяц", callback_data="period_month")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)


async def cb_period_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Сводка за текущий месяц."""
    query = update.callback_query
    await query.answer()
    if not is_admin_check(query):
        await query.answer("⛔ Только для руководителей.", show_alert=True); return
    today = date.today()
    month_start = today.replace(day=1).strftime("%Y-%m-%d")
    month_end   = today.strftime("%Y-%m-%d")
    month_label = today.strftime("%B %Y")

    import calendar
    _, last = calendar.monthrange(today.year, today.month)
    wd = sum(1 for d in range(1, today.day+1)
             if date(today.year, today.month, d).weekday() < 5)

    employees = emp_employees()
    reps      = reports_for_period(month_start, month_end)
    all_t     = tasks_all()
    done_m    = [t for t in all_t if t["status"] == "done"
                 and t.get("done_at","")[:10] >= month_start]
    over      = tasks_overdue()

    lines = [
        f"📆 <b>Сводка за {month_label}</b>\n",
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
                     f"отчётов {r_c}/{wd} ({pct}%), задач ✅ {d_c}")

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("📅 За неделю", callback_data="period_week")],
        [InlineKeyboardButton("◀️ Назад", callback_data="back_main")],
    ])
    await query.message.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=keyboard)


def build_admin_keyboard():
    """Главная клавиатура для группы руководителей."""
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📋 Задачи сегодня", callback_data="tasks_today"),
            InlineKeyboardButton("📋 Все активные", callback_data="show_all_tasks"),
        ],
        [
            InlineKeyboardButton("📊 Сводка по отделам", callback_data="summary_depts"),
            InlineKeyboardButton("👤 По сотруднику", callback_data="summary_person_list"),
        ],
        [
            InlineKeyboardButton("📅 Период: неделя", callback_data="period_week"),
            InlineKeyboardButton("📅 Период: месяц", callback_data="period_month"),
        ],
    ])


async def job_daily_digest(bot: Bot):
    today = datetime.now().strftime("%d.%m.%Y")
    employees = emp_employees()
    p_list = plans_today(); r_list = reports_today()
    plan_ids   = {str(p["tg_id"]) for p in p_list}
    report_ids = {str(r["tg_id"]) for r in r_list}

    submitted = len(report_ids)
    total = len(employees)
    no_plan   = [e["full_name"] for e in employees if str(e["tg_id"]) not in plan_ids]
    no_report = [e["full_name"] for e in employees if str(e["tg_id"]) not in report_ids]
    over = tasks_overdue()

    lines = [
        f"📊 <b>Ежедневная сводка — {today}</b>",
        f"Планов: {len(plan_ids)}/{total}   Отчётов: {submitted}/{total}\n",
    ]
    if no_plan:
        lines.append(f"❌ <b>Не сдали план:</b> {', '.join(no_plan)}")
    if no_report:
        lines.append(f"❌ <b>Не сдали отчёт:</b> {', '.join(no_report)}")
    if over:
        lines.append(f"\n⚠️ <b>Просроченных задач: {len(over)}</b>")
        for t in over[:3]:
            lines.append(f"  🔴 {t['assigned_to_name']}: {t['title']}")
        if len(over) > 3:
            lines.append(f"  ...и ещё {len(over)-3}")

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
    app.add_handler(CallbackQueryHandler(cb_back_main,            pattern="^back_main$"))
    app.add_handler(CallbackQueryHandler(cb_tasks_today,           pattern="^tasks_today$"))
    app.add_handler(CallbackQueryHandler(cb_show_all_tasks,        pattern="^show_all_tasks$"))
    app.add_handler(CallbackQueryHandler(cb_summary_depts,         pattern="^summary_depts$"))
    app.add_handler(CallbackQueryHandler(cb_summary_person_list,   pattern="^summary_person_list$"))
    app.add_handler(CallbackQueryHandler(cb_summary_person,        pattern="^person_"))
    app.add_handler(CallbackQueryHandler(cb_period_week,           pattern="^period_week$"))
    app.add_handler(CallbackQueryHandler(cb_period_month,          pattern="^period_month$"))

    logger.info("Bot starting...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
