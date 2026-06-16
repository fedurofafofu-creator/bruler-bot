from datetime import datetime
from telegram import Update
from telegram.ext import ContextTypes, CommandHandler
import sheets.employees as emp
import sheets.plans_reports as pr
import sheets.tasks as ts


def _admin_only(func):
    async def wrapper(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        if not emp.is_admin(update.effective_user.id):
            await update.message.reply_text("⛔ Только для руководителей.")
            return
        return await func(update, ctx)
    return wrapper


async def makeadmin_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Первый запуск: /makeadmin — делает себя админом если нет ни одного."""
    admins = emp.get_admins()
    tg_id = update.effective_user.id
    if admins and not emp.is_admin(tg_id):
        await update.message.reply_text("⛔ Уже есть администраторы. Обратись к ним.")
        return
    if not emp.is_registered(tg_id):
        await update.message.reply_text("Сначала зарегистрируйся через /start.")
        return
    emp.set_admin(tg_id)
    await update.message.reply_text("✅ Ты теперь администратор!")


@_admin_only
async def team_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Кто сдал план и отчёт сегодня."""
    employees = emp.get_employees()
    plans = pr.get_plans_today()
    reports = pr.get_reports_today()

    plan_ids = {str(p["tg_id"]) for p in plans}
    report_ids = {str(r["tg_id"]) for r in reports}

    lines = [f"📊 <b>Команда сегодня, {datetime.now().strftime('%d.%m.%Y')}:</b>\n"]
    for e in employees:
        has_p = "✅" if str(e["tg_id"]) in plan_ids else "❌"
        has_r = "✅" if str(e["tg_id"]) in report_ids else "—"
        lines.append(f"{e['full_name']}  план {has_p}  отчёт {has_r}")

    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


@_admin_only
async def tasks_all_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Все активные задачи команды."""
    open_tasks = ts.get_open_tasks()
    overdue = ts.get_overdue_tasks()

    lines = ["📋 <b>Все активные задачи:</b>\n"]
    if not open_tasks and not overdue:
        lines.append("Нет активных задач.")
    else:
        for t in overdue:
            dl = t["deadline"]
            dl_fmt = dl[8:] + "." + dl[5:7] if dl else "—"
            lines.append(f"🔴 <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']} (просрочено, срок {dl_fmt})")
        for t in open_tasks:
            if t in overdue:
                continue
            dl = t["deadline"]
            dl_fmt = dl[8:] + "." + dl[5:7] if dl else "—"
            lines.append(f"🔵 <code>{t['task_id']}</code> {t['assigned_to_name']}: {t['title']} — до {dl_fmt}")

    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


@_admin_only
async def setadmin_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/ setadmin @username"""
    if not ctx.args:
        await update.message.reply_text("Укажи: /setadmin @username")
        return
    username = ctx.args[0].lstrip("@")
    for e in emp.get_all():
        if e["username"].lstrip("@").lower() == username.lower():
            emp.set_admin(int(e["tg_id"]))
            await update.message.reply_text(f"✅ {e['full_name']} теперь администратор.")
            return
    await update.message.reply_text(f"❌ @{username} не найден.")


def build_handlers():
    return [
        CommandHandler("makeadmin", makeadmin_cmd),
        CommandHandler("team", team_cmd),
        CommandHandler("tasks_all", tasks_all_cmd),
        CommandHandler("setadmin", setadmin_cmd),
    ]
