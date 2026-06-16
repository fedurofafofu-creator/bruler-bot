from datetime import datetime
from telegram import Bot
import sheets.employees as emp
import sheets.plans_reports as pr
import sheets.tasks as ts
from config import ADMIN_CHAT_ID


async def request_plans(bot: Bot):
    """11:00 — запросить планы у всех сотрудников."""
    employees = emp.get_employees()
    for e in employees:
        if pr.has_plan_today(int(e["tg_id"])):
            continue
        try:
            await bot.send_message(
                chat_id=int(e["tg_id"]),
                text=(
                    f"☀️ Доброе утро, {e['full_name']}!\n\n"
                    "Напиши план на сегодня командой /plan\n"
                    "или просто ответь на это сообщение своим планом."
                )
            )
        except Exception as ex:
            print(f"[request_plans] Ошибка для {e['full_name']}: {ex}")


async def remind_plans(bot: Bot):
    """11:30 — напомнить кто не сдал план."""
    employees = emp.get_employees()
    for e in employees:
        if not pr.has_plan_today(int(e["tg_id"])):
            try:
                await bot.send_message(
                    chat_id=int(e["tg_id"]),
                    text="🔔 Напоминание: не забудь написать план на день /plan"
                )
            except Exception as ex:
                print(f"[remind_plans] {ex}")


async def request_reports(bot: Bot):
    """19:00 — запросить отчёты."""
    employees = emp.get_employees()
    for e in employees:
        try:
            await bot.send_message(
                chat_id=int(e["tg_id"]),
                text=(
                    f"🌆 {e['full_name']}, рабочий день заканчивается!\n\n"
                    "Напиши отчёт за день: /report\n"
                    "Что сделал, что в процессе, что переносится."
                )
            )
        except Exception as ex:
            print(f"[request_reports] {ex}")


async def remind_reports(bot: Bot):
    """19:30 — напомнить кто не сдал отчёт."""
    employees = emp.get_employees()
    for e in employees:
        if not pr.has_report_today(int(e["tg_id"])):
            try:
                await bot.send_message(
                    chat_id=int(e["tg_id"]),
                    text="🔔 Напоминание: не забудь написать отчёт /report"
                )
            except Exception as ex:
                print(f"[remind_reports] {ex}")


async def send_daily_digest(bot: Bot):
    """19:30 — ежедневная сводка в чат руководителей."""
    today = datetime.now().strftime("%d.%m.%Y")
    employees = emp.get_employees()
    plans = pr.get_plans_today()
    reports = pr.get_reports_today()

    plan_map = {str(p["tg_id"]): p["plan_text"] for p in plans}
    report_map = {str(r["tg_id"]): r["report_text"] for r in reports}

    submitted_count = len(reports)
    total = len(employees)

    lines = [
        f"📊 <b>Сводка команды за {today}</b>",
        f"Отчётов сдано: {submitted_count}/{total}\n",
    ]

    for e in employees:
        tid = str(e["tg_id"])
        name = e["full_name"]
        has_plan = tid in plan_map
        has_report = tid in report_map

        if has_report:
            lines.append(f"✅ <b>{name}</b>")
            if has_plan:
                lines.append(f"  📋 <i>План:</i> {plan_map[tid][:200]}")
            lines.append(f"  📝 <i>Отчёт:</i> {report_map[tid][:300]}")
        elif has_plan:
            lines.append(f"🟡 <b>{name}</b> — план есть, отчёт не сдан")
            lines.append(f"  📋 {plan_map[tid][:200]}")
        else:
            lines.append(f"❌ <b>{name}</b> — ни плана, ни отчёта")

        lines.append("")

    # просроченные задачи
    overdue = ts.get_overdue_tasks()
    if overdue:
        lines.append(f"⚠️ <b>Просроченные задачи ({len(overdue)}):</b>")
        for t in overdue[:5]:
            lines.append(f"  🔴 {t['assigned_to_name']}: {t['title']} (срок {t['deadline']})")

    try:
        await bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text="\n".join(lines),
            parse_mode="HTML"
        )
    except Exception as ex:
        print(f"[daily_digest] {ex}")


async def ping_task_deadlines(bot: Bot):
    """Пинг исполнителей по дедлайнам (запускается утром)."""
    due_tomorrow = ts.get_tasks_due_tomorrow()
    for t in due_tomorrow:
        try:
            await bot.send_message(
                chat_id=int(t["assigned_to_id"]),
                text=(
                    f"⏰ Напоминание о задаче!\n\n"
                    f"<b>{t['title']}</b>\n"
                    f"Срок: завтра ({t['deadline'][8:]}.{t['deadline'][5:7]})\n"
                    f"ID: <code>{t['task_id']}</code>\n\n"
                    f"Не забудь отметить /done {t['task_id']}"
                ),
                parse_mode="HTML"
            )
        except Exception as ex:
            print(f"[ping_deadlines] {ex}")

    due_today = ts.get_tasks_due_today()
    for t in due_today:
        try:
            await bot.send_message(
                chat_id=int(t["assigned_to_id"]),
                text=(
                    f"🚨 Срок задачи истекает сегодня!\n\n"
                    f"<b>{t['title']}</b>\n"
                    f"ID: <code>{t['task_id']}</code>\n\n"
                    f"/done {t['task_id']} — отметить выполненной"
                ),
                parse_mode="HTML"
            )
        except Exception as ex:
            print(f"[ping_today] {ex}")
