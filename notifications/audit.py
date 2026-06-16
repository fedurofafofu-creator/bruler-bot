from datetime import datetime, timedelta
from telegram import Bot
import sheets.employees as emp
import sheets.plans_reports as pr
import sheets.tasks as ts
from config import ADMIN_CHAT_ID


async def send_weekly_audit(bot: Bot):
    """Пятница 18:00 — еженедельный аудит."""
    today = datetime.now()
    iso_week = today.strftime("%Y-W%W")
    week_start = (today - timedelta(days=today.weekday())).strftime("%d.%m")
    week_end = today.strftime("%d.%m.%Y")

    employees = emp.get_employees()
    plans = pr.get_plans_for_week(iso_week)
    reports = pr.get_reports_for_week(iso_week)

    # задачи
    all_tasks = ts.get_all_tasks()
    done_this_week = [
        t for t in all_tasks
        if t["status"] == "done" and t.get("done_at", "").startswith(today.strftime("%Y"))
        and _is_this_week(t.get("done_at", ""), iso_week)
    ]
    overdue = ts.get_overdue_tasks()

    lines = [
        f"📅 <b>Еженедельный аудит ({week_start} — {week_end})</b>\n",
        f"Сотрудников: {len(employees)}",
        f"Планов сдано: {len(set(p['tg_id'] for p in plans))} / {len(employees)}",
        f"Отчётов сдано: {len(set(r['tg_id'] for r in reports))} / {len(employees)}",
        f"Задач выполнено за неделю: {len(done_this_week)}",
        f"Просроченных задач: {len(overdue)}\n",
    ]

    lines.append("<b>По сотрудникам:</b>")
    for e in employees:
        tid = str(e["tg_id"])
        p_count = sum(1 for p in plans if str(p["tg_id"]) == tid)
        r_count = sum(1 for r in reports if str(r["tg_id"]) == tid)
        done_count = sum(1 for t in done_this_week if str(t["assigned_to_id"]) == tid)
        overdue_count = sum(1 for t in overdue if str(t["assigned_to_id"]) == tid)

        status = "✅" if r_count >= 4 else ("🟡" if r_count >= 2 else "🔴")
        lines.append(
            f"{status} <b>{e['full_name']}</b>: "
            f"планов {p_count}, отчётов {r_count}, "
            f"задач выполнено {done_count}"
            + (f", просрочено {overdue_count}" if overdue_count else "")
        )

    if overdue:
        lines.append(f"\n⚠️ <b>Просроченные задачи:</b>")
        for t in overdue:
            lines.append(f"  🔴 {t['assigned_to_name']}: {t['title']} (срок {t['deadline']})")

    try:
        await bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text="\n".join(lines),
            parse_mode="HTML"
        )
    except Exception as ex:
        print(f"[weekly_audit] {ex}")


async def send_monthly_audit(bot: Bot):
    """Последний рабочий день — ежемесячный аудит."""
    today = datetime.now()
    month_name = today.strftime("%B %Y")
    current_month = today.strftime("%Y-%m")

    employees = emp.get_employees()
    all_plans = pr.get_plans_for_week("ALL")   # получим все и отфильтруем
    all_reports = pr.get_reports_for_week("ALL")

    # фильтруем по месяцу
    month_plans = [p for p in _all_plans_raw() if p["date"].startswith(current_month)]
    month_reports = [r for r in _all_reports_raw() if r["date"].startswith(current_month)]

    all_tasks = ts.get_all_tasks()
    done_month = [t for t in all_tasks
                  if t["status"] == "done"
                  and t.get("done_at", "").startswith(current_month)]
    overdue = ts.get_overdue_tasks()

    working_days = _working_days_this_month(today)

    lines = [
        f"📆 <b>Ежемесячный аудит — {month_name}</b>\n",
        f"Рабочих дней: {working_days}",
        f"Задач выполнено: {len(done_month)}",
        f"Просроченных: {len(overdue)}\n",
        "<b>По сотрудникам:</b>"
    ]

    for e in employees:
        tid = str(e["tg_id"])
        p_count = sum(1 for p in month_plans if str(p["tg_id"]) == tid)
        r_count = sum(1 for r in month_reports if str(r["tg_id"]) == tid)
        done_count = sum(1 for t in done_month if str(t["assigned_to_id"]) == tid)
        pct = round(r_count / max(working_days, 1) * 100)
        status = "✅" if pct >= 80 else ("🟡" if pct >= 50 else "🔴")
        lines.append(
            f"{status} <b>{e['full_name']}</b>: "
            f"отчётов {r_count}/{working_days} ({pct}%), "
            f"задач выполнено {done_count}"
        )

    try:
        await bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text="\n".join(lines),
            parse_mode="HTML"
        )
    except Exception as ex:
        print(f"[monthly_audit] {ex}")


def _is_this_week(dt_str: str, iso_week: str) -> bool:
    try:
        d = datetime.strptime(dt_str[:10], "%Y-%m-%d")
        return d.strftime("%Y-W%W") == iso_week
    except Exception:
        return False


def _working_days_this_month(today: datetime) -> int:
    import calendar
    year, month = today.year, today.month
    _, last_day = calendar.monthrange(year, month)
    count = 0
    for day in range(1, last_day + 1):
        if datetime(year, month, day).weekday() < 5:
            count += 1
    return count


def _all_plans_raw():
    from sheets.client import get_sheet
    s = get_sheet("plans")
    return s.get_all_records()


def _all_reports_raw():
    from sheets.client import get_sheet
    s = get_sheet("reports")
    return s.get_all_records()
