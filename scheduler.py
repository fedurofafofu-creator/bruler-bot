from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
from telegram import Bot
import pytz

from notifications.digest import (
    request_plans, remind_plans,
    request_reports, remind_reports,
    send_daily_digest, ping_task_deadlines,
)
from notifications.audit import send_weekly_audit, send_monthly_audit
from config import TIMEZONE

tz = pytz.timezone(TIMEZONE)


def build_scheduler(bot: Bot) -> AsyncIOScheduler:
    scheduler = AsyncIOScheduler(timezone=tz)

    # 11:00 — запрос плана (пн-пт)
    scheduler.add_job(
        request_plans, CronTrigger(day_of_week="mon-fri", hour=11, minute=0, timezone=tz),
        args=[bot], id="request_plans"
    )
    # 11:30 — напоминание кто не сдал план
    scheduler.add_job(
        remind_plans, CronTrigger(day_of_week="mon-fri", hour=11, minute=30, timezone=tz),
        args=[bot], id="remind_plans"
    )
    # 10:00 — пинг по дедлайнам (сегодня и завтра)
    scheduler.add_job(
        ping_task_deadlines, CronTrigger(day_of_week="mon-fri", hour=10, minute=0, timezone=tz),
        args=[bot], id="ping_deadlines"
    )
    # 19:00 — запрос отчёта
    scheduler.add_job(
        request_reports, CronTrigger(day_of_week="mon-fri", hour=19, minute=0, timezone=tz),
        args=[bot], id="request_reports"
    )
    # 19:30 — напоминание + ежедневная сводка руководителям
    scheduler.add_job(
        remind_reports, CronTrigger(day_of_week="mon-fri", hour=19, minute=30, timezone=tz),
        args=[bot], id="remind_reports"
    )
    scheduler.add_job(
        send_daily_digest, CronTrigger(day_of_week="mon-fri", hour=19, minute=35, timezone=tz),
        args=[bot], id="daily_digest"
    )
    # Пятница 18:00 — еженедельный аудит
    scheduler.add_job(
        send_weekly_audit, CronTrigger(day_of_week="fri", hour=18, minute=0, timezone=tz),
        args=[bot], id="weekly_audit"
    )
    # Последний рабочий день месяца 17:00 — ежемесячный аудит
    scheduler.add_job(
        send_monthly_audit, CronTrigger(day=28, hour=17, minute=0, timezone=tz),
        args=[bot], id="monthly_audit"
    )

    return scheduler
