import asyncio
import logging
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters
)
from config import BOT_TOKEN
from scheduler import build_scheduler
from handlers.start import build_handler as start_handler
from handlers.plans_reports import build_plan_handler, build_report_handler
from handlers.tasks import (
    build_task_handler, done_cmd, status_cmd, mytasks_cmd
)
from handlers.admin import build_handlers as admin_handlers

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)


async def post_init(application):
    """Запускаем планировщик после старта бота."""
    scheduler = build_scheduler(application.bot)
    scheduler.start()
    application.bot_data["scheduler"] = scheduler
    logger.info("Scheduler started.")

    # Устанавливаем меню команд
    await application.bot.set_my_commands([
        ("start",    "Регистрация / главное меню"),
        ("plan",     "Написать план на день"),
        ("report",   "Написать отчёт за день"),
        ("task",     "Поставить задачу @username Название | ДД.ММ.ГГГГ"),
        ("mytasks",  "Мои активные задачи"),
        ("done",     "Отметить задачу выполненной: /done ID"),
        ("status",   "Статус задачи: /status ID"),
        ("team",     "Сводка команды (для руководителей)"),
        ("tasks_all","Все задачи команды (для руководителей)"),
        ("makeadmin","Стать администратором (первый запуск)"),
    ])


def main():
    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .post_init(post_init)
        .build()
    )

    # Регистрируем обработчики
    app.add_handler(start_handler())
    app.add_handler(build_plan_handler())
    app.add_handler(build_report_handler())
    app.add_handler(build_task_handler())

    app.add_handler(CommandHandler("done",      done_cmd))
    app.add_handler(CommandHandler("status",    status_cmd))
    app.add_handler(CommandHandler("mytasks",   mytasks_cmd))

    for h in admin_handlers():
        app.add_handler(h)

    logger.info("Bot starting...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
