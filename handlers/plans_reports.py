from telegram import Update
from telegram.ext import (ContextTypes, ConversationHandler,
                           CommandHandler, MessageHandler, filters)
import sheets.employees as emp
import sheets.plans_reports as pr

WAITING_PLAN = 10
WAITING_REPORT = 11


def _check_registered(update):
    return emp.is_registered(update.effective_user.id)


async def plan_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not _check_registered(update):
        await update.message.reply_text("Сначала напиши /start для регистрации.")
        return ConversationHandler.END
    user = emp.get_by_tg_id(update.effective_user.id)
    if pr.has_plan_today(update.effective_user.id):
        await update.message.reply_text(
            f"✅ {user['full_name']}, план на сегодня уже записан.\n"
            "Если хочешь обновить — напиши план заново:"
        )
    else:
        await update.message.reply_text(
            f"📋 {user['full_name']}, напиши план на сегодня.\n\n"
            "Можно списком, в свободной форме — как удобно:"
        )
    return WAITING_PLAN


async def receive_plan(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    user = emp.get_by_tg_id(tg_id)
    text = update.message.text.strip()
    pr.save_plan(tg_id, user["full_name"], text)
    await update.message.reply_text(
        "✅ План записан! Увидимся в 19:00 — напишу про отчёт."
    )
    return ConversationHandler.END


async def report_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not _check_registered(update):
        await update.message.reply_text("Сначала напиши /start для регистрации.")
        return ConversationHandler.END
    user = emp.get_by_tg_id(update.effective_user.id)
    await update.message.reply_text(
        f"📊 {user['full_name']}, напиши отчёт за день.\n\n"
        "Что сделал, что в процессе, что переносится:"
    )
    return WAITING_REPORT


async def receive_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    user = emp.get_by_tg_id(tg_id)
    text = update.message.text.strip()
    pr.save_report(tg_id, user["full_name"], text)
    await update.message.reply_text("✅ Отчёт записан! Хорошего вечера 🌙")
    return ConversationHandler.END


async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отменено.")
    return ConversationHandler.END


def build_plan_handler():
    return ConversationHandler(
        entry_points=[CommandHandler("plan", plan_cmd)],
        states={WAITING_PLAN: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_plan)]},
        fallbacks=[CommandHandler("cancel", cancel)],
    )


def build_report_handler():
    return ConversationHandler(
        entry_points=[CommandHandler("report", report_cmd)],
        states={WAITING_REPORT: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_report)]},
        fallbacks=[CommandHandler("cancel", cancel)],
    )
