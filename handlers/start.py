from telegram import Update
from telegram.ext import ContextTypes, ConversationHandler, CommandHandler, MessageHandler, filters
import sheets.employees as emp

WAITING_NAME = 1


async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    if emp.is_registered(tg_id):
        user = emp.get_by_tg_id(tg_id)
        role_label = "руководитель" if user["role"] == "admin" else "сотрудник"
        await update.message.reply_text(
            f"👋 Привет, {user['full_name']}!\n"
            f"Ты зарегистрирован как {role_label}.\n\n"
            f"Команды:\n"
            f"/plan — написать план на день\n"
            f"/report — написать отчёт за день\n"
            f"/task — поставить задачу\n"
            f"/mytasks — мои задачи\n"
            f"/done ID — отметить задачу выполненной\n"
            f"/status ID — статус задачи"
        )
        return ConversationHandler.END

    await update.message.reply_text(
        "👋 Привет! Это трекер команды Brûler d'Amour.\n\n"
        "Напиши своё имя и фамилию (как тебя записать в системе):"
    )
    return WAITING_NAME


async def receive_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    username = update.effective_user.username or ""
    full_name = update.message.text.strip()

    emp.register(tg_id, username, full_name, role="employee")

    await update.message.reply_text(
        f"✅ Готово, {full_name}! Ты зарегистрирован.\n\n"
        f"Каждый день в 11:00 я попрошу план, в 19:00 — отчёт.\n"
        f"Можешь написать их прямо сейчас:\n\n"
        f"/plan — план на день\n"
        f"/report — отчёт за день"
    )
    return ConversationHandler.END


def build_handler():
    return ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={WAITING_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_name)]},
        fallbacks=[CommandHandler("start", start)],
    )
