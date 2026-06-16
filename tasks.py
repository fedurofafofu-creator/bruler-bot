import re
from telegram import Update
from telegram.ext import (ContextTypes, ConversationHandler,
                           CommandHandler, MessageHandler, filters)
import sheets.employees as emp
import sheets.tasks as ts

WAITING_TASK_TEXT = 20


async def task_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    /task @username Название задачи | ДД.ММ.ГГГГ
    или просто /task — войдёт в диалог
    """
    if not emp.is_registered(update.effective_user.id):
        await update.message.reply_text("Сначала /start.")
        return ConversationHandler.END

    args = " ".join(ctx.args) if ctx.args else ""
    # пробуем парсить сразу из аргументов
    parsed = _parse_task_args(args)
    if parsed:
        return await _create_task(update, ctx, parsed)

    await update.message.reply_text(
        "📌 Напиши задачу в формате:\n\n"
        "<b>@username Название задачи | ДД.ММ.ГГГГ</b>\n\n"
        "Пример:\n"
        "<code>@tanya Собрать контент-план | 20.06.2026</code>",
        parse_mode="HTML"
    )
    return WAITING_TASK_TEXT


async def receive_task_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    parsed = _parse_task_args(update.message.text.strip())
    if not parsed:
        await update.message.reply_text(
            "Не могу разобрать. Формат:\n"
            "<code>@username Название | ДД.ММ.ГГГГ</code>",
            parse_mode="HTML"
        )
        return WAITING_TASK_TEXT
    return await _create_task(update, ctx, parsed)


def _parse_task_args(text: str):
    """Возвращает (username, title, deadline) или None."""
    m = re.match(r"@(\S+)\s+(.+?)\s*\|\s*(\d{2}\.\d{2}\.\d{4})", text)
    if not m:
        return None
    username, title, deadline_raw = m.groups()
    # конвертируем ДД.ММ.ГГГГ → ГГГГ-ММ-ДД
    parts = deadline_raw.split(".")
    deadline = f"{parts[2]}-{parts[1]}-{parts[0]}"
    return username, title.strip(), deadline


async def _create_task(update, ctx, parsed):
    username, title, deadline = parsed
    creator = emp.get_by_tg_id(update.effective_user.id)

    # ищем исполнителя по username
    assignee = None
    for e in emp.get_all():
        if e["username"].lstrip("@").lower() == username.lower():
            assignee = e
            break

    if not assignee:
        await update.message.reply_text(
            f"❌ Пользователь @{username} не найден в системе.\n"
            "Убедись, что он зарегистрирован через /start."
        )
        return ConversationHandler.END

    task_id = ts.create_task(
        created_by_id=update.effective_user.id,
        created_by_name=creator["full_name"],
        assigned_to_id=assignee["tg_id"],
        assigned_to_name=assignee["full_name"],
        title=title,
        deadline=deadline,
    )

    # уведомляем исполнителя
    deadline_fmt = deadline[8:] + "." + deadline[5:7] + "." + deadline[:4]
    try:
        await ctx.bot.send_message(
            chat_id=int(assignee["tg_id"]),
            text=(
                f"📌 Тебе поставлена задача!\n\n"
                f"<b>{title}</b>\n"
                f"От: {creator['full_name']}\n"
                f"Срок: {deadline_fmt}\n"
                f"ID задачи: <code>{task_id}</code>\n\n"
                f"Когда выполнишь — напиши /done {task_id}"
            ),
            parse_mode="HTML"
        )
    except Exception:
        pass

    await update.message.reply_text(
        f"✅ Задача создана!\n\n"
        f"<b>{title}</b>\n"
        f"Исполнитель: {assignee['full_name']}\n"
        f"Срок: {deadline_fmt}\n"
        f"ID: <code>{task_id}</code>",
        parse_mode="HTML"
    )
    return ConversationHandler.END


async def done_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ctx.args:
        await update.message.reply_text("Укажи ID задачи: /done ABCD1234")
        return
    task_id = ctx.args[0].upper()
    comment = " ".join(ctx.args[1:]) if len(ctx.args) > 1 else ""

    task = ts.get_task(task_id)
    if not task:
        await update.message.reply_text(f"❌ Задача {task_id} не найдена.")
        return

    tg_id = update.effective_user.id
    if str(task["assigned_to_id"]) != str(tg_id) and not emp.is_admin(tg_id):
        await update.message.reply_text("❌ Это не твоя задача.")
        return

    ts.mark_done(task_id, comment)

    # уведомляем постановщика
    try:
        await ctx.bot.send_message(
            chat_id=int(task["created_by_id"]),
            text=(
                f"✅ Задача выполнена!\n\n"
                f"<b>{task['title']}</b>\n"
                f"Выполнил: {task['assigned_to_name']}\n"
                f"ID: <code>{task_id}</code>"
                + (f"\nКомментарий: {comment}" if comment else "")
            ),
            parse_mode="HTML"
        )
    except Exception:
        pass

    await update.message.reply_text(f"✅ Задача <code>{task_id}</code> отмечена выполненной!", parse_mode="HTML")


async def status_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ctx.args:
        await update.message.reply_text("Укажи ID: /status ABCD1234")
        return
    task_id = ctx.args[0].upper()
    task = ts.get_task(task_id)
    if not task:
        await update.message.reply_text(f"❌ Задача {task_id} не найдена.")
        return

    status_labels = {"open": "🔵 Открыта", "in_progress": "🟡 В работе",
                     "done": "✅ Выполнена", "overdue": "🔴 Просрочена"}
    status = status_labels.get(task["status"], task["status"])
    dl = task["deadline"]
    dl_fmt = dl[8:] + "." + dl[5:7] + "." + dl[:4] if dl else "—"

    await update.message.reply_text(
        f"📌 <b>{task['title']}</b>\n\n"
        f"Статус: {status}\n"
        f"Исполнитель: {task['assigned_to_name']}\n"
        f"Постановщик: {task['created_by_name']}\n"
        f"Срок: {dl_fmt}\n"
        f"ID: <code>{task_id}</code>",
        parse_mode="HTML"
    )


async def mytasks_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tg_id = update.effective_user.id
    tasks = ts.get_tasks_for_user(tg_id)
    if not tasks:
        await update.message.reply_text("✅ Нет активных задач!")
        return

    lines = ["📋 <b>Твои активные задачи:</b>\n"]
    status_labels = {"open": "🔵", "in_progress": "🟡", "overdue": "🔴"}
    for t in tasks:
        dl = t["deadline"]
        dl_fmt = dl[8:] + "." + dl[5:7] if dl else "—"
        icon = status_labels.get(t["status"], "⚪")
        lines.append(f"{icon} <code>{t['task_id']}</code> {t['title']} — до {dl_fmt}")

    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


def build_task_handler():
    return ConversationHandler(
        entry_points=[CommandHandler("task", task_cmd)],
        states={WAITING_TASK_TEXT: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_task_text)]},
        fallbacks=[CommandHandler("cancel", lambda u, c: ConversationHandler.END)],
    )
