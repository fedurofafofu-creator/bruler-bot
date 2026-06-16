# Bruler Team Bot

Telegram-бот для трекинга планов, отчётов и задач команды Brûler d'Amour.

## Стек
- Python + python-telegram-bot 21
- Google Sheets (gspread) — база данных
- APScheduler — расписание
- Railway — хостинг
- GitHub — хранение кода + автодеплой

## Переменные окружения (Railway → Variables)

| Переменная | Значение |
|---|---|
| `BOT_TOKEN` | токен от @BotFather |
| `SPREADSHEET_ID` | `1LCmoydfS73DKwjQcwMSnpOZxRHiKMAMjOZRRk0AwYnE` |
| `ADMIN_CHAT_ID` | `-5486975908` |
| `GOOGLE_CREDENTIALS_JSON` | полный JSON сервисного аккаунта (одной строкой) |

## Расписание

| Время (МСК) | Действие |
|---|---|
| 10:00 пн–пт | Пинг по дедлайнам задач |
| 11:00 пн–пт | Запрос плана у всех сотрудников |
| 11:30 пн–пт | Напоминание кто не сдал план |
| 19:00 пн–пт | Запрос отчёта за день |
| 19:30 пн–пт | Напоминание кто не сдал отчёт |
| 19:35 пн–пт | Ежедневная сводка → руководители |
| Пятница 18:00 | Еженедельный аудит → руководители |
| 28-е число 17:00 | Ежемесячный аудит → руководители |

## Команды сотрудников

```
/start   — регистрация
/plan    — план на день
/report  — отчёт за день
/task    — поставить задачу: /task @username Название | ДД.ММ.ГГГГ
/done    — отметить выполненной: /done ID
/status  — статус задачи: /status ID
/mytasks — мои активные задачи
```

## Команды руководителей

```
/makeadmin  — стать админом (только первый раз)
/setadmin   — назначить админа: /setadmin @username
/team       — кто сдал план/отчёт сегодня
/tasks_all  — все активные задачи команды
```

## Деплой на Railway

1. Залей код в GitHub-репозиторий
2. Создай проект на railway.app → Connect GitHub
3. Добавь переменные окружения в Variables
4. Railway автоматически задеплоит при push в main

## Получить GOOGLE_CREDENTIALS_JSON

1. Google Cloud Console → IAM → Service Accounts
2. Выбери `bot-sheets@brulerbudgetingbot.iam.gserviceaccount.com`
3. Keys → Add Key → JSON → скачай файл
4. Содержимое файла целиком вставь как значение `GOOGLE_CREDENTIALS_JSON`
5. Таблицу расшарь на этот email с правами редактора

## Первый запуск

1. Задеплой бота
2. Напиши боту /start → зарегистрируйся
3. Напиши /makeadmin → станешь администратором
4. Остальные сотрудники пишут /start и регистрируются
5. Ты назначаешь других руководителей через /setadmin @username
