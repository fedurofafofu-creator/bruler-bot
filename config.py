import os
import json

BOT_TOKEN = os.getenv("BOT_TOKEN", "8767537310:AAHK1-RmvgH6yF6ZShQFq3A1DeLU0uFsAMs")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1LCmoydfS73DKwjQcwMSnpOZxRHiKMAMjOZRRk0AwYnE")
ADMIN_CHAT_ID = int(os.getenv("ADMIN_CHAT_ID", "-5486975908"))
TIMEZONE = "Europe/Moscow"

GOOGLE_SERVICE_ACCOUNT_EMAIL = "bot-sheets@brulerbudgetingbot.iam.gserviceaccount.com"

# Загружается из env как JSON-строка (Railway → Variables)
# Формат: полный JSON сервисного аккаунта Google
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "")

# Листы таблицы
SHEET_EMPLOYEES = "employees"
SHEET_PLANS     = "plans"
SHEET_REPORTS   = "reports"
SHEET_TASKS     = "tasks"
SHEET_AUDIT     = "audit"

# Расписание (МСК)
PLAN_HOUR    = 11
PLAN_MINUTE  = 0
REPORT_HOUR  = 19
REPORT_MINUTE = 0
DIGEST_HOUR  = 19
DIGEST_MINUTE = 30
REMINDER_PLAN_HOUR   = 11
REMINDER_PLAN_MINUTE = 30
REMINDER_REPORT_HOUR   = 19
REMINDER_REPORT_MINUTE = 30
WEEKLY_AUDIT_DAY  = 4   # пятница (0=пн)
WEEKLY_AUDIT_HOUR = 18
