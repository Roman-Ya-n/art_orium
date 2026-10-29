"""
Debate tournament participation bot.

Install:
    pip install "python-telegram-bot>=21,<23" gspread google-auth-oauthlib

Run:
    Set TELEGRAM_BOT_TOKEN in your environment, then:
    python bot.py

This uses the same Google Sheets OAuth files as acc_check.py:
    credentials.json
    token.json

Important:
- Participants must open the bot and press Start before it can message them.
- The bot matches their Telegram username against the configured columns.
- It adds tracking columns at the right side of each worksheet; it does not overwrite
  the original form-response columns.
"""

import logging
import os
from datetime import datetime, timezone

import gspread
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
)

# ---------------- Configuration ----------------

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
SHEET_ID = os.getenv("SHEET_ID", "").strip()

# Worksheet names and columns containing Telegram usernames (1-based).
# Based on the existing acc_check.py.
SHEETS_CONFIG = {
    "Команди": {"tag_cols": [5, 11], "header_row": 1},
    "Одиночки": {"tag_cols": [3], "header_row": 1},
}

CREDENTIALS_FILE = "credentials.json"
AUTHORIZED_USER_FILE = "token.json"

# Optional: comma-separated Telegram numeric user IDs allowed to run /broadcast.
# Example on Windows PowerShell:
#   $env:ADMIN_IDS="123456789,987654321"
ADMIN_IDS = {
    item.strip()
    for item in os.getenv("ADMIN_IDS", "").split(",")
    if item.strip()
}

STATUS_PENDING = "Очікує відповіді"
STATUS_YES = "Підтверджую участь"
STATUS_NO = "Не зможу бути"

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger("debate-confirmation-bot")

# Cache worksheets and row mappings in memory. They are refreshed when needed.
_gc = None
_spreadsheet = None


# ---------------- Google Sheets helpers ----------------

def get_spreadsheet():
    """Authorize with the existing gspread OAuth setup and open the spreadsheet."""
    global _gc, _spreadsheet
    if _spreadsheet is None:
        _gc = gspread.oauth(
            credentials_filename=CREDENTIALS_FILE,
            authorized_user_filename=AUTHORIZED_USER_FILE,
        )
        _spreadsheet = _gc.open_by_key(SHEET_ID)
    return _spreadsheet


def tracking_headers(tag_col: int) -> tuple[str, str, str]:
    """Distinct tracking columns for each username column in a row."""
    return (
        f"Bot chat_id (tag col {tag_col})",
        f"Підтвердження (tag col {tag_col})",
        f"Час відповіді (tag col {tag_col})",
    )


def ensure_tracking_columns(ws, header_row: int, tag_cols: list[int]) -> dict[int, tuple[int, int, int]]:
    """
    Ensure each username column has its own chat_id, status, and response-time columns.
    Returns {tag_column: (chat_id_column, status_column, response_time_column)}.
    """
    headers = ws.row_values(header_row)
    while len(headers) < ws.col_count:
        headers.append("")

    result = {}
    changed = False

    for tag_col in tag_cols:
        cols = []
        for label in tracking_headers(tag_col):
            try:
                col = headers.index(label) + 1
            except ValueError:
                # Add new tracking fields at the end, never between existing columns.
                col = len(headers) + 1
                headers.append(label)
                changed = True
            cols.append(col)
        result[tag_col] = tuple(cols)

    if changed:
        # Update only header cells that were newly appended.
        for tag_col, cols in result.items():
            for col, label in zip(cols, tracking_headers(tag_col)):
                if col > ws.col_count:
                    ws.add_cols(col - ws.col_count)
                current = ws.cell(header_row, col).value
                if current != label:
                    ws.update_cell(header_row, col, label)

    return result


def get_configured_worksheets():
    """Return configured worksheets with their tracking column indexes."""
    spreadsheet = get_spreadsheet()
    configured = []
    for title, cfg in SHEETS_CONFIG.items():
        try:
            ws = spreadsheet.worksheet(title)
        except gspread.WorksheetNotFound:
            logger.warning("Worksheet not found, skipping: %s", title)
            continue

        tracking = ensure_tracking_columns(ws, cfg["header_row"], cfg["tag_cols"])
        configured.append((ws, cfg, tracking))
    return configured


def normalize_username(value: str | None) -> str:
    if not value:
        return ""
    return value.strip().lstrip("@").lower()


def find_user_rows(username: str):
    """
    Find exact username matches in configured columns.
    Each result identifies a particular username cell, including separate columns
    for two team members on the same spreadsheet row.
    """
    target = normalize_username(username)
    if not target:
        return []

    matches = []
    for ws, cfg, tracking in get_configured_worksheets():
        for tag_col in cfg["tag_cols"]:
            values = ws.col_values(tag_col)
            start_row = cfg["header_row"] + 1
            for row_number, raw in enumerate(values[start_row - 1 :], start=start_row):
                if normalize_username(raw) == target:
                    chat_col, status_col, time_col = tracking[tag_col]
                    matches.append({
                        "worksheet": ws,
                        "worksheet_title": ws.title,
                        "row": row_number,
                        "tag_col": tag_col,
                        "chat_col": chat_col,
                        "status_col": status_col,
                        "time_col": time_col,
                        "username": normalize_username(raw),
                    })
    return matches


def save_chat_id(match: dict, chat_id: int):
    ws = match["worksheet"]
    ws.update_cell(match["row"], match["chat_col"], str(chat_id))
    current_status = ws.cell(match["row"], match["status_col"]).value
    if not current_status:
        ws.update_cell(match["row"], match["status_col"], STATUS_PENDING)


def save_response(match: dict, status: str):
    ws = match["worksheet"]
    ws.update_cell(match["row"], match["status_col"], status)
    ws.update_cell(
        match["row"],
        match["time_col"],
        datetime.now(timezone.utc).astimezone().strftime("%Y-%m-%d %H:%M:%S %Z"),
    )


def iter_registered_participants():
    """Yield each row/username cell with a saved chat_id and pending status."""
    for ws, cfg, tracking in get_configured_worksheets():
        for tag_col in cfg["tag_cols"]:
            chat_col, status_col, time_col = tracking[tag_col]
            usernames = ws.col_values(tag_col)
            chat_ids = ws.col_values(chat_col)
            statuses = ws.col_values(status_col)
            start_row = cfg["header_row"] + 1
            max_len = max(len(usernames), len(chat_ids), len(statuses))
            for row in range(start_row, max_len + 1):
                username = normalize_username(usernames[row - 1] if row <= len(usernames) else "")
                chat_id = chat_ids[row - 1].strip() if row <= len(chat_ids) else ""
                status = statuses[row - 1].strip() if row <= len(statuses) else ""
                if username and chat_id and status != STATUS_YES and status != STATUS_NO:
                    yield {
                        "worksheet": ws,
                        "worksheet_title": ws.title,
                        "row": row,
                        "tag_col": tag_col,
                        "chat_col": chat_col,
                        "status_col": status_col,
                        "time_col": time_col,
                        "username": username,
                        "chat_id": chat_id,
                    }


# ---------------- Telegram handlers ----------------

def confirmation_keyboard():
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Підтверджую участь", callback_data="confirm:yes"),
            InlineKeyboardButton("❌ Не зможу бути", callback_data="confirm:no"),
        ]
    ])


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    message = update.effective_message

    if not user or not message:
        return

    username = normalize_username(user.username)
    if not username:
        await message.reply_text(
            "У тебе не встановлений Telegram username.\n"
            "Будь ласка, встанови його в налаштуваннях Telegram і натисни /start ще раз."
        )
        return

    try:
        matches = find_user_rows(username)
    except Exception:
        logger.exception("Google Sheets lookup failed")
        await message.reply_text(
            "Не вдалося перевірити реєстрацію через Google Sheets. "
            "Спробуй трохи пізніше або напиши організаторам."
        )
        return

    if not matches:
        await message.reply_text(
            f"Не знайшов реєстрацію для @{username}.\n"
            "Перевір, чи саме цей Telegram username вказаний у формі. "
            "Якщо все правильно — напиши організаторам."
        )
        return

    if len(matches) > 1:
        logger.warning("Username %s matched multiple rows", username)
        await message.reply_text(
            "Знайшов кілька записів із цим username, тому не можу безпечно "
            "прив'язати відповідь. Будь ласка, звернися до організаторів."
        )
        return

    match = matches[0]
    try:
        save_chat_id(match, user.id)
    except Exception:
        logger.exception("Could not save chat_id")
        await message.reply_text(
            "Знайшов твою реєстрацію, але не зміг зберегти прив'язку в таблиці. "
            "Напиши організаторам."
        )
        return

    context.user_data["participant_match"] = {
        "worksheet_title": match["worksheet_title"],
        "row": match["row"],
        "tag_col": match["tag_col"],
        "username": username,
    }

    await message.reply_text(
        "Привіт! 👋\n"
        "Знайшов твою реєстрацію на дебатний турнір.\n\n"
        "Будь ласка, підтверди, чи плануєш брати участь:",
        reply_markup=confirmation_keyboard(),
    )


async def handle_confirmation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if not query:
        return
    await query.answer()

    user = query.from_user
    username = normalize_username(user.username)
    if not username:
        await query.edit_message_text("Не бачу твого Telegram username. Напиши організаторам.")
        return

    try:
        matches = find_user_rows(username)
    except Exception:
        logger.exception("Google Sheets lookup failed during confirmation")
        await query.edit_message_text(
            "Не вдалося записати відповідь через технічну помилку. "
            "Спробуй пізніше або напиши організаторам."
        )
        return

    if len(matches) != 1:
        await query.edit_message_text(
            "Не вдалося однозначно знайти твою реєстрацію. Напиши організаторам."
        )
        return

    status = STATUS_YES if query.data == "confirm:yes" else STATUS_NO
    try:
        match = matches[0]
        save_chat_id(match, user.id)
        save_response(match, status)
    except Exception:
        logger.exception("Could not save participant response")
        await query.edit_message_text(
            "Не вдалося зберегти відповідь. Спробуй ще раз або напиши організаторам."
        )
        return

    if status == STATUS_YES:
        await query.edit_message_text("Дякую! ✅ Участь підтверджено. До зустрічі на турнірі!")
    else:
        await query.edit_message_text(
            "Дякую, що повідомив(-ла). ❌ Зафіксував, що ти не зможеш бути."
        )


async def broadcast(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin command: /broadcast sends the confirmation prompt to registered participants."""
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return

    if str(user.id) not in ADMIN_IDS:
        await message.reply_text("Ця команда доступна лише організаторам.")
        return

    try:
        participants = list(iter_registered_participants())
    except Exception:
        logger.exception("Could not load registered participants")
        await message.reply_text("Не вдалося завантажити список із Google Sheets.")
        return

    sent = 0
    failed = 0
    for participant in participants:
        try:
            await context.bot.send_message(
                chat_id=int(participant["chat_id"]),
                text=(
                    "Привіт! 👋 Нагадуємо про дебатний турнір.\n\n"
                    "Будь ласка, підтверди, чи плануєш брати участь:"
                ),
                reply_markup=confirmation_keyboard(),
            )
            sent += 1
        except Exception:
            logger.exception(
                "Failed to message @%s (%s)",
                participant["username"],
                participant["chat_id"],
            )
            failed += 1

    await message.reply_text(
        f"Розсилку завершено.\n"
        f"Надіслано: {sent}\n"
        f"Не вдалося надіслати: {failed}"
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_message:
        await update.effective_message.reply_text(
            "Цей бот збирає підтвердження участі в дебатному турнірі.\n"
            "Натисни /start, щоб знайти свою реєстрацію та відповісти."
        )


def main():
    if not TELEGRAM_BOT_TOKEN:
        raise SystemExit(
            "Не знайдено TELEGRAM_BOT_TOKEN. Задай токен у змінній середовища "
            "TELEGRAM_BOT_TOKEN і запусти програму ще раз."
        )

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("broadcast", broadcast))
    app.add_handler(CallbackQueryHandler(handle_confirmation, pattern=r"^confirm:(yes|no)$"))

    logger.info("Bot is running")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
