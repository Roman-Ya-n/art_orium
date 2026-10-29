"""
Debate tournament participation bot.

Uses the existing Google Sheets OAuth files:
    credentials.json
    token.json

Important:
- Participants must open the bot and press Start before it can message them.
- Usernames are matched against configured columns.
- Invalid usernames are detected by background cell colors set by acc_check.py.
- Tracking columns are added at the right side of each worksheet.
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
SHEET_ID = os.getenv(
    "SHEET_ID",
    "1ksl4Y2-RgQ1TSkNoCvCnpouaUiFRQ9M4XFjNkznH_gs",
).strip()

SHEETS_CONFIG = {
    "Команди": {"tag_cols": [5, 11], "header_row": 1},
    "Одиночки": {"tag_cols": [3], "header_row": 1},
}

CREDENTIALS_FILE = "credentials.json"
AUTHORIZED_USER_FILE = "token.json"

ADMIN_IDS = {
    item.strip()
    for item in os.getenv("ADMIN_IDS", "").split(",")
    if item.strip()
}

STATUS_PENDING = "Очікує відповіді"
STATUS_YES = "Підтверджую участь"
STATUS_NO = "Не зможу бути"

# Ті самі кольори, що й у acc_check.py.
INVALID_COLORS = {
    "NOT_FOUND": {"red": 0.96, "green": 0.70, "blue": 0.70},
    "BAD_FORMAT": {"red": 0.99, "green": 0.80, "blue": 0.55},
    "OTHER": {"red": 1.00, "green": 0.95, "blue": 0.60},
}

CONFIRMED_COLOR = {
    "red": 0.71,
    "green": 0.84,
    "blue": 0.66,
}

REFUSED_COLOR = {
    "red": 0.96,
    "green": 0.70,
    "blue": 0.70,
}

COLOR_TOLERANCE = 0.035

INVALID_REASONS = {
    "NOT_FOUND": "username не знайдено",
    "BAD_FORMAT": "некоректний формат username",
    "OTHER": "бот, канал або видалений акаунт",
}

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger("debate-confirmation-bot")

_gc = None
_spreadsheet = None


# ---------------- Google Sheets helpers ----------------

def get_spreadsheet():
    """Authorize using the existing gspread OAuth setup."""
    global _gc, _spreadsheet

    if _spreadsheet is None:
        _gc = gspread.oauth(
            credentials_filename=CREDENTIALS_FILE,
            authorized_user_filename=AUTHORIZED_USER_FILE,
        )
        _spreadsheet = _gc.open_by_key(SHEET_ID)

    return _spreadsheet


def tracking_headers(tag_col: int) -> tuple[str, str, str]:
    return (
        f"Bot chat_id (tag col {tag_col})",
        f"Підтвердження (tag col {tag_col})",
        f"Час відповіді (tag col {tag_col})",
    )


def ensure_tracking_columns(
    ws, header_row: int, tag_cols: list[int]
) -> dict[int, tuple[int, int, int]]:
    """Create tracking columns if they don't exist."""
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
                col = len(headers) + 1
                headers.append(label)
                changed = True

            cols.append(col)

        result[tag_col] = tuple(cols)

    if changed:
        for tag_col, cols in result.items():
            for col, label in zip(cols, tracking_headers(tag_col)):
                if col > ws.col_count:
                    ws.add_cols(col - ws.col_count)

                current = ws.cell(header_row, col).value
                if current != label:
                    ws.update_cell(header_row, col, label)

    return result


def get_configured_worksheets():
    spreadsheet = get_spreadsheet()
    configured = []

    for title, cfg in SHEETS_CONFIG.items():
        try:
            ws = spreadsheet.worksheet(title)
        except gspread.WorksheetNotFound:
            logger.warning("Worksheet not found, skipping: %s", title)
            continue

        tracking = ensure_tracking_columns(
            ws, cfg["header_row"], cfg["tag_cols"]
        )
        configured.append((ws, cfg, tracking))

    return configured


def normalize_username(value: str | None) -> str:
    if not value:
        return ""
    return value.strip().lstrip("@").lower()


def _color_matches(actual: dict | None, expected: dict) -> bool:
    """Compare an RGB background color with a small rounding tolerance."""
    if not actual:
        return False

    for channel in ("red", "green", "blue"):
        actual_value = actual.get(channel, 0.0)
        expected_value = expected[channel]

        if abs(actual_value - expected_value) > COLOR_TOLERANCE:
            return False

    return True


def _get_invalid_reason(cell_data: dict) -> str | None:
    """
    Determine whether a cell has one of the invalid colors from acc_check.py.
    Prefer the effective color, but also inspect the explicitly entered format.
    """
    formats = [
        cell_data.get("effectiveFormat", {}).get("backgroundColor"),
        cell_data.get("userEnteredFormat", {}).get("backgroundColor"),
    ]

    for reason, expected_color in INVALID_COLORS.items():
        if any(_color_matches(color, expected_color) for color in formats):
            return reason

    return None


def get_invalid_cells(ws, cfg: dict) -> dict[tuple[int, int], str]:
    """
    Return {(row_number, column_number): reason} for invalid-colored usernames.

    Reads formatting from Google Sheets metadata. If the metadata request fails,
    an exception is raised so the bot does not accidentally process those rows.
    """
    username_cols = cfg["tag_cols"]
    header_row = cfg["header_row"]

    # Only read as far as the username data extends.
    last_row = header_row
    for col in username_cols:
        last_row = max(last_row, len(ws.col_values(col)))

    if last_row <= header_row:
        return {}

    first_col = min(username_cols)
    last_col = max(username_cols)

    first_letter = gspread.utils.rowcol_to_a1(1, first_col).rstrip("1")
    last_letter = gspread.utils.rowcol_to_a1(1, last_col).rstrip("1")

    # Quote the worksheet title safely for an A1 range.
    safe_title = ws.title.replace("'", "''")
    cell_range = (
        f"'{safe_title}'!"
        f"{first_letter}{header_row + 1}:{last_letter}{last_row}"
    )

    metadata = ws.spreadsheet.fetch_sheet_metadata(
        params={
            "includeGridData": "true",
            "ranges": cell_range,
            "fields": (
                "sheets(properties(title),"
                "data(startRow,startColumn,"
                "rowData(values(userEnteredFormat(backgroundColor),"
                "effectiveFormat(backgroundColor)))))"
            ),
        }
    )

    invalid = {}
    sheets = metadata.get("sheets", [])

    # No grid data is unexpected for a requested range; fail safely.
    if not sheets or not sheets[0].get("data"):
        raise RuntimeError(
            f"Google Sheets returned no formatting data for {ws.title}"
        )

    for grid in sheets[0].get("data", []):
        start_row = grid.get("startRow", 0)  # Google API uses zero-based indexes.
        start_col = grid.get("startColumn", 0)

        for row_offset, row_data in enumerate(grid.get("rowData", [])):
            row_number = start_row + row_offset + 1

            for col_offset, cell_data in enumerate(row_data.get("values", [])):
                col_number = start_col + col_offset + 1

                if col_number not in username_cols:
                    continue

                reason = _get_invalid_reason(cell_data)
                if reason:
                    invalid[(row_number, col_number)] = reason

    return invalid


def get_invalid_cells_by_sheet(configured=None):
    """Load invalid username cells once per worksheet."""
    if configured is None:
        configured = get_configured_worksheets()

    invalid_by_sheet = {}

    for ws, cfg, _tracking in configured:
        invalid_by_sheet[ws.title] = get_invalid_cells(ws, cfg)

    return invalid_by_sheet


def find_user_rows(username: str):
    """
    Find exact username matches. Results include an 'invalid_reason' field
    when the matched username cell is colored as invalid by acc_check.py.
    """
    target = normalize_username(username)
    if not target:
        return []

    configured = get_configured_worksheets()
    invalid_by_sheet = get_invalid_cells_by_sheet(configured)
    matches = []

    for ws, cfg, tracking in configured:
        invalid_cells = invalid_by_sheet[ws.title]

        for tag_col in cfg["tag_cols"]:
            values = ws.col_values(tag_col)
            start_row = cfg["header_row"] + 1

            for row_number, raw in enumerate(
                values[start_row - 1:], start=start_row
            ):
                if normalize_username(raw) != target:
                    continue

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
                    "invalid_reason": invalid_cells.get(
                        (row_number, tag_col)
                    ),
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
        datetime.now(timezone.utc).astimezone().strftime(
            "%Y-%m-%d %H:%M:%S %Z"
        ),
    )


def mark_confirmation(match: dict, color: dict):
    """Color the cell immediately left of the username."""
    ws = match["worksheet"]
    row = match["row"]
    username_col = match["tag_col"]

    if username_col <= 1:
        raise ValueError("Username column has no cell to its left")

    cell = gspread.utils.rowcol_to_a1(row, username_col - 1)
    ws.format(cell, {"backgroundColor": color})


def iter_registered_participants(configured=None, invalid_by_sheet=None):
    """Yield valid registered participants who have not yet answered."""
    if configured is None:
        configured = get_configured_worksheets()

    if invalid_by_sheet is None:
        invalid_by_sheet = get_invalid_cells_by_sheet(configured)

    for ws, cfg, tracking in configured:
        invalid_cells = invalid_by_sheet[ws.title]

        for tag_col in cfg["tag_cols"]:
            chat_col, status_col, time_col = tracking[tag_col]

            usernames = ws.col_values(tag_col)
            chat_ids = ws.col_values(chat_col)
            statuses = ws.col_values(status_col)

            start_row = cfg["header_row"] + 1
            max_len = max(len(usernames), len(chat_ids), len(statuses))

            for row in range(start_row, max_len + 1):
                username = normalize_username(
                    usernames[row - 1] if row <= len(usernames) else ""
                )
                chat_id = (
                    chat_ids[row - 1].strip()
                    if row <= len(chat_ids)
                    else ""
                )
                status = (
                    statuses[row - 1].strip()
                    if row <= len(statuses)
                    else ""
                )

                if not username or not chat_id:
                    continue

                if status in (STATUS_YES, STATUS_NO):
                    continue

                invalid_reason = invalid_cells.get((row, tag_col))
                if invalid_reason:
                    continue

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
            InlineKeyboardButton(
                "✅ Підтверджую участь",
                callback_data="confirm:yes",
            ),
            InlineKeyboardButton(
                "❌ Не зможу бути",
                callback_data="confirm:no",
            ),
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
            "Будь ласка, встанови його в налаштуваннях Telegram "
            "і натисни /start ще раз."
        )
        return

    try:
        matches = find_user_rows(username)
    except Exception:
        logger.exception("Google Sheets lookup/format check failed")
        await message.reply_text(
            "Не вдалося перевірити реєстрацію та кольори запису "
            "через Google Sheets. Спробуй пізніше або напиши організаторам."
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
            "Знайшов кілька записів із цим username, тому не можу "
            "безпечно прив'язати відповідь. Будь ласка, звернися до організаторів."
        )
        return

    match = matches[0]

    if match["invalid_reason"]:
        reason = INVALID_REASONS.get(
            match["invalid_reason"], "username позначений як некоректний"
        )
        await message.reply_text(
            "Знайшов твій запис, але username позначений як некоректний "
            f"у таблиці ({reason}).\n"
            "Будь ласка, звернися до організаторів, щоб вони перевірили запис."
        )
        return

    try:
        save_chat_id(match, user.id)
    except Exception:
        logger.exception("Could not save chat_id")
        await message.reply_text(
            "Знайшов твою реєстрацію, але не зміг зберегти прив'язку "
            "в таблиці. Напиши організаторам."
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


async def handle_confirmation(
    update: Update, context: ContextTypes.DEFAULT_TYPE
):
    query = update.callback_query
    if not query:
        return

    await query.answer()

    user = query.from_user
    username = normalize_username(user.username)

    if not username:
        await query.edit_message_text(
            "Не бачу твого Telegram username. Напиши організаторам."
        )
        return

    try:
        matches = find_user_rows(username)
    except Exception:
        logger.exception("Google Sheets lookup/format check failed during confirmation")
        await query.edit_message_text(
            "Не вдалося перевірити запис і зберегти відповідь "
            "через технічну помилку. Спробуй пізніше або напиши організаторам."
        )
        return

    if len(matches) != 1:
        await query.edit_message_text(
            "Не вдалося однозначно знайти твою реєстрацію. Напиши організаторам."
        )
        return

    match = matches[0]

    if match["invalid_reason"]:
        reason = INVALID_REASONS.get(
            match["invalid_reason"], "username позначений як некоректний"
        )
        await query.edit_message_text(
            "Не можу зафіксувати відповідь: username позначений як "
            f"некоректний у таблиці ({reason}). Напиши організаторам."
        )
        return

    status = STATUS_YES if query.data == "confirm:yes" else STATUS_NO

    try:
        save_chat_id(match, user.id)
        save_response(match, status)

        try:
            if status == STATUS_YES:
                mark_confirmation(match, CONFIRMED_COLOR)
            elif status == STATUS_NO:
                mark_confirmation(match, REFUSED_COLOR)
        except Exception:
            logger.exception(
                "Response saved, but could not color confirmation cell"
            )

    except Exception:
        logger.exception("Failed to save confirmation response")
        # Тут залиш свій наявний код обробки помилки
        # збереження відповіді без змін.
        logger.exception("Could not save participant response")
        await query.edit_message_text(
            "Не вдалося зберегти відповідь. Спробуй ще раз або напиши організаторам."
        )
        return

    if status == STATUS_YES:
        await query.edit_message_text(
            "Дякую! ✅ Участь підтверджено. До зустрічі на турнірі!"
        )
    else:
        await query.edit_message_text(
            "Дякую, що повідомив(-ла). ❌ Зафіксував, що ти не зможеш бути."
        )


def split_telegram_text(text: str, limit: int = 3500) -> list[str]:
    """Split a long admin report into Telegram-safe messages."""
    chunks = []
    current = ""

    for line in text.splitlines():
        if len(current) + len(line) + 1 > limit:
            if current:
                chunks.append(current)
            current = line
        else:
            current = f"{current}\n{line}" if current else line

    if current:
        chunks.append(current)

    return chunks


async def broadcast(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin command: send the confirmation prompt to valid pending participants."""
    user = update.effective_user
    message = update.effective_message

    if not user or not message:
        return

    if str(user.id) not in ADMIN_IDS:
        await message.reply_text("Ця команда доступна лише організаторам.")
        return

    try:
        configured = get_configured_worksheets()
        invalid_by_sheet = get_invalid_cells_by_sheet(configured)

        skipped = []
        for ws, cfg, _tracking in configured:
            for (row, col), reason in invalid_by_sheet[ws.title].items():
                raw = ws.cell(row, col).value or ""
                skipped.append(
                    f"• {ws.title}, рядок {row}, @{normalize_username(raw) or '—'}"
                    f" — {INVALID_REASONS.get(reason, reason)}"
                )

        participants = list(
            iter_registered_participants(configured, invalid_by_sheet)
        )

    except Exception:
        logger.exception("Could not load participants or username formatting")
        await message.reply_text(
            "Розсилку скасовано: не вдалося надійно перевірити "
            "таблицю та кольори username. Перевір логи сервера."
        )
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

    summary = (
        "Розсилку завершено.\n"
        f"Надіслано: {sent}\n"
        f"Не вдалося надіслати: {failed}\n"
        f"Пропущено некоректних кольорових записів: {len(skipped)}"
    )
    await message.reply_text(summary)

    if skipped:
        report = "Записи, пропущені через кольори username:\n" + "\n".join(skipped)
        for chunk in split_telegram_text(report):
            await message.reply_text(chunk)


async def help_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
):
    if update.effective_message:
        await update.effective_message.reply_text(
            "Цей бот збирає підтвердження участі в дебатному турнірі.\n"
            "Натисни /start, щоб знайти свою реєстрацію та відповісти."
        )


def main():
    if not TELEGRAM_BOT_TOKEN:
        raise SystemExit(
            "Не знайдено TELEGRAM_BOT_TOKEN. Задай токен у змінній "
            "середовища TELEGRAM_BOT_TOKEN і запусти програму ще раз."
        )

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("broadcast", broadcast))
    app.add_handler(
        CallbackQueryHandler(
            handle_confirmation,
            pattern=r"^confirm:(yes|no)$",
        )
    )

    logger.info("Bot is running")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()