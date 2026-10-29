import asyncio
import re
import gspread
from telethon import TelegramClient
from telethon.errors import FloodWaitError, UsernameInvalidError, UsernameNotOccupiedError
from telethon.tl.types import User
from google_auth_oauthlib.flow import InstalledAppFlow

import os

API_ID = os.environ.get("API_ID")
API_HASH = os.environ.get("API_HASH")
SHEET_ID = os.environ.get("SHEET_ID")
DEF_COLS = [5, 11]
DEF_HDR = 1

# Окремі налаштування для конкретних аркушів (назви як на вкладках).
# Якщо аркуша тут немає, береться DEF_COLS / DEF_HDR.
# Якщо хочеш перевіряти ТІЛЬКИ перелічені аркуші, постав ONLY_CFG = True.
CFG = {
    "Команди": {"cols": [5, 11], "hdr": 1},
    "Одиночки": {"cols": [3],    "hdr": 1},
}
ONLY_CFG = False

PAUSE = 2   # секунд між запитами до Telegram
# ==============================

# кольори RGB 0..1
CLR = {
    "NOT_FOUND":  {"red": 0.96, "green": 0.70, "blue": 0.70},  # червоний
    "BAD_FORMAT": {"red": 0.99, "green": 0.80, "blue": 0.55},  # помаранчевий
    "OTHER":      {"red": 1.00, "green": 0.95, "blue": 0.60},  # жовтий (бот, канал, видалений)
}

USR_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{3,30}[A-Za-z0-9]$")


def is_valid_fmt(usr: str) -> bool:
    return bool(USR_RE.match(usr)) and "__" not in usr


def my_flow(client_config, scopes):
    flow = InstalledAppFlow.from_client_config(client_config, scopes)
    return flow.run_local_server(port=8080, timeout_seconds=300)


async def check_usr(client, usr: str) -> str:
    usr = usr.strip().lstrip("@")
    if not is_valid_fmt(usr):
        return "BAD_FORMAT"
    try:
        ent = await client.get_entity(usr)
    except (UsernameNotOccupiedError, UsernameInvalidError, ValueError):
        return "NOT_FOUND"
    except FloodWaitError as e:
        print(f"FloodWait: чекаю {e.seconds} с")
        await asyncio.sleep(e.seconds + 1)
        return await check_usr(client, usr)

    if not isinstance(ent, User) or ent.bot or ent.deleted:
        return "OTHER"
    return "OK"


async def main():
    gc = gspread.oauth(
        credentials_filename="credentials.json",
        authorized_user_filename="token.json",
        flow=my_flow,
    )
    sh = gc.open_by_key(SHEET_ID)

    sheets = sh.worksheets()
    if ONLY_CFG:
        sheets = [ws for ws in sheets if ws.title in CFG]

    cache = {}   # спільний кеш на всі аркуші

    async with TelegramClient("sess", API_ID, API_HASH) as client:
        for ws in sheets:
            cfg = CFG.get(ws.title, {"cols": DEF_COLS, "hdr": DEF_HDR})
            cols, hdr = cfg["cols"], cfg["hdr"]
            print(f"--- Аркуш: {ws.title} ---")

            fmts = []
            for col in cols:
                vals = ws.col_values(col)[hdr:]
                for i, raw in enumerate(vals):
                    key = raw.strip().lstrip("@").lower()
                    if not key:
                        continue
                    if key not in cache:
                        cache[key] = await check_usr(client, key)
                        await asyncio.sleep(PAUSE)
                        print(ws.title, key, cache[key])

                    if cache[key] == "OK":
                        continue   # існуючі ніки не чіпаємо

                    row = hdr + 1 + i
                    cell = gspread.utils.rowcol_to_a1(row, col)
                    fmts.append({
                        "range": cell,
                        "format": {"backgroundColor": CLR[cache[key]]},
                    })

            if fmts:
                ws.batch_format(fmts)   # фарбуємо весь аркуш одним запитом
                print(f"Зафарбовано клітинок: {len(fmts)}")


asyncio.run(main())