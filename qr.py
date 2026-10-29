import asyncio
import qrcode
from telethon import TelegramClient
from telethon.errors import SessionPasswordNeededError

import os

API_ID = os.environ.get("API_ID")
API_HASH = os.environ.get("API_HASH")

async def main():
    cl = TelegramClient("sess", API_ID, API_HASH)
    await cl.connect()
    if not await cl.is_user_authorized():
        qr = await cl.qr_login()
        img = qrcode.make(qr.url)
        img.save("qr.png")
        print("Відкрий qr.png і скануй з телефону (120 с)")
        try:
            await qr.wait(120)
        except SessionPasswordNeededError:
            await cl.sign_in(password=input("2FA password: "))
    me = await cl.get_me()
    print("OK:", me.first_name, me.id)
    await cl.disconnect()


asyncio.run(main())