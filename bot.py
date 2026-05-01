import os
import logging
import asyncio
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import Message
from aiohttp import web

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    raise ValueError("BOT_TOKEN is not set!")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# Health check сервер для Render
async def health_check(request):
    return web.Response(text="OK")

async def run_health_check():
    app = web.Application()
    app.router.add_get("/health", health_check)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", 8080)
    await site.start()
    logging.info("Health check server started on http://0.0.0.0:8080/health")

@dp.message(Command("start"))
async def cmd_start(message: Message):
    await message.answer("Hello! I'm alive!")

@dp.message(F.text)
async def echo(message: Message):
    await message.answer(f"Echo: {message.text}")

async def main():
    # Запускаем health-check сервер в фоне
    asyncio.create_task(run_health_check())
    logging.info("Starting bot polling...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
