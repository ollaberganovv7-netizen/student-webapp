import asyncio
import logging
import sys

from aiogram import Bot, Dispatcher
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from aiogram.fsm.storage.memory import MemoryStorage

from config import BOT_TOKEN
from database.db import init_db
from middlewares.register import RegisterMiddleware

from handlers import start, presentation, documents, payment, admin, coursework, tezis, maqola, uslubiy
from utils.scanner import update_templates_json

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    stream=sys.stdout,
)
logger = logging.getLogger(__name__)


async def main():
    if not BOT_TOKEN:
        logger.error("BOT_TOKEN topilmadi! .env faylini tekshiring.")
        sys.exit(1)

    try:
        # Initialize database
        logger.info("Ma'lumotlar bazasi ishga tushirilmoqda...")
        await init_db()
        logger.info("Ma'lumotlar bazasi tayyor ✅")

        # Update templates catalog
        logger.info("Taqdimot dizaynlari katalogi yangilanmoqda...")
        update_templates_json()
        logger.info("Katalog yangilandi ✅")
    except Exception as e:
        logger.error(f"Xatolik yuz berdi: {e}", exc_info=True)
        sys.exit(1)

    # Create bot & dispatcher
    bot = Bot(
        token=BOT_TOKEN,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML),
    )
    dp = Dispatcher(storage=MemoryStorage())

    # Register middleware on all updates
    dp.message.middleware(RegisterMiddleware())
    dp.callback_query.middleware(RegisterMiddleware())

    # Include routers
    dp.include_router(admin.router)       # Admin first (priority)
    dp.include_router(start.router)
    dp.include_router(presentation.router)
    dp.include_router(documents.router)
    dp.include_router(coursework.router)
    dp.include_router(tezis.router)       # Tezis handler
    dp.include_router(maqola.router)      # Maqola handler
    dp.include_router(uslubiy.router)     # Uslubiy ishlanma handler
    dp.include_router(payment.router)

    # Set bot commands menu
    from aiogram.types import BotCommand
    commands = [
        BotCommand(command="start", description="Botni ishga tushirish"),
        BotCommand(command="stop", description="Jarayonni to'xtatish"),
        BotCommand(command="buy", description="Balansni to'ldirish"),
        BotCommand(command="help", description="Yordam va qo'llab-quvvatlash"),
    ]
    await bot.set_my_commands(commands)

    # Start polling
    try:
        logger.info("Bot ishga tushdi 🚀")
        await bot.delete_webhook(drop_pending_updates=True)
        await dp.start_polling(bot)
    except KeyboardInterrupt:
        logger.info("Bot to'xtatildi (Ctrl+C)")
    finally:
        await bot.session.close()
        logger.info("Bot sessiyasi yopildi")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Dastur to'xtatildi")
        sys.exit(0)
