import asyncio
import os
import logging
from aiogram import Bot, Dispatcher
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from dotenv import load_dotenv
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from datetime import time
from zoneinfo import ZoneInfo

from bot.config import Settings


def _configure_logging() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")


async def main() -> None:
    load_dotenv()
    _configure_logging()
    settings = Settings.load()

    # Make sure Google client picks credentials
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = settings.google_credentials_path

    # Initialize DI container early
    from bot.services.di import Container
    Container.init(settings)

    bot = Bot(token=settings.bot_token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher()

    from bot.handlers.admin import admin_router
    from bot.handlers.morning import morning_router
    from bot.handlers.evening import evening_router
    from bot.handlers.summary import summary_router
    from bot.handlers.callbacks import callbacks_router

    dp.include_router(admin_router)
    dp.include_router(morning_router)
    dp.include_router(evening_router)
    dp.include_router(summary_router)
    dp.include_router(callbacks_router)

    # Scheduler: reminders and daily summary
    scheduler = AsyncIOScheduler(timezone=ZoneInfo(settings.default_timezone))

    async def send_morning_reminders():
        try:
            container = Container.get()
            chat_id = container.sheets.get_group_chat_id()
            if not chat_id:
                return
            # Send menu to each manager topic
            from bot.keyboards.main import get_main_menu_keyboard
            for binding in container.sheets._bindings.get_all_records():
                topic_id_raw = str(binding.get("topic_id", "")).strip()
                if not topic_id_raw.isdigit():
                    continue
                topic_id = int(topic_id_raw)
                manager = binding.get("manager")
                if topic_id and manager:
                    await bot.send_message(
                        chat_id, 
                        f"🌅 Утреннее напоминание для <b>{manager}</b>\nВремя заполнить утренний отчет!", 
                        message_thread_id=topic_id,
                        reply_markup=get_main_menu_keyboard()
                    )
        except Exception as e:
            logging.getLogger(__name__).warning(f"Morning reminder error: {e}")

    async def send_evening_reminders():
        try:
            container = Container.get()
            chat_id = container.sheets.get_group_chat_id()
            if not chat_id:
                return
            # Send menu to each manager topic
            from bot.keyboards.main import get_main_menu_keyboard
            for binding in container.sheets._bindings.get_all_records():
                topic_id_raw = str(binding.get("topic_id", "")).strip()
                if not topic_id_raw.isdigit():
                    continue
                topic_id = int(topic_id_raw)
                manager = binding.get("manager")
                if topic_id and manager:
                    await bot.send_message(
                        chat_id, 
                        f"🌆 Вечернее напоминание для <b>{manager}</b>\nВремя заполнить вечерний отчет!", 
                        message_thread_id=topic_id,
                        reply_markup=get_main_menu_keyboard()
                    )
        except Exception as e:
            logging.getLogger(__name__).warning(f"Evening reminder error: {e}")

    async def post_daily_summary():
        try:
            container = Container.get()
            chat_id = container.sheets.get_group_chat_id()
            topic_id = container.sheets.get_summary_topic_id()
            if chat_id and topic_id:
                # Build and send summary text directly
                from bot.utils.time_utils import date_str_for_today
                from bot.services.summary_builder import build_summary_text
                day = date_str_for_today(container.settings)
                text = build_summary_text(container.settings, container.sheets, day)
                await bot.send_message(chat_id, text, message_thread_id=topic_id)
        except Exception as e:
            logging.getLogger(__name__).warning(f"Daily summary error: {e}")

    hh, mm = map(int, settings.morning_reminder.split(":"))
    scheduler.add_job(send_morning_reminders, "cron", hour=hh, minute=mm)
    hh, mm = map(int, settings.evening_reminder.split(":"))
    scheduler.add_job(send_evening_reminders, "cron", hour=hh, minute=mm)
    # Автосводку отключаем по требованию клиента
    scheduler.start()

    logging.getLogger(__name__).info("Bot is starting polling...")
    await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        pass
