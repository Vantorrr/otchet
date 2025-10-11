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
from bot.utils.time_utils import now_in_tz


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
    logging.getLogger(__name__).info(
        "Scheduler timezone set to %s", settings.default_timezone
    )

    def _parse_hhmm(value: str, fallback: tuple[int, int]) -> tuple[int, int]:
        try:
            text = (value or "").strip()
            hh_str, mm_str = [p.strip() for p in text.split(":", 1)]
            return int(hh_str), int(mm_str)
        except Exception:
            logging.getLogger(__name__).warning(
                "Invalid time format '%s'. Falling back to %02d:%02d", value, fallback[0], fallback[1]
            )
            return fallback

    def _is_quiet_time() -> bool:
        """Pro Core: check if current time is in quiet hours."""
        try:
            now_h = now_in_tz(settings).hour
            now_m = now_in_tz(settings).minute
            q_start = _parse_hhmm(settings.reminder_quiet_start, (22, 0))
            q_end = _parse_hhmm(settings.reminder_quiet_end, (8, 0))
            cur = now_h * 60 + now_m
            start = q_start[0] * 60 + q_start[1]
            end = q_end[0] * 60 + q_end[1]
            if start < end:
                return start <= cur < end
            else:
                return cur >= start or cur < end
        except Exception:
            return False

    def _is_weekend() -> bool:
        """Check if today is Saturday or Sunday."""
        return now_in_tz(settings).weekday() in (5, 6)

    async def send_morning_reminders():
        try:
            if _is_weekend():
                logging.getLogger(__name__).info("Morning reminder skipped: weekend")
                return
            if _is_quiet_time():
                logging.getLogger(__name__).info("Morning reminder skipped: quiet hours")
                return
            container = Container.get()
            # Iterate through all office chats, skip HQ
            from bot.offices_config import is_hq
            from bot.keyboards.main import get_main_menu_keyboard
            all_chat_ids = container.sheets.get_all_group_chat_ids()
            if not all_chat_ids:
                logging.getLogger(__name__).warning("Morning reminder: no group_chat_ids configured")
                return
            total = 0
            sent = 0
            records = container.sheets._bindings.get_all_records()
            for chat_id in all_chat_ids:
                if is_hq(chat_id):
                    continue  # do not send reminders in HQ
                for binding in records:
                    if str(binding.get("chat_id")) != str(chat_id):
                        continue
                    topic_id_raw = str(binding.get("topic_id", "")).strip()
                    if not topic_id_raw.isdigit():
                        continue
                    topic_id = int(topic_id_raw)
                    manager = binding.get("manager")
                    if not (topic_id and manager):
                        continue
                    total += 1
                    try:
                        await bot.send_message(
                            chat_id,
                            f"🌅 Утреннее напоминание для <b>{manager}</b>\nВремя заполнить утренний отчет!",
                            message_thread_id=topic_id,
                            reply_markup=get_main_menu_keyboard(),
                        )
                        sent += 1
                    except Exception as err:
                        logging.getLogger(__name__).warning(
                            f"Failed to send morning reminder to {manager} (chat {chat_id}, topic {topic_id}): {err}"
                        )
            logging.getLogger(__name__).info("Morning reminder sent: %d/%d", sent, total)
        except Exception as e:
            logging.getLogger(__name__).warning(f"Morning reminder error: {e}")

    async def send_evening_reminders():
        try:
            if _is_weekend():
                logging.getLogger(__name__).info("Evening reminder skipped: weekend")
                return
            if _is_quiet_time():
                logging.getLogger(__name__).info("Evening reminder skipped: quiet hours")
                return
            container = Container.get()
            from bot.offices_config import is_hq
            from bot.keyboards.main import get_main_menu_keyboard
            all_chat_ids = container.sheets.get_all_group_chat_ids()
            if not all_chat_ids:
                logging.getLogger(__name__).warning("Evening reminder: no group_chat_ids configured")
                return
            total = 0
            sent = 0
            records = container.sheets._bindings.get_all_records()
            for chat_id in all_chat_ids:
                if is_hq(chat_id):
                    continue  # do not send reminders in HQ
                for binding in records:
                    if str(binding.get("chat_id")) != str(chat_id):
                        continue
                    topic_id_raw = str(binding.get("topic_id", "")).strip()
                    if not topic_id_raw.isdigit():
                        continue
                    topic_id = int(topic_id_raw)
                    manager = binding.get("manager")
                    if not (topic_id and manager):
                        continue
                    total += 1
                    try:
                        await bot.send_message(
                            chat_id,
                            f"🌆 Вечернее напоминание для <b>{manager}</b>\nВремя заполнить вечерний отчет!",
                            message_thread_id=topic_id,
                            reply_markup=get_main_menu_keyboard(),
                        )
                        sent += 1
                    except Exception as err:
                        logging.getLogger(__name__).warning(
                            f"Failed to send evening reminder to {manager} (chat {chat_id}, topic {topic_id}): {err}"
                        )
            logging.getLogger(__name__).info("Evening reminder sent: %d/%d", sent, total)
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

    mh, mm = _parse_hhmm(settings.morning_reminder, (9, 30))
    eh, em = _parse_hhmm(settings.evening_reminder, (17, 30))
    logging.getLogger(__name__).info(
        "Scheduling reminders: morning %02d:%02d, evening %02d:%02d (TZ %s)",
        mh, mm, eh, em, settings.default_timezone,
    )

    scheduler.add_job(
        send_morning_reminders,
        "cron",
        hour=mh,
        minute=mm,
        id="morning_reminders",
        misfire_grace_time=600,
        coalesce=True,
        replace_existing=True,
    )
    scheduler.add_job(
        send_evening_reminders,
        "cron",
        hour=eh,
        minute=em,
        id="evening_reminders",
        misfire_grace_time=600,
        coalesce=True,
        replace_existing=True,
    )
    # Автосводку отключаем по требованию клиента
    scheduler.start()
    try:
        m_job = scheduler.get_job("morning_reminders")
        e_job = scheduler.get_job("evening_reminders")
        logging.getLogger(__name__).info(
            "Next runs: morning=%s, evening=%s",
            getattr(m_job, "next_run_time", None),
            getattr(e_job, "next_run_time", None),
        )
    except Exception:
        pass

    logging.getLogger(__name__).info("Bot is starting polling...")
    await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        pass
