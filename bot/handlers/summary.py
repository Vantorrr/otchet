from __future__ import annotations

from aiogram import Router, types, F
from aiogram.filters import Command
from aiogram.enums import ChatType

from bot.services.di import Container
from bot.utils.time_utils import parse_date_or_today
from bot.services.summary_builder import build_summary_text

summary_router = Router()


def _int_or_zero(value: object) -> int:
    try:
        return int(value)
    except Exception:
        return 0


@summary_router.message(Command("summary"), F.chat.type == ChatType.SUPERGROUP)
async def cmd_summary(message: types.Message) -> None:
    container = Container.get()
    args = message.text.split(maxsplit=1)
    day = parse_date_or_today(args[1] if len(args) > 1 else None, container.settings)

    summary_text = build_summary_text(container.settings, container.sheets, day)

    summary_topic_id = container.sheets.get_summary_topic_id()
    if summary_topic_id and message.chat.type == ChatType.SUPERGROUP:
        await message.bot.send_message(
            chat_id=message.chat.id,
            text=summary_text,
            message_thread_id=summary_topic_id,
        )
        if message.message_thread_id != summary_topic_id:
            await message.reply("Сводка опубликована в теме сводок.")
    else:
        await message.reply(summary_text)
