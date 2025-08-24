from __future__ import annotations

from aiogram import Router, types, F
from aiogram.filters import Command
from aiogram.filters.command import CommandObject
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
    cmd = CommandObject(message=message)
    day = parse_date_or_today((cmd.args or None), container.settings)

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


@summary_router.message(Command("summary_range"))
async def cmd_summary_range(message: types.Message, command: CommandObject) -> None:
    container = Container.get()
    if not command.args:
        await message.reply("Использование: /summary_range YYYY-MM-DD YYYY-MM-DD")
        return
    parts = command.args.split()
    if len(parts) != 2:
        await message.reply("Использование: /summary_range YYYY-MM-DD YYYY-MM-DD")
        return
    start, end = parts[0], parts[1]
    from bot.services.summary_builder import build_summary_text
    text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end)
    await message.reply(text)
