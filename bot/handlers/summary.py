from __future__ import annotations

from aiogram import Router, types, F
from aiogram.filters import Command
from aiogram.filters.command import CommandObject
from aiogram.enums import ChatType

from bot.services.di import Container
from bot.utils.time_utils import parse_date_or_today
from bot.services.summary_builder import build_summary_text

summary_router = Router()


def split_long_message(text: str, max_length: int = 4000) -> list[str]:
    """Разбивает длинное сообщение на части, не превышающие max_length символов"""
    if len(text) <= max_length:
        return [text]
    
    parts = []
    current_part = ""
    
    lines = text.split('\n')
    for line in lines:
        # Если даже одна строка слишком длинная, принудительно обрезаем
        if len(line) > max_length:
            if current_part:
                parts.append(current_part.strip())
                current_part = ""
            # Разбиваем длинную строку на части
            while len(line) > max_length:
                parts.append(line[:max_length])
                line = line[max_length:]
            if line:
                current_part = line + '\n'
        # Проверяем поместится ли строка в текущую часть
        elif len(current_part) + len(line) + 1 <= max_length:
            current_part += line + '\n'
        else:
            # Строка не помещается, начинаем новую часть
            if current_part:
                parts.append(current_part.strip())
            current_part = line + '\n'
    
    # Добавляем последнюю часть
    if current_part:
        parts.append(current_part.strip())
    
    return parts


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
    parts = split_long_message(summary_text)

    summary_topic_id = container.sheets.get_summary_topic_id()
    if summary_topic_id and message.chat.type == ChatType.SUPERGROUP:
        for i, part in enumerate(parts):
            if i == 0:
                await message.bot.send_message(
                    chat_id=message.chat.id,
                    text=part,
                    message_thread_id=summary_topic_id,
                )
            else:
                await message.bot.send_message(
                    chat_id=message.chat.id,
                    text=f"📄 Часть {i + 1}:\n\n{part}",
                    message_thread_id=summary_topic_id,
                )
        if message.message_thread_id != summary_topic_id:
            await message.reply("Сводка опубликована в теме сводок.")
    else:
        for i, part in enumerate(parts):
            if i == 0:
                await message.reply(part)
            else:
                await message.answer(f"📄 Часть {i + 1}:\n\n{part}")


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
    text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end)
    
    # Разбиваем длинное сообщение на части
    message_parts = split_long_message(text)
    
    for i, part in enumerate(message_parts):
        if i == 0:
            await message.reply(part)
        else:
            await message.answer(f"📄 Часть {i + 1}:\n\n{part}")
