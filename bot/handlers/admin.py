from __future__ import annotations

from aiogram import Router, types, F
from aiogram.filters import Command
from aiogram.filters.command import CommandObject
from aiogram.enums import ChatType

from bot.config import Settings
from bot.services.di import Container
from bot.keyboards.main import get_main_menu_keyboard, get_admin_menu_keyboard

admin_router = Router()


@admin_router.message(Command("start"))
async def cmd_start(message: types.Message) -> None:
    # Initialize DI container when bot receives /start in any chat
    settings = Settings.load()
    Container.init(settings)
    # Remember group chat id for scheduled jobs
    if message.chat.id and message.chat.type == ChatType.SUPERGROUP:
        try:
            Container.get().sheets.set_group_chat_id(message.chat.id)
        except Exception:
            pass
    await message.reply("Бот запущен. Используйте /bind_manager и /set_summary_topic в нужных темах.")


@admin_router.message(Command("bind_manager"), F.chat.type == ChatType.SUPERGROUP)
async def cmd_bind_manager(message: types.Message) -> None:
    if not message.message_thread_id:
        await message.reply("Команда должна выполняться внутри темы.")
        return
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        await message.reply("Укажите ФИО менеджера: /bind_manager <ФИО>")
        return
    manager = args[1].strip()
    container = Container.get()
    container.sheets.set_manager_binding(message.message_thread_id, manager)
    await message.reply(f"Тема привязана к менеджеру: {manager}")


@admin_router.message(Command("set_summary_topic"), F.chat.type == ChatType.SUPERGROUP)
async def cmd_set_summary_topic(message: types.Message) -> None:
    if not message.message_thread_id:
        await message.reply("Команда должна выполняться внутри темы.")
        return
    container = Container.get()
    container.sheets.set_summary_topic(message.message_thread_id)
    await message.reply("Эта тема установлена для сводки.")


@admin_router.message(Command("menu"), F.chat.type == ChatType.SUPERGROUP)
async def cmd_menu(message: types.Message) -> None:
    if not message.message_thread_id:
        await message.reply("Команда должна выполняться внутри темы.")
        return
    
    container = Container.get()
    manager = container.sheets.get_manager_by_topic(message.message_thread_id)
    summary_topic_id = container.sheets.get_summary_topic_id()
    
    if manager:
        # Тема менеджера
        await message.reply(
            f"Меню для менеджера <b>{manager}</b>:",
            reply_markup=get_main_menu_keyboard()
        )
    elif message.message_thread_id == summary_topic_id:
        # Тема сводки
        await message.reply(
            "Меню администратора:",
            reply_markup=get_admin_menu_keyboard()
        )
    else:
        # Неопределенная тема
        await message.reply(
            "Эта тема не настроена. Используйте:\n"
            "• /bind_manager ФИО - для привязки к менеджеру\n"
            "• /set_summary_topic - для темы сводки"
        )


@admin_router.message(Command("purge_manager"))
async def cmd_purge_manager(message: types.Message, command: CommandObject) -> None:
    # Prefer parsing via CommandObject to be robust with mentions: /purge_manager@bot args
    argline = (command.args or "").strip()
    if not argline:
        await message.reply("Укажите ФИО менеджера: /purge_manager <ФИО> [YYYY-MM-DD]")
        return
    tail = argline.split()
    manager = tail[0]
    date = tail[1] if len(tail) > 1 else None
    container = Container.get()
    deleted_reports = container.sheets.delete_reports_by_manager(manager, date)
    deleted_bindings = container.sheets.delete_bindings_by_manager(manager)
    await message.reply(
        f"Удалено записей: Reports={deleted_reports}, Bindings={deleted_bindings} для менеджера {manager}"
    )

# Fallback in case Command filter misses due to client quirks (e.g. slash with extra spaces)
@admin_router.message(F.text.regexp(r"^/purge_manager(?:@\w+)?(\s+.*)?$"))
async def cmd_purge_manager_fallback(message: types.Message) -> None:
    parts = message.text.split(maxsplit=1) if message.text else []
    if len(parts) < 2:
        await message.reply("Укажите ФИО менеджера: /purge_manager <ФИО> [YYYY-MM-DD]")
        return
    tail = parts[1].split()
    manager = tail[0]
    date = tail[1] if len(tail) > 1 else None
    container = Container.get()
    deleted_reports = container.sheets.delete_reports_by_manager(manager, date)
    deleted_bindings = container.sheets.delete_bindings_by_manager(manager)
    await message.reply(
        f"Удалено записей: Reports={deleted_reports}, Bindings={deleted_bindings} для менеджера {manager}"
    )
