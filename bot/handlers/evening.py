from __future__ import annotations

from aiogram import Router, types, F
from aiogram.filters import Command
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.enums import ChatType

from bot.services.di import Container
from bot.services.sheets import EveningData
from bot.utils.time_utils import date_str_for_today
from bot.keyboards.main import get_main_menu_keyboard


class EveningStates(StatesGroup):
    waiting_calls_success = State()
    waiting_leads_units = State()
    waiting_leads_volume = State()
    waiting_new_calls = State()


evening_router = Router()


@evening_router.message(Command("evening"), F.chat.type == ChatType.SUPERGROUP)
async def cmd_evening(message: types.Message, state: FSMContext) -> None:
    if not message.message_thread_id:
        await message.reply("Эта команда должна выполняться в теме менеджера.")
        return
    container = Container.get()
    manager = container.sheets.get_manager_by_topic(message.message_thread_id)
    if not manager:
        await message.reply("Тема не привязана к менеджеру. Используйте /bind_manager <ФИО>.")
        return
    await state.update_data(manager=manager)
    await state.set_state(EveningStates.waiting_calls_success)
    await message.reply("Вечерний отчет. Перезвоны успешно (число):")


@evening_router.message(EveningStates.waiting_calls_success, F.text.regexp(r"^\d+$"))
async def evening_calls_success(message: types.Message, state: FSMContext) -> None:
    await state.update_data(calls_success=int(message.text))
    await state.set_state(EveningStates.waiting_leads_units)
    await message.reply("Завел заявок за сегодня, штуки:")


@evening_router.message(EveningStates.waiting_leads_units, F.text.regexp(r"^\d+$"))
async def evening_leads_units(message: types.Message, state: FSMContext) -> None:
    await state.update_data(leads_units=int(message.text))
    await state.set_state(EveningStates.waiting_leads_volume)
    await message.reply("Завел заявок за сегодня, объем:")


@evening_router.message(EveningStates.waiting_leads_volume, F.text.regexp(r"^\d+$"))
async def evening_leads_volume(message: types.Message, state: FSMContext) -> None:
    await state.update_data(leads_volume=int(message.text))
    await state.set_state(EveningStates.waiting_new_calls)
    await message.reply("Количество новых дозвонов:")


@evening_router.message(EveningStates.waiting_new_calls, F.text.regexp(r"^\d+$"))
async def evening_new_calls(message: types.Message, state: FSMContext) -> None:
    data = await state.get_data()
    calls_success = int(data["calls_success"])  # type: ignore[index]
    leads_units = int(data["leads_units"])  # type: ignore[index]
    leads_volume = int(data["leads_volume"])  # type: ignore[index]
    new_calls = int(message.text)

    container = Container.get()
    manager = data["manager"]  # type: ignore[index]
    date_str = date_str_for_today(container.settings)

    container.sheets.upsert_report(
        date_str,
        manager,
        evening=EveningData(
            calls_success=calls_success,
            leads_units=leads_units,
            leads_volume=leads_volume,
            new_calls=new_calls,
        ),
    )

    await state.clear()
    await message.reply("Вечерний отчет сохранен. Спасибо!", reply_markup=get_main_menu_keyboard())
    # Auto-post summary into summary topic if configured
    summary_topic_id = container.sheets.get_summary_topic_id()
    if summary_topic_id and message.chat:
        from bot.services.summary_builder import build_summary_text
        text = build_summary_text(container.settings, container.sheets, date_str_for_today(container.settings))
        await message.bot.send_message(
            chat_id=message.chat.id,
            text=text,
            message_thread_id=summary_topic_id,
        )


@evening_router.message(EveningStates.waiting_calls_success)
@evening_router.message(EveningStates.waiting_leads_units)
@evening_router.message(EveningStates.waiting_leads_volume)
@evening_router.message(EveningStates.waiting_new_calls)
async def evening_invalid(message: types.Message) -> None:
    await message.reply("Введите число.")
