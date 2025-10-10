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
    waiting_approved_volume = State()
    waiting_issued_volume = State()
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
    await message.reply("Завел заявок за сегодня, объем (млн):")


@evening_router.message(EveningStates.waiting_leads_volume, F.text.regexp(r"^\d+$"))
async def evening_leads_volume(message: types.Message, state: FSMContext) -> None:
    await state.update_data(leads_volume=int(message.text))
    await state.set_state(EveningStates.waiting_approved_volume)
    await message.reply("Одобрено за сегодня, объем (млн):")


@evening_router.message(EveningStates.waiting_approved_volume, F.text.regexp(r"^\d+$"))
async def evening_approved(message: types.Message, state: FSMContext) -> None:
    await state.update_data(approved_volume=int(message.text))
    await state.set_state(EveningStates.waiting_issued_volume)
    await message.reply("Выдано за сегодня, объем (млн):")


@evening_router.message(EveningStates.waiting_issued_volume, F.text.regexp(r"^\d+$"))
async def evening_issued(message: types.Message, state: FSMContext) -> None:
    await state.update_data(issued_volume=int(message.text))
    await state.set_state(EveningStates.waiting_new_calls)
    await message.reply("Количество новых дозвонов:")


@evening_router.message(EveningStates.waiting_new_calls, F.text.regexp(r"^\d+$"))
async def evening_new_calls(message: types.Message, state: FSMContext) -> None:
    data = await state.get_data()
    calls_success = int(data["calls_success"])  # type: ignore[index]
    leads_units = int(data["leads_units"])  # type: ignore[index]
    leads_volume = int(data["leads_volume"])  # type: ignore[index]
    approved_volume = int(data.get("approved_volume", 0))  # type: ignore[index]
    issued_volume = int(data.get("issued_volume", 0))  # type: ignore[index]
    new_calls = int(message.text)

    container = Container.get()
    manager = data["manager"]  # type: ignore[index]
    date_str = date_str_for_today(container.settings)
    
    # Determine office by chat_id
    from bot.offices_config import get_office_by_chat_id
    office = get_office_by_chat_id(message.chat.id)

    container.sheets.upsert_report(
        date_str,
        manager,
        evening=EveningData(
            calls_success=calls_success,
            leads_units=leads_units,
            leads_volume=leads_volume,
            approved_volume=approved_volume,
            issued_volume=issued_volume,
            new_calls=new_calls,
        ),
        office=office,
    )

    await state.clear()
    await message.reply("Вечерний отчет сохранен. Спасибо!", reply_markup=get_main_menu_keyboard())


@evening_router.message(EveningStates.waiting_calls_success)
@evening_router.message(EveningStates.waiting_leads_units)
@evening_router.message(EveningStates.waiting_leads_volume)
@evening_router.message(EveningStates.waiting_approved_volume)
@evening_router.message(EveningStates.waiting_issued_volume)
@evening_router.message(EveningStates.waiting_new_calls)
async def evening_invalid(message: types.Message) -> None:
    await message.reply("Введите число.")
