from __future__ import annotations

from aiogram import Router, types, F
from aiogram.filters import Command
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.enums import ChatType

from bot.services.di import Container
from bot.services.sheets import MorningData
from bot.utils.time_utils import date_str_for_today
from bot.keyboards.main import get_main_menu_keyboard


class MorningStates(StatesGroup):
    waiting_calls_planned = State()
    waiting_new_calls_planned = State()


morning_router = Router()


@morning_router.message(Command("morning"), F.chat.type == ChatType.SUPERGROUP)
async def cmd_morning(message: types.Message, state: FSMContext) -> None:
    if not message.message_thread_id:
        await message.reply("Эта команда должна выполняться в теме менеджера.")
        return
    container = Container.get()
    manager = container.sheets.get_manager_by_topic(message.message_thread_id)
    if not manager:
        await message.reply("Тема не привязана к менеджеру. Используйте /bind_manager <ФИО>.")
        return
    await state.update_data(manager=manager)
    await state.set_state(MorningStates.waiting_calls_planned)
    await message.reply("Утренний отчет. Введите количество перезвонов (план на сегодня):")


@morning_router.message(MorningStates.waiting_calls_planned, F.text.regexp(r"^\d+$"))
async def morning_calls_planned(message: types.Message, state: FSMContext) -> None:
    await state.update_data(calls_planned=int(message.text))
    await state.set_state(MorningStates.waiting_new_calls_planned)
    await message.reply("Количество новых звонков (план):")

@morning_router.message(MorningStates.waiting_new_calls_planned, F.text.regexp(r"^\d+$"))
async def morning_new_calls_planned(message: types.Message, state: FSMContext) -> None:
    data = await state.get_data()
    calls_planned = int(data["calls_planned"])  # type: ignore[index]
    new_calls_planned = int(message.text)

    container = Container.get()
    manager = data["manager"]  # type: ignore[index]
    date_str = date_str_for_today(container.settings)
    
    # Determine office by chat_id
    from bot.offices_config import get_office_by_chat_id
    office = get_office_by_chat_id(message.chat.id)

    container.sheets.upsert_report(
        date_str,
        manager,
        morning=MorningData(
            calls_planned=calls_planned,
            leads_units_planned=0,
            leads_volume_planned=0,
            new_calls_planned=new_calls_planned,
        ),
        office=office,
    )

    await state.clear()
    await message.reply("Утренний отчет сохранен. Спасибо!", reply_markup=get_main_menu_keyboard())


@morning_router.message(MorningStates.waiting_calls_planned)
@morning_router.message(MorningStates.waiting_new_calls_planned)
async def morning_invalid(message: types.Message) -> None:
    # Проверяем, не медиафайл ли это
    if message.photo or message.video or message.document or message.voice or message.audio:
        await message.reply("❌ Отправлен медиафайл вместо числа.\n\nПожалуйста, введите число.")
    else:
        await message.reply("Введите число.")
