from __future__ import annotations

from aiogram import Router, types, F
from aiogram.enums import ChatType
from aiogram.fsm.context import FSMContext

from bot.services.di import Container
from bot.handlers.morning import MorningStates
from bot.handlers.evening import EveningStates
from bot.utils.time_utils import (
    date_str_for_today,
    parse_date_or_today,
    start_end_of_week_today,
    start_end_of_month_today,
    start_end_of_quarter_today,
)
from bot.services.summary_builder import build_summary_text

callbacks_router = Router()


@callbacks_router.callback_query(F.data == "morning_report")
async def callback_morning_report(callback: types.CallbackQuery, state: FSMContext) -> None:
    if not callback.message or not callback.message.message_thread_id:
        await callback.answer("Ошибка: не в теме")
        return
    
    container = Container.get()
    manager = container.sheets.get_manager_by_topic(callback.message.message_thread_id)
    if not manager:
        await callback.answer("Тема не привязана к менеджеру")
        return
    
    await state.update_data(manager=manager)
    await state.set_state(MorningStates.waiting_calls_planned)
    await callback.message.answer("Утренний отчет. Введите количество перезвонов (план на сегодня):")
    await callback.answer()


@callbacks_router.callback_query(F.data == "evening_report")
async def callback_evening_report(callback: types.CallbackQuery, state: FSMContext) -> None:
    if not callback.message or not callback.message.message_thread_id:
        await callback.answer("Ошибка: не в теме")
        return
    
    container = Container.get()
    manager = container.sheets.get_manager_by_topic(callback.message.message_thread_id)
    if not manager:
        await callback.answer("Тема не привязана к менеджеру")
        return
    
    await state.update_data(manager=manager)
    await state.set_state(EveningStates.waiting_calls_success)
    await callback.message.answer("Вечерний отчет. Перезвоны успешно (число):")
    await callback.answer()


@callbacks_router.callback_query(F.data == "summary_today")
async def callback_summary_today(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    container = Container.get()
    day = date_str_for_today(container.settings)
    summary_text = build_summary_text(container.settings, container.sheets, day)
    
    summary_topic_id = container.sheets.get_summary_topic_id()
    if summary_topic_id and callback.message.chat.type == ChatType.SUPERGROUP:
        await callback.bot.send_message(
            chat_id=callback.message.chat.id,
            text=summary_text,
            message_thread_id=summary_topic_id,
        )
        if callback.message.message_thread_id != summary_topic_id:
            await callback.message.answer("Сводка опубликована в теме сводок.")
    else:
        await callback.message.answer(summary_text)
    
    await callback.answer()


@callbacks_router.callback_query(F.data == "summary_week")
async def callback_summary_week(callback: types.CallbackQuery) -> None:
    container = Container.get()
    start, end = start_end_of_week_today(container.settings)
    text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end)
    await callback.message.answer(text)
    await callback.answer()


@callbacks_router.callback_query(F.data == "summary_month")
async def callback_summary_month(callback: types.CallbackQuery) -> None:
    container = Container.get()
    start, end = start_end_of_month_today(container.settings)
    text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end)
    await callback.message.answer(text)
    await callback.answer()


@callbacks_router.callback_query(F.data == "summary_quarter")
async def callback_summary_quarter(callback: types.CallbackQuery) -> None:
    container = Container.get()
    start, end = start_end_of_quarter_today(container.settings)
    text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end)
    await callback.message.answer(text)
    await callback.answer()


@callbacks_router.callback_query(F.data == "summary_date")
async def callback_summary_date(callback: types.CallbackQuery) -> None:
    await callback.message.answer(
        "Для сводки за конкретную дату используйте команду:\n"
        "<code>/summary YYYY-MM-DD</code>\n\n"
        "Например: <code>/summary 2024-01-15</code>"
    )
    await callback.answer()


@callbacks_router.callback_query(F.data == "setup_topic")
async def callback_setup_topic(callback: types.CallbackQuery) -> None:
    await callback.message.answer(
        "Настройка тем:\n\n"
        "<b>Для привязки темы к менеджеру:</b>\n"
        "<code>/bind_manager ФИО</code>\n\n"
        "<b>Для установки темы сводки:</b>\n"
        "<code>/set_summary_topic</code>"
    )
    await callback.answer()


@callbacks_router.callback_query(F.data == "summary_period")
async def callback_summary_period(callback: types.CallbackQuery) -> None:
    await callback.message.answer(
        "Сводка за произвольный период:\n"
        "<code>/summary_range YYYY-MM-DD YYYY-MM-DD</code>\n\n"
        "Например: <code>/summary_range 2025-08-01 2025-08-14</code>"
    )
    await callback.answer()
