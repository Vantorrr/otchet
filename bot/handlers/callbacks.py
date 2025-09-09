from __future__ import annotations

import io
from aiogram import Router, types, F
from aiogram.enums import ChatType
from aiogram.fsm.context import FSMContext

from bot.services.di import Container
from bot.handlers.morning import MorningStates
from bot.handlers.evening import EveningStates
from bot.services.data_aggregator import DataAggregatorService
from bot.services.presentation import PresentationService
from bot.services.tempo_analytics import TempoAnalyticsService
from bot.utils.time_utils import (
    date_str_for_today,
    parse_date_or_today,
    start_end_of_week_today,
    start_end_of_month_today,
    start_end_of_quarter_today,
)
from bot.services.summary_builder import build_summary_text
from bot.keyboards.main import (
    get_admin_menu_keyboard,
    get_admin_summaries_keyboard,
    get_admin_ai_keyboard,
)
from aiogram.fsm.state import StatesGroup, State


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

callbacks_router = Router()


class AskAIStates(StatesGroup):
    waiting_question = State()


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


@callbacks_router.callback_query(F.data == "ask_ai")
async def callback_ask_ai(callback: types.CallbackQuery, state: FSMContext) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    await state.set_state(AskAIStates.waiting_question)
    await callback.message.answer("Напишите вопрос для ИИ одним сообщением. Чтобы отменить — отправьте /cancel.")
    await callback.answer()


@callbacks_router.message(AskAIStates.waiting_question)
async def handle_ai_question(message: types.Message, state: FSMContext) -> None:
    container = Container.get()
    from bot.services.yandex_gpt import YandexGPTService
    svc = YandexGPTService(container.settings)
    await message.answer("🤖 Думаю...")
    answer = await svc.generate_answer(message.text or "")
    for part in split_long_message(answer):
        await message.answer(part)
    # Не выходим из состояния, чтобы можно было продолжать диалог.
    # Выйти можно командой /cancel.


@callbacks_router.message(F.text == "/cancel")
async def handle_ai_cancel(message: types.Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer("Диалог с ИИ завершён. Чтобы начать снова — нажмите ‘Спроси у ИИ’.")

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
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    try:
        container = Container.get()
        start, end = start_end_of_week_today(container.settings)
        
        text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end)
        
        # Разбиваем длинное сообщение на части
        parts = split_long_message(text)
        
        for i, part in enumerate(parts):
            if i == 0:
                await callback.message.answer(part)
            else:
                await callback.message.answer(f"📄 Часть {i + 1}:\n\n{part}")
        
        await callback.answer("Готово!")
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при генерации недельной сводки: {str(e)}")
        await callback.answer("Ошибка!")


@callbacks_router.callback_query(F.data == "summary_month")
async def callback_summary_month(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    try:
        container = Container.get()
        start, end = start_end_of_month_today(container.settings)
        text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end)
        
        # Разбиваем длинное сообщение на части
        parts = split_long_message(text)
        
        for i, part in enumerate(parts):
            if i == 0:
                await callback.message.answer(part)
            else:
                await callback.message.answer(f"📄 Часть {i + 1}:\n\n{part}")
        
        await callback.answer("Готово!")
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при генерации месячной сводки: {str(e)}")
        await callback.answer("Ошибка!")


@callbacks_router.callback_query(F.data == "summary_quarter")
async def callback_summary_quarter(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    try:
        container = Container.get()
        start, end = start_end_of_quarter_today(container.settings)
        text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end)
        
        # Разбиваем длинное сообщение на части
        parts = split_long_message(text)
        
        for i, part in enumerate(parts):
            if i == 0:
                await callback.message.answer(part)
            else:
                await callback.message.answer(f"📄 Часть {i + 1}:\n\n{part}")
        
        await callback.answer("Готово!")
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при генерации квартальной сводки: {str(e)}")
        await callback.answer("Ошибка!")


@callbacks_router.callback_query(F.data == "summary_date")
async def callback_summary_date(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    await callback.message.answer(
        "Для сводки за конкретную дату используйте команду:\n"
        "<code>/summary YYYY-MM-DD</code>\n\n"
        "Например: <code>/summary 2024-01-15</code>"
    )
    await callback.answer()


@callbacks_router.callback_query(F.data == "setup_topic")
async def callback_setup_topic(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    await callback.message.answer(
        "Настройка тем:\n\n"
        "<b>Для привязки темы к менеджеру:</b>\n"
        "<code>/bind_manager ФИО</code>\n\n"
        "<b>Для установки темы сводки:</b>\n"
        "<code>/set_summary_topic</code>"
    )
    await callback.answer()


@callbacks_router.callback_query(F.data == "admin_section_summaries")
async def callback_admin_section_summaries(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    await callback.message.edit_text("Меню администратора → Сводки:", reply_markup=get_admin_summaries_keyboard())
    await callback.answer()


@callbacks_router.callback_query(F.data == "admin_section_ai")
async def callback_admin_section_ai(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    await callback.message.edit_text("Меню администратора → AI:", reply_markup=get_admin_ai_keyboard())
    await callback.answer()


@callbacks_router.callback_query(F.data == "admin_back")
async def callback_admin_back(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    await callback.message.edit_text("Меню администратора:", reply_markup=get_admin_menu_keyboard())
    await callback.answer()

@callbacks_router.callback_query(F.data == "summary_period")
async def callback_summary_period(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    await callback.message.answer(
        "Сводка за произвольный период:\n"
        "<code>/summary_range YYYY-MM-DD YYYY-MM-DD</code>\n\n"
        "Например: <code>/summary_range 2025-08-01 2025-08-14</code>"
    )
    await callback.answer()


@callbacks_router.callback_query(F.data == "presentation_week")
async def callback_presentation_week(callback: types.CallbackQuery) -> None:
    """Generate weekly AI presentation."""
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    await callback.message.answer("🔄 Генерирую недельную AI-презентацию...")
    
    try:
        container = Container.get()
        
        # Initialize services
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = PresentationService(container.settings)
        
        # Get weekly data
        period_data, period_name, start_date, end_date = await aggregator.aggregate_weekly_data()
        
        if not period_data:
            await callback.message.answer("❌ Нет данных за эту неделю.")
            await callback.answer()
            return
        
        # Generate presentation
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, period_name, start_date, end_date
        )
        
        # Send as document
        document = types.BufferedInputFile(
            pptx_bytes,
            filename=f"AI_Отчет_{period_name.replace(' ', '_')}.pptx"
        )
        
        await callback.message.answer_document(
            document,
            caption=f"📊 {period_name}\n🤖 AI-презентация готова!"
        )
        
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при генерации презентации: {str(e)}")
    
    await callback.answer()


@callbacks_router.callback_query(F.data == "presentation_month")
async def callback_presentation_month(callback: types.CallbackQuery) -> None:
    """Generate monthly AI presentation."""
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    await callback.message.answer("🔄 Генерирую месячную AI-презентацию...")
    
    try:
        container = Container.get()
        
        # Initialize services
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = PresentationService(container.settings)
        
        # Get monthly data
        period_data, period_name, start_date, end_date = await aggregator.aggregate_monthly_data()
        
        if not period_data:
            await callback.message.answer("❌ Нет данных за этот месяц.")
            await callback.answer()
            return
        
        # Generate presentation
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, period_name, start_date, end_date
        )
        
        # Send as document
        document = types.BufferedInputFile(
            pptx_bytes,
            filename=f"AI_Отчет_{period_name.replace(' ', '_')}.pptx"
        )
        
        await callback.message.answer_document(
            document,
            caption=f"📊 {period_name}\n🤖 AI-презентация готова!"
        )
        
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при генерации презентации: {str(e)}")
    
    await callback.answer()


@callbacks_router.callback_query(F.data == "presentation_quarter")
async def callback_presentation_quarter(callback: types.CallbackQuery) -> None:
    """Generate quarterly AI presentation."""
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    await callback.message.answer("🔄 Генерирую квартальную AI-презентацию...")
    
    try:
        container = Container.get()
        
        # Initialize services
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = PresentationService(container.settings)
        
        # Get quarterly data
        period_data, period_name, start_date, end_date = await aggregator.aggregate_quarterly_data()
        
        if not period_data:
            await callback.message.answer("❌ Нет данных за этот квартал.")
            await callback.answer()
            return
        
        # Generate presentation
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, period_name, start_date, end_date
        )
        
        # Send as document
        document = types.BufferedInputFile(
            pptx_bytes,
            filename=f"AI_Отчет_{period_name.replace(' ', '_')}.pptx"
        )
        
        await callback.message.answer_document(
            document,
            caption=f"📊 {period_name}\n🤖 AI-презентация готова!"
        )
        
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при генерации презентации: {str(e)}")
    
    await callback.answer()


@callbacks_router.callback_query(F.data == "tempo_check")
async def callback_tempo_check(callback: types.CallbackQuery) -> None:
    """Check managers falling behind tempo."""
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    await callback.message.answer("🔍 Анализирую темп выполнения планов...")
    
    try:
        container = Container.get()
        
        # Initialize tempo analytics
        tempo_service = TempoAnalyticsService(container.sheets)
        
        # Get tempo alerts
        alerts = await tempo_service.analyze_monthly_tempo()
        
        if not alerts:
            await callback.message.answer("✅ Все менеджеры работают в рамках плана!")
            await callback.answer()
            return
        
        # Format alerts
        response = "⚠️ Менеджеры, отстающие от плана:\n\n"
        
        critical_alerts = [a for a in alerts if a.alert_level == "critical"]
        warning_alerts = [a for a in alerts if a.alert_level == "warning"]
        
        if critical_alerts:
            response += "🔴 КРИТИЧНО:\n"
            for alert in critical_alerts:
                response += f"{alert.message}\n\n"
        
        if warning_alerts:
            response += "🟡 ПРЕДУПРЕЖДЕНИЕ:\n"
            for alert in warning_alerts:
                response += f"{alert.message}\n\n"
        
        await callback.message.answer(response)
        
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при анализе темпа: {str(e)}")
    
    await callback.answer()
