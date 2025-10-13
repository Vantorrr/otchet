from __future__ import annotations

import io
from aiogram import Router, types, F
from aiogram.enums import ChatType
from aiogram.fsm.context import FSMContext

from bot.services.di import Container
from bot.handlers.morning import MorningStates
from bot.handlers.evening import EveningStates
from bot.services.data_aggregator import DataAggregatorService
from bot.services.simple_presentation import SimplePresentationService
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
from bot.offices_config import get_office_by_chat_id, is_hq
from aiogram.fsm.state import StatesGroup, State
from aiogram.filters.command import CommandObject


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
    manager = container.sheets.get_manager_by_topic(callback.message.chat.id, callback.message.message_thread_id)
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
    manager = container.sheets.get_manager_by_topic(callback.message.chat.id, callback.message.message_thread_id)
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
    
    summary_topic_id = container.sheets.get_summary_topic_id(callback.message.chat.id)
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
    
    # Answer early to avoid Telegram timeout on long operations
    try:
        await callback.answer("Генерация началась")
    except Exception:
        pass

    try:
        container = Container.get()
        start, end = start_end_of_week_today(container.settings)
        
        # Determine office filter for non-HQ chats
        office_filter = None
        if callback.message and callback.message.chat and not is_hq(callback.message.chat.id):
            office_filter = get_office_by_chat_id(callback.message.chat.id)
            if office_filter == "Unknown":
                office_filter = None
        
        text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end, office_filter=office_filter)
        
        # Разбиваем длинное сообщение на части
        parts = split_long_message(text)
        
        for i, part in enumerate(parts):
            if i == 0:
                await callback.message.answer(part)
            else:
                await callback.message.answer(f"📄 Часть {i + 1}:\n\n{part}")
        
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при генерации недельной сводки: {str(e)}")


@callbacks_router.callback_query(F.data == "summary_month")
async def callback_summary_month(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    # Answer early to avoid Telegram timeout on long operations
    try:
        await callback.answer("Генерация началась")
    except Exception:
        pass

    try:
        container = Container.get()
        start, end = start_end_of_month_today(container.settings)
        
        # Determine office filter for non-HQ chats
        office_filter = None
        if callback.message and callback.message.chat and not is_hq(callback.message.chat.id):
            office_filter = get_office_by_chat_id(callback.message.chat.id)
            if office_filter == "Unknown":
                office_filter = None
        
        text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end, office_filter=office_filter)
        
        # Разбиваем длинное сообщение на части
        parts = split_long_message(text)
        
        for i, part in enumerate(parts):
            if i == 0:
                await callback.message.answer(part)
            else:
                await callback.message.answer(f"📄 Часть {i + 1}:\n\n{part}")
        
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при генерации месячной сводки: {str(e)}")


@callbacks_router.callback_query(F.data == "summary_quarter")
async def callback_summary_quarter(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    # Answer early to avoid Telegram timeout on long operations
    try:
        await callback.answer("Генерация началась")
    except Exception:
        pass

    try:
        container = Container.get()
        start, end = start_end_of_quarter_today(container.settings)
        
        # Determine office filter for non-HQ chats
        office_filter = None
        if callback.message and callback.message.chat and not is_hq(callback.message.chat.id):
            office_filter = get_office_by_chat_id(callback.message.chat.id)
            if office_filter == "Unknown":
                office_filter = None
        
        text = build_summary_text(container.settings, container.sheets, day=start, start=start, end=end, office_filter=office_filter)
        
        # Разбиваем длинное сообщение на части
        parts = split_long_message(text)
        
        for i, part in enumerate(parts):
            if i == 0:
                await callback.message.answer(part)
            else:
                await callback.message.answer(f"📄 Часть {i + 1}:\n\n{part}")
        
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка при генерации квартальной сводки: {str(e)}")


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


@callbacks_router.callback_query(F.data.in_({"admin_remind_morning", "admin_remind_evening"}))
async def callback_admin_reminders(callback: types.CallbackQuery) -> None:
    # Работает только в офисных чатах (не HQ)
    if not callback.message:
        await callback.answer("Ошибка")
        return
    from bot.offices_config import is_hq
    if is_hq(callback.message.chat.id):
        await callback.answer("Доступно только в офисных чатах", show_alert=False)
        return
    container = Container.get()
    chat_id = callback.message.chat.id
    mode = "morning" if callback.data == "admin_remind_morning" else "evening"
    from bot.keyboards.main import get_main_menu_keyboard
    sent = 0
    for binding in container.sheets._bindings.get_all_records():
        bind_chat = str(binding.get("chat_id", "")).strip()
        # Поддержка старых записей без chat_id: считаем их текущим чатом
        if bind_chat and bind_chat != str(chat_id):
            continue
        topic_id_raw = str(binding.get("topic_id", "")).strip()
        if not topic_id_raw.isdigit():
            continue
        topic_id = int(topic_id_raw)
        manager = binding.get("manager")
        if not (topic_id and manager):
            continue
        text = (
            f"🌅 Утреннее напоминание для <b>{manager}</b>\nВремя заполнить утренний отчет!"
            if mode == "morning"
            else f"🌆 Вечернее напоминание для <b>{manager}</b>\nВремя заполнить вечерний отчет!"
        )
        try:
            await callback.bot.send_message(
                chat_id,
                text,
                message_thread_id=topic_id,
                reply_markup=get_main_menu_keyboard(),
            )
            sent += 1
        except Exception:
            continue
    await callback.message.answer(f"✅ Отправлено напоминаний: {sent} ({mode}).")
    await callback.answer()


@callbacks_router.callback_query(F.data.in_({"admin_remind_all_morning", "admin_remind_all_evening"}))
async def callback_admin_reminders_all(callback: types.CallbackQuery) -> None:
    # Доступно только в HQ: рассылает напоминания всем офисам
    if not callback.message:
        await callback.answer("Ошибка")
        return
    from bot.offices_config import is_hq
    if not is_hq(callback.message.chat.id):
        await callback.answer("Доступно только в HQ", show_alert=False)
        return
    container = Container.get()
    mode = "morning" if callback.data == "admin_remind_all_morning" else "evening"
    from bot.keyboards.main import get_main_menu_keyboard
    sent = 0
    total = 0
    all_chat_ids = container.sheets.get_all_group_chat_ids()
    for chat_id in all_chat_ids:
        if is_hq(chat_id):
            continue
        for binding in container.sheets._bindings.get_all_records():
            if str(binding.get("chat_id")) != str(chat_id):
                continue
            topic_id_raw = str(binding.get("topic_id", "")).strip()
            if not topic_id_raw.isdigit():
                continue
            topic_id = int(topic_id_raw)
            manager = binding.get("manager")
            if not (topic_id and manager):
                continue
            text = (
                f"🌅 Утреннее напоминание для <b>{manager}</b>\nВремя заполнить утренний отчет!"
                if mode == "morning"
                else f"🌆 Вечернее напоминание для <b>{manager}</b>\nВремя заполнить вечерний отчет!"
            )
            try:
                await callback.bot.send_message(
                    chat_id,
                    text,
                    message_thread_id=topic_id,
                    reply_markup=get_main_menu_keyboard(),
                )
                sent += 1
            except Exception:
                continue
            total += 1
    await callback.message.answer(f"✅ Отправлено напоминаний: {sent}/{total} ({mode}).")
    await callback.answer()


@callbacks_router.callback_query(F.data == "presentation_week")
async def callback_presentation_week(callback: types.CallbackQuery) -> None:
    """Generate weekly AI presentation."""
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    # Answer early to avoid Telegram timeout on long operations
    try:
        await callback.answer("Генерация началась")
    except Exception:
        pass

    await callback.message.answer("🔄 Генерирую недельную AI-презентацию...")
    
    try:
        container = Container.get()
        
        # Initialize services
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = SimplePresentationService(container.settings)
        
        # Determine office filter for non-HQ chats
        office_filter = None
        if callback.message and callback.message.chat and not is_hq(callback.message.chat.id):
            office_filter = get_office_by_chat_id(callback.message.chat.id)
            if office_filter == "Unknown":
                office_filter = None
        
        # Get weekly data (with previous for comparison)
        period_data, previous_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_weekly_data_with_previous(office_filter=office_filter)
        
        if not period_data:
            await callback.message.answer("❌ Нет данных за эту неделю.")
            await callback.answer()
            return
        
        # Generate presentation
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, previous_data, period_name, start_date, end_date, prev_start, prev_end, office_filter=office_filter
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
    


@callbacks_router.callback_query(F.data == "presentation_month")
async def callback_presentation_month(callback: types.CallbackQuery) -> None:
    """Generate monthly AI presentation."""
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    # Answer early to avoid Telegram timeout on long operations
    try:
        await callback.answer("Генерация началась")
    except Exception:
        pass

    await callback.message.answer("🔄 Генерирую месячную AI-презентацию...")
    
    try:
        container = Container.get()
        
        # Initialize services
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = SimplePresentationService(container.settings)
        
        # Determine office filter for non-HQ chats
        office_filter = None
        if callback.message and callback.message.chat and not is_hq(callback.message.chat.id):
            office_filter = get_office_by_chat_id(callback.message.chat.id)
            if office_filter == "Unknown":
                office_filter = None
        
        # Get monthly data (with previous for comparison)
        period_data, previous_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_monthly_data_with_previous(office_filter=office_filter)
        
        if not period_data:
            await callback.message.answer("❌ Нет данных за этот месяц.")
            await callback.answer()
            return
        
        # Generate presentation
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, previous_data, period_name, start_date, end_date, prev_start, prev_end, office_filter=office_filter
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
    


@callbacks_router.callback_query(F.data == "presentation_quarter")
async def callback_presentation_quarter(callback: types.CallbackQuery) -> None:
    """Generate quarterly AI presentation."""
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    # Answer early to avoid Telegram timeout on long operations
    try:
        await callback.answer("Генерация началась")
    except Exception:
        pass

    await callback.message.answer("🔄 Генерирую квартальную AI-презентацию...")
    
    try:
        container = Container.get()
        
        # Initialize services
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = SimplePresentationService(container.settings)
        
        # Determine office filter for non-HQ chats
        office_filter = None
        if callback.message and callback.message.chat and not is_hq(callback.message.chat.id):
            office_filter = get_office_by_chat_id(callback.message.chat.id)
            if office_filter == "Unknown":
                office_filter = None
        
        # Get quarterly data (with previous for comparison)
        period_data, previous_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_quarterly_data_with_previous(office_filter=office_filter)
        
        if not period_data:
            await callback.message.answer("❌ Нет данных за этот квартал.")
            await callback.answer()
            return
        
        # Generate presentation
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, previous_data, period_name, start_date, end_date, prev_start, prev_end, office_filter=office_filter
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
    


@callbacks_router.callback_query(F.data == "presentation_period")
async def callback_presentation_period(callback: types.CallbackQuery) -> None:
    if not callback.message:
        await callback.answer("Ошибка")
        return
    await callback.message.answer(
        "Укажи период в формате:\n"
        "<code>/presentation_range YYYY-MM-DD YYYY-MM-DD</code>\n\n"
        "Или сравни два периода:\n"
        "<code>/presentation_compare YYYY-MM-DD YYYY-MM-DD YYYY-MM-DD YYYY-MM-DD</code>\n\n"
        "Например: <code>/presentation_compare 2025-08-01 2025-08-07 2025-09-01 2025-09-07</code>"
    )
    await callback.answer()


@callbacks_router.callback_query(F.data == "tempo_check")
async def callback_tempo_check(callback: types.CallbackQuery) -> None:
    """Check managers falling behind tempo."""
    if not callback.message:
        await callback.answer("Ошибка")
        return
    
    # Answer early to avoid Telegram timeout on long operations
    try:
        await callback.answer("Анализ начался")
    except Exception:
        pass

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
    
