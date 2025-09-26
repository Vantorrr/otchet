from __future__ import annotations

import io
from aiogram import Router, types, F
from aiogram.filters import Command
from aiogram.filters.command import CommandObject
from aiogram.enums import ChatType

from bot.config import Settings
from bot.services.di import Container
from bot.services.data_aggregator import DataAggregatorService
from bot.services.presentation import PresentationService
from bot.services.google_slides import GoogleSlidesService
from bot.services.tempo_analytics import TempoAnalyticsService
from bot.keyboards.main import get_main_menu_keyboard, get_admin_menu_keyboard
from aiogram.filters.command import CommandObject
from bot.utils.time_utils import parse_date_or_today

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


@admin_router.message(Command("generate_weekly_presentation"))
async def cmd_generate_weekly_presentation(message: types.Message) -> None:
    """Generate weekly presentation with AI analysis."""
    await message.reply("🔄 Генерирую недельную презентацию...")
    
    try:
        container = Container.get()
        
        # Initialize services
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = PresentationService(container.settings)
        
        # Get weekly data with previous
        period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_weekly_data_with_previous()
        
        if not period_data:
            await message.reply("❌ Нет данных за эту неделю.")
            return
        
        # Check if previous period has data
        if not prev_data:
            await message.reply(
                f"⚠️ Нет данных за предыдущую неделю ({prev_start.strftime('%d.%m.%Y')}—{prev_end.strftime('%d.%m.%Y')}).\n"
                f"\nГенерирую презентацию только за текущий период..."
            )
            # Generate without comparison
            prev_data = {}
            
        # Generate presentation with comparison
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, period_name, start_date, end_date, prev_data, prev_start, prev_end
        )
        
        # Send as document
        document = types.BufferedInputFile(
            pptx_bytes,
            filename=f"Отчет_{period_name.replace(' ', '_')}.pptx"
        )
        
        await message.reply_document(
            document,
            caption=f"📊 {period_name}\n🤖 Презентация с AI-анализом готова!"
        )
        
    except Exception as e:
        await message.reply(f"❌ Ошибка при генерации презентации: {str(e)}")


@admin_router.message(Command("generate_monthly_presentation"))
async def cmd_generate_monthly_presentation(message: types.Message) -> None:
    """Generate monthly presentation with AI analysis."""
    await message.reply("🔄 Генерирую месячную презентацию...")
    
    try:
        container = Container.get()
        
        # Initialize services
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = PresentationService(container.settings)
        
        # Get monthly data with previous
        period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_monthly_data_with_previous()
        
        if not period_data:
            await message.reply("❌ Нет данных за этот месяц.")
            return
        
        # Check if previous period has data
        if not prev_data:
            await message.reply(
                f"⚠️ Нет данных за предыдущий месяц ({prev_start.strftime('%d.%m.%Y')}—{prev_end.strftime('%d.%m.%Y')}).\n"
                f"\nГенерирую презентацию только за текущий период..."
            )
            # Generate without comparison
            prev_data = {}
            
        # Generate presentation with comparison
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, period_name, start_date, end_date, prev_data, prev_start, prev_end
        )
        
        # Send as document
        document = types.BufferedInputFile(
            pptx_bytes,
            filename=f"Отчет_{period_name.replace(' ', '_')}.pptx"
        )
        
        await message.reply_document(
            document,
            caption=f"📊 {period_name}\n🤖 Презентация с AI-анализом готова!"
        )
        
    except Exception as e:
        await message.reply(f"❌ Ошибка при генерации презентации: {str(e)}")


@admin_router.message(Command("generate_quarterly_presentation"))
async def cmd_generate_quarterly_presentation(message: types.Message) -> None:
    """Generate quarterly presentation with AI analysis."""
    await message.reply("🔄 Генерирую квартальную презентацию...")
    
    try:
        container = Container.get()
        
        # Initialize services
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = PresentationService(container.settings)
        
        # Get quarterly data with previous
        period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_quarterly_data_with_previous()
        
        if not period_data:
            await message.reply("❌ Нет данных за этот квартал.")
            return
        
        # Check if previous period has data
        if not prev_data:
            await message.reply(
                f"⚠️ Нет данных за предыдущий квартал ({prev_start.strftime('%d.%m.%Y')}—{prev_end.strftime('%d.%m.%Y')}).\n"
                f"\nГенерирую презентацию только за текущий период..."
            )
            # Generate without comparison
            prev_data = {}
            
        # Generate presentation with comparison
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, period_name, start_date, end_date, prev_data, prev_start, prev_end
        )
        
        # Send as document
        document = types.BufferedInputFile(
            pptx_bytes,
            filename=f"Отчет_{period_name.replace(' ', '_')}.pptx"
        )
        
        await message.reply_document(
            document,
            caption=f"📊 {period_name}\n🤖 Презентация с AI-анализом готова!"
        )
        
    except Exception as e:
        await message.reply(f"❌ Ошибка при генерации презентации: {str(e)}")


@admin_router.message(Command("presentation_range"))
async def cmd_presentation_range(message: types.Message, command: CommandObject) -> None:
    """Generate AI presentation for custom date range: /presentation_range YYYY-MM-DD YYYY-MM-DD"""
    args = (command.args or "").split()
    if len(args) != 2:
        await message.reply("Укажите период: /presentation_range YYYY-MM-DD YYYY-MM-DD")
        return
    try:
        from datetime import datetime as _dt
        start = _dt.strptime(args[0], "%Y-%m-%d").date()
        end = _dt.strptime(args[1], "%Y-%m-%d").date()
    except Exception:
        await message.reply("Неверный формат дат. Пример: /presentation_range 2025-08-01 2025-08-07")
        return
    await message.reply("🔄 Генерирую презентацию за указанный период...")
    try:
        container = Container.get()
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = PresentationService(container.settings)
        period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_custom_with_previous(start, end)
        if not period_data:
            await message.reply("❌ Нет данных за этот период.")
            return
        
        # Check if previous period has data
        if not prev_data:
            # Proceed without comparison
            prev_data = {}
            
        pptx_bytes = await presentation_service.generate_presentation(period_data, period_name, start_date, end_date, prev_data, prev_start, prev_end)
        document = types.BufferedInputFile(pptx_bytes, filename=f"AI_Отчет_{period_name.replace(' ', '_')}.pptx")
        await message.reply_document(document, caption=f"📊 {period_name}\n🤖 AI-презентация готова!")
    except Exception as e:
        await message.reply(f"❌ Ошибка: {str(e)}")


@admin_router.message(Command("slides_range"))
async def cmd_slides_range(message: types.Message, command: CommandObject) -> None:
    """Generate Google Slides deck for custom date range and export PDF to Drive folder.
    Usage: /slides_range YYYY-MM-DD YYYY-MM-DD
    """
    args = (command.args or "").split()
    if len(args) != 2:
        await message.reply("Укажите период: /slides_range YYYY-MM-DD YYYY-MM-DD")
        return
    try:
        from datetime import datetime as _dt
        start = _dt.strptime(args[0], "%Y-%m-%d").date()
        end = _dt.strptime(args[1], "%Y-%m-%d").date()
    except Exception:
        await message.reply("Неверный формат дат. Пример: /slides_range 2025-08-01 2025-08-07")
        return

    await message.reply("🔄 Генерирую Google Slides и PDF…")
    try:
        container = Container.get()
        aggregator = DataAggregatorService(container.sheets)
        period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_custom_with_previous(start, end)
        if not period_data:
            await message.reply("❌ Нет данных за этот период.")
            return

        slides = GoogleSlidesService(container.settings)
        deck_id = slides.create_presentation(f"Отчет {period_name}")
        slides.move_presentation_to_folder(deck_id)
        pdf_bytes = slides.export_pdf(deck_id)
        document = types.BufferedInputFile(pdf_bytes, filename=f"Отчет_{period_name.replace(' ', '_')}.pdf")
        await message.reply_document(document, caption=f"📄 PDF экспортировано. Презентация создана в папке Drive.")
    except Exception as e:
        await message.reply(f"❌ Ошибка Slides: {str(e)}")


@admin_router.message(Command("presentation_compare"))
async def cmd_presentation_compare(message: types.Message, command: CommandObject) -> None:
    """Generate AI presentation comparing two custom periods: /presentation_compare YYYY-MM-DD YYYY-MM-DD YYYY-MM-DD YYYY-MM-DD"""
    args = (command.args or "").split()
    if len(args) != 4:
        await message.reply("Формат: /presentation_compare YYYY-MM-DD YYYY-MM-DD YYYY-MM-DD YYYY-MM-DD\n(первый период A: start end, второй период B: start end)")
        return
    try:
        from datetime import datetime as _dt
        a_start = _dt.strptime(args[0], "%Y-%m-%d").date()
        a_end = _dt.strptime(args[1], "%Y-%m-%d").date()
        b_start = _dt.strptime(args[2], "%Y-%m-%d").date()
        b_end = _dt.strptime(args[3], "%Y-%m-%d").date()
    except Exception:
        await message.reply("Неверные даты. Пример: /presentation_compare 2025-08-01 2025-08-07 2025-09-01 2025-09-07")
        return
    await message.reply("🔄 Генерирую презентацию (сравнение двух периодов)...")
    try:
        container = Container.get()
        aggregator = DataAggregatorService(container.sheets)
        presentation_service = PresentationService(container.settings)
        data_a, data_b, title, start_a, end_a, start_b, end_b = await aggregator.aggregate_two_periods(a_start, a_end, b_start, b_end)
        if not data_a and not data_b:
            await message.reply("❌ Нет данных за выбранные периоды.")
            return
        pptx_bytes = await presentation_service.generate_presentation(data_a, title, start_a, end_a, data_b, start_b, end_b)
        document = types.BufferedInputFile(pptx_bytes, filename=f"AI_Отчет_{title.replace(' ', '_')}.pptx")
        await message.reply_document(document, caption=f"📊 {title}\n🤖 AI-презентация готова!")
    except Exception as e:
        await message.reply(f"❌ Ошибка: {str(e)}")


@admin_router.message(Command("tempo_check"))
async def cmd_tempo_check(message: types.Message) -> None:
    """Check managers falling behind tempo."""
    await message.reply("🔍 Анализирую темп выполнения планов...")
    
    try:
        container = Container.get()
        
        # Initialize tempo analytics
        tempo_service = TempoAnalyticsService(container.sheets)
        
        # Get tempo alerts
        alerts = await tempo_service.analyze_monthly_tempo()
        
        if not alerts:
            await message.reply("✅ Все менеджеры работают в рамках плана!")
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
        
        await message.reply(response)
        
    except Exception as e:
        await message.reply(f"❌ Ошибка при анализе темпа: {str(e)}")


@admin_router.message(Command("remind_now"), F.chat.type == ChatType.SUPERGROUP)
async def cmd_remind_now(message: types.Message) -> None:
    """Manually trigger morning/evening reminders now (for testing). Usage: /remind_now morning|evening"""
    args = (message.text or "").split()
    mode = args[1].lower() if len(args) > 1 else "morning"
    if mode not in {"morning", "evening"}:
        await message.reply("Использование: /remind_now morning|evening")
        return

    container = Container.get()
    chat_id = container.sheets.get_group_chat_id()
    if not chat_id:
        await message.reply("❌ Не задан group_chat_id. Отправьте /start в группе, чтобы сохранить его.")
        return

    from bot.keyboards.main import get_main_menu_keyboard

    sent = 0
    for binding in container.sheets._bindings.get_all_records():
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
            await message.bot.send_message(
                chat_id,
                text,
                message_thread_id=topic_id,
                reply_markup=get_main_menu_keyboard(),
            )
            sent += 1
        except Exception as e:
            # Continue sending to others even if one fails
            continue

    await message.reply(f"✅ Отправлено напоминаний: {sent} ({mode}).")
