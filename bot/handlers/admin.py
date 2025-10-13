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
from bot.offices_config import get_office_by_chat_id
from aiogram.types import CallbackQuery
from aiogram.filters.command import CommandObject
from bot.utils.time_utils import parse_date_or_today
from typing import Dict, Tuple, Optional

admin_router = Router()

# In-memory flow state for "Отчёт по дате": (chat_id, thread_id) -> {'awaiting': 'start'|'end', 'start': 'YYYY-MM-DD'}
REPORT_FLOW_STATE: Dict[Tuple[int, Optional[int]], Dict[str, str]] = {}


@admin_router.message(Command("start"))
async def cmd_start(message: types.Message) -> None:
    # Initialize DI container when bot receives /start in any chat
    settings = Settings.load()
    Container.init(settings)
    # Remember group chat id for scheduled jobs
    chat_id = message.chat.id
    if chat_id and message.chat.type == ChatType.SUPERGROUP:
        try:
            Container.get().sheets.set_group_chat_id(chat_id)
        except Exception:
            pass
        await message.reply(f"Бот запущен. Используйте /bind_manager и /set_summary_topic в нужных темах.\n\n📍 Chat ID: {chat_id}")
    else:
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
    container.sheets.set_manager_binding(message.chat.id, message.message_thread_id, manager)
    await message.reply(f"Тема привязана к менеджеру: {manager}")


@admin_router.message(Command("set_summary_topic"), F.chat.type == ChatType.SUPERGROUP)
async def cmd_set_summary_topic(message: types.Message) -> None:
    if not message.message_thread_id:
        await message.reply("Команда должна выполняться внутри темы.")
        return
    container = Container.get()
    container.sheets.set_summary_topic(message.chat.id, message.message_thread_id)
    await message.reply("Эта тема установлена для сводки.")


@admin_router.message(Command("menu"), F.chat.type == ChatType.SUPERGROUP)
async def cmd_menu(message: types.Message) -> None:
    if not message.message_thread_id:
        await message.reply("Команда должна выполняться внутри темы.")
        return
    
    container = Container.get()
    manager = container.sheets.get_manager_by_topic(message.chat.id, message.message_thread_id)
    summary_topic_id = container.sheets.get_summary_topic_id(message.chat.id)
    
    if manager:
        # Тема менеджера
        await message.reply(
            f"Меню для менеджера <b>{manager}</b>:",
            reply_markup=get_main_menu_keyboard()
        )
    elif message.message_thread_id == summary_topic_id:
        # Тема сводки
        from bot.offices_config import is_hq
        is_hq_chat = is_hq(message.chat.id)
        await message.reply(
            "Меню администратора:" + (" (Головной офис)" if is_hq_chat else ""),
            reply_markup=get_admin_menu_keyboard(is_hq=is_hq_chat)
        )
    else:
        # Неопределенная тема
        await message.reply(
            "Эта тема не настроена. Используйте:\n"
            "• /bind_manager ФИО — привязать менеджера к ЭТОЙ теме\n"
            "• /set_summary_topic — сделать ЭТУ тему темой сводки"
        )


# ====== Inline flow: Отчёт по дате ======

@admin_router.callback_query(F.data == "admin_report_by_date")
async def cb_report_by_date(query: CallbackQuery) -> None:
    key = (query.message.chat.id, getattr(query.message, 'message_thread_id', None))
    REPORT_FLOW_STATE[key] = {'awaiting': 'start'}
    await query.message.answer("📅 Введите дату начала в формате YYYY-MM-DD")
    await query.answer()


@admin_router.message(F.text.regexp(r"^\d{4}-\d{2}-\d{2}$"))
async def msg_report_by_date_flow(message: types.Message) -> None:
    key = (message.chat.id, getattr(message, 'message_thread_id', None))
    state = REPORT_FLOW_STATE.get(key)
    if not state:
        return
    from datetime import datetime as _dt
    try:
        if state.get('awaiting') == 'start':
            _dt.strptime(message.text.strip(), "%Y-%m-%d")
            state['start'] = message.text.strip()
            state['awaiting'] = 'end'
            await message.reply("📅 Теперь введите дату окончания в формате YYYY-MM-DD")
            return
        if state.get('awaiting') == 'end':
            start = _dt.strptime(state.get('start', ''), "%Y-%m-%d").date()
            end = _dt.strptime(message.text.strip(), "%Y-%m-%d").date()
            REPORT_FLOW_STATE.pop(key, None)
            container = Container.get()
            aggregator = DataAggregatorService(container.sheets)
            from bot.offices_config import is_hq, get_all_offices
            
            # Check if this is HQ
            if is_hq(message.chat.id):
                # For HQ - generate office summary instead of full presentation
                all_offices = get_all_offices()
                response = f"📊 Отчет по офисам: {start.strftime('%d.%m.%Y')}—{end.strftime('%d.%m.%Y')}\n\n"
                
                total_calls = 0
                total_new = 0
                total_leads = 0
                total_volume = 0.0
                total_issued = 0.0
                
                for office in all_offices:
                    office_data = await aggregator._aggregate_data_for_period(start, end, office_filter=office)
                    if office_data:
                        office_calls = sum(m.calls_fact for m in office_data.values())
                        office_new = sum(m.new_calls for m in office_data.values())
                        office_leads = sum(m.leads_units_fact for m in office_data.values())
                        office_volume = sum(m.leads_volume_fact for m in office_data.values())
                        office_issued = sum(m.issued_volume for m in office_data.values())
                        
                        response += f"\n🏢 <b>{office}</b>\n"
                        response += f"👥 Менеджеров: {len(office_data)}\n"
                        response += f"📞 Перезвоны: {office_calls}\n"
                        response += f"🆕 Новые звонки: {office_new}\n"
                        response += f"📋 Заявки (шт): {office_leads}\n"
                        response += f"💰 Заявки (млн): {office_volume:.1f}\n"
                        response += f"✅ Выдано (млн): {office_issued:.1f}\n"
                        
                        total_calls += office_calls
                        total_new += office_new
                        total_leads += office_leads
                        total_volume += office_volume
                        total_issued += office_issued
                
                response += f"\n======================\n"
                response += f"<b>📊 ИТОГО ПО ВСЕМ ОФИСАМ</b>\n"
                response += f"📞 Перезвоны: {total_calls}\n"
                response += f"🆕 Новые звонки: {total_new}\n"
                response += f"📋 Заявки (шт): {total_leads}\n"
                response += f"💰 Заявки (млн): {total_volume:.1f}\n"
                response += f"✅ Выдано (млн): {total_issued:.1f}\n"
                
                await message.reply(response)
            else:
                # For regular offices - generate presentation as before
                from bot.services.simple_presentation import SimplePresentationService
                presentation_service = SimplePresentationService(container.settings)
                office_filter = get_office_by_chat_id(message.chat.id) if message.chat.id else None
                period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_custom_with_previous(start, end, office_filter=office_filter)
                if not period_data:
                    await message.reply("❌ Нет данных за этот период.")
                    return
                pptx_bytes = await presentation_service.generate_presentation(period_data, period_name, start_date, end_date, prev_data, prev_start, prev_end, office_filter=office_filter)
                document = types.BufferedInputFile(pptx_bytes, filename=f"Отчет_По_Активности_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pptx")
                await message.reply_document(document, caption=f"📊 {period_name}\n🤖 Простой отчёт готов!")
    except Exception as e:
        REPORT_FLOW_STATE.pop(key, None)
        await message.reply(f"❌ Ошибка: {str(e)}")


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
    """Generate simple presentation for custom date range: /presentation_range YYYY-MM-DD YYYY-MM-DD"""
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
        from bot.services.simple_presentation import SimplePresentationService
        presentation_service = SimplePresentationService(container.settings)
        period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_custom_with_previous(start, end)
        if not period_data:
            await message.reply("❌ Нет данных за этот период.")
            return
        
        # Check if previous period has data
        if not prev_data:
            # Proceed without comparison
            prev_data = {}
            
        pptx_bytes = await presentation_service.generate_presentation(period_data, period_name, start_date, end_date, prev_data, prev_start, prev_end)
        document = types.BufferedInputFile(pptx_bytes, filename=f"Отчет_По_Активности_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pptx")
        await message.reply_document(document, caption=f"📊 {period_name}\n🤖 Презентация готова!")
    except Exception as e:
        await message.reply(f"❌ Ошибка: {str(e)}")


@admin_router.message(Command("simple_range"))
async def cmd_simple_range(message: types.Message, command: CommandObject) -> None:
    """Generate simple PPTX (per-manager table + AI comment + logo): /simple_range YYYY-MM-DD YYYY-MM-DD"""
    args = (command.args or "").split()
    if len(args) != 2:
        await message.reply("Укажите период: /simple_range YYYY-MM-DD YYYY-MM-DD")
        return
    try:
        from datetime import datetime as _dt
        start = _dt.strptime(args[0], "%Y-%m-%d").date()
        end = _dt.strptime(args[1], "%Y-%m-%d").date()
    except Exception:
        await message.reply("Неверный формат дат. Пример: /simple_range 2025-08-01 2025-08-07")
        return
    await message.reply("🔄 Генерирую простой PPTX (по менеджерам + комментарии ИИ)…")
    try:
        container = Container.get()
        aggregator = DataAggregatorService(container.sheets)
        from bot.services.simple_presentation import SimplePresentationService
        presentation_service = SimplePresentationService(container.settings)
        period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_custom_with_previous(start, end)
        if not period_data:
            await message.reply("❌ Нет данных за этот период.")
            return
        pptx_bytes = await presentation_service.generate_presentation(period_data, period_name, start_date, end_date, prev_data, prev_start, prev_end)
        document = types.BufferedInputFile(pptx_bytes, filename=f"Отчет_По_Активности_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pptx")
        await message.reply_document(document, caption=f"📊 {period_name}\n🤖 Простой отчёт готов!")
    except Exception as e:
        await message.reply(f"❌ Ошибка: {str(e)}")


@admin_router.message(Command("slides_range"))
async def cmd_slides_range(message: types.Message, command: CommandObject) -> None:
    """Generate premium presentation for custom date range (PPTX or Google Slides based on config).
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

    container = Container.get()
    
    # Check if Google Slides is enabled
    if container.settings.use_google_slides:
        await message.reply("⚠️ Google Slides требует настройки Google Workspace с Общими дисками для снятия квот API.\n\n"
                           "Пока генерирую премиум PPTX (9 слайдов, графики, AI, светофор) — это займёт ~15 сек…")
    else:
        await message.reply("🔄 Генерирую премиум презентацию (9 слайдов, графики, AI-анализ)…")
    
    try:
        aggregator = DataAggregatorService(container.sheets)
        
        period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = await aggregator.aggregate_custom_with_previous(start, end)
        if not period_data:
            await message.reply("❌ Нет данных за этот период.")
            return
        
        # Get daily series for charts
        daily_series = await aggregator.get_daily_series(start_date, end_date)
        
        # Generate premium PPTX
        from bot.services.premium_presentation import PremiumPresentationService
        presentation_service = PremiumPresentationService(container.settings)
        pptx_bytes = await presentation_service.generate_presentation(
            period_data, period_name, start_date, end_date, prev_data, prev_start, prev_end, daily_series
        )
        
        document = types.BufferedInputFile(
            pptx_bytes,
            filename=f"Отчет_{container.settings.office_name}_Неделя_{period_name.replace(' ', '_')}.pptx"
        )
        
        await message.reply_document(
            document,
            caption=f"✅ Премиум отчёт готов!\n📊 {period_name}\n🤖 AI-комментарии, графики, светофор KPI, рейтинг ТОП/АнтиТОП"
        )
        
    except Exception as e:
        import html
        safe_err = html.escape(str(e))
        await message.reply(f"❌ Ошибка генерации: {safe_err}", parse_mode=None)


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
        if str(binding.get("chat_id")) != str(chat_id) or not topic_id_raw.isdigit():
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


# ====== Callback handlers for office management (HQ only) ======

@admin_router.callback_query(F.data == "admin_section_offices")
async def cb_offices_menu(query: types.CallbackQuery) -> None:
    from bot.keyboards.main import get_admin_offices_keyboard
    await query.message.edit_text("🏢 Управление офисами:", reply_markup=get_admin_offices_keyboard())
    await query.answer()


@admin_router.callback_query(F.data == "admin_back")
async def cb_admin_back(query: types.CallbackQuery) -> None:
    from bot.offices_config import is_hq
    is_hq_chat = is_hq(query.message.chat.id)
    await query.message.edit_text("Меню администратора:" + (" (Головной офис)" if is_hq_chat else ""), reply_markup=get_admin_menu_keyboard(is_hq=is_hq_chat))
    await query.answer()


# Office-specific presentation handlers
@admin_router.callback_query(F.data.in_(["presentation_office4", "presentation_sanzharovsky", "presentation_baturlov", "presentation_savela", "presentation_all_offices"]))
async def cb_presentation_office(query: types.CallbackQuery) -> None:
    # Early ack to avoid Telegram timeout on long generation
    try:
        await query.answer("⏳ Генерирую презентацию…", show_alert=False)
    except Exception:
        pass
    office_map = {
        "presentation_office4": "Офис 4",
        "presentation_sanzharovsky": "Санжаровский",
        "presentation_baturlov": "Батурлов",
        "presentation_savela": "Савела",
        "presentation_all_offices": None  # All offices
    }
    office = office_map.get(query.data)
    
    # Send status message
    await query.message.answer("🔄 Генерирую презентацию...")
    
    try:
        container = Container.get()
        aggregator = DataAggregatorService(container.sheets)
        from bot.services.simple_presentation import SimplePresentationService
        presentation_service = SimplePresentationService(container.settings)
        
        # Use current week
        from datetime import datetime as _dt, timedelta
        from bot.utils.time_utils import start_end_of_week_today
        start_str, end_str = start_end_of_week_today(container.settings)
        start = _dt.strptime(start_str, "%Y-%m-%d").date()
        end = _dt.strptime(end_str, "%Y-%m-%d").date()
        
        # Aggregate with office filter
        if office:
            period_data = await aggregator._aggregate_data_for_period(start, end, office_filter=office)
            delta = end - start
            prev_end = start - timedelta(days=1)
            prev_start = prev_end - delta
            prev_data = await aggregator._aggregate_data_for_period(prev_start, prev_end, office_filter=office)
            period_name = f"{office}: Неделя {start.strftime('%d.%m')}—{end.strftime('%d.%m.%Y')}"
        else:
            period_data = await aggregator._aggregate_data_for_period(start, end)
            delta = end - start
            prev_end = start - timedelta(days=1)
            prev_start = prev_end - delta
            prev_data = await aggregator._aggregate_data_for_period(prev_start, prev_end)
            period_name = f"Все офисы: Неделя {start.strftime('%d.%m')}—{end.strftime('%d.%m.%Y')}"
        
        if not period_data:
            await query.message.answer("❌ Нет данных за этот период.")
            await query.answer()
            return
        
        pptx_bytes = await presentation_service.generate_presentation(period_data, period_name, start, end, prev_data or {}, prev_start, prev_end, office_filter=office)
        document = types.BufferedInputFile(pptx_bytes, filename=f"Отчет_{office or 'Все'}_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pptx")
        await query.message.answer_document(document, caption=f"📊 {period_name}\n🤖 Презентация готова!")
        try:
            await query.answer()
        except Exception:
            pass
    except Exception as e:
        await query.message.answer(f"❌ Ошибка: {str(e)}")
        try:
            await query.answer()
        except Exception:
            pass


# Office-specific summary handlers
@admin_router.callback_query(F.data.in_(["summary_office4", "summary_sanzharovsky", "summary_baturlov", "summary_savela", "summary_all_offices"]))
async def cb_summary_office(query: types.CallbackQuery) -> None:
    # Early ack to avoid Telegram timeout on longer aggregations
    try:
        await query.answer("⏳ Считаю сводку…", show_alert=False)
    except Exception:
        pass
    office_map = {
        "summary_office4": "Офис 4",
        "summary_sanzharovsky": "Санжаровский",
        "summary_baturlov": "Батурлов",
        "summary_savela": "Савела",
        "summary_all_offices": None  # All offices
    }
    office = office_map.get(query.data)
    try:
        container = Container.get()
        aggregator = DataAggregatorService(container.sheets)
        
        # Use current week
        from datetime import datetime as _dt
        from bot.utils.time_utils import start_end_of_week_today
        start_str, end_str = start_end_of_week_today(container.settings)
        start = _dt.strptime(start_str, "%Y-%m-%d").date()
        end = _dt.strptime(end_str, "%Y-%m-%d").date()
        
        # Aggregate with office filter
        if office:
            period_data = await aggregator._aggregate_data_for_period(start, end, office_filter=office)
            period_name = f"{office}: {start.strftime('%d.%m')}—{end.strftime('%d.%m.%Y')}"
        else:
            period_data = await aggregator._aggregate_data_for_period(start, end)
            period_name = f"Все офисы: {start.strftime('%d.%m')}—{end.strftime('%d.%m.%Y')}"
        
        if not period_data:
            await query.message.answer(f"❌ Нет данных за {period_name}.")
            await query.answer()
            return
        
        # Build summary text
        total_calls = sum(m.calls_fact for m in period_data.values())
        total_new = sum(m.new_calls for m in period_data.values())
        total_leads = sum(m.leads_units_fact for m in period_data.values())
        total_volume = sum(m.leads_volume_fact for m in period_data.values())
        total_issued = sum(m.issued_volume for m in period_data.values())
        
        response = f"📊 Сводка: {period_name}\n\n"
        response += f"👥 Менеджеров: {len(period_data)}\n"
        response += f"📞 Перезвоны: {total_calls}\n"
        response += f"🆕 Новые звонки: {total_new}\n"
        response += f"📋 Заявки (шт): {total_leads}\n"
        response += f"💰 Заявки (млн): {total_volume:.1f}\n"
        response += f"✅ Выдано (млн): {total_issued:.1f}\n\n"
        response += "Топ-3 менеджера:\n"
        
        # Rank by issued
        ranked = sorted(period_data.items(), key=lambda x: x[1].issued_volume, reverse=True)[:3]
        for i, (name, m) in enumerate(ranked, start=1):
            response += f"{i}. {name}: {m.issued_volume:.1f} млн выдано\n"
        
        await query.message.answer(response)
        try:
            await query.answer()
        except Exception:
            pass
    except Exception as e:
        await query.message.answer(f"❌ Ошибка: {str(e)}")
        try:
            await query.answer()
        except Exception:
            pass


# Compare offices
@admin_router.callback_query(F.data == "compare_offices")
async def cb_compare_offices(query: types.CallbackQuery) -> None:
    # Early ack to avoid Telegram timeout
    try:
        await query.answer("⏳ Сравниваю офисы…", show_alert=False)
    except Exception:
        pass
    try:
        container = Container.get()
        aggregator = DataAggregatorService(container.sheets)
        
        from datetime import datetime as _dt
        from bot.utils.time_utils import start_end_of_week_today
        start_str, end_str = start_end_of_week_today(container.settings)
        start = _dt.strptime(start_str, "%Y-%m-%d").date()
        end = _dt.strptime(end_str, "%Y-%m-%d").date()
        
        offices = ["Офис 4", "Санжаровский", "Батурлов", "Савела"]
        response = f"📈 Сравнение офисов: {start.strftime('%d.%m')}—{end.strftime('%d.%m.%Y')}\n\n"
        
        office_stats = []
        for office in offices:
            data = await aggregator._aggregate_data_for_period(start, end, office_filter=office)
            if data:
                total_calls = sum(m.calls_fact for m in data.values())
                total_issued = sum(m.issued_volume for m in data.values())
                office_stats.append((office, len(data), total_calls, total_issued))
        
        # Sort by issued
        office_stats.sort(key=lambda x: x[3], reverse=True)
        
        for i, (office, mgrs, calls, issued) in enumerate(office_stats, start=1):
            medal = "🥇" if i==1 else "🥈" if i==2 else "🥉"
            response += f"{medal} {office}:\n"
            response += f"   👥 Менеджеров: {mgrs}\n"
            response += f"   📞 Звонков: {calls}\n"
            response += f"   ✅ Выдано: {issued:.1f} млн\n\n"
        
        await query.message.answer(response)
        try:
            await query.answer()
        except Exception:
            pass
    except Exception as e:
        await query.message.answer(f"❌ Ошибка: {str(e)}")
        try:
            await query.answer()
        except Exception:
            pass
