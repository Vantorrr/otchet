"""Build office-grouped summaries for HQ."""

from typing import Dict, List
from datetime import date

from bot.config import Settings
from bot.services.sheets import SheetsClient
from bot.services.data_aggregator import DataAggregatorService
from bot.offices_config import get_all_offices


async def build_office_summary_text(
    settings: Settings,
    sheets: SheetsClient,
    start: str,
    end: str,
) -> str:
    """Build summary grouped by offices for HQ."""
    aggregator = DataAggregatorService(sheets)
    all_offices = get_all_offices()
    
    # Parse dates
    from datetime import datetime as _dt
    start_date = _dt.strptime(start, "%Y-%m-%d").date()
    end_date = _dt.strptime(end, "%Y-%m-%d").date()
    
    response = f"📊 <b>Сводка по офисам: {start_date.strftime('%d.%m.%Y')}—{end_date.strftime('%d.%m.%Y')}</b>\n"
    
    total_calls_plan = 0
    total_new_calls_plan = 0
    total_leads_plan_units = 0
    total_leads_plan_volume = 0.0
    total_calls_fact = 0
    total_new_calls = 0
    total_leads_units = 0
    total_leads_volume = 0.0
    total_approved = 0.0
    total_issued = 0.0
    
    for office in all_offices:
        office_data = await aggregator._aggregate_data_for_period(start_date, end_date, office_filter=office)
        if office_data:
            # Calculate office totals (fields from ManagerData in presentation.py)
            office_calls_plan = sum(m.calls_plan for m in office_data.values())
            office_new_calls_plan = sum(m.new_calls_plan for m in office_data.values())
            office_leads_plan_units = sum(m.leads_units_plan for m in office_data.values())
            office_leads_plan_volume = sum(m.leads_volume_plan for m in office_data.values())
            
            office_calls_fact = sum(m.calls_fact for m in office_data.values())
            office_new_calls = sum(m.new_calls for m in office_data.values())
            office_leads_units = sum(m.leads_units_fact for m in office_data.values())
            office_leads_volume = sum(m.leads_volume_fact for m in office_data.values())
            office_approved = sum(m.approved_volume for m in office_data.values())
            office_issued = sum(m.issued_volume for m in office_data.values())
            
            response += f"\n\n🏢 <b>{office}</b>\n"
            response += f"👥 Менеджеров: {len(office_data)}\n"
            response += "<b>План</b>\n"
            response += f"• 📲 Перезвоны: <b>{office_calls_plan}</b>\n"
            response += f"• ☎️ Новые звонки: <b>{office_new_calls_plan}</b>\n"
            response += f"• 📝 Заявки, шт: <b>{office_leads_plan_units}</b>\n"
            response += f"• 💰 Заявки, млн: <b>{office_leads_plan_volume:.1f}</b>\n"
            response += "\n<b>Факт</b>\n"
            response += f"• 📲 Перезвоны: <b>{office_calls_fact}</b> из <b>{office_calls_plan}</b>"
            if office_calls_plan > 0:
                response += f" ({office_calls_fact/office_calls_plan*100:.1f}%)"
            response += "\n"
            response += f"• ☎️ Новые звонки: <b>{office_new_calls}</b>\n"
            response += f"• 📝 Заявки, шт: <b>{office_leads_units}</b>\n"
            response += f"• 💰 Заявки, млн: <b>{office_leads_volume:.1f}</b>\n"
            response += f"• ✅ Одобрено, млн: <b>{office_approved:.1f}</b>\n"
            response += f"• ✅ Выдано, млн: <b>{office_issued:.1f}</b>\n"
            
            # Add to totals
            total_calls_plan += office_calls_plan
            total_new_calls_plan += office_new_calls_plan
            total_leads_plan_units += office_leads_plan_units
            total_leads_plan_volume += office_leads_plan_volume
            total_calls_fact += office_calls_fact
            total_new_calls += office_new_calls
            total_leads_units += office_leads_units
            total_leads_volume += office_leads_volume
            total_approved += office_approved
            total_issued += office_issued
    
    # Add totals
    response += "\n" + "="*40 + "\n"
    response += "<b>📊 ИТОГО ПО ВСЕМ ОФИСАМ</b>\n"
    response += "<b>План</b>\n"
    response += f"• 📲 Перезвоны: <b>{total_calls_plan}</b>\n"
    response += f"• ☎️ Новые звонки: <b>{total_new_calls_plan}</b>\n"
    response += f"• 📝 Заявки, шт: <b>{total_leads_plan_units}</b>\n"
    response += f"• 💰 Заявки, млн: <b>{total_leads_plan_volume:.1f}</b>\n"
    response += "\n<b>Факт</b>\n"
    response += f"• 📲 Перезвоны: <b>{total_calls_fact}</b> из <b>{total_calls_plan}</b>"
    if total_calls_plan > 0:
        response += f" ({total_calls_fact/total_calls_plan*100:.1f}%)"
    response += "\n"
    response += f"• ☎️ Новые звонки: <b>{total_new_calls}</b>\n"
    response += f"• 📝 Заявки, шт: <b>{total_leads_units}</b>\n"
    response += f"• 💰 Заявки, млн: <b>{total_leads_volume:.1f}</b>\n"
    response += f"• ✅ Одобрено, млн: <b>{total_approved:.1f}</b>\n"
    response += f"• ✅ Выдано, млн: <b>{total_issued:.1f}</b>\n"
    response += "="*40
    
    return response
