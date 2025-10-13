from __future__ import annotations

from typing import List, Dict, Any, Optional
from datetime import datetime, timedelta

from bot.config import Settings
from bot.services.sheets import SheetsClient


def _int_or_zero(value: object) -> int:
    try:
        return int(value)
    except Exception:
        return 0


def _normalize_date(value: Any) -> Optional[str]:
    """Return YYYY-MM-DD or None if cannot parse."""
    if value is None:
        return None
    # Google Sheets may return strings, or numbers (serial), rarely datetime
    if isinstance(value, str):
        s = value.strip()
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.strftime("%Y-%m-%d")
            except Exception:
                pass
        return None
    if isinstance(value, (int, float)):
        # Google Sheets serial date (epoch 1899-12-30)
        try:
            base = datetime(1899, 12, 30)
            dt = base + timedelta(days=float(value))
            return dt.strftime("%Y-%m-%d")
        except Exception:
            return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return None


def _within(value: Any, start: Optional[str], end: Optional[str]) -> bool:
    ds = _normalize_date(value)
    if ds is None:
        return False
    if start and ds < start:
        return False
    if end and ds > end:
        return False
    return True


def build_summary_text(settings: Settings, sheets: SheetsClient, day: str, *, start: str | None = None, end: str | None = None, office_filter: str | None = None) -> str:
    all_records = sheets._reports.get_all_records()
    
    # Debug logging
    import logging
    logger = logging.getLogger(__name__)
    logger.info(f"build_summary_text called with office_filter={office_filter}")
    logger.info(f"Total records before filtering: {len(all_records)}")
    
    # First filter by office if needed
    if office_filter:
        # Debug: show what offices we have in data
        offices_in_data = set(r.get("office", "NO_OFFICE") for r in all_records)
        logger.info(f"Offices found in data: {sorted(offices_in_data)}")
        
        all_records = [r for r in all_records if r.get("office") == office_filter]
        logger.info(f"Records after office filtering: {len(all_records)}")
        
        # Debug: show sample of filtered records
        if all_records:
            sample_managers = set(r.get("manager", "NO_MANAGER") for r in all_records[:10])
            logger.info(f"Sample managers in filtered data: {sorted(sample_managers)}")
    
    # Then filter by date
    if start or end:
        reports: List[Dict[str, Any]] = [r for r in all_records if _within(r.get("date"), start, end)]
        title = f"📊 <b>Сводка за период {start} — {end}</b>"
        if office_filter:
            title = f"📊 <b>Сводка ({office_filter}) за период {start} — {end}</b>"
    else:
        # Используем ту же логику _within для одной даты
        reports: List[Dict[str, Any]] = [r for r in all_records if _within(r.get("date"), day, day)]
        title = f"📊 <b>Сводка за {day}</b>"
        if office_filter:
            title = f"📊 <b>Сводка ({office_filter}) за {day}</b>"
    
    logger.info(f"Records after date filtering: {len(reports)}")
    if reports and office_filter:
        # Show what offices and managers we have in final data
        final_offices = set(r.get("office", "NO_OFFICE") for r in reports)
        final_managers = set(r.get("manager", "NO_MANAGER") for r in reports)
        logger.info(f"Final offices in reports: {sorted(final_offices)}")
        logger.info(f"Final managers in reports: {sorted(final_managers)}")

    lines: List[str] = [title]

    # Группируем данные по менеджерам и суммируем
    managers_data: Dict[str, Dict[str, int]] = {}
    
    for r in reports:
        manager = r.get("manager", "?")
        if manager not in managers_data:
            managers_data[manager] = {
                "calls_planned": 0,
                "leads_planned_units": 0,
                "leads_planned_volume": 0,
                "new_calls_planned": 0,
                "calls_success": 0,
                "leads_units": 0,
                "leads_volume": 0,
                "approved_volume": 0,
                "issued_volume": 0,
                "new_calls": 0,
            }
        
        # Суммируем данные
        managers_data[manager]["calls_planned"] += _int_or_zero(r.get("morning_calls_planned", 0))
        managers_data[manager]["leads_planned_units"] += _int_or_zero(r.get("morning_leads_planned_units", 0))
        managers_data[manager]["leads_planned_volume"] += _int_or_zero(r.get("morning_leads_planned_volume", 0))
        managers_data[manager]["new_calls_planned"] += _int_or_zero(r.get("morning_new_calls_planned", 0))
        
        managers_data[manager]["calls_success"] += _int_or_zero(r.get("evening_calls_success", 0))
        managers_data[manager]["leads_units"] += _int_or_zero(r.get("evening_leads_units", 0))
        managers_data[manager]["leads_volume"] += _int_or_zero(r.get("evening_leads_volume", 0))
        managers_data[manager]["approved_volume"] += _int_or_zero(r.get("evening_approved_volume", 0))
        managers_data[manager]["issued_volume"] += _int_or_zero(r.get("evening_issued_volume", 0))
        managers_data[manager]["new_calls"] += _int_or_zero(r.get("evening_new_calls", 0))

    # Выводим данные по каждому менеджеру
    for manager, data in managers_data.items():
        calls_planned = data["calls_planned"]
        leads_planned_units = data["leads_planned_units"]
        leads_planned_volume = data["leads_planned_volume"]
        new_calls_planned = data["new_calls_planned"]
        
        calls_success = data["calls_success"]
        leads_units = data["leads_units"]
        leads_volume = data["leads_volume"]
        approved_volume = data["approved_volume"]
        issued_volume = data["issued_volume"]
        new_calls = data["new_calls"]

        # Прогнозность: перезвоны и заявки (объём)
        if calls_planned > 0:
            calls_forecast_pct = f"{(calls_success / calls_planned * 100):.1f}%"
            calls_forecast_pair = f"{calls_success}/{calls_planned} ({calls_forecast_pct})"
        elif calls_success > 0:
            calls_forecast_pair = f"{calls_success}/{calls_planned} (план был 0)"
        else:
            calls_forecast_pair = f"{calls_success}/{calls_planned}"
            
        if leads_planned_volume > 0:
            vol_forecast_pct = f"{(leads_volume / leads_planned_volume * 100):.1f}%"
            vol_forecast_pair = f"{leads_volume}/{leads_planned_volume} ({vol_forecast_pct})"
        elif leads_volume > 0:
            vol_forecast_pair = f"{leads_volume}/{leads_planned_volume} (план был 0)"
        else:
            vol_forecast_pair = f"{leads_volume}/{leads_planned_volume}"

        lines.append(
            "\n".join(
                [
                    f"\n<b>👤 {manager}</b>",
                    "<b>План</b>",
                    f"• 📲 Перезвоны: <b>{calls_planned}</b>",
                    f"   ☎️ Новые звонки: <b>{new_calls_planned}</b>",
                    f"• 📝 Заявки, шт: <b>{leads_planned_units}</b>",
                    f"• 💰 Заявки, млн: <b>{leads_planned_volume}</b>",
                    "",
                    "<b>Факт</b>",
                    f"• 📲 Перезвоны: <b>{calls_success}</b> из <b>{calls_planned}</b>",
                    f"•  ☎️ Новые звонки: <b>{new_calls}</b>",
                    f"• 📝Заявки, шт: <b>{leads_units}</b>",
                    f"• 💰 Заявки, млн: <b>{leads_volume}</b>",
                    f"• ✅ Одобрено, млн: <b>{approved_volume}</b>",
                    f"• ✅ Выдано, млн: <b>{issued_volume}</b>",
                    "<b>Прогнозность</b>",
                    f"• 🔮 Перезвоны (факт/план): <b>{calls_forecast_pair}</b>",
                    f"• 🔮 Заявки (объём) факт/план: <b>{vol_forecast_pair}</b>",
                ]
            )
        )

    if not managers_data:
        if start or end:
            lines.append("Нет данных за этот период.")
        else:
            lines.append(f"Нет данных за {day}.")
    else:
        # Добавляем суммарный итог по всем менеджерам
        total_calls_planned = sum(data["calls_planned"] for data in managers_data.values())
        total_leads_planned_units = sum(data["leads_planned_units"] for data in managers_data.values())
        total_leads_planned_volume = sum(data["leads_planned_volume"] for data in managers_data.values())
        total_new_calls_planned = sum(data["new_calls_planned"] for data in managers_data.values())
        
        total_calls_success = sum(data["calls_success"] for data in managers_data.values())
        total_leads_units = sum(data["leads_units"] for data in managers_data.values())
        total_leads_volume = sum(data["leads_volume"] for data in managers_data.values())
        total_approved_volume = sum(data["approved_volume"] for data in managers_data.values())
        total_issued_volume = sum(data["issued_volume"] for data in managers_data.values())
        total_new_calls = sum(data["new_calls"] for data in managers_data.values())

        # Прогнозность итоговая
        if total_calls_planned > 0:
            total_calls_forecast_pct = f"{(total_calls_success / total_calls_planned * 100):.1f}%"
            total_calls_forecast_pair = f"{total_calls_success}/{total_calls_planned} ({total_calls_forecast_pct})"
        elif total_calls_success > 0:
            total_calls_forecast_pair = f"{total_calls_success}/{total_calls_planned} (план был 0)"
        else:
            total_calls_forecast_pair = f"{total_calls_success}/{total_calls_planned}"
            
        if total_leads_planned_volume > 0:
            total_vol_forecast_pct = f"{(total_leads_volume / total_leads_planned_volume * 100):.1f}%"
            total_vol_forecast_pair = f"{total_leads_volume}/{total_leads_planned_volume} ({total_vol_forecast_pct})"
        elif total_leads_volume > 0:
            total_vol_forecast_pair = f"{total_leads_volume}/{total_leads_planned_volume} (план был 0)"
        else:
            total_vol_forecast_pair = f"{total_leads_volume}/{total_leads_planned_volume}"

        lines.append(
            "\n".join(
                [
                    "\n" + "="*40,
                    f"<b>📊 ИТОГО ПО КОМАНДЕ</b>",
                    "<b>План</b>",
                    f"• 📲 Перезвоны: <b>{total_calls_planned}</b>",
                    f"   ☎️ Новые звонки: <b>{total_new_calls_planned}</b>",
                    f"• 📝 Заявки, шт: <b>{total_leads_planned_units}</b>",
                    f"• 💰 Заявки, млн: <b>{total_leads_planned_volume}</b>",
                    "",
                    "<b>Факт</b>",
                    f"• 📲 Перезвоны: <b>{total_calls_success}</b> из <b>{total_calls_planned}</b>",
                    f"•  ☎️ Новые звонки: <b>{total_new_calls}</b>",
                    f"• 📝Заявки, шт: <b>{total_leads_units}</b>",
                    f"• 💰 Заявки, млн: <b>{total_leads_volume}</b>",
                    f"• ✅ Одобрено, млн: <b>{total_approved_volume}</b>",
                    f"• ✅ Выдано, млн: <b>{total_issued_volume}</b>",
                    "<b>Прогнозность</b>",
                    f"• 🔮 Перезвоны (факт/план): <b>{total_calls_forecast_pair}</b>",
                    f"• 🔮 Заявки (объём) факт/план: <b>{total_vol_forecast_pair}</b>",
                    "="*40,
                ]
            )
        )

    return "\n".join(lines)


