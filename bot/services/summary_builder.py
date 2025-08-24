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


def build_summary_text(settings: Settings, sheets: SheetsClient, day: str, *, start: str | None = None, end: str | None = None) -> str:
    if start or end:
        all_records = sheets._reports.get_all_records()
        reports: List[Dict[str, Any]] = [r for r in all_records if _within(r.get("date"), start, end)]
        title = f"📊 <b>Сводка за период {start} — {end}</b>"
    else:
        reports = sheets.get_reports_by_date(day)
        title = f"📊 <b>Сводка за {day}</b>"

    lines: List[str] = [title]

    total_calls_planned = 0
    total_calls_success = 0

    for r in reports:
        manager = r.get("manager", "?")
        calls_planned = _int_or_zero(r.get("morning_calls_planned", 0))
        leads_planned_units = _int_or_zero(r.get("morning_leads_planned_units", 0))
        leads_planned_volume = _int_or_zero(r.get("morning_leads_planned_volume", 0))
        new_calls_planned = _int_or_zero(r.get("morning_new_calls_planned", 0))

        calls_success = _int_or_zero(r.get("evening_calls_success", 0))
        leads_units = _int_or_zero(r.get("evening_leads_units", 0))
        leads_volume = _int_or_zero(r.get("evening_leads_volume", 0))
        approved_volume = _int_or_zero(r.get("evening_approved_volume", 0))
        issued_volume = _int_or_zero(r.get("evening_issued_volume", 0))
        new_calls = _int_or_zero(r.get("evening_new_calls", 0))

        # Прогнозность: перезвоны и заявки (объём)
        if calls_planned > 0:
            calls_forecast_pct = f"{(calls_success / calls_planned * 100):.1f}%"
            calls_forecast_pair = f"{calls_success}/{calls_planned} ({calls_forecast_pct})"
        else:
            calls_forecast_pair = "—"
        if leads_planned_volume > 0:
            vol_forecast_pct = f"{(leads_volume / leads_planned_volume * 100):.1f}%"
            vol_forecast_pair = f"{leads_volume}/{leads_planned_volume} ({vol_forecast_pct})"
        else:
            vol_forecast_pair = "—"

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
        total_calls_planned += calls_planned
        total_calls_success += calls_success

    if not reports:
        lines.append("Нет данных на этот день.")

    return "\n".join(lines)


