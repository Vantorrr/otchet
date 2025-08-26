from __future__ import annotations

from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from bot.config import Settings


def now_in_tz(settings: Settings) -> datetime:
    return datetime.now(ZoneInfo(settings.default_timezone))


def date_str_for_today(settings: Settings) -> str:
    return now_in_tz(settings).strftime("%Y-%m-%d")


def parse_date_or_today(arg: str | None, settings: Settings) -> str:
    """Parse user-provided date into YYYY-MM-DD.

    Accepts common formats and normalizes unicode dashes. Falls back to today on failure.
    Supported examples:
    - 2025-08-22
    - 22.08.2025
    - 2025/08/22
    - 22/08/2025
    - 2025–08–22 (en/em/figure dash)
    """
    if not arg:
        return date_str_for_today(settings)

    s = str(arg).strip()
    # Replace various unicode dashes with ASCII hyphen
    for dash in ("–", "—", "−", "‒", "―"):
        s = s.replace(dash, "-")
    # Only first token if extra text provided
    s = s.split()[0]

    # Try multiple formats; return normalized
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y/%m/%d", "%d/%m/%Y"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue

    return date_str_for_today(settings)


def start_end_of_week_today(settings: Settings) -> tuple[str, str]:
    now = now_in_tz(settings)
    # Monday as start
    start = (now - timedelta(days=now.weekday())).date()
    end = start + timedelta(days=6)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def start_end_of_month_today(settings: Settings) -> tuple[str, str]:
    now = now_in_tz(settings)
    start = now.replace(day=1).date()
    # next month start - 1 day
    if start.month == 12:
        next_month = start.replace(year=start.year + 1, month=1, day=1)
    else:
        next_month = start.replace(month=start.month + 1, day=1)
    end = (datetime.combine(next_month, datetime.min.time()) - timedelta(days=1)).date()
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def start_end_of_quarter_today(settings: Settings) -> tuple[str, str]:
    now = now_in_tz(settings)
    q = (now.month - 1) // 3  # 0..3
    start_month = q * 3 + 1
    start = now.replace(month=start_month, day=1).date()
    # quarter end month
    end_month = start_month + 2
    if end_month == 12:
        next_month = start.replace(year=start.year + 1, month=1, day=1)
    else:
        next_month = start.replace(month=end_month + 1, day=1)
    end = (datetime.combine(next_month, datetime.min.time()) - timedelta(days=1)).date()
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
