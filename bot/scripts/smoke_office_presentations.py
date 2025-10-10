from __future__ import annotations

"""CLI smoke test: generate weekly PPTX per office and for all offices.

Usage on server:
  .venv/bin/python -m bot.scripts.smoke_office_presentations
"""

import os
from pathlib import Path
from datetime import datetime as _dt, timedelta

from bot.config import Settings
from bot.services.di import Container
from bot.services.data_aggregator import DataAggregatorService
from bot.services.simple_presentation import SimplePresentationService
from bot.utils.time_utils import start_end_of_week_today


def main() -> int:
    # Load settings and init DI (uses .env)
    settings = Settings.load()
    Container.init(settings)

    # Ensure Google creds env for gspread
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = settings.google_credentials_path

    container = Container.get()
    aggregator = DataAggregatorService(container.sheets)
    pres = SimplePresentationService(container.settings)

    # Current week boundaries
    start_str, end_str = start_end_of_week_today(container.settings)
    start = _dt.strptime(start_str, "%Y-%m-%d").date()
    end = _dt.strptime(end_str, "%Y-%m-%d").date()
    delta = end - start
    prev_end = start - timedelta(days=1)
    prev_start = prev_end - delta

    offices = ["Офис 4", "Санжаровский", "Батурлов", "Савела", None]  # None = all offices

    out_dir = Path("out")
    out_dir.mkdir(parents=True, exist_ok=True)

    # Use asyncio.run to execute async blocks sequentially

    import asyncio

    async def run_one(office: str | None):
        if office:
            cur = await aggregator._aggregate_data_for_period(start, end, office_filter=office)
            prev = await aggregator._aggregate_data_for_period(prev_start, prev_end, office_filter=office)
            title = f"{office}: Неделя {start.strftime('%d.%m')}—{end.strftime('%d.%m.%Y')}"
            label = office
        else:
            cur = await aggregator._aggregate_data_for_period(start, end)
            prev = await aggregator._aggregate_data_for_period(prev_start, prev_end)
            title = f"Все офисы: Неделя {start.strftime('%d.%m')}—{end.strftime('%d.%m.%Y')}"
            label = "Все"

        if not cur:
            return f"{label}: данных нет — пропуск"

        pptx_bytes = await pres.generate_presentation(cur, title, start, end, prev or {}, prev_start, prev_end)
        out_path = out_dir / f"Отчет_{label}_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pptx"
        out_path.write_bytes(pptx_bytes)
        # Short textual confirmation
        mgrs = len(cur)
        total_calls = sum(m.calls_fact for m in cur.values())
        total_issued = sum(m.issued_volume for m in cur.values())
        return f"{label}: OK → {out_path.name} (менеджеров={mgrs}, звонков={total_calls}, выдано={total_issued:.1f} млн)"

    async def run_all():
        msgs = []
        for off in offices:
            msg = await run_one(off)
            msgs.append(msg)
        return msgs

    msgs = __import__("asyncio").run(run_all())
    for m in msgs:
        print(m)

    print(f"\nГотово. Файлы в: {out_dir.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


