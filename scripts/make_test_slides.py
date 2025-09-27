#!/usr/bin/env python3
from __future__ import annotations

import os
import sys
from datetime import datetime, timedelta, date

from dotenv import load_dotenv

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from bot.config import Settings
from bot.services.di import Container
from bot.services.data_aggregator import DataAggregatorService
from bot.services.presentation import PresentationService
from bot.services.google_slides import GoogleSlidesService


def parse_dates(args: list[str]) -> tuple[date, date]:
    if len(args) >= 2:
        a = datetime.strptime(args[0], "%Y-%m-%d").date()
        b = datetime.strptime(args[1], "%Y-%m-%d").date()
        return a, b
    # default: last full week Mon-Sun
    today = date.today()
    last_sun = today - timedelta(days=today.weekday() + 1)
    start = last_sun - timedelta(days=6)
    return start, last_sun


def main():
    load_dotenv()
    # Ensure creds
    os.environ.setdefault("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
    os.environ.setdefault("BOT_TOKEN", "dummy-token-for-offline")
    os.environ.setdefault("USE_GOOGLE_SLIDES", "true")
    # DRIVE_FOLDER_ID should be set in env; otherwise will raise later

    settings = Settings.load()
    Container.init(settings)

    start, end = parse_dates(sys.argv[1:])
    print(f"Generating Slides for {start}..{end}")

    aggregator = DataAggregatorService(Container.get().sheets)
    period_data, prev_data, period_name, start_date, end_date, prev_start, prev_end = (
        __import__("asyncio").get_event_loop().run_until_complete(
            aggregator.aggregate_custom_with_previous(start, end)
        )
    )
    if not period_data:
        print("No data for the requested period.")
        return 1

    slides = GoogleSlidesService(settings)
    prs_service = PresentationService(settings)
    deck_id = slides.create_presentation(f"Отчет {period_name}")
    # Move to configured folder if set
    try:
        slides.move_presentation_to_folder(deck_id)
    except Exception:
        pass

    try:
        logo_id = slides.upload_logo_to_drive()
        slides.apply_branding(deck_id, logo_id)
    except Exception:
        pass

    totals = prs_service._calculate_totals(period_data)
    __import__("asyncio").get_event_loop().run_until_complete(
        slides.build_title_and_summary(deck_id, "Офис", f"{period_name} — {start.strftime('%d.%m.%Y')}—{end.strftime('%d.%m.%Y')}", totals)
    )

    if prev_data:
        prev_totals = prs_service._calculate_totals(prev_data)
        __import__("asyncio").get_event_loop().run_until_complete(
            slides.add_comparison_with_ai(deck_id, prev_totals, totals, "Динамика: предыдущий vs текущий")
        )

    # Simple ranking
    try:
        kpi = {}
        for m in period_data.values():
            kpi[m.name] = {
                'calls_plan': m.calls_plan,
                'calls_fact': m.calls_fact,
                'leads_units_plan': m.leads_units_plan,
                'leads_units_fact': m.leads_units_fact,
                'leads_volume_plan': m.leads_volume_plan,
                'leads_volume_fact': m.leads_volume_fact,
            }
        scored = []
        for name, v in kpi.items():
            calls_pct = (v['calls_fact']/v['calls_plan']*100) if v['calls_plan'] else 0
            vol_pct = (v['leads_volume_fact']/v['leads_volume_plan']*100) if v['leads_volume_plan'] else 0
            scored.append((0.5*calls_pct+0.5*vol_pct, name))
        scored.sort(reverse=True)
        best = [n for _, n in scored[:2]]
        worst = [n for _, n in list(reversed(scored[-2:]))]
        ranking = {"best": best, "worst": worst, "reasons": {}}
        __import__("asyncio").get_event_loop().run_until_complete(slides.add_top2_antitop2(deck_id, ranking))
    except Exception:
        pass

    # Export PDF locally for quick check
    pdf_bytes = slides.export_pdf(deck_id)
    out = f"test_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pdf"
    with open(out, "wb") as f:
        f.write(pdf_bytes)
    print(f"PDF saved: {out}")
    print("Deck is created in your Drive folder as well.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


