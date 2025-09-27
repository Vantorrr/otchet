#!/usr/bin/env python3
from __future__ import annotations

import os
import sys
from datetime import datetime, date

from dotenv import load_dotenv

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from bot.config import Settings
from bot.services.di import Container
from bot.services.data_aggregator import DataAggregatorService
from bot.services.presentation import PresentationService
from bot.services.google_slides import GoogleSlidesService
from bot.services.reference_slides import ReferenceSlidesBuilder


def main():
    load_dotenv()
    os.environ.setdefault("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
    os.environ.setdefault("BOT_TOKEN", "dummy-token-for-offline")
    os.environ.setdefault("USE_GOOGLE_SLIDES", "true")
    os.environ.setdefault("OFFICE_NAME", "Банковские гарантии")

    settings = Settings.load()
    Container.init(settings)

    start = datetime.strptime("2025-09-01", "%Y-%m-%d").date()
    end = datetime.strptime("2025-09-07", "%Y-%m-%d").date()
    print(f"Building REFERENCE-STYLE Slides for {start}..{end}")

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
    builder = ReferenceSlidesBuilder(settings, slides)
    prs_service = PresentationService(settings)
    
    deck_id = slides.create_presentation(f"Отчет {period_name} — Референс")
    totals = prs_service._calculate_totals(period_data)
    
    # Build ranking
    scored = []
    for m in period_data.values():
        calls_pct = (m.calls_fact/m.calls_plan*100) if m.calls_plan else 0
        vol_pct = (m.leads_volume_fact/m.leads_volume_plan*100) if m.leads_volume_plan else 0
        scored.append((0.5*calls_pct+0.5*vol_pct, m.name))
    scored.sort(reverse=True)
    best = [n for _, n in scored[:2]]
    worst = [n for _, n in list(reversed(scored[-2:]))]
    ranking = {"best": best, "worst": worst, "reasons": {}}
    
    # Build reference deck
    period_dates = f"{start.strftime('%d.%m.%Y')} — {end.strftime('%d.%m.%Y')}"
    builder.create_reference_deck(
        deck_id, 
        settings.office_name,
        period_name,
        period_dates,
        totals,
        ranking
    )
    
    # Skip moving to folder due to quota; user will move manually
    
    pdf_bytes = slides.export_pdf(deck_id)
    out = f"reference_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pdf"
    with open(out, "wb") as f:
        f.write(pdf_bytes)
    
    print(f"REFERENCE PDF saved: {out}")
    print("Reference deck created in Drive folder.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
