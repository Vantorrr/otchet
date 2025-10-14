#!/usr/bin/env python3
"""Debug office assignments in Google Sheets"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from bot.services.summary_builder import build_summary_text
from bot.config import Settings
from bot.services.di import Container

# Load settings and initialize container
settings = Settings.load()
Container.init(settings)
container = Container.get()

# Test summary generation for Savela office
print("=" * 60)
print("TESTING SUMMARY FOR SAVELA OFFICE")
print("=" * 60)

# Generate summary with Savela filter
summary_text = build_summary_text(
    container.settings, 
    container.sheets, 
    day="2025-10-13", 
    office_filter="Савела"
)

print("\nSummary with office_filter='Савела':")
print("-" * 60)
print(summary_text)
print("-" * 60)

# Check raw data
print("\n\nCHECKING RAW DATA:")
print("=" * 60)

all_records = container.sheets._reports.get_all_records()
today_records = [r for r in all_records if r.get('date') == '2025-10-13']

# Group by office
office_counts = {}
for r in today_records:
    office = r.get('office', 'NO_OFFICE')
    if office not in office_counts:
        office_counts[office] = []
    office_counts[office].append(r.get('manager', 'NO_MANAGER'))

print(f"\nTotal records for 2025-10-13: {len(today_records)}")
print("\nRecords by office:")
for office, managers in sorted(office_counts.items()):
    print(f"\n{office}:")
    manager_counts = {}
    for m in managers:
        manager_counts[m] = manager_counts.get(m, 0) + 1
    for manager, count in sorted(manager_counts.items()):
        print(f"  - {manager}: {count} records")

