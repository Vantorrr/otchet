#!/usr/bin/env python3
"""Quick check of Google Sheets data without DI container."""

import gspread
from google.oauth2.service_account import Credentials
import os

# Setup credentials
creds = Credentials.from_service_account_file(
    'service_account.json',
    scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
)
gc = gspread.authorize(creds)

# Open spreadsheet
sheet = gc.open("Sales Reports")
reports = sheet.worksheet("Reports")

# Get today's data
all_records = reports.get_all_records()
today_records = [r for r in all_records if r.get('date') == '2025-10-13']

# Group by office
print("=" * 60)
print("ДАННЫЕ ЗА 2025-10-13")
print("=" * 60)

office_groups = {}
for r in today_records:
    office = r.get('office', 'NO_OFFICE')
    manager = r.get('manager', 'NO_MANAGER')
    if office not in office_groups:
        office_groups[office] = []
    office_groups[office].append(manager)

for office, managers in sorted(office_groups.items()):
    print(f"\nОфис: {office}")
    print(f"Менеджеры ({len(managers)}): {', '.join(sorted(set(managers)))}")

# Check if all managers are assigned to Savela
savela_managers = office_groups.get('Савела', [])
print("\n" + "=" * 60)
print(f"В офисе Савела найдено {len(savela_managers)} записей")
if savela_managers:
    print(f"Менеджеры: {', '.join(sorted(set(savela_managers)))}")

