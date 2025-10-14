#!/usr/bin/env python3
"""Check Savela office data in Google Sheets"""

import os
import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from bot.services.sheets import SheetsClient
from bot.config import Settings
from bot.services.di import Container

# Initialize container
from bot.config import Settings
settings = Settings.load()
Container.init(settings)
container = Container.get()

# Get all records
all_records = container.sheets._reports.get_all_records()
today_records = [r for r in all_records if r.get('date') == '2025-10-13']

# Group by offices
office_managers = {}
for r in today_records:
    office = r.get('office', 'NO_OFFICE')
    manager = r.get('manager', 'NO_MANAGER')
    if office not in office_managers:
        office_managers[office] = set()
    office_managers[office].add(manager)

print('Данные за 2025-10-13:')
for office, managers in sorted(office_managers.items()):
    print(f'\nОфис: {office}')
    print(f'Менеджеры: {sorted(managers)}')

# Check Savela records specifically
print('\n\nЗаписи офиса Савела:')
savela_records = [r for r in today_records if r.get('office') == 'Савела']
print(f'Найдено записей: {len(savela_records)}')
for r in savela_records[:5]:  # первые 5 записей для примера
    print(f'  - {r.get("manager")}: дата={r.get("date")} офис={r.get("office")}')

# Show what office filter should be
from bot.offices_config import get_office_by_chat_id
savela_chat_id = -1002755506700
print(f'\n\nОфис для chat_id {savela_chat_id}: {get_office_by_chat_id(savela_chat_id)}')
