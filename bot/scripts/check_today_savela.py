#!/usr/bin/env python3
"""Check today's Savela data directly via gspread."""

import os
import gspread
from google.oauth2.service_account import Credentials

# Get credentials
creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")
creds = Credentials.from_service_account_file(creds_path, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
gc = gspread.authorize(creds)

# Open spreadsheet
spreadsheet_name = os.getenv("SPREADSHEET_NAME", "Sales Reports")
sheet = gc.open(spreadsheet_name)
reports_sheet = sheet.worksheet("Reports")

# Get all records
all_records = reports_sheet.get_all_records()

# Filter today's records
today_records = [r for r in all_records if r.get('date') == '2025-10-13']

# Group by office
office_data = {}
for r in today_records:
    office = r.get('office', 'NO_OFFICE')
    manager = r.get('manager', 'NO_MANAGER')
    if office not in office_data:
        office_data[office] = []
    office_data[office].append(manager)

print("=" * 60)
print("ДАННЫЕ ЗА 2025-10-13:")
print("=" * 60)

for office, managers in sorted(office_data.items()):
    print(f"\nОфис: {office}")
    print(f"Количество записей: {len(managers)}")
    print(f"Менеджеры: {sorted(set(managers))}")

# Specific check for Savela
print("\n" + "=" * 60)
print("КОНКРЕТНО ОФИС САВЕЛА:")
print("=" * 60)

savela_records = [r for r in today_records if r.get('office') == 'Савела']
print(f"\nНайдено записей с офисом 'Савела': {len(savela_records)}")

if savela_records:
    managers_in_savela = sorted(set(r.get('manager', 'NO_MANAGER') for r in savela_records))
    print(f"Менеджеры в Савеле: {managers_in_savela}")
    
    # Show first few records
    print("\nПримеры записей:")
    for i, r in enumerate(savela_records[:5]):
        print(f"  {i+1}. Менеджер: {r.get('manager')}, Офис: {r.get('office')}, Дата: {r.get('date')}")
else:
    print("Записей с офисом 'Савела' не найдено!")

# Check if managers are assigned to wrong offices
print("\n" + "=" * 60)
print("ПРОВЕРКА НЕПРАВИЛЬНЫХ НАЗНАЧЕНИЙ:")
print("=" * 60)

# Expected mapping
expected_mapping = {
    "Абдрахманов": "Батурлов",
    "Тест": "Батурлов",
    "test": "Батурлов",
    "Ягудаев": "Батурлов",
    "Василенко": "Батурлов",
    "Ефросинин": "Батурлов",
    "Воробьев": "Савела",
    "Чертыковцев": "Савела",
    "Романченко": "Савела",
    "Бариев": "Савела"
}

wrong_assignments = []
for r in today_records:
    manager = r.get('manager', '')
    actual_office = r.get('office', '')
    expected_office = expected_mapping.get(manager, '')
    
    if expected_office and actual_office != expected_office:
        wrong_assignments.append({
            'manager': manager,
            'actual': actual_office,
            'expected': expected_office
        })

if wrong_assignments:
    print("\nНайдены неправильные назначения офисов:")
    for wa in wrong_assignments:
        print(f"  - {wa['manager']}: сейчас в '{wa['actual']}', должен быть в '{wa['expected']}'")
else:
    print("\nВсе менеджеры назначены в правильные офисы!")

