"""Fix office assignments for all managers based on current knowledge."""
import gspread
import os
from collections import defaultdict


def fix_all_office_assignments() -> None:
    """Assign correct offices to all managers in Reports sheet."""
    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "/opt/otchet/service_account.json")
    sheet_name = os.getenv("SPREADSHEET_NAME", "Sales Reports")
    
    gc = gspread.service_account(filename=creds_path)
    spread = gc.open(sheet_name)
    reports = spread.worksheet("Reports")
    
    # Complete manager to office mapping
    manager_to_office = {
        # Савела
        "Бариев": "Савела",
        "Туробов": "Савела",
        "Романченко": "Савела",
        "Шевченко": "Савела",
        "Чертыковцев": "Савела",
        "Воробьев": "Савела",
        
        # Батурлов
        "Ефросинин": "Батурлов",
        "test": "Батурлов",
        "Тест": "Батурлов",
        
        # Офис 4
        "Абдрахманов": "Офис 4",
        
        # Санжаровский
        "Ягудаев": "Санжаровский",
        "Василенко": "Санжаровский",
        
        # Add more mappings as needed
    }
    
    # Get all records
    all_records = reports.get_all_records()
    print(f"📊 Найдено записей: {len(all_records)}")
    
    # Count managers without office assignment
    managers_without_office = defaultdict(int)
    
    # Get column index for 'office' and 'manager'
    headers = reports.row_values(1)
    if 'office' not in headers:
        print("❌ Колонка 'office' не найдена!")
        return
    
    office_col_idx = headers.index('office') + 1  # 1-based
    manager_col_idx = headers.index('manager') + 1 if 'manager' in headers else 2
    
    updates = []
    updated_count = 0
    
    for idx, record in enumerate(all_records, start=2):  # start=2 because row 1 is header
        manager = str(record.get('manager', '')).strip()
        current_office = str(record.get('office', '')).strip()
        
        if manager and not current_office:
            managers_without_office[manager] += 1
        
        # Get correct office for this manager
        correct_office = manager_to_office.get(manager, "")
        
        # Update if office is empty or incorrect
        if manager and correct_office and current_office != correct_office:
            updates.append({
                'range': f'{chr(64 + office_col_idx)}{idx}',
                'values': [[correct_office]]
            })
            updated_count += 1
            print(f"  ✏️ {manager}: '{current_office}' → '{correct_office}'")
        
        # Batch update every 100 rows
        if len(updates) >= 100:
            reports.batch_update(updates)
            print(f"✅ Обновлено {len(updates)} строк...")
            updates = []
    
    # Final batch
    if updates:
        reports.batch_update(updates)
        print(f"✅ Обновлено {len(updates)} строк...")
    
    print(f"\n🎉 Готово! Исправлено {updated_count} записей")
    
    # Report managers without office mapping
    unknown_managers = set()
    for idx, record in enumerate(all_records, start=2):
        manager = str(record.get('manager', '')).strip()
        if manager and manager not in manager_to_office:
            unknown_managers.add(manager)
    
    if unknown_managers:
        print("\n⚠️ Менеджеры без привязки к офису:")
        for manager in sorted(unknown_managers):
            count = managers_without_office.get(manager, 0)
            print(f"  - {manager} ({count} записей)")
    
    print("\n📊 Статистика по офисам:")
    office_stats = defaultdict(set)
    for manager, office in manager_to_office.items():
        office_stats[office].add(manager)
    
    for office, managers in sorted(office_stats.items()):
        print(f"\n🏢 {office}:")
        for manager in sorted(managers):
            print(f"  - {manager}")


if __name__ == "__main__":
    fix_all_office_assignments()
