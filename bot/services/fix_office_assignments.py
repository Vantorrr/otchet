#!/usr/bin/env python3
"""Fix office assignments in Google Sheets."""

import os
import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from bot.config import Settings
from bot.services.sheets import SheetsClient
from bot.services.di import Container

# Correct manager to office mapping
MANAGER_TO_OFFICE = {
    # Батурлов office
    "Абдрахманов": "Батурлов",
    "Тест": "Батурлов", 
    "test": "Батурлов",
    "Ягудаев": "Батурлов",
    "Василенко": "Батурлов",
    "Ефросинин": "Батурлов",
    
    # Савела office
    "Воробьев": "Савела",
    "Чертыковцев": "Савела",
    "Романченко": "Савела",
    "Бариев": "Савела",
    
    # Санжаровский office
    "Санжаровский": "Санжаровский",
    "Кобылянский": "Санжаровский",
    "Сафонов": "Санжаровский",
    "Туробов": "Санжаровский",
    "Шевченко": "Санжаровский",
    "Лазарев": "Санжаровский",
    
    # Офис 4
    "Соколов": "Офис 4",
    "Рыбалка": "Офис 4",
    "Корниенко": "Офис 4",
    "Николаев": "Офис 4",
}

def main():
    # Load settings
    settings = Settings.load()
    Container.init(settings)
    container = Container.get()
    
    # Get Reports sheet
    reports_sheet = container.sheets._reports
    
    # Get all records
    all_records = reports_sheet.get_all_records()
    
    print(f"Total records: {len(all_records)}")
    
    # Count assignments
    assignment_counts = {}
    fixed_count = 0
    
    # Check and fix each record
    for i, record in enumerate(all_records):
        row_num = i + 2  # +2 because sheets are 1-indexed and have header
        manager = record.get('manager', '')
        current_office = record.get('office', '')
        expected_office = MANAGER_TO_OFFICE.get(manager, '')
        
        if expected_office:
            if current_office != expected_office:
                # Update office
                office_col = 13  # Column M (office)
                reports_sheet.update_cell(row_num, office_col, expected_office)
                fixed_count += 1
                print(f"Fixed row {row_num}: {manager} from '{current_office}' to '{expected_office}'")
            
            # Count for statistics
            key = f"{manager}: {expected_office}"
            assignment_counts[key] = assignment_counts.get(key, 0) + 1
        else:
            print(f"Warning: Unknown manager '{manager}' in row {row_num}")
    
    print(f"\nFixed {fixed_count} records")
    print("\nFinal assignments:")
    for assignment, count in sorted(assignment_counts.items()):
        print(f"  {assignment}: {count} records")

if __name__ == "__main__":
    main()

