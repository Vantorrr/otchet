"""Backfill office column for existing data based on manager names."""
import gspread
import os


def backfill_office_column() -> None:
    """Fill office column for all existing records in Reports sheet."""
    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "/opt/otchet/service_account.json")
    sheet_name = os.getenv("SPREADSHEET_NAME", "Sales Reports")
    
    gc = gspread.service_account(filename=creds_path)
    spread = gc.open(sheet_name)
    reports = spread.worksheet("Reports")
    
    # Get all records
    all_records = reports.get_all_records()
    
    # Manager to office mapping (adjust based on your actual managers per office)
    # For now, assign all existing data to "Савела" (existing office)
    # You can manually specify manager-to-office mapping here
    manager_to_office = {
        # Савела managers (existing)
        "Бариев": "Савела",
        "Туробов": "Савела",
        "Романченко": "Савела",
        "Шевченко": "Савела",
        "Чертыковцев": "Савела",
        "Воробьев": "Савела",
        "test": "Савела",
        # Add other offices' managers as they come
    }
    
    print(f"📊 Найдено записей: {len(all_records)}")
    
    # Get column index for 'office'
    headers = reports.row_values(1)
    if 'office' not in headers:
        print("❌ Колонка 'office' не найдена в заголовках!")
        return
    
    office_col_idx = headers.index('office') + 1  # 1-based
    manager_col_idx = headers.index('manager') + 1 if 'manager' in headers else 2
    
    updates = []
    updated_count = 0
    
    for idx, record in enumerate(all_records, start=2):  # start=2 because row 1 is header
        manager = str(record.get('manager', '')).strip()
        current_office = str(record.get('office', '')).strip()
        
        # Skip if office already filled
        if current_office:
            continue
        
        # Determine office by manager
        office = manager_to_office.get(manager, "Савела")  # default to Савела for unknown
        
        # Prepare update (row, col, value)
        updates.append({
            'range': f'{chr(64 + office_col_idx)}{idx}',
            'values': [[office]]
        })
        updated_count += 1
        
        # Batch update every 100 rows
        if len(updates) >= 100:
            reports.batch_update(updates)
            print(f"✅ Обновлено {len(updates)} строк...")
            updates = []
    
    # Final batch
    if updates:
        reports.batch_update(updates)
        print(f"✅ Обновлено {len(updates)} строк...")
    
    print(f"\n🎉 Готово! Заполнено {updated_count} записей колонкой 'office'")


if __name__ == "__main__":
    backfill_office_column()

