"""Setup professional multi-office Google Sheets structure."""
import gspread
import os


def setup_office_sheets() -> None:
    """Create professional sheet structure with separate views per office."""
    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "/opt/otchet/service_account.json")
    sheet_name = os.getenv("SPREADSHEET_NAME", "Sales Reports")
    
    gc = gspread.service_account(filename=creds_path)
    spread = gc.open(sheet_name)
    
    offices = ["Офис 4", "Санжаровский", "Батурлов", "Савела"]
    
    # Create office-specific sheets with FILTER formulas
    for office in offices:
        try:
            sheet = spread.worksheet(office)
            print(f"⚠️ Лист '{office}' уже существует — пересоздаём для единообразия")
            spread.del_worksheet(sheet)
            sheet = spread.add_worksheet(title=office, rows=1000, cols=20)
            print(f"✅ Пересоздан лист '{office}'")
        except gspread.exceptions.WorksheetNotFound:
            sheet = spread.add_worksheet(title=office, rows=1000, cols=20)
            print(f"✅ Создан лист '{office}'")
        
        # Headers for office view (always set after creation/recreation)
        headers = [
            "Дата", "Менеджер", "План перезвоны", "Факт перезвоны",
            "План новые", "Факт новые",
            "Заявки шт", "Заявки млн", "Одобрено млн", "Выдано млн"
        ]
        sheet.update([headers], range_name="A1:J1")
        
        # Format header row
        sheet.format("A1:J1", {
            "backgroundColor": {"red": 0.89, "green": 0.95, "blue": 1.0},
            "textFormat": {"bold": True, "fontSize": 11},
            "horizontalAlignment": "CENTER",
        })
        
        # Add QUERY formula (uses ColN indexing; RU locale -> semicolons)
        # Select columns: date (Col1), manager (Col2), morning_calls_planned (Col4),
        # evening_calls_success (Col7), morning_new_calls_planned (Col5), evening_new_calls (Col12),
        # evening_leads_units (Col8), evening_leads_volume (Col9), approved_volume (Col10), issued_volume (Col11)
        # Filter by office (Col13)
        query_formula = (
            f'=QUERY(Reports!A:M; "select Col1, Col2, Col4, Col7, Col5, Col12, Col8, Col9, Col10, Col11 '
            f'where Col13 = \"{office}\" order by Col1 desc"; 1)'
        )
        sheet.update([[query_formula]], range_name="A2")
        
        print(f"✅ Настроен лист '{office}' с формулой QUERY")
    
    # Create HQ summary sheet
    try:
        hq_sheet = spread.worksheet("Сводная HQ")
        print("✅ Лист 'Сводная HQ' уже существует")
    except gspread.exceptions.WorksheetNotFound:
        hq_sheet = spread.add_worksheet(title="Сводная HQ", rows=1000, cols=30)
        print("✅ Создан лист 'Сводная HQ'")
        
        # HQ summary headers
        headers = [
            "Офис", "Менеджеров", 
            "План перезвоны", "Факт перезвоны",
            "План новые", "Факт новые",
            "Заявки шт", "Заявки млн", "Одобрено млн", "Выдано млн"
        ]
        hq_sheet.update([headers], range_name="A1:J1")
        hq_sheet.format("A1:J1", {
            "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.83},
            "textFormat": {"bold": True, "fontSize": 12},
            "horizontalAlignment": "CENTER",
        })
        
        print("✅ Настроен лист 'Сводная HQ'")
    
    print("\n🎉 Структура Google Sheets готова!")
    print("📌 Листы: Reports (сырые данные), Офис 4, Санжаровский, Батурлов, Сводная HQ")


if __name__ == "__main__":
    setup_office_sheets()

