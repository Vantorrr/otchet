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
    
    # Create office-specific sheets with QUERY formulas
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
            
            # Headers for office view
            headers = [
                "Дата", "Менеджер", "План перезвоны", "Факт перезвоны", "% перезвоны",
                "План новые", "Факт новые", "% новые",
                "Заявки шт", "Заявки млн", "Одобрено млн", "Выдано млн"
            ]
            sheet.update("A1:L1", [headers])
            
            # Format header row
            sheet.format("A1:L1", {
                "backgroundColor": {"red": 0.89, "green": 0.95, "blue": 1.0},
                "textFormat": {"bold": True, "fontSize": 11},
                "horizontalAlignment": "CENTER",
            })
            
            # Add QUERY formula to pull data from Reports sheet
            # Row 2: =QUERY(Reports!A:M, "SELECT A, B, C, G, G/C, F, L, L/F, H, I, J, K WHERE B = 'OFFICE_NAME' ORDER BY A DESC", 1)
            query_formula = f'=QUERY(Reports!A:M, "SELECT A, B, C, G, G/C, F, L, L/F, H, I, J, K WHERE C = \'{office}\' ORDER BY A DESC", 0)'
            sheet.update("A2", [[query_formula]])
            
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
            "Офис", "Период", "Менеджеров", 
            "План перезвоны", "Факт перезвоны", "% перезвоны",
            "План новые", "Факт новые", "% новые",
            "Заявки шт", "Заявки млн", "Одобрено млн", "Выдано млн"
        ]
        hq_sheet.update("A1:M1", [headers])
        hq_sheet.format("A1:M1", {
            "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.83},
            "textFormat": {"bold": True, "fontSize": 12},
            "horizontalAlignment": "CENTER",
        })
        
        print("✅ Настроен лист 'Сводная HQ'")
    
    print("\n🎉 Структура Google Sheets готова!")
    print("📌 Листы: Reports (сырые данные), Офис 4, Санжаровский, Батурлов, Сводная HQ")


if __name__ == "__main__":
    setup_office_sheets()

