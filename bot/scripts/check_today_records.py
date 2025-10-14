"""Check today's records to find office issues."""
import os
from datetime import date
from dotenv import load_dotenv
import gspread

load_dotenv()


def check_today_records():
    """Check all records for today and their office assignments."""
    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "/opt/otchet/service_account.json")
    sheet_name = os.getenv("SPREADSHEET_NAME", "Sales Reports")
    
    gc = gspread.service_account(filename=creds_path)
    spread = gc.open(sheet_name)
    reports = spread.worksheet("Reports")
    
    # Get all records
    all_records = reports.get_all_records()
    
    # Today's date
    today = date.today().strftime("%Y-%m-%d")
    print(f"📅 Checking records for today: {today}\n")
    
    # Find today's records
    today_records = []
    for record in all_records:
        date_str = str(record.get('date', ''))
        if date_str.startswith(today) or date_str == today:
            today_records.append(record)
    
    print(f"📊 Found {len(today_records)} records for today\n")
    
    # Group by manager and check office
    from collections import defaultdict
    by_manager = defaultdict(list)
    
    for record in today_records:
        manager = record.get('manager', 'NO_MANAGER')
        by_manager[manager].append(record)
    
    # Show each manager's records
    for manager, records in sorted(by_manager.items()):
        print(f"👤 {manager}:")
        for r in records:
            office = r.get('office', '')
            print(f"  - Office: '{office}' (empty: {not office})")
            print(f"    Calls: {r.get('morning_calls_planned', 0)} planned, {r.get('evening_calls_success', 0)} done")
            print(f"    Row data: date={r.get('date')}, office='{office}'")
        print()


if __name__ == "__main__":
    check_today_records()

