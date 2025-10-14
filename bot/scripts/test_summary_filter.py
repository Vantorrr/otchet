"""Test summary filtering logic."""
import os
from datetime import date
from dotenv import load_dotenv

load_dotenv()


def test_summary_filter():
    """Test the exact filtering logic from build_summary_text."""
    from bot.config import Settings
    from bot.services.di import Container
    from bot.services.summary_builder import build_summary_text
    
    settings = Settings.load()
    Container.init(settings)
    container = Container.get()
    
    # Test for today with Savela filter
    today = date.today().strftime("%Y-%m-%d")
    print(f"📅 Testing summary for: {today}")
    print(f"🏢 Office filter: Савела\n")
    
    # Get raw data
    all_records = container.sheets._reports.get_all_records()
    print(f"Total records in sheet: {len(all_records)}")
    
    # Test filtering manually
    savela_records = [r for r in all_records if r.get("office") == "Савела"]
    print(f"Records with office='Савела': {len(savela_records)}")
    
    # Today's Savela records
    today_savela = [r for r in savela_records if str(r.get("date", "")).startswith(today)]
    print(f"Today's Савела records: {len(today_savela)}")
    print("Managers:", [r.get('manager') for r in today_savela])
    
    # Now test the actual function
    print("\n🔍 Testing build_summary_text with office_filter='Савела':")
    summary = build_summary_text(settings, container.sheets, today, office_filter="Савела")
    
    # Count managers in summary
    import re
    managers = re.findall(r"👤 (\w+)", summary)
    print(f"\nManagers found in summary: {managers}")
    print(f"Total managers: {len(managers)}")
    
    # Check if wrong managers are included
    wrong_managers = [m for m in managers if m not in ['Воробьев', 'Чертыковцев', 'Романченко', 'Бариев', 'Туробов', 'Шевченко']]
    if wrong_managers:
        print(f"\n❌ WRONG MANAGERS IN SUMMARY: {wrong_managers}")
    else:
        print("\n✅ Only Савела managers in summary")


if __name__ == "__main__":
    test_summary_filter()

