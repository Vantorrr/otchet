"""Test office filtering in summaries."""
import asyncio
import os
from datetime import datetime, date
from dotenv import load_dotenv

load_dotenv()


async def test_office_filtering():
    """Test how office filtering works."""
    # Initialize
    from bot.config import Settings
    from bot.services.di import Container
    
    settings = Settings.load()
    Container.init(settings)
    container = Container.get()
    
    # Test office mapping
    from bot.offices_config import get_office_by_chat_id, OFFICE_MAPPING
    
    print("📍 Office mappings:")
    for chat_id, office in OFFICE_MAPPING.items():
        print(f"  {chat_id}: {office}")
    
    # Test Savela chat ID
    savela_chat_id = -1002755506700
    office = get_office_by_chat_id(savela_chat_id)
    print(f"\n🏢 Office for Savela chat ({savela_chat_id}): {office}")
    
    # Get all records and check offices
    all_records = container.sheets._reports.get_all_records()
    print(f"\n📊 Total records: {len(all_records)}")
    
    # Count by office
    from collections import defaultdict
    office_counts = defaultdict(int)
    managers_by_office = defaultdict(set)
    
    for record in all_records:
        office = record.get('office', 'NO_OFFICE')
        manager = record.get('manager', 'NO_MANAGER')
        office_counts[office] += 1
        managers_by_office[office].add(manager)
    
    print("\n📈 Records by office:")
    for office, count in sorted(office_counts.items()):
        print(f"  {office}: {count} records")
        managers = sorted(managers_by_office[office])
        print(f"    Managers: {', '.join(managers[:5])}" + (" ..." if len(managers) > 5 else ""))
    
    # Check for empty office records
    empty_office_records = [r for r in all_records if not r.get('office') or r.get('office') == '']
    print(f"\n⚠️ Records without office: {len(empty_office_records)}")
    if empty_office_records:
        managers_no_office = set(r.get('manager', 'NO_MANAGER') for r in empty_office_records)
        print(f"  Managers without office: {', '.join(sorted(managers_no_office))}")
    
    # Test filtering for today
    today = date.today().strftime("%Y-%m-%d")
    print(f"\n📅 Testing filter for today ({today}) and office 'Савела':")
    
    # Filter by office
    savela_records = [r for r in all_records if r.get("office") == "Савела"]
    print(f"  Records with office='Савела': {len(savela_records)}")
    
    # Filter by date and office
    today_savela = [r for r in savela_records if str(r.get("date", "")).startswith(today)]
    print(f"  Records for today in Савела: {len(today_savela)}")
    
    if today_savela:
        print("  Managers:")
        for r in today_savela[:10]:
            print(f"    - {r.get('manager')}: office={r.get('office')}")


if __name__ == "__main__":
    asyncio.run(test_office_filtering())
