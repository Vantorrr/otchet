"""Debug Savela office issue."""
from bot.offices_config import OFFICE_MAPPING, get_office_by_chat_id

# Check Savela chat ID
savela_chat_id = -1002755506700
office = get_office_by_chat_id(savela_chat_id)
print(f"Chat ID {savela_chat_id} -> Office: '{office}'")
print(f"Expected: 'Савела'")
print(f"Match: {office == 'Савела'}")

# Show all mappings
print("\nAll mappings:")
for cid, name in OFFICE_MAPPING.items():
    print(f"  {cid}: {name}")

