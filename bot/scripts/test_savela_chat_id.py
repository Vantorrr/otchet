#!/usr/bin/env python3
"""Test Savela chat_id mapping"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from bot.offices_config import get_office_by_chat_id, OFFICE_MAPPING

# Test Savela chat_id
savela_chat_id = -1002755506700
office_name = get_office_by_chat_id(savela_chat_id)

print(f"Chat ID: {savela_chat_id}")
print(f"Office name: {office_name}")
print(f"\nAll office mappings:")
for chat_id, office in OFFICE_MAPPING.items():
    print(f"  {chat_id}: {office}")

