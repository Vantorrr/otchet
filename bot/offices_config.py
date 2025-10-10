"""Office configuration mapping chat_id to office name."""

OFFICE_MAPPING = {
    -1002511898620: "Офис 4",
    -1002963422273: "Санжаровский",
    -1003190087300: "Батурлов",
    -1002497067319: "Савела",  # Existing office (need chat_id confirmation)
    -1003164833460: "Руководительская",  # HQ
}

HQ_CHAT_ID = -1003164833460

def get_office_by_chat_id(chat_id: int) -> str:
    """Get office name by chat_id, default to 'Unknown'."""
    return OFFICE_MAPPING.get(chat_id, "Unknown")

def is_hq(chat_id: int) -> bool:
    """Check if chat_id is HQ."""
    return chat_id == HQ_CHAT_ID

def get_all_offices() -> list:
    """Get list of all office names excluding HQ."""
    return [name for cid, name in OFFICE_MAPPING.items() if cid != HQ_CHAT_ID]

