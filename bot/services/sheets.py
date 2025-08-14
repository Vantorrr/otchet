from __future__ import annotations

import gspread
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from bot.config import Settings


REPORTS_SHEET = "Reports"
BINDINGS_SHEET = "Bindings"
CONFIG_SHEET = "Config"

REPORT_HEADERS = [
    "date",
    "manager",
    "morning_calls_planned",
    "morning_leads_planned_units",
    "morning_leads_planned_volume",
    "morning_new_calls_planned",
    "evening_calls_success",
    "evening_leads_units",
    "evening_leads_volume",
    "evening_new_calls",
]

BINDINGS_HEADERS = ["topic_id", "manager"]
CONFIG_HEADERS = ["key", "value"]


@dataclass
class MorningData:
    calls_planned: int
    leads_units_planned: int
    leads_volume_planned: int
    new_calls_planned: int


@dataclass
class EveningData:
    calls_success: int
    leads_units: int
    leads_volume: int
    new_calls: int


class SheetsClient:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        # Use service account file explicitly for clarity
        self._gc = gspread.service_account(filename=settings.google_credentials_path)
        self._spread = self._open_or_create_spreadsheet(settings.spreadsheet_name)
        self._reports = self._get_or_create_worksheet(REPORTS_SHEET, REPORT_HEADERS)
        self._bindings = self._get_or_create_worksheet(BINDINGS_SHEET, BINDINGS_HEADERS)
        self._config = self._get_or_create_worksheet(CONFIG_SHEET, CONFIG_HEADERS)

    def _open_or_create_spreadsheet(self, name: str):
        try:
            return self._gc.open(name)
        except gspread.SpreadsheetNotFound:
            return self._gc.create(name)

    def _get_or_create_worksheet(self, title: str, headers: List[str]):
        try:
            ws = self._spread.worksheet(title)
            # Ensure headers exist and include newly added columns
            existing = ws.row_values(1)
            if not existing:
                ws.append_row(headers)
            else:
                missing = [h for h in headers if h not in existing]
                if missing:
                    new_headers = existing + missing
                    ws.update("1:1", [new_headers])
            return ws
        except gspread.WorksheetNotFound:
            ws = self._spread.add_worksheet(title=title, rows=1000, cols=max(10, len(headers)))
            ws.append_row(headers)
            return ws

    # Bindings
    def set_manager_binding(self, topic_id: int, manager: str) -> None:
        records = self._bindings.get_all_records()
        # Update if exists
        updated = False
        for idx, row in enumerate(records, start=2):
            if str(row.get("topic_id")) == str(topic_id):
                self._bindings.update_cell(idx, 2, manager)
                updated = True
                break
        if not updated:
            self._bindings.append_row([str(topic_id), manager])

    def get_manager_by_topic(self, topic_id: int) -> Optional[str]:
        records = self._bindings.get_all_records()
        for row in records:
            if str(row.get("topic_id")) == str(topic_id):
                return str(row.get("manager")) if row.get("manager") else None
        return None

    def set_summary_topic(self, topic_id: int) -> None:
        self._set_config("summary_topic_id", str(topic_id))

    def get_summary_topic_id(self) -> Optional[int]:
        value = self._get_config("summary_topic_id")
        return int(value) if value else None

    def set_group_chat_id(self, chat_id: int) -> None:
        self._set_config("group_chat_id", str(chat_id))

    def get_group_chat_id(self) -> Optional[int]:
        value = self._get_config("group_chat_id")
        return int(value) if value else None

    def _set_config(self, key: str, value: str) -> None:
        records = self._config.get_all_records()
        for idx, row in enumerate(records, start=2):
            if str(row.get("key")) == key:
                self._config.update_cell(idx, 2, value)
                return
        self._config.append_row([key, value])

    def _get_config(self, key: str) -> Optional[str]:
        records = self._config.get_all_records()
        for row in records:
            if str(row.get("key")) == key:
                return str(row.get("value")) if row.get("value") else None
        return None

    # Reports
    def upsert_report(self, date_str: str, manager: str, morning: MorningData | None = None, evening: EveningData | None = None) -> None:
        records = self._reports.get_all_records()
        row_index: Optional[int] = None
        for idx, row in enumerate(records, start=2):
            if str(row.get("date")) == date_str and str(row.get("manager")) == manager:
                row_index = idx
                break
        # Prepare existing values if any
        existing: Dict[str, Any] = {h: "" for h in REPORT_HEADERS}
        if row_index is not None:
            row_vals = self._reports.row_values(row_index)
            for i, h in enumerate(REPORT_HEADERS):
                if i < len(row_vals):
                    existing[h] = row_vals[i]
        existing["date"] = date_str
        existing["manager"] = manager
        if morning:
            existing["morning_calls_planned"] = str(morning.calls_planned)
            existing["morning_leads_planned_units"] = str(morning.leads_units_planned)
            existing["morning_leads_planned_volume"] = str(morning.leads_volume_planned)
            existing["morning_new_calls_planned"] = str(morning.new_calls_planned)
        if evening:
            existing["evening_calls_success"] = str(evening.calls_success)
            existing["evening_leads_units"] = str(evening.leads_units)
            existing["evening_leads_volume"] = str(evening.leads_volume)
            existing["evening_new_calls"] = str(evening.new_calls)
        row_values = [existing.get(h, "") for h in REPORT_HEADERS]
        if row_index is None:
            self._reports.append_row(row_values)
        else:
            # Update range to include potentially extra headers
            end_col = chr(ord('A') + max(len(REPORT_HEADERS), len(row_values)) - 1)
            self._reports.update(f"A{row_index}:{end_col}{row_index}", [row_values])

    def get_reports_by_date(self, date_str: str) -> List[Dict[str, Any]]:
        records = self._reports.get_all_records()
        return [r for r in records if str(r.get("date")) == date_str]
