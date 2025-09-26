from __future__ import annotations

import io
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.service_account import Credentials

from bot.config import Settings
from bot.services.yandex_gpt import YandexGPTService


SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]


@dataclass
class SlidesResources:
    drive: Any
    slides: Any


class GoogleSlidesService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        creds = Credentials.from_service_account_file(settings.google_credentials_path, scopes=SCOPES)
        self._resources = SlidesResources(
            drive=build("drive", "v3", credentials=creds),
            slides=build("slides", "v1", credentials=creds),
        )
        self._ai = YandexGPTService(settings)

    # --- Helpers ---
    def _get_folder_id(self) -> str:
        if not self._settings.drive_folder_id:
            raise RuntimeError("DRIVE_FOLDER_ID is not set in environment")
        return self._settings.drive_folder_id

    # --- Public API ---
    def create_presentation(self, title: str) -> str:
        body = {"title": title}
        pres = self._resources.slides.presentations().create(body=body).execute()
        return pres["presentationId"]

    def move_presentation_to_folder(self, presentation_id: str) -> None:
        folder_id = self._get_folder_id()
        # Get current parents
        file = self._resources.drive.files().get(fileId=presentation_id, fields="parents").execute()
        prev_parents = ",".join(file.get("parents", []))
        self._resources.drive.files().update(
            fileId=presentation_id,
            addParents=folder_id,
            removeParents=prev_parents or None,
            fields="id, parents",
        ).execute()

    def export_pdf(self, presentation_id: str) -> bytes:
        request = self._resources.drive.files().export_media(fileId=presentation_id, mimeType="application/pdf")
        buf = io.BytesIO()
        downloader = request
        # MediaIoBaseDownload is heavier; the export_media returns bytes via .execute()
        data = downloader.execute()
        buf.write(data)
        return buf.getvalue()

    # --- High-level deck builder (phase 1) ---
    async def build_title_and_summary(
        self,
        presentation_id: str,
        office_name: str,
        period_title: str,
        totals: Dict[str, float],
    ) -> None:
        # Title slide
        self.set_title_slide(presentation_id, f"{office_name} — Отчет по продажам", period_title)
        # Add summary slide: simple table via createSlide + text boxes (fast MVP)
        create = self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"}}}]}
        ).execute()
        # Get last page id (the slide we just created)
        pres = self._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        # Set title
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{
                "replaceAllText": {
                    "containsText": {"text": "Click to add title", "matchCase": False},
                    "replaceText": "Общие показатели команды"
                }
            }]}
        ).execute()
        # Create a simple 4x7 text grid using text boxes (left aligned for names, centered for numbers)
        x0, y0, row_h, col_w = 40, 120, 24, 150  # pt
        headers = ["Показатель", "План", "Факт", "Конв (%)"]
        metrics = [
            ("Повторные звонки", totals.get("calls_plan", 0), totals.get("calls_fact", 0), totals.get("calls_percentage", 0.0)),
            ("Заявки, шт", totals.get("leads_units_plan", 0), totals.get("leads_units_fact", 0), totals.get("leads_units_percentage", 0.0)),
            ("Заявки, млн", totals.get("leads_volume_plan", 0.0), totals.get("leads_volume_fact", 0.0), totals.get("leads_volume_percentage", 0.0)),
            ("Одобрено, млн", "-", totals.get("approved_volume", 0.0), "-"),
            ("Выдано, млн", "-", totals.get("issued_volume", 0.0), "-"),
            ("Новые звонки", "-", totals.get("new_calls", 0), "-")
        ]
        # Header row
        for c, text in enumerate(headers):
            self.add_textbox(presentation_id, page_id, f"hdr_{c}", text, x0 + c * col_w, y0, col_w, row_h, font_size=12)
        # Data rows
        for r, (name, plan, fact, conv) in enumerate(metrics, start=1):
            self.add_textbox(presentation_id, page_id, f"n_{r}", name, x0 + 0 * col_w, y0 + r * row_h, col_w, row_h, 11)
            self.add_textbox(presentation_id, page_id, f"p_{r}", f"{plan}", x0 + 1 * col_w, y0 + r * row_h, col_w, row_h, 11)
            self.add_textbox(presentation_id, page_id, f"f_{r}", f"{fact}", x0 + 2 * col_w, y0 + r * row_h, col_w, row_h, 11)
            self.add_textbox(presentation_id, page_id, f"c_{r}", f"{conv}", x0 + 3 * col_w, y0 + r * row_h, col_w, row_h, 11)

        # AI team comment under the table
        comment_title_id = "team_comment_title"
        comment_body_id = "team_comment_body"
        self.add_textbox(presentation_id, page_id, comment_title_id, "Комментарий ИИ — Команда", x0, y0 + (len(metrics)+1) * row_h + 20, col_w * 4, 22, 13)
        team_comment = await self._ai.generate_team_comment(totals, period_title)
        self.add_textbox(presentation_id, page_id, comment_body_id, team_comment, x0, y0 + (len(metrics)+1) * row_h + 44, col_w * 4, 100, 11)

    # --- Content helpers (basic) ---
    def set_title_slide(self, presentation_id: str, title: str, subtitle: str) -> None:
        # Create a title slide and set text
        requests: List[Dict[str, Any]] = []
        # Create slide
        requests.append({
            "createSlide": {
                "slideLayoutReference": {"predefinedLayout": "TITLE"}
            }
        })
        # After createSlide, we cannot know objectIds beforehand. We'll replace texts in placeholders.
        # Use replaceAllText on {{TITLE}}/{{SUBTITLE}} tokens; simpler: insert then replace defaults
        # Replace default 'Click to add title' and 'subtitle' with provided text
        requests.append({
            "replaceAllText": {
                "containsText": {"text": "Click to add title", "matchCase": False},
                "replaceText": title
            }
        })
        requests.append({
            "replaceAllText": {
                "containsText": {"text": "Click to add subtitle", "matchCase": False},
                "replaceText": subtitle
            }
        })
        self._resources.slides.presentations().batchUpdate(presentationId=presentation_id, body={"requests": requests}).execute()

    def add_textbox(self, presentation_id: str, page_id: str, object_id: str, text: str, x: int, y: int, w: int, h: int, font_size: int = 12) -> None:
        # Position in EMUs (1 pt ~ 12700 emu). Slides API uses magnitude + unit.
        requests: List[Dict[str, Any]] = [
            {"createShape": {
                "objectId": object_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": page_id,
                    "size": {"width": {"magnitude": w, "unit": "PT"}, "height": {"magnitude": h, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": x, "translateY": y, "unit": "PT"}
                }
            }},
            {"insertText": {"objectId": object_id, "text": text}},
            {"updateTextStyle": {
                "objectId": object_id,
                "fields": "fontSize",
                "style": {"fontSize": {"magnitude": font_size, "unit": "PT"}}
            }}
        ]
        self._resources.slides.presentations().batchUpdate(presentationId=presentation_id, body={"requests": requests}).execute()


