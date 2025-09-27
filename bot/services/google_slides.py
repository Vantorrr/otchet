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
    "https://www.googleapis.com/auth/spreadsheets",
]


@dataclass
class SlidesResources:
    drive: Any
    slides: Any
    sheets: Any


class GoogleSlidesService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        creds = Credentials.from_service_account_file(settings.google_credentials_path, scopes=SCOPES)
        self._resources = SlidesResources(
            drive=build("drive", "v3", credentials=creds),
            slides=build("slides", "v1", credentials=creds),
            sheets=build("sheets", "v4", credentials=creds),
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

    # --- Branding helpers ---
    def _hex_to_rgb01(self, hex_color: str) -> Dict[str, float]:
        try:
            hex_color = hex_color.lstrip('#')
            r = int(hex_color[0:2], 16) / 255.0
            g = int(hex_color[2:4], 16) / 255.0
            b = int(hex_color[4:6], 16) / 255.0
            return {"red": r, "green": g, "blue": b}
        except Exception:
            return {"red": 0.8, "green": 0.0, "blue": 0.0}

    def upload_logo_to_drive(self) -> Optional[str]:
        path = getattr(self._settings, 'pptx_logo_path', '')
        if not path or not os.path.exists(path):
            return None
        metadata = {"name": "logo.png", "mimeType": "image/png"}
        folder_id = self._settings.drive_folder_id or None
        if folder_id:
            metadata["parents"] = [folder_id]
        with open(path, 'rb') as f:
            media = MediaIoBaseUpload(f, mimetype="image/png")
            file = self._resources.drive.files().create(body=metadata, media_body=media, fields="id").execute()
            return file.get("id")
        return None

    def apply_branding(self, presentation_id: str, logo_drive_id: Optional[str]) -> None:
        pres = self._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_w = int(pres.get('pageSize', {}).get('width', {}).get('magnitude', 960))
        requests: List[Dict[str, Any]] = []
        band_color = self._hex_to_rgb01(getattr(self._settings, 'slides_card_bg_color', '#F5F5F5'))
        for s in pres.get('slides', []):
            sid = s.get('objectId')
            # Top band
            band_id = f"band_{sid}"
            requests.append({
                "createShape": {
                    "objectId": band_id,
                    "shapeType": "RECTANGLE",
                    "elementProperties": {
                        "pageObjectId": sid,
                        "size": {"width": {"magnitude": page_w, "unit": "PT"}, "height": {"magnitude": 36, "unit": "PT"}},
                        "transform": {"scaleX": 1, "scaleY": 1, "translateX": 0, "translateY": 0, "unit": "PT"}
                    }
                }
            })
            requests.append({
                "updateShapeProperties": {
                    "objectId": band_id,
                    "fields": "shapeBackgroundFill.solidFill.color",
                    "shapeProperties": {
                        "shapeBackgroundFill": {
                            "solidFill": {"color": {"rgbColor": band_color}}
                        }
                    }
                }
            })
            # Logo
            if logo_drive_id:
                requests.append({
                    "createImage": {
                        "url": None,
                        "driveImageId": logo_drive_id,
                        "elementProperties": {
                            "pageObjectId": sid,
                            "size": {"width": {"magnitude": 110, "unit": "PT"}, "height": {"magnitude": 40, "unit": "PT"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": page_w - 130, "translateY": 0, "unit": "PT"}
                        }
                    }
                })
        if requests:
            self._resources.slides.presentations().batchUpdate(presentationId=presentation_id, body={"requests": requests}).execute()

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

    async def add_comparison_with_ai(
        self,
        presentation_id: str,
        prev_totals: Dict[str, float],
        cur_totals: Dict[str, float],
        title: str,
    ) -> None:
        # Create slide
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"}}}]}
        ).execute()
        pres = self._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        # Title
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{
                "replaceAllText": {
                    "containsText": {"text": "Click to add title", "matchCase": False},
                    "replaceText": title
                }
            }]}
        ).execute()

        # Two-column text boxes with totals (простая версия)
        x0, y0, row_h, col_w = 40, 120, 22, 220
        def write_col(prefix: str, totals: Dict[str, float], x: int):
            lines = [
                f"Повторные звонки: {int(totals.get('calls_fact',0))} из {int(totals.get('calls_plan',0))}",
                f"Заявки, шт: {int(totals.get('leads_units_fact',0))} из {int(totals.get('leads_units_plan',0))}",
                f"Заявки, млн: {totals.get('leads_volume_fact',0.0):.1f} из {totals.get('leads_volume_plan',0.0):.1f}",
                f"Одобрено, млн: {totals.get('approved_volume',0.0):.1f}",
                f"Выдано, млн: {totals.get('issued_volume',0.0):.1f}",
                f"Новые звонки: {int(totals.get('new_calls',0))} из {int(totals.get('new_calls_plan',0))}",
            ]
            for i, t in enumerate([prefix] + lines):
                oid = f"{prefix}_{i}"
                self.add_textbox(presentation_id, page_id, oid, t, x, y0 + i*row_h, col_w, row_h, 11 if i>0 else 13)

        write_col("Предыдущий", prev_totals, x0)
        write_col("Текущий", cur_totals, x0 + col_w + 40)

        # AI comparison comment
        ai_text = await self._ai.generate_comparison_comment(prev_totals, cur_totals, title)
        self.add_textbox(presentation_id, page_id, "cmp_ai", ai_text, x0, y0 + 9*row_h, col_w*2 + 40, 100, 11)

    async def add_top2_antitop2(self, presentation_id: str, ranking: Dict[str, Any]) -> None:
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "TITLE_AND_TWO_COLUMNS"}}}]}
        ).execute()
        pres = self._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        # Title
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{
                "replaceAllText": {
                    "containsText": {"text": "Click to add title", "matchCase": False},
                    "replaceText": "ТОП-2 и АнтиТОП-2"
                }
            }]}
        ).execute()
        best = ranking.get("best", [])[:2]
        worst = ranking.get("worst", [])[:2]
        reasons = ranking.get("reasons", {}) or {}
        x_left, x_right, y0, lh = 40, 360, 120, 22
        self.add_textbox(presentation_id, page_id, "best_hdr", "Лучшие:", x_left, y0, 260, lh, 13)
        for i, name in enumerate(best, start=1):
            self.add_textbox(presentation_id, page_id, f"best_{i}", f"🏆 {name}: {reasons.get(name,'отрыв по KPI')}", x_left, y0 + i*lh, 300, lh, 11)
        self.add_textbox(presentation_id, page_id, "worst_hdr", "Ниже темпа:", x_right, y0, 260, lh, 13)
        for i, name in enumerate(worst, start=1):
            self.add_textbox(presentation_id, page_id, f"worst_{i}", f"⚠️ {name}: {reasons.get(name,'просадка по KPI')}", x_right, y0 + i*lh, 300, lh, 11)

    # --- Sheets data and charts helpers ---
    def upsert_values_sheet(self, spreadsheet_id: str, sheet_title: str, headers: List[str], rows: List[List[Any]]) -> None:
        # Ensure sheet exists
        ss = self._resources.sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = None
        for s in ss.get("sheets", []):
            if s.get("properties", {}).get("title") == sheet_title:
                sheet_id = s.get("properties", {}).get("sheetId")
                break
        if sheet_id is None:
            # Add sheet
            self._resources.sheets.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": [{"addSheet": {"properties": {"title": sheet_title}}}]}
            ).execute()
        # Write values
        values = [headers] + rows
        self._resources.sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_title}!A1",
            valueInputOption="RAW",
            body={"values": values}
        ).execute()

    def ensure_basic_chart(self, spreadsheet_id: str, sheet_title: str, chart_title: str) -> int:
        # Find sheetId
        ss = self._resources.sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = None
        for s in ss.get("sheets", []):
            if s.get("properties", {}).get("title") == sheet_title:
                sheet_id = s.get("properties", {}).get("sheetId")
                break
        if sheet_id is None:
            raise RuntimeError("Data sheet not found for chart")
        # Create a simple column chart for range A1:D7
        add_chart_req = {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": chart_title,
                        "basicChart": {
                            "chartType": "COLUMN",
                            "legendPosition": "RIGHT_LEGEND",
                            "axis": [
                                {"position": "BOTTOM_AXIS", "title": "Показатель"},
                                {"position": "LEFT_AXIS", "title": "Значение"}
                            ],
                            "domains": [
                                {"domain": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 7, "startColumnIndex": 0, "endColumnIndex": 1}]}}}
                            ],
                            "series": [
                                {"series": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 7, "startColumnIndex": 1, "endColumnIndex": 2}]}}, "targetAxis": "LEFT_AXIS"},
                                {"series": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 7, "startColumnIndex": 2, "endColumnIndex": 3}]}}, "targetAxis": "LEFT_AXIS"}
                            ],
                        }
                    },
                    "position": {"newSheet": False, "overlayPosition": {"anchorCell": {"sheetId": sheet_id, "rowIndex": 9, "columnIndex": 0}, "widthPixels": 800, "heightPixels": 300}}
                }
            }
        }
        resp = self._resources.sheets.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [add_chart_req]}).execute()
        chart_id = resp.get("replies", [{}])[0].get("addChart", {}).get("chart", {}).get("chartId")
        if chart_id is None:
            # Fallback: try reading last charts
            ss = self._resources.sheets.spreadsheets().get(spreadsheetId=spreadsheet_id, includeGridData=False).execute()
            for s in ss.get("sheets", []):
                if s.get("charts"):
                    chart_id = s["charts"][-1]["chartId"]
        if chart_id is None:
            raise RuntimeError("Failed to create chart")
        return int(chart_id)

    def embed_sheets_chart(self, presentation_id: str, page_id: str, spreadsheet_id: str, chart_id: int, x: int, y: int, w: int, h: int) -> None:
        req = {
            "createSheetsChart": {
                "spreadsheetId": spreadsheet_id,
                "chartId": chart_id,
                "linkingMode": "LINKED",
                "elementProperties": {
                    "pageObjectId": page_id,
                    "size": {"width": {"magnitude": w, "unit": "PT"}, "height": {"magnitude": h, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": x, "translateY": y, "unit": "PT"}
                }
            }
        }
        self._resources.slides.presentations().batchUpdate(presentationId=presentation_id, body={"requests": [req]}).execute()

    # High-level: add line chart Plan→Issued and per-manager columns
    def add_charts_from_series(
        self,
        presentation_id: str,
        spreadsheet_id: str,
        series_sheet: str,
        daily_rows: List[List[Any]],
        managers_sheet: str,
        managers_rows: List[List[Any]],
    ) -> None:
        # Daily series write
        self.upsert_values_sheet(spreadsheet_id, series_sheet, ["Дата", "План, млн", "Факт, млн", "Выдано, млн"], daily_rows)
        # Create chart on series_sheet (columns B-D vs A)
        chart_id = self.ensure_basic_chart(spreadsheet_id, series_sheet, "План → Выдано (млн)")
        pres = self._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        self.embed_sheets_chart(presentation_id, page_id, spreadsheet_id, chart_id, 40, 580, 800, 260)

        # Managers columns (шт или млн)
        self.upsert_values_sheet(spreadsheet_id, managers_sheet, ["Менеджер", "Заявки, шт", "Звонки, повторные"], managers_rows)
        chart_id2 = self.ensure_basic_chart(spreadsheet_id, managers_sheet, "По менеджерам")
        # New slide for managers chart
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "BLANK"}}}]}
        ).execute()
        pres = self._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page2 = pres["slides"][-1]["objectId"]
        self.embed_sheets_chart(presentation_id, page2, spreadsheet_id, chart_id2, 40, 120, 800, 420)

    # Radar (Spider) chart: manager vs department average
    def ensure_radar_chart(self, spreadsheet_id: str, sheet_title: str, chart_title: str, row_count: int) -> int:
        ss = self._resources.sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = None
        for s in ss.get("sheets", []):
            if s.get("properties", {}).get("title") == sheet_title:
                sheet_id = s.get("properties", {}).get("sheetId")
                break
        if sheet_id is None:
            raise RuntimeError("Radar data sheet not found")
        add_chart_req = {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": chart_title,
                        "basicChart": {
                            "chartType": "RADAR",
                            "legendPosition": "RIGHT_LEGEND",
                            "domains": [
                                {"domain": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 0, "endColumnIndex": 1}]}}}
                            ],
                            "series": [
                                {"series": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 1, "endColumnIndex": 2}]}}},
                                {"series": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 2, "endColumnIndex": 3}]}}}
                            ]
                        }
                    },
                    "position": {"newSheet": False, "overlayPosition": {"anchorCell": {"sheetId": sheet_id, "rowIndex": 12, "columnIndex": 0}, "widthPixels": 800, "heightPixels": 400}}
                }
            }
        }
        resp = self._resources.sheets.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [add_chart_req]}).execute()
        chart_id = resp.get("replies", [{}])[0].get("addChart", {}).get("chart", {}).get("chartId")
        if chart_id is None:
            raise RuntimeError("Failed to create radar chart")
        return int(chart_id)

    def add_radar_slide(self, presentation_id: str, spreadsheet_id: str, sheet_title: str, rows: List[List[Any]], manager_name: str) -> None:
        self.upsert_values_sheet(spreadsheet_id, sheet_title, ["Метрика", "Среднее отдела", manager_name], rows)
        chart_id = self.ensure_radar_chart(spreadsheet_id, sheet_title, f"{manager_name} vs отдел", len(rows))
        # New slide
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"}}}]}
        ).execute()
        pres = self._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        # Replace title
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{
                "replaceAllText": {
                    "containsText": {"text": "Click to add title", "matchCase": False},
                    "replaceText": f"Сравнение — {manager_name}"
                }
            }]}
        ).execute()
        self.embed_sheets_chart(presentation_id, page_id, spreadsheet_id, chart_id, 40, 140, 800, 400)

    def add_gap_table(self, presentation_id: str, rows: List[List[Any]]) -> None:
        # New slide with table constructed from text boxes (compact)
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"}}}]}
        ).execute()
        pres = self._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        # Title
        self._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{
                "replaceAllText": {
                    "containsText": {"text": "Click to add title", "matchCase": False},
                    "replaceText": "GAP: отставание от плана (млн)"
                }
            }]}
        ).execute()
        # Grid
        x0, y0, row_h, col_w = 40, 140, 22, 180
        headers = ["Менеджер", "План, млн", "Выдано, млн", "GAP, млн"]
        for c, text in enumerate(headers):
            self.add_textbox(presentation_id, page_id, f"gap_h_{c}", text, x0 + c*col_w, y0, col_w, row_h, 12)
        for r, (name, plan, issued, gap) in enumerate(rows, start=1):
            self.add_textbox(presentation_id, page_id, f"gap_n_{r}", str(name), x0 + 0*col_w, y0 + r*row_h, col_w, row_h, 11)
            self.add_textbox(presentation_id, page_id, f"gap_p_{r}", f"{plan:.1f}", x0 + 1*col_w, y0 + r*row_h, col_w, row_h, 11)
            self.add_textbox(presentation_id, page_id, f"gap_i_{r}", f"{issued:.1f}", x0 + 2*col_w, y0 + r*row_h, col_w, row_h, 11)
            self.add_textbox(presentation_id, page_id, f"gap_g_{r}", f"{gap:.1f}", x0 + 3*col_w, y0 + r*row_h, col_w, row_h, 11)

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
        # Apply brand colors and font via text replacements (simple)
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


