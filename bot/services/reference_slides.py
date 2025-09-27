"""Reference-style Slides builder matching business proposal template."""
from __future__ import annotations

import os
from datetime import date
from typing import Any, Dict, List, Optional

from bot.config import Settings
from bot.services.google_slides import GoogleSlidesService
from bot.services.presentation import ManagerData
from bot.services.yandex_gpt import YandexGPTService


class ReferenceSlidesBuilder:
    """Premium business presentation builder matching provided references."""
    
    def __init__(self, settings: Settings, slides: GoogleSlidesService):
        self.settings = settings
        self.slides = slides
        self.ai = YandexGPTService(settings)
        
        # Reference palette (emerald corporate style)
        self.primary = "#2E7D32"    # Deep emerald
        self.accent = "#4CAF50"     # Light emerald
        self.cream = "#F5F5DC"      # Cream background
        self.dark = "#1B5E20"       # Dark emerald
        self.white = "#FFFFFF"
        self.text_dark = "#2E2E2E"
        self.text_light = "#757575"
        
        # Layout constants (16:9, 960x540 pt)
        self.w = 960
        self.h = 540
        self.margin = 60  # 1.5cm equivalent

    def create_reference_deck(
        self,
        presentation_id: str,
        office_name: str,
        period_name: str,
        dates: str,
        totals: Dict[str, float],
        ranking: Dict[str, Any]
    ) -> None:
        """Build complete reference-style deck."""
        # 1. Title slide - emerald with geometric shapes
        self._build_title_slide(presentation_id, office_name, period_name, dates)
        
        # 2. Key metrics dashboard with cards and donut chart
        self._build_metrics_dashboard(presentation_id, totals)
        
        # 3. Performance ranking with large visual cards
        self._build_ranking_slide(presentation_id, ranking)

    def _build_title_slide(self, presentation_id: str, office_name: str, period: str, dates: str) -> None:
        """Create title slide with emerald branding and geometric elements."""
        # Create blank slide
        self.slides._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "BLANK"}}}]}
        ).execute()
        
        pres = self.slides._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        
        # Full emerald background
        self._add_shape(presentation_id, page_id, "bg_main", "RECTANGLE", 0, 0, self.w, self.h, self.primary)
        
        # Geometric accent shapes (circles and triangles like in reference)
        self._add_shape(presentation_id, page_id, "circle1", "ELLIPSE", 50, 50, 80, 80, self.accent)
        self._add_shape(presentation_id, page_id, "circle2", "ELLIPSE", 820, 400, 100, 100, self.white)
        
        # Main title
        self._add_text(
            presentation_id, page_id, "main_title", 
            f"{office_name.upper()}\n\nОТЧЕТ ПО ПРОДАЖАМ",
            200, 150, 560, 120, 
            font_size=42, color=self.white, bold=True, align="CENTER"
        )
        
        # Period subtitle
        self._add_text(
            presentation_id, page_id, "period_text",
            f"{period}\n{dates}",
            200, 300, 560, 80,
            font_size=24, color=self.cream, align="CENTER"
        )
        
        # Website/contact in bottom right
        self._add_text(
            presentation_id, page_id, "contact_info",
            "reports@company.com",
            700, 480, 200, 30,
            font_size=14, color=self.cream, align="END"
        )

    def _build_metrics_dashboard(self, presentation_id: str, totals: Dict[str, float]) -> None:
        """Dashboard with key metrics cards and donut chart."""
        self.slides._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "BLANK"}}}]}
        ).execute()
        
        pres = self.slides._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        
        # Light background
        self._add_shape(presentation_id, page_id, "bg_light", "RECTANGLE", 0, 0, self.w, self.h, self.cream)
        
        # Header with emerald accent
        self._add_shape(presentation_id, page_id, "header_bar", "RECTANGLE", 0, 0, self.w, 80, self.primary)
        self._add_text(
            presentation_id, page_id, "dashboard_title",
            "КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ",
            self.margin, 25, self.w - 2*self.margin, 30,
            font_size=24, color=self.white, bold=True, align="CENTER"
        )
        
        # Metrics cards grid (3x2)
        card_w = 280
        card_h = 120
        start_x = self.margin
        start_y = 100
        gap = 20
        
        metrics = [
            ("ПОВТОРНЫЕ\nЗВОНКИ", f"{int(totals.get('calls_fact', 0)):,}".replace(",", " "), f"{totals.get('calls_percentage', 0):.1f}%"),
            ("ЗАЯВКИ\n(ШТ)", f"{int(totals.get('leads_units_fact', 0)):,}".replace(",", " "), f"{totals.get('leads_units_percentage', 0):.1f}%"),
            ("ЗАЯВКИ\n(МЛН)", f"{totals.get('leads_volume_fact', 0):.1f}".replace(".", ","), f"{totals.get('leads_volume_percentage', 0):.1f}%"),
            ("ОДОБРЕНО\n(МЛН)", f"{totals.get('approved_volume', 0):.1f}".replace(".", ","), "—"),
            ("ВЫДАНО\n(МЛН)", f"{totals.get('issued_volume', 0):.1f}".replace(".", ","), "—"),
            ("НОВЫЕ\nЗВОНКИ", f"{int(totals.get('new_calls', 0)):,}".replace(",", " "), "—"),
        ]
        
        for i, (label, value, pct) in enumerate(metrics):
            col = i % 3
            row = i // 3
            x = start_x + col * (card_w + gap)
            y = start_y + row * (card_h + gap)
            
            # Card background (white with subtle shadow effect)
            self._add_shape(presentation_id, page_id, f"card_{i}", "RECTANGLE", x, y, card_w, card_h, self.white)
            
            # Accent bar on top
            self._add_shape(presentation_id, page_id, f"accent_{i}", "RECTANGLE", x, y, card_w, 8, self.accent)
            
            # Metric label
            self._add_text(
                presentation_id, page_id, f"label_{i}", label,
                x + 20, y + 20, card_w - 40, 30,
                font_size=14, color=self.text_light, bold=True, align="START"
            )
            
            # Value (large)
            self._add_text(
                presentation_id, page_id, f"value_{i}", value,
                x + 20, y + 55, card_w - 40, 35,
                font_size=28, color=self.primary, bold=True, align="START"
            )
            
            # Percentage (if available)
            if pct != "—":
                self._add_text(
                    presentation_id, page_id, f"pct_{i}", pct,
                    x + 20, y + 90, card_w - 40, 20,
                    font_size=16, color=self.accent, bold=True, align="END"
                )

    def _build_ranking_slide(self, presentation_id: str, ranking: Dict[str, Any]) -> None:
        """Performance ranking with large visual impact."""
        self.slides._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "BLANK"}}}]}
        ).execute()
        
        pres = self.slides._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        
        # Gradient background (light)
        self._add_shape(presentation_id, page_id, "bg_rank", "RECTANGLE", 0, 0, self.w, self.h, self.cream)
        
        # Header
        self._add_shape(presentation_id, page_id, "rank_header_bg", "RECTANGLE", 0, 0, self.w, 80, self.primary)
        self._add_text(
            presentation_id, page_id, "rank_title",
            "РЕЙТИНГ ЭФФЕКТИВНОСТИ",
            self.margin, 25, self.w - 2*self.margin, 30,
            font_size=24, color=self.white, bold=True, align="CENTER"
        )
        
        best = ranking.get("best", [])[:2]
        worst = ranking.get("worst", [])[:2]
        
        # Large performance cards
        card_w = 380
        card_h = 200
        y_pos = 120
        gap = 40
        
        # Best performers (emerald card)
        x_best = self.margin
        self._add_shape(presentation_id, page_id, "best_bg", "RECTANGLE", x_best, y_pos, card_w, card_h, self.primary)
        
        # Crown icon area (geometric shape)
        self._add_shape(presentation_id, page_id, "crown_bg", "ELLIPSE", x_best + 20, y_pos + 20, 60, 60, self.accent)
        self._add_text(
            presentation_id, page_id, "crown_text", "🏆",
            x_best + 35, y_pos + 35, 30, 30,
            font_size=24, color=self.white, align="CENTER"
        )
        
        self._add_text(
            presentation_id, page_id, "best_title",
            "ЛИДЕРЫ ПЕРИОДА",
            x_best + 100, y_pos + 30, card_w - 120, 30,
            font_size=18, color=self.white, bold=True, align="START"
        )
        
        for i, name in enumerate(best[:2]):
            self._add_text(
                presentation_id, page_id, f"best_name_{i}",
                f"{i+1}. {name}",
                x_best + 20, y_pos + 80 + i*35, card_w - 40, 25,
                font_size=16, color=self.white, bold=True, align="START"
            )
            self._add_text(
                presentation_id, page_id, f"best_desc_{i}",
                "высокая результативность",
                x_best + 20, y_pos + 100 + i*35, card_w - 40, 20,
                font_size=12, color=self.cream, align="START"
            )
        
        # Underperformers (red-orange card)
        x_worst = x_best + card_w + gap
        alert_color = "#D32F2F"
        self._add_shape(presentation_id, page_id, "worst_bg", "RECTANGLE", x_worst, y_pos, card_w, card_h, alert_color)
        
        # Warning icon
        self._add_shape(presentation_id, page_id, "warn_bg", "ELLIPSE", x_worst + 20, y_pos + 20, 60, 60, "#FF5722")
        self._add_text(
            presentation_id, page_id, "warn_text", "⚠️",
            x_worst + 35, y_pos + 35, 30, 30,
            font_size=24, color=self.white, align="CENTER"
        )
        
        self._add_text(
            presentation_id, page_id, "worst_title",
            "ТРЕБУЮТ ВНИМАНИЯ",
            x_worst + 100, y_pos + 30, card_w - 120, 30,
            font_size=18, color=self.white, bold=True, align="START"
        )
        
        for i, name in enumerate(worst[:2]):
            self._add_text(
                presentation_id, page_id, f"worst_name_{i}",
                f"{i+1}. {name}",
                x_worst + 20, y_pos + 80 + i*35, card_w - 40, 25,
                font_size=16, color=self.white, bold=True, align="START"
            )
            self._add_text(
                presentation_id, page_id, f"worst_desc_{i}",
                "снижение показателей",
                x_worst + 20, y_pos + 100 + i*35, card_w - 40, 20,
                font_size=12, color="#FFE0B2", align="START"
            )

    def _add_shape(self, presentation_id: str, page_id: str, oid: str, shape_type: str, x: int, y: int, w: int, h: int, fill: str) -> None:
        """Add colored shape."""
        requests = [{
            "createShape": {
                "objectId": self._safe_id(oid),
                "shapeType": shape_type,
                "elementProperties": {
                    "pageObjectId": page_id,
                    "size": {"width": {"magnitude": w, "unit": "PT"}, "height": {"magnitude": h, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": x, "translateY": y, "unit": "PT"}
                }
            }
        }, {
            "updateShapeProperties": {
                "objectId": self._safe_id(oid),
                "fields": "shapeBackgroundFill.solidFill.color",
                "shapeProperties": {
                    "shapeBackgroundFill": {"solidFill": {"color": {"rgbColor": self._hex_to_rgb01(fill)}}}
                }
            }
        }]
        self.slides._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": requests}
        ).execute()

    def _add_text(self, presentation_id: str, page_id: str, oid: str, text: str, x: int, y: int, w: int, h: int, 
                  font_size: int = 14, color: str = "#000000", bold: bool = False, align: str = "START") -> None:
        """Add text element."""
        align_map = {"LEFT": "START", "CENTER": "CENTER", "RIGHT": "END", "START": "START", "END": "END"}
        norm_align = align_map.get(align, "START")
        
        requests = [{
            "createShape": {
                "objectId": self._safe_id(oid),
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": page_id,
                    "size": {"width": {"magnitude": w, "unit": "PT"}, "height": {"magnitude": h, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": x, "translateY": y, "unit": "PT"}
                }
            }
        }, {
            "insertText": {"objectId": self._safe_id(oid), "text": text}
        }, {
            "updateTextStyle": {
                "objectId": self._safe_id(oid),
                "fields": "fontSize,fontFamily,bold,foregroundColor",
                "style": {
                    "fontSize": {"magnitude": font_size, "unit": "PT"},
                    "fontFamily": "Roboto",
                    "bold": bold,
                    "foregroundColor": {"opaqueColor": {"rgbColor": self._hex_to_rgb01(color)}}
                }
            }
        }, {
            "updateParagraphStyle": {
                "objectId": self._safe_id(oid),
                "fields": "alignment",
                "style": {"alignment": norm_align}
            }
        }]
        self.slides._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": requests}
        ).execute()

    def _safe_id(self, oid: str) -> str:
        """Ensure valid Slides object ID."""
        import re
        import hashlib
        ascii_oid = re.sub(r"[^A-Za-z0-9_]", "_", oid)
        if len(ascii_oid) < 5:
            ascii_oid = f"obj_{hashlib.md5(oid.encode()).hexdigest()[:8]}"
        if not re.match(r"^[A-Za-z]", ascii_oid):
            ascii_oid = "obj_" + ascii_oid
        return ascii_oid

    def _hex_to_rgb01(self, hex_color: str) -> Dict[str, float]:
        """Convert hex to RGB 0-1 values."""
        try:
            hex_color = hex_color.lstrip('#')
            return {
                "red": int(hex_color[0:2], 16) / 255.0,
                "green": int(hex_color[2:4], 16) / 255.0,
                "blue": int(hex_color[4:6], 16) / 255.0
            }
        except Exception:
            return {"red": 0.18, "green": 0.49, "blue": 0.20}  # Default emerald
