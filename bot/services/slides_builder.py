"""Premium Slides builder for banking guarantees reports."""
from __future__ import annotations

import os
from datetime import date
from typing import Any, Dict, List, Optional
from dataclasses import dataclass

from bot.config import Settings
from bot.services.google_slides import GoogleSlidesService
from bot.services.presentation import ManagerData
from bot.services.yandex_gpt import YandexGPTService
from bot.services.slides_shapes import SlidesShapeHelper


@dataclass
class SlideSpec:
    """Slide configuration for premium business presentation."""
    title: str
    layout: str = "BLANK"
    bg_color: Optional[str] = None


class PremiumSlidesBuilder:
    def __init__(self, settings: Settings, slides: GoogleSlidesService):
        self.settings = settings
        self.slides = slides
        self.ai = YandexGPTService(settings)
        self.shapes = SlidesShapeHelper(slides._resources.slides)
        # 16:9 slide dimensions (pt)
        self.page_w = 960
        self.page_h = 540
        # Grid: 12 columns, 1.5cm margins = 42.5pt
        self.margin = 42.5
        self.content_w = self.page_w - 2 * self.margin
        self.col_w = self.content_w / 12
        
    def _add_slide_bg(self, presentation_id: str, page_id: str) -> None:
        """Add premium gradient background."""
        # Full-page rectangle with gradient from white to light card color
        self.shapes.add_rectangle(
            presentation_id, page_id, f"bg_{page_id}",
            0, 0, self.page_w, self.page_h,
            fill_hex="#FFFFFF"
        )
        # Top brand band
        self.shapes.add_rectangle(
            presentation_id, page_id, f"band_{page_id}",
            0, 0, self.page_w, 50,
            fill_hex=self.settings.slides_primary_color
        )
        # Logo if available
        try:
            if self.settings.pptx_logo_path and os.path.exists(self.settings.pptx_logo_path):
                logo_id = self.slides.upload_logo_to_drive()
                if logo_id:
                    # Embed logo top-right
                    pass  # Logo embedding via createImage
        except Exception:
            pass

    async def build_title_slide(self, presentation_id: str, period_name: str, dates: str) -> None:
        """Premium title slide with branding."""
        self.slides._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "BLANK"}}}]}
        ).execute()
        pres = self.slides._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        
        self._add_slide_bg(presentation_id, page_id)
        
        # Main title: office name + "Отчет по продажам"
        self.slides.add_textbox(
            presentation_id, page_id, "title_main",
            f"{self.settings.office_name}\nОтчет по продажам",
            self.margin, 180, self.content_w, 120,
            font_size=48, align="CENTER", bold=True,
            text_color_hex=self.settings.slides_text_color
        )
        
        # Period subtitle
        self.slides.add_textbox(
            presentation_id, page_id, "title_period", f"{period_name}\n{dates}",
            self.margin, 320, self.content_w, 80,
            font_size=24, align="CENTER",
            text_color_hex=self.settings.slides_muted_color
        )

    async def build_team_summary_slide(self, presentation_id: str, totals: Dict[str, float], period_title: str) -> None:
        """Team metrics with AI analysis in card layout."""
        self.slides._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "BLANK"}}}]}
        ).execute()
        pres = self.slides._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        
        self._add_slide_bg(presentation_id, page_id)
        
        # Header
        self.slides.add_textbox(
            presentation_id, page_id, "sum_title", "Ключевые показатели команды",
            self.margin, 70, self.content_w, 40,
            font_size=28, align="CENTER", bold=True,
            text_color_hex=self.settings.slides_primary_color
        )
        
        # Metrics cards layout (2x3 grid)
        card_w = (self.content_w - 40) / 3
        card_h = 90
        y_row1 = 130
        y_row2 = 240
        
        def _fmt_ru(val, is_mln=False, is_pct=False):
            if is_pct:
                return f"{val:.1f}%".replace(".", ",")
            elif is_mln:
                return f"{val:.1f} млн".replace(".", ",")
            else:
                try:
                    return f"{int(val):,}".replace(",", " ")
                except:
                    return str(val)
        
        cards = [
            ("📞 Повторные звонки", totals.get('calls_fact', 0), False, False),
            ("📝 Заявки (шт)", totals.get('leads_units_fact', 0), False, False),
            ("💰 Заявки (объём)", totals.get('leads_volume_fact', 0), True, False),
            ("✅ Одобрено", totals.get('approved_volume', 0), True, False),
            ("✅ Выдано", totals.get('issued_volume', 0), True, False),
            ("☎️ Новые звонки", totals.get('new_calls', 0), False, False),
        ]
        
        for i, (label, value, is_mln, is_pct) in enumerate(cards):
            col = i % 3
            row = i // 3
            x = self.margin + col * (card_w + 20)
            y = y_row1 if row == 0 else y_row2
            
            # Card background
            self.shapes.add_rectangle(
                presentation_id, page_id, f"card_bg_{i}",
                x, y, card_w, card_h,
                fill_hex=self.settings.slides_card_bg_color
            )
            # Label
            self.slides.add_textbox(
                presentation_id, page_id, f"card_lbl_{i}", label,
                x + 10, y + 10, card_w - 20, 30,
                font_size=14, align="CENTER", bold=True,
                text_color_hex=self.settings.slides_text_color
            )
            # Value
            self.slides.add_textbox(
                presentation_id, page_id, f"card_val_{i}", _fmt_ru(value, is_mln, is_pct),
                x + 10, y + 45, card_w - 20, 35,
                font_size=20, align="CENTER", bold=True,
                text_color_hex=self.settings.slides_primary_color
            )
        
        # AI comment block
        ai_comment = await self.ai.generate_team_comment(totals, period_title)
        self.slides.add_textbox(
            presentation_id, page_id, "ai_comment_title", "Анализ команды",
            self.margin, 360, self.content_w, 25,
            font_size=16, align="START", bold=True,
            text_color_hex=self.settings.slides_text_color
        )
        self.slides.add_textbox(
            presentation_id, page_id, "ai_comment_body", ai_comment[:200] + ("..." if len(ai_comment) > 200 else ""),
            self.margin, 390, self.content_w, 100,
            font_size=14, align="START",
            text_color_hex=self.settings.slides_text_color,
            fill_hex=self.settings.slides_card_bg_color
        )

    async def build_top_ranking_slide(self, presentation_id: str, ranking: Dict[str, Any]) -> None:
        """Premium TOP/AntiTOP ranking cards."""
        self.slides._resources.slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "BLANK"}}}]}
        ).execute()
        pres = self.slides._resources.slides.presentations().get(presentationId=presentation_id).execute()
        page_id = pres["slides"][-1]["objectId"]
        
        self._add_slide_bg(presentation_id, page_id)
        
        # Header
        self.slides.add_textbox(
            presentation_id, page_id, "rank_title", "Рейтинг эффективности",
            self.margin, 70, self.content_w, 40,
            font_size=28, align="CENTER", bold=True,
            text_color_hex=self.settings.slides_primary_color
        )
        
        best = ranking.get("best", [])[:2]
        worst = ranking.get("worst", [])[:2]
        reasons = ranking.get("reasons", {})
        
        # Large cards layout
        card_w = (self.content_w - 60) / 2
        card_h = 140
        y_cards = 140
        
        # Best performers card
        self.shapes.add_rectangle(
            presentation_id, page_id, "best_card_bg",
            self.margin, y_cards, card_w, card_h,
            fill_hex=self.settings.slides_primary_color
        )
        self.slides.add_textbox(
            presentation_id, page_id, "best_header", "🏆 Лидеры периода",
            self.margin + 20, y_cards + 15, card_w - 40, 30,
            font_size=18, align="CENTER", bold=True,
            text_color_hex="#FFFFFF"
        )
        
        for i, name in enumerate(best[:2]):
            reason = reasons.get(name, "высокая результативность")
            self.slides.add_textbox(
                presentation_id, page_id, f"best_name_{i}", f"{i+1}. {name}",
                self.margin + 20, y_cards + 50 + i*35, card_w - 40, 25,
                font_size=16, align="START", bold=True,
                text_color_hex="#FFFFFF"
            )
            self.slides.add_textbox(
                presentation_id, page_id, f"best_reason_{i}", reason,
                self.margin + 20, y_cards + 70 + i*35, card_w - 40, 20,
                font_size=12, align="START",
                text_color_hex="#FFFFFF"
            )
        
        # Underperformers card
        x_right = self.margin + card_w + 40
        self.shapes.add_rectangle(
            presentation_id, page_id, "worst_card_bg",
            x_right, y_cards, card_w, card_h,
            fill_hex=self.settings.slides_alert_color
        )
        self.slides.add_textbox(
            presentation_id, page_id, "worst_header", "⚠️ Требуют внимания",
            x_right + 20, y_cards + 15, card_w - 40, 30,
            font_size=18, align="CENTER", bold=True,
            text_color_hex="#FFFFFF"
        )
        
        for i, name in enumerate(worst[:2]):
            reason = reasons.get(name, "снижение показателей")
            self.slides.add_textbox(
                presentation_id, page_id, f"worst_name_{i}", f"{i+1}. {name}",
                x_right + 20, y_cards + 50 + i*35, card_w - 40, 25,
                font_size=16, align="START", bold=True,
                text_color_hex="#FFFFFF"
            )
            self.slides.add_textbox(
                presentation_id, page_id, f"worst_reason_{i}", reason,
                x_right + 20, y_cards + 70 + i*35, card_w - 40, 20,
                font_size=12, align="START",
                text_color_hex="#FFFFFF"
            )

    def export_to_drive_pdf(self, presentation_id: str, period_name: str) -> str:
        """Export PDF to Drive with naming scheme."""
        # Move to folder first
        try:
            self.slides.move_presentation_to_folder(presentation_id)
        except Exception:
            pass
        
        # Export PDF
        pdf_bytes = self.slides.export_pdf(presentation_id)
        
        # Upload PDF to same folder with naming scheme
        import re
        week_match = re.search(r'(\d+)', period_name)
        week_num = week_match.group(1) if week_match else "XX"
        year = "2025"  # Could extract from dates
        filename = f"Отчет_{self.settings.office_name}_Неделя{week_num}_{year}.pdf"
        
        metadata = {
            "name": filename,
            "mimeType": "application/pdf",
            "parents": [self.settings.drive_folder_id] if self.settings.drive_folder_id else None
        }
        
        from googleapiclient.http import MediaIoBaseUpload
        import io
        media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf")
        file = self.slides._resources.drive.files().create(
            body=metadata, media_body=media, fields="id,webViewLink", supportsAllDrives=True
        ).execute()
        
        return file.get("webViewLink", "")
