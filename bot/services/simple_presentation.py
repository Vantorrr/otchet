"""
Simplified PPTX generator matching client's reference exactly.
Structure:
- Title slide
- Manager statistics table (6 metrics x 5 columns)
- AI commentary (compact numbered list)
"""
from __future__ import annotations
import io
from typing import Dict
from datetime import date
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Cm

from bot.config import Settings
from bot.services.yandex_gpt import YandexGPTService


def hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert hex color to RGBColor."""
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


PRIMARY = "#1565C0"
ACCENT2 = "#42A5F5"
TEXT_MAIN = "#212121"
TEXT_MUTED = "#757575"
CARD_BG = "#F5F5F5"


class SimplePresentationService:
    """Generate simple PPTX presentation matching client's reference."""
    
    def __init__(self, settings: Settings):
        self.settings = settings
        self.ai = YandexGPTService(settings)
    
    async def generate_presentation(
        self,
        period_data: Dict,
        period_name: str,
        start_date: date,
        end_date: date,
        prev_data: Dict,
        prev_start: date,
        prev_end: date,
    ) -> bytes:
        """Generate presentation with title + one slide per manager."""
        prs = Presentation()
        prs.slide_width = Inches(11.69)  # 16:9
        prs.slide_height = Inches(6.58)
        
        margin = Inches(0.6)
        
        # Slide 1: Title
        await self._add_title_slide(prs, period_name, start_date, end_date, margin)
        
        # Calculate team averages for comparison
        total_managers = len(period_data)
        avg = self._calculate_averages(period_data, total_managers)
        
        # One slide per manager
        for manager_name, manager_data in period_data.items():
            await self._add_manager_stats_slide(prs, manager_name, manager_data, avg, margin)
        
        # Save to bytes
        buffer = io.BytesIO()
        prs.save(buffer)
        buffer.seek(0)
        return buffer.read()
    
    def _calculate_averages(self, period_data, total_managers):
        """Calculate team averages."""
        totals = {
            'calls_plan': sum(m.calls_plan for m in period_data.values()),
            'calls_fact': sum(m.calls_fact for m in period_data.values()),
            'new_calls_plan': sum(m.new_calls_plan for m in period_data.values()),
            'new_calls_fact': sum(m.new_calls for m in period_data.values()),
            'leads_units_plan': sum(m.leads_units_plan for m in period_data.values()),
            'leads_units_fact': sum(m.leads_units_fact for m in period_data.values()),
            'leads_volume_plan': sum(m.leads_volume_plan for m in period_data.values()),
            'leads_volume_fact': sum(m.leads_volume_fact for m in period_data.values()),
            'approved_units': sum(getattr(m, 'approved_units', 0) for m in period_data.values()),
            'issued_volume': sum(m.issued_volume for m in period_data.values()),
        }
        return {k: v / total_managers if total_managers > 0 else 0 for k, v in totals.items()}
    
    async def _add_title_slide(self, prs, period_name, start_date, end_date, margin):
        """Title slide."""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        self._add_logo(slide, prs)
        
        # Title
        title_box = slide.shapes.add_textbox(
            margin, Inches(2.5), prs.slide_width - 2*margin, Inches(1.5)
        )
        title_box.text_frame.text = f"Отчет По Активности\n{period_name}"
        for p in title_box.text_frame.paragraphs:
            p.font.name = "Roboto"
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(PRIMARY)
            p.alignment = PP_ALIGN.CENTER
    
    async def _add_manager_stats_slide(self, prs, manager_name, manager_data, avg, margin):
        """Manager statistics table + AI commentary (exactly as in reference)."""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        self._add_logo(slide, prs)
        
        # Title with manager name
        title_box = slide.shapes.add_textbox(margin, Inches(0.3), prs.slide_width - 2*margin, Inches(0.5))
        title_box.text_frame.text = f"Статистика менеджера: {manager_name}"
        title_box.text_frame.paragraphs[0].font.name = "Roboto"
        title_box.text_frame.paragraphs[0].font.size = Pt(22)
        title_box.text_frame.paragraphs[0].font.bold = True
        title_box.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(PRIMARY)
        title_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Table: 7 rows (header + 6 metrics), 5 columns
        rows = 7
        cols = 5
        tbl_width = prs.slide_width - 2*margin
        tbl_height = Inches(3.0)
        tbl = slide.shapes.add_table(rows, cols, margin, Inches(0.9), tbl_width, tbl_height).table
        
        # Set column widths
        col_widths = [Inches(2.2), Inches(1.8), Inches(1.5), Inches(2.2), Inches(2.2)]
        for c in range(cols):
            tbl.columns[c].width = col_widths[c]
        
        # Headers
        headers = ["", "Запланировано", "Факт", "Конверсия", "Средний\nменеджер\nфакт", "Средний\nменеджер\nконверсия"]
        # Actually 6 columns if we include empty first column header
        # But reference shows 5 columns with merged first cell - let's use 6 cols
        # Re-create table with 6 columns
        tbl = slide.shapes.add_table(rows, 6, margin, Inches(0.9), tbl_width, tbl_height).table
        
        col_widths = [Inches(2.2), Inches(1.6), Inches(1.5), Inches(1.7), Inches(1.8), Inches(2.0)]
        for c in range(6):
            tbl.columns[c].width = col_widths[c]
        
        headers = ["", "Запланировано", "Факт", "Конверсия", "Средний\nменеджер\nфакт", "Средний\nменеджер\nконверсия"]
        for c, hdr in enumerate(headers):
            cell = tbl.cell(0, c)
            cell.text = hdr
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb("#E3F2FD")
            for p in cell.text_frame.paragraphs:
                p.font.name = "Roboto"
                p.font.size = Pt(11)
                p.font.bold = True
                p.font.color.rgb = hex_to_rgb(TEXT_MAIN)
                p.alignment = PP_ALIGN.CENTER
        
        # Manager data vs averages
        m = manager_data
        row_data = [
            ["Повторные звонки", m.calls_plan, m.calls_fact, 
             f"{(m.calls_fact/m.calls_plan*100) if m.calls_plan else 0:.0f}%",
             f"{avg['calls_fact']:.1f}",
             f"{(avg['calls_fact']/avg['calls_plan']*100) if avg['calls_plan'] else 0:.0f}%"],
            ["Новые звонки", m.new_calls_plan, m.new_calls,
             f"{(m.new_calls/m.new_calls_plan*100) if m.new_calls_plan else 0:.0f}%",
             f"{avg['new_calls_fact']:.1f}",
             f"{(avg['new_calls_fact']/avg['new_calls_plan']*100) if avg['new_calls_plan'] else 0:.0f}%"],
            ["Заведено заявок шт", m.leads_units_plan, m.leads_units_fact,
             f"{(m.leads_units_fact/m.leads_units_plan*100) if m.leads_units_plan else 0:.0f}%",
             f"{avg['leads_units_fact']:.1f}",
             f"{(avg['leads_units_fact']/avg['leads_units_plan']*100) if avg['leads_units_plan'] else 0:.0f}%"],
            ["Заведено заявок, млн", f"{m.leads_volume_plan:.1f}", f"{m.leads_volume_fact:.1f}",
             f"{(m.leads_volume_fact/m.leads_volume_plan*100) if m.leads_volume_plan else 0:.0f}%",
             f"{avg['leads_volume_fact']:.1f}",
             f"{(avg['leads_volume_fact']/avg['leads_volume_plan']*100) if avg['leads_volume_plan'] else 0:.0f}%"],
            ["Одобрено заявок шт", "", getattr(m, 'approved_units', 0), "", f"{avg['approved_units']:.1f}", ""],
            ["Выдано, млн", "", f"{m.issued_volume:.1f}", "", f"{avg['issued_volume']:.1f}", ""]
        ]
        
        for r, data in enumerate(row_data, start=1):
            for c, val in enumerate(data):
                cell = tbl.cell(r, c)
                cell.text = str(val)
                try:
                    cell.vertical_anchor = 1  # Middle
                except Exception:
                    pass
                if r % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = hex_to_rgb("#F7F9FC")
                for p in cell.text_frame.paragraphs:
                    p.font.name = "Roboto"
                    p.font.size = Pt(10)
                    p.font.color.rgb = hex_to_rgb(TEXT_MAIN)
                    p.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
                    p.space_before = Pt(3)
                    p.space_after = Pt(3)
        
        # AI commentary section
        y_start = Inches(0.9) + tbl_height + Inches(0.2)
        comment_box = slide.shapes.add_textbox(margin, y_start, prs.slide_width - 2*margin, Inches(2.8))
        
        # Generate AI comment
        calls_conv = (m.calls_fact/m.calls_plan*100) if m.calls_plan else 0
        new_calls_conv = (m.new_calls/m.new_calls_plan*100) if m.new_calls_plan else 0
        leads_conv = (m.leads_units_fact/m.leads_units_plan*100) if m.leads_units_plan else 0
        vol_conv = (m.leads_volume_fact/m.leads_volume_plan*100) if m.leads_volume_plan else 0
        
        prompt = f"""Ты — руководитель отдела по банковским гарантиям. Проанализируй работу менеджера {manager_name} за период.

ПОКАЗАТЕЛИ МЕНЕДЖЕРА:
- Повторные звонки: {m.calls_fact} из {m.calls_plan} ({calls_conv:.0f}%) | Средний: {avg['calls_fact']:.1f}
- Новые звонки: {m.new_calls} из {m.new_calls_plan} ({new_calls_conv:.0f}%) | Средний: {avg['new_calls_fact']:.1f}
- Заявки (шт): {m.leads_units_fact} из {m.leads_units_plan} ({leads_conv:.0f}%) | Средний: {avg['leads_units_fact']:.1f}
- Заявки (млн): {m.leads_volume_fact:.1f} из {m.leads_volume_plan:.1f} ({vol_conv:.0f}%) | Средний: {avg['leads_volume_fact']:.1f}
- Одобрено: {getattr(m, 'approved_units', 0)} шт | Средний: {avg['approved_units']:.1f}
- Выдано (млн): {m.issued_volume:.1f} | Средний: {avg['issued_volume']:.1f}

ПРАВИЛА АНАЛИЗА:
1. Конверсия 80%+ = отлично, 60-80% = норма, <60% = нужна работа
2. Одобрено ≠ Выдано — это нормально (выдачи растягиваются по времени, не ругай конверсию)
3. Сравни КАЖДЫЙ показатель менеджера со средним по команде
4. Дай общий вывод: работает лучше/хуже/на уровне команды

ФОРМАТ ОТВЕТА (строго 3 пункта):
1. [Сравнение звонков и активности с командой — выше/ниже среднего, на сколько]
2. [Сравнение заявок и объёмов с командой — выше/ниже среднего, конверсия план/факт]
3. [Общий вывод и рекомендация — продолжить/усилить активность/проработать конверсию]

Пиши КРАТКО, ДЕЛОВЫМ ЯЗЫКОМ, БЕЗ ВОДЫ."""
        
        try:
            ai_text = await self.ai.generate_answer(prompt)
        except Exception as e:
            ai_text = (
                "1. Активность менеджера в работе\n"
                "2. Конверсия по заявкам и объёмам\n"
                "3. Общий вывод и рекомендации"
            )
        
        # Add title and text
        tf = comment_box.text_frame
        tf.clear()
        
        title_p = tf.paragraphs[0]
        title_p.text = "Комментарий ИИ"
        title_p.font.name = "Roboto"
        title_p.font.size = Pt(14)
        title_p.font.bold = True
        title_p.font.color.rgb = hex_to_rgb(TEXT_MAIN)
        title_p.space_after = Pt(8)
        
        # Subtitle
        subtitle_p = tf.add_paragraph()
        subtitle_p.text = "Коротко и по делу:"
        subtitle_p.font.name = "Roboto"
        subtitle_p.font.size = Pt(11)
        subtitle_p.font.color.rgb = hex_to_rgb(TEXT_MUTED)
        subtitle_p.space_after = Pt(6)
        
        # AI text (compact)
        for line in ai_text.split('\n'):
            if line.strip():
                p = tf.add_paragraph()
                p.text = line.strip()
                p.font.name = "Roboto"
                p.font.size = Pt(10)
                p.font.color.rgb = hex_to_rgb(TEXT_MAIN)
                p.space_after = Pt(3)
                p.level = 0

    def _add_logo(self, slide, prs) -> None:
        """Add company logo at top-right if file exists (Логотип.png)."""
        import os
        from pptx.util import Inches
        logo_name_candidates = ["Логотип.png", "logo.png", "Logo.png"]
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        logo_path = None
        for name in logo_name_candidates:
            p = os.path.join(root, name)
            if os.path.exists(p):
                logo_path = p
                break
        if not logo_path:
            return
        try:
            width = Inches(1.4)
            left = prs.slide_width - width - Inches(0.4)
            top = Inches(0.2)
            slide.shapes.add_picture(logo_path, left, top, width=width)
        except Exception:
            pass

