"""Presentation generation service."""
import os
import io
from datetime import datetime, date
from typing import Dict, Any, List, Tuple, Optional
from dataclasses import dataclass

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from bot.services.yandex_gpt import YandexGPTService
from bot.config import Settings


@dataclass
class ManagerData:
    """Data structure for manager statistics."""
    name: str
    calls_plan: int = 0
    calls_fact: int = 0
    leads_units_plan: int = 0
    leads_units_fact: int = 0
    leads_volume_plan: float = 0.0
    leads_volume_fact: float = 0.0
    approved_volume: float = 0.0
    issued_volume: float = 0.0
    new_calls: int = 0
    
    @property
    def calls_percentage(self) -> float:
        """Calculate calls completion percentage."""
        return (self.calls_fact / self.calls_plan * 100) if self.calls_plan > 0 else 0
    
    @property
    def leads_units_percentage(self) -> float:
        """Calculate leads units completion percentage."""
        return (self.leads_units_fact / self.leads_units_plan * 100) if self.leads_units_plan > 0 else 0
    
    @property
    def leads_volume_percentage(self) -> float:
        """Calculate leads volume completion percentage."""
        return (self.leads_volume_fact / self.leads_volume_plan * 100) if self.leads_volume_plan > 0 else 0


class PresentationService:
    """Service for generating PowerPoint presentations."""
    
    def __init__(self, settings: Settings):
        self.settings = settings
        self.gpt_service = YandexGPTService(settings)
    
    async def generate_presentation(
        self,
        period_data: Dict[str, ManagerData],
        period_name: str,
        start_date: date,
        end_date: date
    ) -> bytes:
        """
        Generate PowerPoint presentation with analytics.
        
        Args:
            period_data: Dictionary mapping manager names to their data
            period_name: Human-readable period name (e.g., "Неделя 18-24 августа")
            start_date: Period start date
            end_date: Period end date
            
        Returns:
            PPTX file as bytes
        """
        # Create presentation
        prs = Presentation()
        
        # Set slide size (16:9)
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        # Title slide
        await self._add_title_slide(prs, period_name, start_date, end_date)
        
        # Summary slide
        await self._add_summary_slide(prs, period_data, period_name)
        
        # Manager slides
        for manager_name, manager_data in period_data.items():
            await self._add_manager_slide(prs, manager_data)
        
        # AI Analysis slide
        await self._add_ai_analysis_slide(prs, period_data, period_name)
        
        # Save to bytes
        pptx_buffer = io.BytesIO()
        prs.save(pptx_buffer)
        pptx_buffer.seek(0)
        
        return pptx_buffer.getvalue()
    
    async def _add_title_slide(
        self,
        prs: Presentation,
        period_name: str,
        start_date: date,
        end_date: date
    ):
        """Add title slide."""
        slide_layout = prs.slide_layouts[0]  # Title slide layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Title
        title = slide.shapes.title
        title.text = f"Отчет по продажам"
        title.text_frame.paragraphs[0].font.size = Pt(44)
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(204, 0, 0)  # Red
        
        # Subtitle
        subtitle = slide.placeholders[1]
        subtitle.text = f"{period_name}\n{start_date.strftime('%d.%m.%Y')} — {end_date.strftime('%d.%m.%Y')}"
        subtitle.text_frame.paragraphs[0].font.size = Pt(28)
        subtitle.text_frame.paragraphs[1].font.size = Pt(20)
        subtitle.text_frame.paragraphs[1].font.color.rgb = RGBColor(102, 102, 102)  # Gray
    
    async def _add_summary_slide(
        self,
        prs: Presentation,
        period_data: Dict[str, ManagerData],
        period_name: str
    ):
        """Add summary slide with team totals."""
        slide_layout = prs.slide_layouts[1]  # Title and content layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Title
        title = slide.shapes.title
        title.text = f"Общие показатели команды"
        title.text_frame.paragraphs[0].font.size = Pt(32)
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(204, 0, 0)
        
        # Calculate totals
        totals = self._calculate_totals(period_data)
        
        # Content
        content = slide.placeholders[1]
        content.text = f"""📊 Итоги за {period_name}

📞 Перезвоны: {totals['calls_fact']:,} из {totals['calls_plan']:,} ({totals['calls_percentage']:.1f}%)
📝 Заявки (шт): {totals['leads_units_fact']:,} из {totals['leads_units_plan']:,} ({totals['leads_units_percentage']:.1f}%)
💰 Заявки (млн): {totals['leads_volume_fact']:.1f} из {totals['leads_volume_plan']:.1f} ({totals['leads_volume_percentage']:.1f}%)
✅ Одобрено (млн): {totals['approved_volume']:.1f}
✅ Выдано (млн): {totals['issued_volume']:.1f}
☎️ Новые звонки: {totals['new_calls']:,}

👥 Активных менеджеров: {len(period_data)}"""
        
        # Format content
        for paragraph in content.text_frame.paragraphs:
            paragraph.font.size = Pt(18)
            paragraph.space_after = Pt(6)
    
    async def _add_manager_slide(self, prs: Presentation, manager_data: ManagerData):
        """Add individual manager slide."""
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        
        # Title
        title = slide.shapes.title
        title.text = f"👤 {manager_data.name}"
        title.text_frame.paragraphs[0].font.size = Pt(32)
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(204, 0, 0)
        
        # Performance indicators
        calls_status = "🟢" if manager_data.calls_percentage >= 80 else "🟡" if manager_data.calls_percentage >= 60 else "🔴"
        leads_status = "🟢" if manager_data.leads_volume_percentage >= 80 else "🟡" if manager_data.leads_volume_percentage >= 60 else "🔴"
        
        # Content
        content = slide.placeholders[1]
        content.text = f"""📈 Показатели эффективности

{calls_status} Перезвоны: {manager_data.calls_fact:,} из {manager_data.calls_plan:,} ({manager_data.calls_percentage:.1f}%)

📝 Заявки (шт): {manager_data.leads_units_fact:,} из {manager_data.leads_units_plan:,} ({manager_data.leads_units_percentage:.1f}%)

{leads_status} Заявки (млн): {manager_data.leads_volume_fact:.1f} из {manager_data.leads_volume_plan:.1f} ({manager_data.leads_volume_percentage:.1f}%)

✅ Одобрено (млн): {manager_data.approved_volume:.1f}

✅ Выдано (млн): {manager_data.issued_volume:.1f}

☎️ Новые звонки: {manager_data.new_calls:,}"""
        
        # Format content
        for paragraph in content.text_frame.paragraphs:
            paragraph.font.size = Pt(16)
            paragraph.space_after = Pt(8)
    
    async def _add_ai_analysis_slide(
        self,
        prs: Presentation,
        period_data: Dict[str, ManagerData],
        period_name: str
    ):
        """Add AI analysis slide."""
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        
        # Title
        title = slide.shapes.title
        title.text = "🤖 AI-Анализ и рекомендации"
        title.text_frame.paragraphs[0].font.size = Pt(32)
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(204, 0, 0)
        
        # Generate AI analysis
        analysis_data = {}
        for name, data in period_data.items():
            analysis_data[name] = {
                'calls_plan': data.calls_plan,
                'calls_fact': data.calls_fact,
                'leads_units_plan': data.leads_units_plan,
                'leads_units_fact': data.leads_units_fact,
                'leads_volume_plan': data.leads_volume_plan,
                'leads_volume_fact': data.leads_volume_fact,
                'approved_volume': data.approved_volume,
                'issued_volume': data.issued_volume,
            }
        
        ai_analysis = await self.gpt_service.generate_analysis(analysis_data)
        
        # Content
        content = slide.placeholders[1]
        content.text = ai_analysis
        
        # Format content
        for paragraph in content.text_frame.paragraphs:
            paragraph.font.size = Pt(14)
            paragraph.space_after = Pt(6)
    
    def _calculate_totals(self, period_data: Dict[str, ManagerData]) -> Dict[str, float]:
        """Calculate team totals."""
        totals = {
            'calls_plan': 0,
            'calls_fact': 0,
            'leads_units_plan': 0,
            'leads_units_fact': 0,
            'leads_volume_plan': 0.0,
            'leads_volume_fact': 0.0,
            'approved_volume': 0.0,
            'issued_volume': 0.0,
            'new_calls': 0,
        }
        
        for manager_data in period_data.values():
            totals['calls_plan'] += manager_data.calls_plan
            totals['calls_fact'] += manager_data.calls_fact
            totals['leads_units_plan'] += manager_data.leads_units_plan
            totals['leads_units_fact'] += manager_data.leads_units_fact
            totals['leads_volume_plan'] += manager_data.leads_volume_plan
            totals['leads_volume_fact'] += manager_data.leads_volume_fact
            totals['approved_volume'] += manager_data.approved_volume
            totals['issued_volume'] += manager_data.issued_volume
            totals['new_calls'] += manager_data.new_calls
        
        # Calculate percentages
        totals['calls_percentage'] = (totals['calls_fact'] / totals['calls_plan'] * 100) if totals['calls_plan'] > 0 else 0
        totals['leads_units_percentage'] = (totals['leads_units_fact'] / totals['leads_units_plan'] * 100) if totals['leads_units_plan'] > 0 else 0
        totals['leads_volume_percentage'] = (totals['leads_volume_fact'] / totals['leads_volume_plan'] * 100) if totals['leads_volume_plan'] > 0 else 0
        
        return totals
