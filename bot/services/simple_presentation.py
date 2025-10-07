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
from bot.services.data_aggregator import DataAggregatorService
from bot.services.di import Container


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
        """Generate presentation with title + one slide per manager + team summary."""
        prs = Presentation()
        prs.slide_width = Inches(11.69)  # 16:9
        prs.slide_height = Inches(6.58)
        
        margin = Inches(0.6)
        
        # Slide 1: Title
        await self._add_title_slide(prs, period_name, start_date, end_date, margin)
        
        # Calculate team averages for comparison
        total_managers = len(period_data)
        avg = self._calculate_averages(period_data, total_managers)
        prev_avg = self._calculate_averages(prev_data, len(prev_data)) if prev_data else None

        # Compute previous quarter references (weekly):
        prev_q_team_weekly, prev_q_per_manager_weekly = await self._compute_prev_quarter_refs(end_date)
        # Calls overview slide (second slide)
        await self._add_calls_overview_slide(prs, period_data, prev_data, prev_q_per_manager_weekly, margin, period_name)
        # Leads overview slide (third slide)
        await self._add_leads_overview_slide(prs, period_data, prev_data, prev_q_per_manager_weekly, margin, period_name)
        
        # One slide per manager
        for manager_name, manager_data in period_data.items():
            await self._add_manager_stats_slide(prs, manager_name, manager_data, avg, prev_avg, margin)
        # Team summary slide
        await self._add_team_summary_slide(prs, period_data, avg, period_name, margin)
        
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

    async def _compute_prev_quarter_refs(self, end_date):
        """Compute previous quarter weekly averages: team and per-manager.

        Team weekly avg: sum facts across existing weeks / weeks_with_data.
        Per-manager weekly avg: average of manager weekly avgs among managers with data.
        """
        try:
            container = Container.get()
            aggregator = DataAggregatorService(container.sheets)
            # Determine previous quarter date range using aggregator helper
            from datetime import timedelta
            # Use end_date to infer current quarter; reuse existing helper via quarter aggregation, then shift
            current, _, _, _, prev_start, prev_end = await aggregator.aggregate_quarterly_data_with_previous()
            # prev_start/prev_end already describe previous quarter
            series = await aggregator.get_daily_series(prev_start, prev_end)
            # Build week buckets
            from collections import defaultdict
            week_buckets = defaultdict(lambda: {'calls_fact':0,'new_calls':0,'leads_units_fact':0,'leads_volume_fact':0.0,'approved_volume':0.0,'issued_volume':0.0})
            for item in series:
                from datetime import datetime as _dt
                d = _dt.strptime(item['date'], '%Y-%m-%d').date()
                iso_year, iso_week, _ = d.isocalendar()
                key = f"{iso_year}-W{iso_week:02d}"
                for k in week_buckets[key]:
                    week_buckets[key][k] += float(item.get(k, 0) or 0)
            weeks_with_data = [w for w in week_buckets.values() if any(v>0 for v in w.values())]
            def avg_of_key(key: str) -> float:
                if not weeks_with_data:
                    return 0.0
                return sum(w[key] for w in weeks_with_data) / len(weeks_with_data)
            team_weekly = {
                'calls_fact': avg_of_key('calls_fact'),
                'new_calls_fact': avg_of_key('new_calls'),
                'leads_units_fact': avg_of_key('leads_units_fact'),
                'leads_volume_fact': avg_of_key('leads_volume_fact'),
                'approved_units': avg_of_key('approved_volume'),
                'issued_volume': avg_of_key('issued_volume'),
                'calls_plan': 0,
                'new_calls_plan': 0,
                'leads_units_plan': 0,
                'leads_volume_plan': 0.0,
            }
            # Per-manager weekly average: approximate using per-team divided by count of active managers last quarter
            prev_data = await aggregator._aggregate_data_for_period(prev_start, prev_end)
            active_managers = len(prev_data) if prev_data else 0
            per_manager_weekly = None
            if active_managers > 0:
                per_manager_weekly = {
                    'calls_fact': team_weekly['calls_fact'] / active_managers,
                    'new_calls_fact': team_weekly['new_calls_fact'] / active_managers,
                    'leads_units_fact': team_weekly['leads_units_fact'] / active_managers,
                    'leads_volume_fact': team_weekly['leads_volume_fact'] / active_managers,
                    'approved_units': team_weekly['approved_units'] / active_managers,
                    'issued_volume': team_weekly['issued_volume'] / active_managers,
                    'calls_plan': 0,
                    'new_calls_plan': 0,
                    'leads_units_plan': 0,
                    'leads_volume_plan': 0.0,
                }
            return team_weekly, per_manager_weekly
        except Exception:
            return None, None

    async def _add_calls_overview_slide(self, prs, period_data, prev_data, avg, margin, period_name):
        """Add 'Общие показатели звонков' slide as per reference."""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        self._add_logo(slide, prs)
        # Title
        title = slide.shapes.add_textbox(margin, Inches(0.3), prs.slide_width - 2*margin, Inches(0.5))
        tf = title.text_frame
        tf.text = "Общие показатели звонков"
        p = tf.paragraphs[0]; p.font.name = "Roboto"; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = hex_to_rgb(PRIMARY); p.alignment = PP_ALIGN.CENTER

        # Aggregate current totals
        cur_calls_plan = sum(m.calls_plan for m in period_data.values())
        cur_calls_fact = sum(m.calls_fact for m in period_data.values())
        cur_new_plan = sum(m.new_calls_plan for m in period_data.values())
        cur_new_fact = sum(m.new_calls for m in period_data.values())

        # Previous totals (fallback 0)
        prev_calls_fact = sum((getattr(m, 'calls_fact', 0) for m in (prev_data or {}).values())) if prev_data else 0
        prev_new_fact = sum((getattr(m, 'new_calls', 0) for m in (prev_data or {}).values())) if prev_data else 0
        prev_calls_plan = sum((getattr(m, 'calls_plan', 0) for m in (prev_data or {}).values())) if prev_data else 0
        prev_new_plan = sum((getattr(m, 'new_calls_plan', 0) for m in (prev_data or {}).values())) if prev_data else 0

        # Table with 2 rows + header, columns: Показатель, План, Факт, Конверсия, % к факту, % конверсии, Средний факт
        rows, cols = 3, 7
        tbl = slide.shapes.add_table(rows, cols, margin, Inches(0.9), prs.slide_width - 2*margin, Inches(2.2)).table
        headers = ["Показатель", "План", "Факт", "Конверсия", "% к факту (ПП)", "Δ конверсии, п.п. (ПП)", "Средний факт (ПП)"]
        for c, h in enumerate(headers):
            cell = tbl.cell(0, c); cell.text = h
            cell.fill.solid();
            # Подсветка правого блока колонок, относящихся к ПП
            cell.fill.fore_color.rgb = hex_to_rgb("#E3F2FD" if c < 4 else "#BBDEFB")
            for par in cell.text_frame.paragraphs:
                par.font.name = "Roboto"; par.font.size = Pt(11); par.font.bold = True; par.alignment = PP_ALIGN.CENTER; par.font.color.rgb = hex_to_rgb(TEXT_MAIN)

        # Data rows
        def pct(a, b):
            try:
                return round(a / b * 100) if b else 0
            except Exception:
                return 0

        calls_conv = pct(cur_calls_fact, cur_calls_plan)
        new_conv = pct(cur_new_fact, cur_new_plan)
        prev_calls_conv = pct(prev_calls_fact, prev_calls_plan) if prev_calls_plan else 0
        prev_new_conv = pct(prev_new_fact, prev_new_plan) if prev_new_plan else 0
        vs_calls = pct(cur_calls_fact - prev_calls_fact, prev_calls_fact) if prev_calls_fact else 0
        vs_new = pct(cur_new_fact - prev_new_fact, prev_new_fact) if prev_new_fact else 0
        # conversion deltas in percentage points
        vs_calls_conv = calls_conv - prev_calls_conv
        vs_new_conv = new_conv - prev_new_conv

        data_rows = [
            ["Повторные звонки", cur_calls_plan, cur_calls_fact, f"{calls_conv}%", f"{vs_calls:+d}%", f"{vs_calls_conv:+d}", f"{avg['calls_fact']:.1f}"],
            ["Новые звонки", cur_new_plan, cur_new_fact, f"{new_conv}%", f"{vs_new:+d}%", f"{vs_new_conv:+d}", f"{avg['new_calls_fact']:.1f}"],
        ]
        for r, row in enumerate(data_rows, start=1):
            for c, v in enumerate(row):
                cell = tbl.cell(r, c); cell.text = str(v)
                if r % 2 == 0:
                    cell.fill.solid(); cell.fill.fore_color.rgb = hex_to_rgb("#F7F9FC")
                for par in cell.text_frame.paragraphs:
                    par.font.name = "Roboto"; par.font.size = Pt(10); par.alignment = PP_ALIGN.CENTER if c>0 else PP_ALIGN.LEFT
                    if c >= 4:  # колонки, относящиеся к ПП
                        par.font.color.rgb = hex_to_rgb(PRIMARY)

        # Примечание про ПП
        note = slide.shapes.add_textbox(margin, Inches(3.15), prs.slide_width - 2*margin, Inches(0.3))
        nt = note.text_frame; nt.clear()
        pnt = nt.paragraphs[0]; pnt.text = "Колонки справа (% к факту, Δ конверсии, Средний факт) — сравнение с ПП (предыдущим периодом той же длины)."
        pnt.font.name = "Roboto"; pnt.font.size = Pt(9); pnt.font.color.rgb = hex_to_rgb(TEXT_MUTED)

        # Comment block
        box = slide.shapes.add_textbox(margin, Inches(3.45), prs.slide_width - 2*margin, Inches(2.25))
        prompt = (
            "Сформируй краткий комментарий по звонкам за период '" + period_name + "'. "
            f"Повторные: факт {cur_calls_fact} из {cur_calls_plan} ({calls_conv}%). Новые: факт {cur_new_fact} из {cur_new_plan} ({new_conv}%). "
            "Сравни с прошлым периодом, укажи изменения в %, сделай 2–3 тезиса и 1 рекомендацию."
        )
        try:
            text = await self.ai.generate_answer(prompt)
        except Exception:
            text = "Количество звонков выросло/снизилось относительно прошлого периода. Рекомендация: скорректировать темп и довести план."
        t = box.text_frame; t.clear();
        h = t.paragraphs[0]; h.text = "Комментарии нейросети"; h.font.name = "Roboto"; h.font.size = Pt(14); h.font.bold = True
        p1 = t.add_paragraph(); p1.text = text; p1.font.name = "Roboto"; p1.font.size = Pt(10)

    async def _add_leads_overview_slide(self, prs, period_data, prev_data, avg, margin, period_name):
        """Add 'Общие показатели по заявкам' slide (units, volume, approved, issued)."""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        self._add_logo(slide, prs)
        # Title
        title = slide.shapes.add_textbox(margin, Inches(0.3), prs.slide_width - 2*margin, Inches(0.5))
        tf = title.text_frame
        tf.text = "Общие показатели по заявкам"
        p = tf.paragraphs[0]; p.font.name = "Roboto"; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = hex_to_rgb(PRIMARY); p.alignment = PP_ALIGN.CENTER

        # Current totals
        units_plan = 0  # план по заявкам признан бессмысленным и исключён
        units_fact = sum(m.leads_units_fact for m in period_data.values())
        vol_plan = 0  # план по заявкам (млн) исключён
        vol_fact = sum(m.leads_volume_fact for m in period_data.values())
        approved_plan = sum(getattr(m, 'approved_plan', 0) for m in period_data.values())
        approved_fact = sum(getattr(m, 'approved_volume', 0) for m in period_data.values())
        issued_plan = sum(getattr(m, 'issued_plan', 0) for m in period_data.values())
        issued_fact = sum(m.issued_volume for m in period_data.values())

        # Previous totals
        prev_units_fact = sum((getattr(m, 'leads_units_fact', 0) for m in (prev_data or {}).values())) if prev_data else 0
        prev_vol_fact = sum((getattr(m, 'leads_volume_fact', 0.0) for m in (prev_data or {}).values())) if prev_data else 0
        prev_approved_fact = sum((getattr(m, 'approved_volume', 0.0) for m in (prev_data or {}).values())) if prev_data else 0
        prev_issued_fact = sum((getattr(m, 'issued_volume', 0.0) for m in (prev_data or {}).values())) if prev_data else 0

        # Table
        rows, cols = 5, 7
        tbl = slide.shapes.add_table(rows, cols, margin, Inches(0.9), prs.slide_width - 2*margin, Inches(2.4)).table
        headers = ["Показатель", "План", "Факт", "Конверсия", "% к факту (ПП)", "Δ конверсии, п.п. (ПП)", "Средний факт (ПП)"]
        for c, h in enumerate(headers):
            cell = tbl.cell(0, c); cell.text = h
            cell.fill.solid(); cell.fill.fore_color.rgb = hex_to_rgb("#E3F2FD" if c < 4 else "#BBDEFB")
            for par in cell.text_frame.paragraphs:
                par.font.name = "Roboto"; par.font.size = Pt(11); par.font.bold = True; par.alignment = PP_ALIGN.CENTER; par.font.color.rgb = hex_to_rgb(TEXT_MAIN)

        def pct(a, b):
            try:
                return round(a / b * 100) if b else 0
            except Exception:
                return 0

        units_conv = 0
        vol_conv = 0
        approved_conv = pct(approved_fact, approved_plan) if approved_plan else 0
        issued_conv = pct(issued_fact, issued_plan) if issued_plan else 0

        data_rows = [
            ["Заявки, штук", "—", units_fact, "—", f"{pct(units_fact - prev_units_fact, prev_units_fact) if prev_units_fact else 0:+d}%", "—", f"{avg['leads_units_fact']:.1f}"],
            ["Заявки, млн", "—", f"{vol_fact:.1f}", "—", f"{pct(vol_fact - prev_vol_fact, prev_vol_fact) if prev_vol_fact else 0:+d}%", "—", f"{avg['leads_volume_fact']:.1f}"],
            ["Одобрено, млн", f"{approved_plan:.1f}" if approved_plan else "-", f"{approved_fact:.1f}", f"{approved_conv}%" if approved_plan else "—", f"{pct(approved_fact - prev_approved_fact, prev_approved_fact) if prev_approved_fact else 0:+d}%", "—", f"{avg['approved_units']:.1f}"],
            ["Выдано, млн", f"{issued_plan:.1f}" if issued_plan else "-", f"{issued_fact:.1f}", f"{issued_conv}%" if issued_plan else "—", f"{pct(issued_fact - prev_issued_fact, prev_issued_fact) if prev_issued_fact else 0:+d}%", "—", f"{avg['issued_volume']:.1f}"],
        ]
        for r, row in enumerate(data_rows, start=1):
            for c, v in enumerate(row):
                cell = tbl.cell(r, c); cell.text = str(v)
                if r % 2 == 0:
                    cell.fill.solid(); cell.fill.fore_color.rgb = hex_to_rgb("#F7F9FC")
                for par in cell.text_frame.paragraphs:
                    par.font.name = "Roboto"; par.font.size = Pt(10); par.alignment = PP_ALIGN.CENTER if c>0 else PP_ALIGN.LEFT
                    if c >= 4:
                        par.font.color.rgb = hex_to_rgb(PRIMARY)

        # Comment
        box = slide.shapes.add_textbox(margin, Inches(3.4), prs.slide_width - 2*margin, Inches(2.3))
        prompt = (
            "Сформируй комментарий по заявкам (шт и млн), одобрениям и выдачам за период '" + period_name + "'. "
            f"Факт: units {units_fact}/{units_plan}, volume {vol_fact:.1f}/{vol_plan:.1f}, approved {approved_fact:.1f}, issued {issued_fact:.1f}. "
            "Сравни с прошлым периодом, отметь сильные/слабые места, дай 2–3 рекомендации."
        )
        try:
            text = await self.ai.generate_answer(prompt)
        except Exception:
            text = "Факт заявок и объёмов оценён. Рекомендуется сфокусироваться на стабильности одобрений и доведении выдач."
        t = box.text_frame; t.clear();
        h = t.paragraphs[0]; h.text = "Комментарии нейросети"; h.font.name = "Roboto"; h.font.size = Pt(14); h.font.bold = True
        p1 = t.add_paragraph(); p1.text = text; p1.font.name = "Roboto"; p1.font.size = Pt(10)

    async def _add_team_summary_slide(self, prs, period_data, avg, period_name, margin):
        """Add team totals and concise AI comment."""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        self._add_logo(slide, prs)
        # Title
        t = slide.shapes.add_textbox(margin, Inches(0.3), prs.slide_width - 2*margin, Inches(0.5))
        t.text_frame.text = f"Итоги по команде: {period_name}"
        p0 = t.text_frame.paragraphs[0]
        p0.font.name = "Roboto"
        p0.font.size = Pt(22)
        p0.font.bold = True
        p0.font.color.rgb = hex_to_rgb(PRIMARY)
        p0.alignment = PP_ALIGN.CENTER

        # Totals table (4 строки)
        rows, cols = 5, 4
        tbl = slide.shapes.add_table(rows, cols, margin, Inches(0.9), prs.slide_width - 2*margin, Inches(2.0)).table
        headers = ["Метрика", "Факт", "План", "Выполнение"]
        for c, h in enumerate(headers):
            cell = tbl.cell(0, c)
            cell.text = h
            cell.fill.solid(); cell.fill.fore_color.rgb = hex_to_rgb("#E3F2FD")
            for p in cell.text_frame.paragraphs:
                p.font.name = "Roboto"; p.font.size = Pt(11); p.font.bold = True; p.alignment = PP_ALIGN.CENTER

        # Compute totals
        total_calls_plan = sum(m.calls_plan for m in period_data.values())
        total_calls_fact = sum(m.calls_fact for m in period_data.values())
        total_units_plan = sum(m.leads_units_plan for m in period_data.values())
        total_units_fact = sum(m.leads_units_fact for m in period_data.values())
        total_vol_plan = sum(m.leads_volume_plan for m in period_data.values())
        total_vol_fact = sum(m.leads_volume_fact for m in period_data.values())
        total_issued = sum(m.issued_volume for m in period_data.values())

        data_rows = [
            ["Перезвоны", total_calls_fact, total_calls_plan, f"{(total_calls_fact/total_calls_plan*100) if total_calls_plan else 0:.0f}%"],
            ["Заявки, шт", total_units_fact, total_units_plan, f"{(total_units_fact/total_units_plan*100) if total_units_plan else 0:.0f}%"],
            ["Заявки, млн", f"{total_vol_fact:.1f}", f"{total_vol_plan:.1f}", f"{(total_vol_fact/total_vol_plan*100) if total_vol_plan else 0:.0f}%"],
            ["Выдано, млн", f"{total_issued:.1f}", "—", "—"],
        ]
        for r, row in enumerate(data_rows, start=1):
            for c, v in enumerate(row):
                cell = tbl.cell(r, c); cell.text = str(v)
                if r % 2 == 0:
                    cell.fill.solid(); cell.fill.fore_color.rgb = hex_to_rgb("#F7F9FC")
                for p in cell.text_frame.paragraphs:
                    p.font.name = "Roboto"; p.font.size = Pt(10); p.alignment = PP_ALIGN.CENTER if c>0 else PP_ALIGN.LEFT

        # Team AI comment
        box = slide.shapes.add_textbox(margin, Inches(3.1), prs.slide_width - 2*margin, Inches(2.6))
        from textwrap import dedent
        prompt = dedent(f"""
        Ты — руководитель отдела. Дай краткий вывод по команде (5–7 предложений), деловым стилем.
        Укажи выполнение планов по перезвонам, заявкам (шт и млн) и общий итог по выдачам. Дай 2–3 приоритета на неделю.
        Данные: calls {total_calls_fact}/{total_calls_plan}; units {total_units_fact}/{total_units_plan}; volume {total_vol_fact:.1f}/{total_vol_plan:.1f}; issued {total_issued:.1f}.
        """)
        try:
            comment = await self.ai.generate_answer(prompt)
        except Exception:
            comment = "Итоги сформированы. Фокус: довести планы по перезвонам и объёмам; удерживать выдачи."
        tf = box.text_frame; tf.clear()
        hp = tf.paragraphs[0]; hp.text = "Итоги по команде"
        hp.font.name = "Roboto"; hp.font.size = Pt(14); hp.font.bold = True; hp.font.color.rgb = hex_to_rgb(TEXT_MAIN)
        p = tf.add_paragraph(); p.text = comment; p.font.name = "Roboto"; p.font.size = Pt(10)
    
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
    
    async def _add_manager_stats_slide(self, prs, manager_name, manager_data, avg, prev_avg, margin):
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
        # baseline for "средний менеджер": из прошлого квартала (недельный), если доступно; иначе текущая средняя
        _, prev_q_per_manager_weekly = await self._compute_prev_quarter_refs(Container.get().settings and None)  # settings ignored
        ref = prev_q_per_manager_weekly or prev_avg or avg
        row_data = [
            ["Повторные звонки", m.calls_plan, m.calls_fact, 
             f"{(m.calls_fact/m.calls_plan*100) if m.calls_plan else 0:.0f}%",
             f"{ref['calls_fact']:.1f}",
             f"{(ref['calls_fact']/ref['calls_plan']*100) if ref['calls_plan'] else 0:.0f}%"],
            ["Новые звонки", m.new_calls_plan, m.new_calls,
             f"{(m.new_calls/m.new_calls_plan*100) if m.new_calls_plan else 0:.0f}%",
             f"{ref['new_calls_fact']:.1f}",
             f"{(ref['new_calls_fact']/ref['new_calls_plan']*100) if ref['new_calls_plan'] else 0:.0f}%"],
            ["Заведено заявок шт", "—", m.leads_units_fact,
             "—",
             f"{ref['leads_units_fact']:.1f}",
             f"{(ref['leads_units_fact']/ref['leads_units_plan']*100) if ref['leads_units_plan'] else 0:.0f}%"],
            ["Заведено заявок, млн", "—", f"{m.leads_volume_fact:.1f}",
             "—",
             f"{ref['leads_volume_fact']:.1f}",
             f"{(ref['leads_volume_fact']/ref['leads_volume_plan']*100) if ref['leads_volume_plan'] else 0:.0f}%"],
            ["Одобрено заявок шт", "", getattr(m, 'approved_units', 0), "", f"{ref['approved_units']:.1f}", ""],
            ["Выдано, млн", "", f"{m.issued_volume:.1f}", "", f"{ref['issued_volume']:.1f}", ""]
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
- Повторные звонки: {m.calls_fact} из {m.calls_plan} ({calls_conv:.0f}%) | Средний менеджер: {ref['calls_fact']:.1f}
- Новые звонки: {m.new_calls} из {m.new_calls_plan} ({new_calls_conv:.0f}%) | Средний менеджер: {ref['new_calls_fact']:.1f}
- Заявки (шт): {m.leads_units_fact} из {m.leads_units_plan} ({leads_conv:.0f}%) | Средний менеджер: {ref['leads_units_fact']:.1f}
- Заявки (млн): {m.leads_volume_fact:.1f} из {m.leads_volume_plan:.1f} ({vol_conv:.0f}%) | Средний менеджер: {ref['leads_volume_fact']:.1f}
- Одобрено: {getattr(m, 'approved_units', 0)} шт | Средний менеджер: {ref['approved_units']:.1f}
- Выдано (млн): {m.issued_volume:.1f} | Средний менеджер: {ref['issued_volume']:.1f}

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

