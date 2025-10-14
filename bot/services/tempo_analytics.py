"""Tempo analytics service for tracking performance against daily targets."""
import calendar
from datetime import date, datetime, timedelta
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from zoneinfo import ZoneInfo

from bot.services.sheets import SheetsClient
from bot.services.di import Container
from bot.utils.time_utils import now_in_tz


@dataclass
class TempoAlert:
    """Alert for manager falling behind tempo."""
    manager_name: str
    metric: str  # 'calls', 'issued_volume'
    current_value: float
    expected_value: float
    deviation_percent: float
    alert_level: str  # 'warning' or 'critical'
    message: str


class TempoAnalyticsService:
    """Service for analyzing manager performance tempo."""
    
    WARNING_THRESHOLD = -20.0  # -20%
    CRITICAL_THRESHOLD = -40.0  # -40%
    
    def __init__(self, sheets_service: SheetsClient):
        self.sheets_service = sheets_service
    
    async def analyze_monthly_tempo(
        self,
        target_date: Optional[date] = None,
        office_filter: Optional[str] = None,
    ) -> List[TempoAlert]:
        """
        Analyze managers' tempo against monthly plans.
        
        Args:
            target_date: Date to analyze (defaults to today)
            
        Returns:
            List of tempo alerts for managers falling behind
        """
        if target_date is None:
            target_date = now_in_tz(Container.get().settings).date()
        
        # Get monthly plans (hardcoded for MVP)
        monthly_plans = self._get_monthly_plans(target_date)
        if not monthly_plans:
            return []
        
        # Calculate working days
        working_days_total = self._count_working_days(target_date.year, target_date.month)
        working_days_passed = self._count_working_days_until(target_date)
        
        if working_days_passed == 0:
            return []  # No working days passed yet
        
        # Get actual data from sheets
        month_start = date(target_date.year, target_date.month, 1)
        actual_data = await self._get_actual_data_for_period(month_start, target_date, office_filter=office_filter)
        
        # Analyze each manager
        alerts = []
        for manager_name, plans in monthly_plans.items():
            if manager_name not in actual_data:
                continue
            
            manager_actuals = actual_data[manager_name]
            
            # Analyze calls tempo
            calls_alert = self._analyze_metric_tempo(
                manager_name=manager_name,
                metric="calls",
                plan_total=plans.get('calls_plan', 0),
                actual_total=manager_actuals.get('calls_fact', 0),
                working_days_total=working_days_total,
                working_days_passed=working_days_passed
            )
            if calls_alert:
                alerts.append(calls_alert)
            
            # Analyze issued volume tempo (факт выдано)
            volume_alert = self._analyze_metric_tempo(
                manager_name=manager_name,
                metric="issued_volume",
                plan_total=plans.get('issued_volume_plan', plans.get('leads_volume_plan', 0.0)),
                actual_total=manager_actuals.get('issued_volume_fact', 0.0),
                working_days_total=working_days_total,
                working_days_passed=working_days_passed
            )
            if volume_alert:
                alerts.append(volume_alert)
        
        return alerts

    async def build_monthly_tempo_rows(
        self,
        target_date: Optional[date] = None,
        office_filter: Optional[str] = None,
    ) -> List[str]:
        """Return human-readable lines with tempo per manager (even without alerts)."""
        if target_date is None:
            target_date = now_in_tz(Container.get().settings).date()

        monthly_plans = self._get_monthly_plans(target_date)

        working_days_total = self._count_working_days(target_date.year, target_date.month)
        working_days_passed = self._count_working_days_until(target_date)
        if working_days_total == 0 or working_days_passed == 0:
            return []

        month_start = date(target_date.year, target_date.month, 1)
        actual_data = await self._get_actual_data_for_period(month_start, target_date, office_filter=office_filter)

        lines: List[str] = []
        for manager_name in sorted(actual_data.keys()):
            act = actual_data.get(manager_name, {})
            plan = monthly_plans.get(manager_name, {})

            # Calls
            calls_plan = float(plan.get('calls_plan', 0) or 0)
            calls_fact = float(act.get('calls_fact', 0) or 0)
            calls_expected = (calls_plan / working_days_total * working_days_passed) if calls_plan > 0 else None
            calls_dev_pct = ((calls_fact - calls_expected) / calls_expected * 100) if calls_expected and calls_expected > 0 else None

            # Issued volume (or leads volume plan proxy)
            vol_plan = float(plan.get('issued_volume_plan', plan.get('leads_volume_plan', 0.0)) or 0.0)
            vol_fact = float(act.get('issued_volume_fact', 0.0) or 0.0)
            vol_expected = (vol_plan / working_days_total * working_days_passed) if vol_plan > 0 else None
            vol_dev_pct = ((vol_fact - vol_expected) / vol_expected * 100) if vol_expected and vol_expected > 0 else None

            def fmt_pair(fact: float, expected: Optional[float]) -> str:
                if expected is None:
                    return f"{fact:,.0f} / –"
                if isinstance(fact, float) and not fact.is_integer():
                    return f"{fact:.1f} / {expected:.1f}"
                return f"{int(fact):,} / {expected:.0f}"

            def fmt_dev(dev: Optional[float]) -> str:
                return f"{dev:+.1f}%" if dev is not None else "–"

            line = (
                f"👤 {manager_name}\n"
                f"• 📲 Перезвоны: {fmt_pair(calls_fact, calls_expected)} (∆ {fmt_dev(calls_dev_pct)})\n"
                f"• 💰 Выдано, млн: {vol_fact:.1f} / {vol_expected:.1f}" + (f" (∆ {vol_dev_pct:+.1f}%)" if vol_expected else " (план не задан)")
            )
            lines.append(line)

        return lines
    
    def _analyze_metric_tempo(
        self,
        manager_name: str,
        metric: str,
        plan_total: float,
        actual_total: float,
        working_days_total: int,
        working_days_passed: int
    ) -> Optional[TempoAlert]:
        """Analyze tempo for a specific metric."""
        if plan_total == 0 or working_days_total == 0:
            return None
        
        # Calculate expected value by this date
        daily_plan = plan_total / working_days_total
        expected_value = daily_plan * working_days_passed
        
        # Calculate deviation
        deviation = actual_total - expected_value
        deviation_percent = (deviation / expected_value * 100) if expected_value > 0 else 0
        
        # Check if alert is needed
        if deviation_percent <= self.CRITICAL_THRESHOLD:
            alert_level = "critical"
            emoji = "🔴"
        elif deviation_percent <= self.WARNING_THRESHOLD:
            alert_level = "warning"
            emoji = "🟡"
        else:
            return None  # No alert needed
        
        # Format metric name
        metric_names = {
            'calls': 'Перезвоны',
            'issued_volume': 'Выдано (млн)'
        }
        metric_display = metric_names.get(metric, metric)
        
        # Create alert message
        if metric == 'issued_volume':
            message = (
                f"{emoji} {manager_name}: {metric_display}\n"
                f"Факт (выдано): {actual_total:.1f} млн\n"
                f"Ожидалось к дате: {expected_value:.1f} млн\n"
                f"Отклонение: {deviation:+.1f} млн ({deviation_percent:+.1f}%)"
            )
        else:
            message = (
                f"{emoji} {manager_name}: {metric_display}\n"
                f"Факт: {actual_total:,.0f}\n"
                f"Ожидалось: {expected_value:,.0f}\n"
                f"Отклонение: {deviation:+.0f} ({deviation_percent:+.1f}%)"
            )
        
        return TempoAlert(
            manager_name=manager_name,
            metric=metric,
            current_value=actual_total,
            expected_value=expected_value,
            deviation_percent=deviation_percent,
            alert_level=alert_level,
            message=message
        )
    
    def _get_monthly_plans(self, target_date: date) -> Dict[str, Dict[str, float]]:
        """Get monthly plans for all managers (hardcoded for MVP)."""
        # MVP: Use September plans for all months
        return {
            'Бариев': {
                'calls_plan': 0,  # Will be updated when client provides data
                'leads_volume_plan': 70.0
            },
            'Туробов': {
                'calls_plan': 0,
                'leads_volume_plan': 100.0
            },
            'Романченко': {
                'calls_plan': 0,
                'leads_volume_plan': 70.0
            },
            'Шевченко': {
                'calls_plan': 0,
                'leads_volume_plan': 70.0
            },
            'Чертыковцев': {
                'calls_plan': 0,
                'leads_volume_plan': 25.0
            }
        }
    
    def _count_working_days(self, year: int, month: int) -> int:
        """Count working days in a month (Mon-Fri, excluding Russian holidays)."""
        # Get all days in month
        _, last_day = calendar.monthrange(year, month)
        working_days = 0
        
        for day in range(1, last_day + 1):
            current_date = date(year, month, day)
            if self._is_working_day(current_date):
                working_days += 1
        
        return working_days
    
    def _count_working_days_until(self, target_date: date) -> int:
        """Count working days from month start until target date (inclusive)."""
        month_start = date(target_date.year, target_date.month, 1)
        working_days = 0
        
        current_date = month_start
        while current_date <= target_date:
            if self._is_working_day(current_date):
                working_days += 1
            current_date += timedelta(days=1)
        
        return working_days
    
    def _is_working_day(self, check_date: date) -> bool:
        """Check if date is a working day (Mon-Fri, not a Russian holiday)."""
        # Check if weekend
        if check_date.weekday() >= 5:  # Saturday=5, Sunday=6
            return False
        
        # Check Russian holidays (basic list for MVP)
        holidays_2024 = [
            # New Year holidays
            date(2024, 1, 1), date(2024, 1, 2), date(2024, 1, 3),
            date(2024, 1, 4), date(2024, 1, 5), date(2024, 1, 8),
            # Defender of the Fatherland Day
            date(2024, 2, 23),
            # International Women's Day
            date(2024, 3, 8),
            # Labour Day
            date(2024, 5, 1),
            # Victory Day
            date(2024, 5, 9),
            # Russia Day
            date(2024, 6, 12),
            # Unity Day
            date(2024, 11, 4),
            # November holidays (as mentioned by client)
            date(2024, 11, 2), date(2024, 11, 3), date(2024, 11, 4)
        ]
        
        holidays_2025 = [
            # New Year holidays
            date(2025, 1, 1), date(2025, 1, 2), date(2025, 1, 3),
            date(2025, 1, 4), date(2025, 1, 6), date(2025, 1, 7), date(2025, 1, 8),
            # Defender of the Fatherland Day
            date(2025, 2, 23),
            # International Women's Day
            date(2025, 3, 8),
            # Labour Day
            date(2025, 5, 1),
            # Victory Day
            date(2025, 5, 9),
            # Russia Day
            date(2025, 6, 12),
            # Unity Day
            date(2025, 11, 4)
        ]
        
        all_holidays = holidays_2024 + holidays_2025
        
        return check_date not in all_holidays
    
    async def _get_actual_data_for_period(
        self,
        start_date: date,
        end_date: date,
        office_filter: Optional[str] = None,
    ) -> Dict[str, Dict[str, float]]:
        """Get actual data from sheets for the specified period."""
        try:
            # Get all records from sheets
            worksheet = self.sheets_service._reports
            all_records = worksheet.get_all_records()
            
            # Filter and aggregate by manager
            manager_totals = {}
            
            for record in all_records:
                try:
                    record = {str(k).strip().lower(): v for k, v in record.items()}
                    # Filter by office if provided
                    if office_filter:
                        rec_office = str(record.get('office', '')).strip()
                        if rec_office != office_filter:
                            continue

                    # Parse date
                    date_str = str(record.get('date', '')).strip()
                    if not date_str:
                        continue
                    
                    # Try different date formats
                    record_date = None
                    for fmt in ['%Y-%m-%d', '%d.%m.%Y', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d %H:%M:%S', '%d.%m.%Y %H:%M:%S', '%d.%m.%y', '%d.%m.%y %H:%M:%S', '%d-%m-%Y', '%Y.%m.%d']:
                        try:
                            record_date = datetime.strptime(date_str, fmt).date()
                            break
                        except ValueError:
                            continue
                    
                    if record_date is None or record_date < start_date or record_date > end_date:
                        continue
                    
                    manager_name = str(record.get('manager', '')).strip()
                    if not manager_name:
                        continue
                    
                    if manager_name not in manager_totals:
                        manager_totals[manager_name] = {
                            'calls_fact': 0,
                            'issued_volume_fact': 0.0
                        }
                    
                    # Aggregate data
                    manager_totals[manager_name]['calls_fact'] += int(record.get('evening_calls_success', 0) or 0)
                    manager_totals[manager_name]['issued_volume_fact'] += float(record.get('evening_issued_volume', 0) or 0)
                
                except (ValueError, TypeError):
                    continue  # Skip invalid records
            
            return manager_totals
            
        except Exception:
            return {}  # Return empty dict on any error
