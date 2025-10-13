"""Data aggregation service for presentation generation."""
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple
from collections import defaultdict

from bot.services.sheets import SheetsClient
from bot.services.presentation import ManagerData
from bot.services.di import Container
from bot.utils.time_utils import (
    start_end_of_week_today,
    start_end_of_month_today,
    start_end_of_quarter_today,
)


class DataAggregatorService:
    """Service for aggregating data from Google Sheets for presentations."""
    
    def __init__(self, sheets_service: SheetsClient):
        self.sheets_service = sheets_service
    
    async def aggregate_weekly_data(self, target_date: Optional[date] = None) -> Tuple[Dict[str, ManagerData], str, date, date]:
        """Aggregate data for a week."""
        # Use configured timezone-based helpers for current period
        start_date_str, end_date_str = start_end_of_week_today(Container.get().settings)
        from datetime import datetime as _dt
        start_date = _dt.strptime(start_date_str, "%Y-%m-%d").date()
        end_date = _dt.strptime(end_date_str, "%Y-%m-%d").date()
        period_name = f"Неделя {start_date.strftime('%d.%m')}—{end_date.strftime('%d.%m.%Y')}"
        
        data = await self._aggregate_data_for_period(start_date, end_date)
        return data, period_name, start_date, end_date

    async def aggregate_weekly_data_with_previous(self, office_filter: Optional[str] = None) -> Tuple[Dict[str, ManagerData], Dict[str, ManagerData], str, date, date, date, date]:
        """Aggregate current week and previous week data."""
        current, period_name, start_date, end_date = await self.aggregate_weekly_data()
        prev_end = start_date - timedelta(days=1)
        prev_start = prev_end - timedelta(days=6)
        # Apply office filter to both current and previous periods
        current = await self._aggregate_data_for_period(start_date, end_date, office_filter=office_filter)
        previous = await self._aggregate_data_for_period(prev_start, prev_end, office_filter=office_filter)
        return current, previous, period_name, start_date, end_date, prev_start, prev_end
    
    async def aggregate_monthly_data(self, target_date: Optional[date] = None) -> Tuple[Dict[str, ManagerData], str, date, date]:
        """Aggregate data for a month."""
        start_date_str, end_date_str = start_end_of_month_today(Container.get().settings)
        from datetime import datetime as _dt
        start_date = _dt.strptime(start_date_str, "%Y-%m-%d").date()
        end_date = _dt.strptime(end_date_str, "%Y-%m-%d").date()
        period_name = f"Месяц {start_date.strftime('%B %Y')}"
        
        # Translate month names to Russian
        month_names = {
            'January': 'Январь', 'February': 'Февраль', 'March': 'Март',
            'April': 'Апрель', 'May': 'Май', 'June': 'Июнь',
            'July': 'Июль', 'August': 'Август', 'September': 'Сентябрь',
            'October': 'Октябрь', 'November': 'Ноябрь', 'December': 'Декабрь'
        }
        for en_name, ru_name in month_names.items():
            period_name = period_name.replace(en_name, ru_name)
        
        data = await self._aggregate_data_for_period(start_date, end_date)
        return data, period_name, start_date, end_date

    async def aggregate_monthly_data_with_previous(self, office_filter: Optional[str] = None) -> Tuple[Dict[str, ManagerData], Dict[str, ManagerData], str, date, date, date, date]:
        """Aggregate current month and previous month data."""
        current, period_name, start_date, end_date = await self.aggregate_monthly_data()
        # previous month range
        from datetime import date as _date
        prev_end = start_date - timedelta(days=1)
        prev_start = _date(prev_end.year, prev_end.month, 1)
        # Apply office filter to both current and previous periods
        current = await self._aggregate_data_for_period(start_date, end_date, office_filter=office_filter)
        previous = await self._aggregate_data_for_period(prev_start, prev_end, office_filter=office_filter)
        return current, previous, period_name, start_date, end_date, prev_start, prev_end
    
    async def aggregate_quarterly_data(self, target_date: Optional[date] = None) -> Tuple[Dict[str, ManagerData], str, date, date]:
        """Aggregate data for a quarter."""
        start_date_str, end_date_str = start_end_of_quarter_today(Container.get().settings)
        from datetime import datetime as _dt
        start_date = _dt.strptime(start_date_str, "%Y-%m-%d").date()
        end_date = _dt.strptime(end_date_str, "%Y-%m-%d").date()
        # Derive quarter by start_date to avoid None target_date
        quarter_num = (start_date.month - 1) // 3 + 1
        period_name = f"{quarter_num} квартал {start_date.year}"
        
        data = await self._aggregate_data_for_period(start_date, end_date)
        return data, period_name, start_date, end_date

    async def aggregate_quarterly_data_with_previous(self, office_filter: Optional[str] = None) -> Tuple[Dict[str, ManagerData], Dict[str, ManagerData], str, date, date, date, date]:
        """Aggregate current quarter and previous quarter data."""
        current, period_name, start_date, end_date = await self.aggregate_quarterly_data()
        prev_end = start_date - timedelta(days=1)
        # compute start of that quarter
        q = (prev_end.month - 1) // 3 + 1
        start_month = (q - 1) * 3 + 1
        from datetime import date as _date
        prev_start = _date(prev_end.year, start_month, 1)
        # Apply office filter to both current and previous periods
        current = await self._aggregate_data_for_period(start_date, end_date, office_filter=office_filter)
        previous = await self._aggregate_data_for_period(prev_start, prev_end, office_filter=office_filter)
        return current, previous, period_name, start_date, end_date, prev_start, prev_end
    
    async def _aggregate_data_for_period(self, start_date: date, end_date: date, office_filter: Optional[str] = None) -> Dict[str, ManagerData]:
        """Aggregate data for a specific period, optionally filtered by office."""
        try:
            # Get all records from sheets
            worksheet = self.sheets_service._reports
            all_records = worksheet.get_all_records()
            
            # Initialize manager data
            manager_data = defaultdict(lambda: ManagerData(name=""))
            
            # Process each record
            for record in all_records:
                try:
                    # Normalize keys to lowercase to be robust to header casing
                    record = {str(k).strip().lower(): v for k, v in record.items()}
                    
                    # Filter by office if specified
                    if office_filter:
                        rec_office = str(record.get('office', '')).strip()
                        if rec_office != office_filter:
                            continue
                    
                    # Parse and validate date
                    date_str = str(record.get('date', '')).strip()
                    if not date_str:
                        continue
                    
                    record_date = self._parse_record_date(date_str)
                    if record_date is None or record_date < start_date or record_date > end_date:
                        continue
                    
                    # Get manager name
                    manager_name = str(record.get('manager', '')).strip()
                    if not manager_name:
                        continue
                    
                    # Initialize manager data if needed
                    if manager_data[manager_name].name == "":
                        manager_data[manager_name].name = manager_name
                    
                    # Aggregate morning data
                    self._add_morning_data(manager_data[manager_name], record)
                    
                    # Aggregate evening data
                    self._add_evening_data(manager_data[manager_name], record)
                
                except (ValueError, TypeError) as e:
                    continue  # Skip invalid records
            
            # Convert defaultdict to regular dict and remove empty entries
            result = {}
            for name, data in manager_data.items():
                if data.name:  # Only include managers with actual data
                    result[name] = data
            
            return result
            
        except Exception as e:
            # Return empty dict on any error
            return {}

    async def get_daily_series(self, start_date: date, end_date: date, office_filter: Optional[str] = None) -> List[Dict[str, float]]:
        """Aggregate daily totals for the given period for charts.
        Returns list of dicts sorted by date with keys: date, calls_fact, new_calls, leads_units_fact,
        leads_volume_plan, leads_volume_fact, approved_volume, issued_volume.
        """
        try:
            worksheet = self.sheets_service._reports
            all_records = worksheet.get_all_records()
            from collections import defaultdict
            buckets: Dict[str, Dict[str, float]] = defaultdict(lambda: {
                'calls_fact': 0,
                'new_calls': 0,
                'leads_units_fact': 0,
                'leads_volume_plan': 0.0,
                'leads_volume_fact': 0.0,
                'approved_volume': 0.0,
                'issued_volume': 0.0,
            })

            for record in all_records:
                record = {str(k).strip().lower(): v for k, v in record.items()}
                # Filter by office if provided
                if office_filter:
                    rec_office = str(record.get('office', '')).strip()
                    if rec_office != office_filter:
                        continue
                date_str = str(record.get('date', '')).strip()
                if not date_str:
                    continue
                d = self._parse_record_date(date_str)
                if d is None or d < start_date or d > end_date:
                    continue
                key = d.strftime('%Y-%m-%d')
                try:
                    buckets[key]['calls_fact'] += int(record.get('evening_calls_success', 0) or 0)
                    buckets[key]['new_calls'] += int(record.get('evening_new_calls', 0) or 0)
                    buckets[key]['leads_units_fact'] += int(record.get('evening_leads_units', 0) or 0)
                    buckets[key]['leads_volume_plan'] += float(record.get('morning_leads_planned_volume', 0) or 0)
                    buckets[key]['leads_volume_fact'] += float(record.get('evening_leads_volume', 0) or 0)
                    buckets[key]['approved_volume'] += float(record.get('evening_approved_volume', 0) or 0)
                    buckets[key]['issued_volume'] += float(record.get('evening_issued_volume', 0) or 0)
                except (ValueError, TypeError):
                    continue

            # Sort by date and render list
            items: List[Dict[str, float]] = []
            cur = start_date
            while cur <= end_date:
                k = cur.strftime('%Y-%m-%d')
                v = buckets.get(k, {
                    'calls_fact': 0,
                    'new_calls': 0,
                    'leads_units_fact': 0,
                    'leads_volume_plan': 0.0,
                    'leads_volume_fact': 0.0,
                    'approved_volume': 0.0,
                    'issued_volume': 0.0,
                })
                items.append({'date': k, **v})
                cur = cur + timedelta(days=1)
            return items
        except Exception:
            return []

    async def aggregate_custom_with_previous(self, start_date: date, end_date: date) -> Tuple[Dict[str, ManagerData], Dict[str, ManagerData], str, date, date, date, date]:
        """Aggregate for custom [start_date, end_date] and previous equal-length period."""
        data = await self._aggregate_data_for_period(start_date, end_date)
        period_name = f"Период {start_date.strftime('%d.%m.%Y')}—{end_date.strftime('%d.%m.%Y')}"
        # previous range of same length
        delta = end_date - start_date
        prev_end = start_date - timedelta(days=1)
        prev_start = prev_end - delta
        prev_data = await self._aggregate_data_for_period(prev_start, prev_end)
        return data, prev_data, period_name, start_date, end_date, prev_start, prev_end

    async def aggregate_two_periods(
        self,
        start_a: date,
        end_a: date,
        start_b: date,
        end_b: date,
    ) -> Tuple[Dict[str, ManagerData], Dict[str, ManagerData], str, date, date, date, date]:
        """Aggregate for two arbitrary periods A and B, to compare A vs B."""
        data_a = await self._aggregate_data_for_period(start_a, end_a)
        data_b = await self._aggregate_data_for_period(start_b, end_b)
        title = f"Период {start_a.strftime('%d.%m.%Y')}—{end_a.strftime('%d.%m.%Y')} vs {start_b.strftime('%d.%m.%Y')}—{end_b.strftime('%d.%m.%Y')}"
        return data_a, data_b, title, start_a, end_a, start_b, end_b
    
    def _parse_record_date(self, date_str: str) -> Optional[date]:
        """Parse date from various formats."""
        from datetime import datetime
        
        # Handle Google Sheets serial number format
        try:
            if date_str.replace('.', '').isdigit():
                serial_number = float(date_str)
                if serial_number > 40000:  # Reasonable range for Excel/Sheets dates
                    # Excel epoch: January 1, 1900 (but actually 1899-12-30 due to bug)
                    excel_epoch = date(1899, 12, 30)
                    return excel_epoch + timedelta(days=int(serial_number))
        except (ValueError, TypeError):
            pass
        
        # Try common date formats
        formats = [
            '%Y-%m-%d', '%d.%m.%Y', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y',
            '%Y-%m-%d %H:%M:%S', '%d.%m.%Y %H:%M:%S',
            '%d.%m.%y', '%d.%m.%y %H:%M:%S', '%d-%m-%Y', '%Y.%m.%d'
        ]
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
        
        return None
    
    def _add_morning_data(self, manager_data: ManagerData, record: Dict):
        """Add morning report data to manager data."""
        try:
            manager_data.calls_plan += int(record.get('morning_calls_planned', 0) or 0)
            manager_data.leads_units_plan += int(record.get('morning_leads_planned_units', 0) or 0)
            manager_data.leads_volume_plan += float(record.get('morning_leads_planned_volume', 0) or 0)
            manager_data.new_calls_plan += int(record.get('morning_new_calls_planned', 0) or 0)
        except (ValueError, TypeError):
            pass  # Skip invalid data
    
    def _add_evening_data(self, manager_data: ManagerData, record: Dict):
        """Add evening report data to manager data."""
        try:
            manager_data.calls_fact += int(record.get('evening_calls_success', 0) or 0)
            manager_data.leads_units_fact += int(record.get('evening_leads_units', 0) or 0)
            manager_data.leads_volume_fact += float(record.get('evening_leads_volume', 0) or 0)
            manager_data.approved_volume += float(record.get('evening_approved_volume', 0) or 0)
            manager_data.issued_volume += float(record.get('evening_issued_volume', 0) or 0)
            manager_data.new_calls += int(record.get('evening_new_calls', 0) or 0)
        except (ValueError, TypeError):
            pass  # Skip invalid data
