#!/usr/bin/env python3

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import os
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

from bot.config import Settings
from bot.services.sheets import SheetsClient
from bot.utils.time_utils import start_end_of_week_today
from bot.services.summary_builder import build_summary_text

def main():
    settings = Settings.load()
    sheets = SheetsClient(settings)
    
    print("🔍 Тестирование недельной сводки...")
    
    # Получаем даты недели
    start, end = start_end_of_week_today(settings)
    print(f"📅 Период недели: {start} — {end}")
    
    # Получаем все записи из таблицы
    try:
        all_records = sheets._reports.get_all_records()
        print(f"📊 Всего записей в таблице: {len(all_records)}")
        
        # Показываем первые 3 записи для проверки формата дат
        if all_records:
            print("\n📋 Примеры записей:")
            for i, record in enumerate(all_records[:3]):
                date_value = record.get("date", "НЕТ ДАТЫ")
                manager = record.get("manager", "НЕТ МЕНЕДЖЕРА")
                print(f"  {i+1}. Дата: '{date_value}' | Менеджер: '{manager}'")
        
        # Фильтруем записи по периоду недели
        from bot.services.summary_builder import _within
        filtered_records = [r for r in all_records if _within(r.get("date"), start, end)]
        print(f"📈 Записей за текущую неделю: {len(filtered_records)}")
        
        if filtered_records:
            print("\n✅ Записи за неделю:")
            for i, record in enumerate(filtered_records):
                date_value = record.get("date", "НЕТ ДАТЫ")
                manager = record.get("manager", "НЕТ МЕНЕДЖЕРА")
                print(f"  {i+1}. Дата: '{date_value}' | Менеджер: '{manager}'")
        else:
            print("❌ Нет записей за текущую неделю!")
            print("💡 Возможные причины:")
            print("   - Нет данных за этот период")
            print("   - Неправильный формат дат в таблице")
            print("   - Проблема с фильтрацией дат")
        
        # Пытаемся построить сводку
        print("\n🏗️ Построение сводки...")
        summary_text = build_summary_text(settings, sheets, day=start, start=start, end=end)
        print(f"📝 Длина текста сводки: {len(summary_text)} символов")
        
        if len(summary_text) > 100:
            print("✅ Сводка построена успешно!")
            print("\n📊 Первые 200 символов сводки:")
            print(summary_text[:200] + "...")
        else:
            print("❌ Сводка слишком короткая, возможно ошибка")
            print("📄 Полный текст сводки:")
            print(summary_text)
            
    except Exception as e:
        print(f"❌ Ошибка при работе с таблицей: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
