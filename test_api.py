#!/usr/bin/env python3
"""
Тест подключения к Google APIs
"""

import gspread
from dotenv import load_dotenv
import os

def test_apis():
    load_dotenv()
    
    try:
        print("🔄 Тестируем подключение к Google APIs...")
        
        # Подключаемся к Google
        gc = gspread.service_account(filename="service_account.json")
        print("✅ Подключение к Google APIs успешно")
        
        # Пробуем найти таблицу
        try:
            spreadsheet = gc.open("Sales Reports")
            print("✅ Таблица 'Sales Reports' найдена")
            print(f"📊 URL таблицы: {spreadsheet.url}")
            return True
        except gspread.SpreadsheetNotFound:
            print("❌ Таблица 'Sales Reports' не найдена")
            print("Создай таблицу на https://sheets.google.com/")
            print("И дай доступ: swapcoon-sheets@shum-47422.iam.gserviceaccount.com")
            return False
        except Exception as e:
            print(f"❌ Ошибка доступа к таблице: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Ошибка подключения к Google APIs: {e}")
        return False

if __name__ == "__main__":
    if test_apis():
        print("\n🎉 Все готово для запуска бота!")
    else:
        print("\n❌ Нужно исправить проблемы выше")
