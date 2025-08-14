#!/usr/bin/env python3
"""
Проверка настройки Google Sheets
"""

import os
from dotenv import load_dotenv

def check_setup():
    load_dotenv()
    
    print("🔍 Проверка настройки...")
    
    # Проверяем .env
    bot_token = os.getenv("BOT_TOKEN")
    if bot_token:
        print(f"✅ BOT_TOKEN: {'*' * 20}{bot_token[-10:]}")
    else:
        print("❌ BOT_TOKEN не найден в .env")
        return False
    
    # Проверяем service_account.json
    if os.path.exists("service_account.json"):
        print("✅ service_account.json найден")
        try:
            with open("service_account.json", "r") as f:
                import json
                data = json.load(f)
                client_email = data.get("client_email", "")
                if client_email and "@" in client_email:
                    print(f"✅ Service Account Email: {client_email}")
                    print("\n📋 Следующие шаги:")
                    print("1. Создай Google Таблицу 'Sales Reports'")
                    print(f"2. Поделись с: {client_email}")
                    print("3. Дай права 'Редактор'")
                    print("4. Запусти: ./start_bot.sh")
                    return True
                else:
                    print("❌ Некорректный service_account.json")
                    return False
        except Exception as e:
            print(f"❌ Ошибка чтения service_account.json: {e}")
            return False
    else:
        print("❌ service_account.json не найден")
        print("Скачай JSON ключ из Google Cloud Console")
        return False

if __name__ == "__main__":
    if check_setup():
        print("\n🎉 Настройка почти готова!")
    else:
        print("\n❌ Нужно исправить ошибки")
