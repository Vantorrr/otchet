#!/bin/bash

echo "🚀 Запуск Telegram бота для отчетности..."
echo ""

# Проверяем наличие виртуального окружения
if [ ! -d ".venv" ]; then
    echo "❌ Виртуальное окружение не найдено!"
    echo "Запустите: python3 -m venv .venv"
    exit 1
fi

# Проверяем наличие .env
if [ ! -f ".env" ]; then
    echo "❌ Файл .env не найден!"
    echo "Создайте файл .env с настройками"
    exit 1
fi

# Проверяем наличие service_account.json
if [ ! -f "service_account.json" ]; then
    echo "❌ Файл service_account.json не найден!"
    echo "Скачайте JSON ключ из Google Cloud Console"
    echo "и поместите его как service_account.json"
    echo ""
    echo "Инструкция в файле: create_spreadsheet_guide.md"
    exit 1
fi

echo "✅ Все файлы на месте"
echo "🔄 Активация виртуального окружения..."

# Активируем виртуальное окружение
source .venv/bin/activate

echo "🔄 Запуск бота..."
echo "Для остановки нажмите Ctrl+C"
echo ""

# Запускаем бота
python -m bot.main
