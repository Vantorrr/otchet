# Настройка Google Service Account

## Шаги для создания Service Account:

1. **Перейдите в Google Cloud Console**: https://console.cloud.google.com/

2. **Создайте новый проект или выберите существующий**

3. **Включите Google Sheets API**:
   - Перейдите в "APIs & Services" > "Library"
   - Найдите "Google Sheets API"
   - Нажмите "Enable"

4. **Создайте Service Account**:
   - Перейдите в "APIs & Services" > "Credentials"
   - Нажмите "Create Credentials" > "Service Account"
   - Введите название (например, "telegram-bot")
   - Нажмите "Create and Continue"
   - Роль оставьте пустой или выберите "Editor"
   - Нажмите "Done"

5. **Создайте ключ**:
   - Найдите созданный Service Account в списке
   - Нажмите на него
   - Перейдите на вкладку "Keys"
   - Нажмите "Add Key" > "Create new key"
   - Выберите JSON формат
   - Скачайте файл

6. **Настройте проект**:
   - Переименуйте скачанный файл в `service_account.json`
   - Поместите его в папку `/Users/pavelgalante/TGbot/`

7. **Создайте Google Таблицу**:
   - Создайте новую Google Таблицу с названием "Sales Reports"
   - Поделитесь таблицей с email из service_account.json (client_email)
   - Дайте права "Editor"

## Готово!
После этого бот сможет работать с Google Sheets.
