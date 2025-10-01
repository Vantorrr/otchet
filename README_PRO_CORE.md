# 🚀 PRO CORE — Инструкция по развертыванию

## ✅ Что включено в Pro Core:
- ✅ **OpenAI gpt-5-nano** для AI-комментариев (точнее, без обрезаний)
- ✅ **Светофор KPI**: зеленый ≥90%, желтый ≥70%, красный <70%
- ✅ **Средний менеджер**: базовая линия для сравнения
- ✅ **Гибкие напоминания**: тихие часы (22:00-08:00) и окна отправки
- ✅ **Google Slides + PDF**: генерация отчетов прямо в Drive

---

## 📋 Настройка на сервере (5 минут)

### 1. Обнови код:
```bash
cd /root/TGbot
git pull origin main
```

### 2. Установи новые зависимости:
```bash
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. Обнови `.env` — добавь эти строки:
```bash
# Pro Core - OpenAI
OPENAI_API_KEY=<твой_openai_ключ>
OPENAI_MODEL=gpt-5-nano

# Pro Core - Google Slides
DRIVE_FOLDER_ID=1B2h3QufJcciSxemkEUgxb29CmKY3cKBC
USE_GOOGLE_SLIDES=true

# Pro Core - гибкие напоминания (опционально)
REMINDER_QUIET_START=22:00
REMINDER_QUIET_END=08:00
REMINDER_WINDOW_MORNING=09:00-12:00
REMINDER_WINDOW_EVENING=17:00-20:00
```

### 4. Залей OAuth токен для Slides (если еще не залит):
```bash
# На локалке (MacBook):
cd /Users/pavelgalante/TGbot
python3 scripts/google_oauth_setup.py
# Авторизуйся под аккаунтом, который владеет папкой на Drive

# Скопируй токен на сервер:
scp scripts/token.json root@46.149.66.158:/root/TGbot/scripts/token.json
```

### 5. Перезапусти бота:
```bash
systemctl restart tgbot
systemctl status tgbot
```

---

## 🧪 Тестирование Pro Core

### 1. Проверь, что OpenAI работает:
```bash
curl -X POST https://api.openai.com/v1/chat/completions \
  -H "Authorization: Bearer <твой_openai_ключ>" \
  -H "Content-Type: application/json" \
  -d '{"model":"gpt-4o-mini","messages":[{"role":"user","content":"ping"}]}'
```
Ожидается: `{"id":"chatcmpl-...","choices":[{"message":{"content":"...` — любой ответ означает, что ключ валиден.

### 2. Сгенерируй тестовый отчет в Slides:
В Telegram (админ-группа или личка с ботом):
```
/slides_range 2025-09-01 2025-09-07
```
Бот должен:
- Создать презентацию в папке `https://drive.google.com/drive/folders/1B2h3QufJcciSxemkEUgxb29CmKY3cKBC`
- Вернуть PDF-файл в чат
- Добавить AI-комментарии (от OpenAI)
- Применить светофор к конверсиям

---

## 📊 Что изменилось в отчетах:

### PPTX (локально):
- ✅ Светофор KPI в сводке
- ✅ Строка "📊 Средний менеджер" под таблицей команды

### Google Slides (продакшн):
- ✅ Светофор KPI (цветные проценты конверсий)
- ✅ AI-комментарии от OpenAI (короче, точнее, без обрезаний)
- ✅ Средний менеджер (будет добавлен в следующей итерации Slides)

---

## 🔧 Troubleshooting

### Ошибка: `HttpError 403: The caller does not have permission`
**Решение:** Дай права на папку Drive:
- Открой https://drive.google.com/drive/folders/1B2h3QufJcciSxemkEUgxb29CmKY3cKBC
- Нажми "Настройка доступа" → добавь email из `scripts/token.json` или `service_account.json` → "Редактор"

### Ошибка: `ModuleNotFoundError: No module named 'openai'`
**Решение:**
```bash
source .venv/bin/activate
pip install openai==1.43.0
```

### Напоминания не отправляются в тихие часы:
**Ожидается:** С 22:00 до 08:00 (по умолчанию) бот НЕ шлет напоминания. Это фича Pro Core.

---

## 📞 Контакты
Вопросы/баги → @noface_digital или Telegram-чат проекта.

---

**Готово! Про версия запущена. Сдавай клиенту 🚀**

