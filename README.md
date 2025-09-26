## Telegram Reporting Bot

Collects morning/evening reports from managers in a Telegram group with topics, writes to Google Sheets, and posts a daily summary.

### Features
- Morning and evening report flows via guided prompts
- Topic-to-manager binding, summary topic binding
- Data stored in Google Sheets (can later swap to Yandex Tables)
- Manual summary command by day

### Setup
1. Create a Google Cloud project and a Service Account with Google Sheets API enabled. Download the JSON key.
2. Share your target spreadsheet with the service account email (Editor access).
3. Copy `.env.example` to `.env` and fill values. Use absolute path for `GOOGLE_APPLICATION_CREDENTIALS`.
4. Install deps:
```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```
5. Run the bot:
```bash
python -m bot.main
```

### Telegram Group Setup
- Add the bot to a supergroup with topics enabled.
- Create topics: one per manager, plus one summary topic.
- In each manager topic run `/bind_manager <ФИО>`.
- In the summary topic run `/set_summary_topic`.

### Commands
- `/morning` — start morning report (перезвоны план, заявки план: штуки и объем)
- `/evening` — start evening report (перезвоны успешно, заявки заведено: штуки и объем, новые дозвоны)
- `/summary [YYYY-MM-DD]` — post summary for the day (default today)
- `/bind_manager <ФИО>` — bind current topic to a manager
- `/set_summary_topic` — bind current topic as the summary topic

### Notes
- Default timezone: `Europe/Moscow` (configurable via env)
- Initial managers: Бариев, Туробов, Романченко, Шевченко, Чертыковцев

---

### Credits
Разработано командой N0FACE — Digital Legends | [noface.digital](https://noface.digital)

Контакт: Telegram `@pavel_xdev`

## Slides (про-версия)

Добавьте в `.env` переменные бренда и папку Drive:

```env
USE_GOOGLE_SLIDES=true
DRIVE_FOLDER_ID=<your_drive_folder_id>
SLIDES_FONT_FAMILY=Roboto
SLIDES_PRIMARY_COLOR=#2E7D32
SLIDES_ALERT_COLOR=#C62828
SLIDES_ACCENT2_COLOR=#FF8A65
SLIDES_TEXT_COLOR=#222222
SLIDES_MUTED_COLOR=#6B6B6B
SLIDES_CARD_BG_COLOR=#F5F5F5
```

После изменения `.env` перезапустите бота и используйте команду `/slides_range YYYY-MM-DD YYYY-MM-DD`.
