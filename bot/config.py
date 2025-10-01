import os
from dataclasses import dataclass
from typing import List


def get_env(name: str, default: str | None = None) -> str:
    value = os.getenv(name, default)
    if value is None:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


@dataclass
class Settings:
    bot_token: str
    spreadsheet_name: str
    google_credentials_path: str
    default_timezone: str
    managers: List[str]
    morning_reminder: str
    evening_reminder: str
    daily_summary_time: str
    yandex_api_key: str
    yandex_folder_id: str
    pptx_font_family: str
    pptx_primary_color: str
    pptx_secondary_color: str
    pptx_logo_path: str
    pptx_emoji_font: str
    drive_folder_id: str
    use_google_slides: bool
    slides_font_family: str
    slides_primary_color: str
    slides_alert_color: str
    slides_accent2_color: str
    slides_text_color: str
    slides_muted_color: str
    slides_card_bg_color: str
    office_name: str
    reminder_quiet_start: str
    reminder_quiet_end: str
    reminder_window_morning: str
    reminder_window_evening: str

    @staticmethod
    def load() -> "Settings":
        bot_token = get_env("BOT_TOKEN")
        spreadsheet_name = get_env("SPREADSHEET_NAME", "Sales Reports")
        google_credentials_path = get_env("GOOGLE_APPLICATION_CREDENTIALS")
        default_timezone = get_env("DEFAULT_TIMEZONE", "Europe/Moscow")
        managers_raw = get_env("MANAGERS", "Бариев,Туробов,Романченко,Шевченко,Чертыковцев")
        managers = [m.strip() for m in managers_raw.split(",") if m.strip()]
        morning_reminder = get_env("MORNING_REMINDER", "09:30")
        evening_reminder = get_env("EVENING_REMINDER", "17:30")
        daily_summary_time = get_env("DAILY_SUMMARY_TIME", "20:30")
        yandex_api_key = get_env("YANDEX_API_KEY", "")
        yandex_folder_id = get_env("YANDEX_FOLDER_ID", "")
        pptx_font_family = get_env("PPTX_FONT_FAMILY", "Montserrat")
        pptx_primary_color = get_env("PPTX_PRIMARY_COLOR", "#CC0000")
        pptx_secondary_color = get_env("PPTX_SECONDARY_COLOR", "#F3F4F6")
        pptx_logo_path = get_env("PPTX_LOGO_PATH", "")
        pptx_emoji_font = get_env("PPTX_EMOJI_FONT", "Segoe UI Emoji")
        drive_folder_id = get_env("DRIVE_FOLDER_ID", "")
        use_google_slides = get_env("USE_GOOGLE_SLIDES", "false").lower() in ("1", "true", "yes")
        slides_font_family = get_env("SLIDES_FONT_FAMILY", "Roboto")
        slides_primary_color = get_env("SLIDES_PRIMARY_COLOR", "#2E7D32")
        slides_alert_color = get_env("SLIDES_ALERT_COLOR", "#C62828")
        slides_accent2_color = get_env("SLIDES_ACCENT2_COLOR", "#FF8A65")
        slides_text_color = get_env("SLIDES_TEXT_COLOR", "#222222")
        slides_muted_color = get_env("SLIDES_MUTED_COLOR", "#6B6B6B")
        slides_card_bg_color = get_env("SLIDES_CARD_BG_COLOR", "#F5F5F5")
        office_name = get_env("OFFICE_NAME", "Банковские гарантии")
        reminder_quiet_start = get_env("REMINDER_QUIET_START", "22:00")
        reminder_quiet_end = get_env("REMINDER_QUIET_END", "08:00")
        reminder_window_morning = get_env("REMINDER_WINDOW_MORNING", "09:00-12:00")
        reminder_window_evening = get_env("REMINDER_WINDOW_EVENING", "17:00-20:00")
        return Settings(
            bot_token=bot_token,
            spreadsheet_name=spreadsheet_name,
            google_credentials_path=google_credentials_path,
            default_timezone=default_timezone,
            managers=managers,
            morning_reminder=morning_reminder,
            evening_reminder=evening_reminder,
            daily_summary_time=daily_summary_time,
            yandex_api_key=yandex_api_key,
            yandex_folder_id=yandex_folder_id,
            pptx_font_family=pptx_font_family,
            pptx_primary_color=pptx_primary_color,
            pptx_secondary_color=pptx_secondary_color,
            pptx_logo_path=pptx_logo_path,
            pptx_emoji_font=pptx_emoji_font,
            drive_folder_id=drive_folder_id,
            use_google_slides=use_google_slides,
            slides_font_family=slides_font_family,
            slides_primary_color=slides_primary_color,
            slides_alert_color=slides_alert_color,
            slides_accent2_color=slides_accent2_color,
            slides_text_color=slides_text_color,
            slides_muted_color=slides_muted_color,
            slides_card_bg_color=slides_card_bg_color,
            office_name=office_name,
            reminder_quiet_start=reminder_quiet_start,
            reminder_quiet_end=reminder_quiet_end,
            reminder_window_morning=reminder_window_morning,
            reminder_window_evening=reminder_window_evening,
        )
