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
        )
