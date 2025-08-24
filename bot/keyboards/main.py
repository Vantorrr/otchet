from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup


def get_main_menu_keyboard() -> InlineKeyboardMarkup:
    """Главное меню для менеджера"""
    keyboard = [
        [InlineKeyboardButton(text="🌅 Утренний отчет", callback_data="morning_report")],
        [InlineKeyboardButton(text="🌆 Вечерний отчет", callback_data="evening_report")],
    ]
    return InlineKeyboardMarkup(inline_keyboard=keyboard)


def get_admin_menu_keyboard() -> InlineKeyboardMarkup:
    """Меню для администратора"""
    keyboard = [
        [InlineKeyboardButton(text="📊 Сводка за сегодня", callback_data="summary_today")],
        [InlineKeyboardButton(text="📆 Сводка: неделя", callback_data="summary_week")],
        [InlineKeyboardButton(text="🗓️ Сводка: месяц", callback_data="summary_month")],
        [InlineKeyboardButton(text="📣 Сводка: квартал", callback_data="summary_quarter")],
        [InlineKeyboardButton(text="📋 Сводка за дату", callback_data="summary_date")],
        [InlineKeyboardButton(text="⚙️ Настроить тему", callback_data="setup_topic")],
    ]
    return InlineKeyboardMarkup(inline_keyboard=keyboard)
