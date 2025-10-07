from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup


def get_main_menu_keyboard() -> InlineKeyboardMarkup:
    """Главное меню для менеджера"""
    keyboard = [
        [InlineKeyboardButton(text="🌅 Утренний отчет", callback_data="morning_report")],
        [InlineKeyboardButton(text="🌆 Вечерний отчет", callback_data="evening_report")],
    ]
    return InlineKeyboardMarkup(inline_keyboard=keyboard)


def get_admin_menu_keyboard() -> InlineKeyboardMarkup:
    """Корневое меню администратора (две секции)."""
    keyboard = [
        [InlineKeyboardButton(text="📊 Сводки", callback_data="admin_section_summaries")],
        [InlineKeyboardButton(text="🤖 AI-Презентации", callback_data="admin_section_ai")],
        [InlineKeyboardButton(text="📅 Отчёт по дате", callback_data="admin_report_by_date")],
        [InlineKeyboardButton(text="🌅 Утреннее напоминание", callback_data="admin_remind_morning")],
        [InlineKeyboardButton(text="🌆 Вечернее напоминание", callback_data="admin_remind_evening")],
        [InlineKeyboardButton(text="⚙️ Настроить тему", callback_data="setup_topic")],
    ]
    return InlineKeyboardMarkup(inline_keyboard=keyboard)


def get_admin_summaries_keyboard() -> InlineKeyboardMarkup:
    """Подменю: сводки"""
    keyboard = [
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="admin_back")],
        [InlineKeyboardButton(text="📊 Сводка за сегодня", callback_data="summary_today")],
        [InlineKeyboardButton(text="📆 Сводка: неделя", callback_data="summary_week")],
        [InlineKeyboardButton(text="🗓️ Сводка: месяц", callback_data="summary_month")],
        [InlineKeyboardButton(text="📣 Сводка: квартал", callback_data="summary_quarter")],
        [InlineKeyboardButton(text="📅 Сводка: период", callback_data="summary_period")],
        [InlineKeyboardButton(text="📋 Сводка за дату", callback_data="summary_date")],
    ]
    return InlineKeyboardMarkup(inline_keyboard=keyboard)


def get_admin_ai_keyboard() -> InlineKeyboardMarkup:
    """Подменю: AI-презентации/аналитика"""
    keyboard = [
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="admin_back")],
        [InlineKeyboardButton(text="🤖 AI-Презентация: неделя", callback_data="presentation_week")],
        [InlineKeyboardButton(text="🤖 AI-Презентация: месяц", callback_data="presentation_month")],
        [InlineKeyboardButton(text="🤖 AI-Презентация: квартал", callback_data="presentation_quarter")],
        [InlineKeyboardButton(text="🤖 AI-Презентация: период", callback_data="presentation_period")],
        [InlineKeyboardButton(text="💬 Спроси у ИИ", callback_data="ask_ai")],
        [InlineKeyboardButton(text="⚠️ Проверка темпа", callback_data="tempo_check")],
    ]
    return InlineKeyboardMarkup(inline_keyboard=keyboard)
