from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup


def get_main_menu_keyboard() -> InlineKeyboardMarkup:
    """Главное меню для менеджера"""
    keyboard = [
        [InlineKeyboardButton(text="🌅 Утренний отчет", callback_data="morning_report")],
        [InlineKeyboardButton(text="🌆 Вечерний отчет", callback_data="evening_report")],
    ]
    return InlineKeyboardMarkup(inline_keyboard=keyboard)


def get_admin_menu_keyboard(is_hq: bool = False) -> InlineKeyboardMarkup:
    """Корневое меню администратора (с дополнительным меню для HQ)."""
    keyboard = [
        [InlineKeyboardButton(text="📊 Сводки", callback_data="admin_section_summaries")],
        [InlineKeyboardButton(text="🤖 AI-Презентации", callback_data="admin_section_ai")],
        [InlineKeyboardButton(text="📅 Отчёт по дате", callback_data="admin_report_by_date")],
    ]
    if is_hq:
        keyboard.append([InlineKeyboardButton(text="🏢 Управление офисами", callback_data="admin_section_offices")])
        # Глобальные напоминания (только в HQ)
        keyboard.append([InlineKeyboardButton(text="🌅 Напомнить всем (утро)", callback_data="admin_remind_all_morning")])
        keyboard.append([InlineKeyboardButton(text="🌆 Напомнить всем (вечер)", callback_data="admin_remind_all_evening")])
    else:
        # Напоминания доступны только в офисных чатах, не в HQ
        keyboard.append([InlineKeyboardButton(text="🌅 Утреннее напоминание", callback_data="admin_remind_morning")])
        keyboard.append([InlineKeyboardButton(text="🌆 Вечернее напоминание", callback_data="admin_remind_evening")])
    keyboard.append([InlineKeyboardButton(text="⚙️ Настроить тему", callback_data="setup_topic")])
    return InlineKeyboardMarkup(inline_keyboard=keyboard)


def get_admin_offices_keyboard() -> InlineKeyboardMarkup:
    """Подменю: управление офисами (только для HQ)."""
    keyboard = [
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="admin_back")],
        [InlineKeyboardButton(text="📊 Сводка: Все офисы", callback_data="summary_all_offices")],
        [InlineKeyboardButton(text="📊 Офис 4", callback_data="summary_office4")],
        [InlineKeyboardButton(text="📊 Санжаровский", callback_data="summary_sanzharovsky")],
        [InlineKeyboardButton(text="📊 Батурлов", callback_data="summary_baturlov")],
        [InlineKeyboardButton(text="📊 Савела", callback_data="summary_savela")],
        [InlineKeyboardButton(text="🤖 Презентация: Все офисы", callback_data="presentation_all_offices")],
        [InlineKeyboardButton(text="🤖 Презентация: Офис 4", callback_data="presentation_office4")],
        [InlineKeyboardButton(text="🤖 Презентация: Санжаровский", callback_data="presentation_sanzharovsky")],
        [InlineKeyboardButton(text="🤖 Презентация: Батурлов", callback_data="presentation_baturlov")],
        [InlineKeyboardButton(text="🤖 Презентация: Савела", callback_data="presentation_savela")],
        [InlineKeyboardButton(text="📈 Сравнить офисы", callback_data="compare_offices")],
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
