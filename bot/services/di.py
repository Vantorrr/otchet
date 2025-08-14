from __future__ import annotations

from typing import Optional

from bot.config import Settings
from bot.services.sheets import SheetsClient


class Container:
    _instance: Optional["Container"] = None

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.sheets = SheetsClient(settings)

    @classmethod
    def init(cls, settings: Settings) -> "Container":
        if cls._instance is None:
            cls._instance = Container(settings)
        return cls._instance

    @classmethod
    def get(cls) -> "Container":
        if cls._instance is None:
            raise RuntimeError("Container is not initialized")
        return cls._instance
