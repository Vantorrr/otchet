from __future__ import annotations

import io
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.service_account import Credentials

from bot.config import Settings


SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]


@dataclass
class SlidesResources:
    drive: Any
    slides: Any


class GoogleSlidesService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        creds = Credentials.from_service_account_file(settings.google_credentials_path, scopes=SCOPES)
        self._resources = SlidesResources(
            drive=build("drive", "v3", credentials=creds),
            slides=build("slides", "v1", credentials=creds),
        )

    # --- Helpers ---
    def _get_folder_id(self) -> str:
        if not self._settings.drive_folder_id:
            raise RuntimeError("DRIVE_FOLDER_ID is not set in environment")
        return self._settings.drive_folder_id

    # --- Public API ---
    def create_presentation(self, title: str) -> str:
        body = {"title": title}
        pres = self._resources.slides.presentations().create(body=body).execute()
        return pres["presentationId"]

    def move_presentation_to_folder(self, presentation_id: str) -> None:
        folder_id = self._get_folder_id()
        # Get current parents
        file = self._resources.drive.files().get(fileId=presentation_id, fields="parents").execute()
        prev_parents = ",".join(file.get("parents", []))
        self._resources.drive.files().update(
            fileId=presentation_id,
            addParents=folder_id,
            removeParents=prev_parents or None,
            fields="id, parents",
        ).execute()

    def export_pdf(self, presentation_id: str) -> bytes:
        request = self._resources.drive.files().export_media(fileId=presentation_id, mimeType="application/pdf")
        buf = io.BytesIO()
        downloader = request
        # MediaIoBaseDownload is heavier; the export_media returns bytes via .execute()
        data = downloader.execute()
        buf.write(data)
        return buf.getvalue()


