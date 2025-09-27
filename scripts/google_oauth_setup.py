#!/usr/bin/env python3
"""
One-time OAuth setup to authorize under YOUR Google account (with 15GB).
Creates token.json that GoogleSlidesService can use instead of service account.

Usage:
  python3 scripts/google_oauth_setup.py

Prereqs:
  - Create OAuth Client ID (Desktop) in Google Cloud Console
  - Download credentials.json and put near this script
"""
from __future__ import annotations

import os
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/spreadsheets",
]


def main() -> int:
    creds_path = Path(__file__).with_name("credentials.json")
    if not creds_path.exists():
        print("❌ Put your OAuth credentials.json next to this script.")
        return 1

    flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
    creds = flow.run_local_server(port=8765, prompt='consent')
    token_path = Path(__file__).with_name("token.json")
    token_path.write_text(creds.to_json())
    print(f"✅ token.json saved at {token_path}")

    # quick sanity: list files
    drive = build("drive", "v3", credentials=creds)
    drive.files().list(pageSize=1).execute()
    print("✅ Drive access OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


