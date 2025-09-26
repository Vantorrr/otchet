#!/usr/bin/env bash
set -euo pipefail

python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

export BOT_TOKEN="${BOT_TOKEN:-}"
export AUTH_JWT_SECRET="${AUTH_JWT_SECRET:-change-me}"
export AUTH_JWT_TTL_SECONDS="${AUTH_JWT_TTL_SECONDS:-86400}"

uvicorn auth_service.main:app --host 0.0.0.0 --port 8081



