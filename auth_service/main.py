from __future__ import annotations

import hmac
import hashlib
import time
from typing import Dict, Any

import jwt
from fastapi import FastAPI, HTTPException, Depends
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel
import os


app = FastAPI(title="Auth Service", version="1.0.0")


class TelegramLoginPayload(BaseModel):
    # Telegram Login Widget params
    id: int
    first_name: str | None = None
    last_name: str | None = None
    username: str | None = None
    photo_url: str | None = None
    auth_date: int
    hash: str


class MiniAppPayload(BaseModel):
    # Mini App initData string from Telegram
    init_data: str


class TokenResponse(BaseModel):
    token: str
    expires_in: int


def _get_env(name: str, default: str | None = None) -> str:
    value = os.getenv(name, default)
    if value is None:
        raise RuntimeError(f"Missing env {name}")
    return value


BOT_TOKEN = _get_env("BOT_TOKEN", "")
JWT_SECRET = _get_env("AUTH_JWT_SECRET", "change-me")
JWT_EXPIRES = int(_get_env("AUTH_JWT_TTL_SECONDS", "86400"))  # 1 day default


def _check_telegram_login(data: Dict[str, Any]) -> None:
    # Validate per https://core.telegram.org/widgets/login#checking-authorization
    if not BOT_TOKEN:
        raise HTTPException(status_code=500, detail="BOT_TOKEN is not configured")
    received_hash = data.pop("hash", None)
    if not received_hash:
        raise HTTPException(status_code=400, detail="Missing hash")
    # Sort by key
    data_check_string = "\n".join(f"{k}={data[k]}" for k in sorted(data.keys()))
    secret_key = hashlib.sha256(BOT_TOKEN.encode()).digest()
    calculated_hash = hmac.new(secret_key, data_check_string.encode(), hashlib.sha256).hexdigest()
    if calculated_hash != received_hash:
        raise HTTPException(status_code=401, detail="Invalid signature")
    # Stale check
    auth_date = int(data.get("auth_date", 0))
    if auth_date and time.time() - auth_date > 86400:
        raise HTTPException(status_code=401, detail="Auth data expired")


def _issue_jwt(subject: str, extra: Dict[str, Any] | None = None) -> str:
    payload = {
        "sub": subject,
        "iat": int(time.time()),
        "exp": int(time.time()) + JWT_EXPIRES,
    }
    if extra:
        payload.update(extra)
    return jwt.encode(payload, JWT_SECRET, algorithm="HS256")


@app.post("/auth/telegram", response_model=TokenResponse)
def auth_telegram(payload: TelegramLoginPayload):
    data = payload.dict()
    _check_telegram_login(data.copy())
    user_id = str(payload.id)
    token = _issue_jwt(user_id, {"username": payload.username or ""})
    return TokenResponse(token=token, expires_in=JWT_EXPIRES)


@app.post("/auth/miniapp", response_model=TokenResponse)
def auth_miniapp(body: MiniAppPayload):
    # Minimal validation of initData using same HMAC principle
    # Parse querystring-like init_data into dict
    try:
        from urllib.parse import parse_qsl
        parts = dict(parse_qsl(body.init_data, keep_blank_values=True))
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid init_data")
    # Telegram docs: the data-check-string uses all fields except 'hash'
    if "user" in parts and isinstance(parts["user"], str):
        # keep as string in data-check-string
        pass
    _check_telegram_login(parts)
    user_id = "unknown"
    try:
        import json
        if "user" in parts:
            user = json.loads(parts["user"]) if isinstance(parts["user"], str) else parts["user"]
            user_id = str(user.get("id", "unknown"))
    except Exception:
        pass
    token = _issue_jwt(user_id)
    return TokenResponse(token=token, expires_in=JWT_EXPIRES)


security = HTTPBearer(auto_error=False)


@app.get("/auth/verify")
def verify_token(creds: HTTPAuthorizationCredentials | None = Depends(security)):
    if not creds:
        raise HTTPException(status_code=401, detail="Missing token")
    try:
        payload = jwt.decode(creds.credentials, JWT_SECRET, algorithms=["HS256"])
        return {"ok": True, "sub": payload.get("sub"), "exp": payload.get("exp")}
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token expired")
    except Exception:
        raise HTTPException(status_code=401, detail="Invalid token")



