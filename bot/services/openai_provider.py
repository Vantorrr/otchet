from __future__ import annotations

import os
from typing import Dict, Any

import requests


class OpenAIProvider:
    """Thin wrapper around OpenAI Chat Completions for gpt-5-nano-compatible API."""

    def __init__(self) -> None:
        self.api_key = os.getenv("OPENAI_API_KEY", "")
        # Allow overriding base URL if needed
        self.base_url = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
        # Model per Pro Core spec
        self.model = os.getenv("OPENAI_MODEL", "gpt-5-nano")

    def _ensure(self) -> None:
        if not self.api_key:
            raise RuntimeError("OPENAI_API_KEY is not set")

    def generate_text(self, prompt: str, temperature: float = 0.2, max_tokens: int = 700) -> str:
        self._ensure()
        url = f"{self.base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        payload: Dict[str, Any] = {
            "model": self.model,
            "temperature": temperature,
            "max_tokens": max_tokens,
            "messages": [
                {"role": "system", "content": "Вы опытный руководитель отдела банковских гарантий. Пишите кратко, по делу, без Markdown."},
                {"role": "user", "content": prompt},
            ],
        }
        resp = requests.post(url, json=payload, headers=headers, timeout=30)
        if resp.status_code != 200:
            raise RuntimeError(f"OpenAI error {resp.status_code}: {resp.text}")
        data = resp.json()
        try:
            return data["choices"][0]["message"]["content"].strip()
        except Exception as exc:
            raise RuntimeError(f"Unexpected OpenAI response: {data}") from exc


