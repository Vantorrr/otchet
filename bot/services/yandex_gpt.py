"""YandexGPT integration service."""
import json
import requests
from typing import Dict, Any, Optional
from bot.config import Settings


class YandexGPTService:
    """Service for interacting with YandexGPT API."""
    
    def __init__(self, settings: Settings):
        self.api_key = settings.yandex_api_key
        self.folder_id = settings.yandex_folder_id
        self.base_url = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
    
    async def generate_analysis(self, data: Dict[str, Any]) -> str:
        """
        Generate AI analysis and recommendations based on sales data.
        
        Args:
            data: Dictionary containing aggregated sales data
            
        Returns:
            AI-generated analysis text
        """
        if not self.api_key or not self.folder_id:
            return "❌ YandexGPT не настроен. Добавьте YANDEX_API_KEY и YANDEX_FOLDER_ID в .env"
        
        prompt = self._build_analysis_prompt(data)
        
        try:
            response = await self._make_request(prompt)
            return response
        except Exception as e:
            return f"❌ Ошибка при генерации анализа: {str(e)}"

    async def generate_answer(self, question: str) -> str:
        """Generic Q&A generation for free-form questions."""
        if not self.api_key or not self.folder_id:
            return "❌ YandexGPT не настроен. Добавьте YANDEX_API_KEY и YANDEX_FOLDER_ID в .env"

        prompt = (
            "Ты опытный бизнес-аналитик. Отвечай кратко и по делу, на русском.\n"
            "Если спрашивают про наши отчёты/планы/сводки — учитывай, что данные приходят из Google Sheets, а цифры без ПДн.\n\n"
            f"Вопрос: {question}"
        )
        try:
            return await self._make_request(prompt)
        except Exception as e:
            return f"❌ Ошибка YandexGPT: {str(e)}"
    
    def _build_analysis_prompt(self, data: Dict[str, Any]) -> str:
        """Build prompt for AI analysis."""
        prompt = """Ты - аналитик по продажам. Проанализируй данные и дай краткие выводы.

Данные по менеджерам:
"""
        
        for manager, stats in data.items():
            prompt += f"\n{manager}:\n"
            prompt += f"- Перезвоны: {stats.get('calls_fact', 0)}/{stats.get('calls_plan', 0)}\n"
            prompt += f"- Заявки (шт): {stats.get('leads_units_fact', 0)}/{stats.get('leads_units_plan', 0)}\n"
            prompt += f"- Заявки (млн): {stats.get('leads_volume_fact', 0)}/{stats.get('leads_volume_plan', 0)}\n"
            prompt += f"- Одобрено (млн): {stats.get('approved_volume', 0)}\n"
            prompt += f"- Выдано (млн): {stats.get('issued_volume', 0)}\n"
        
        prompt += """

Дай краткий анализ (до 500 слов):
1. Общие итоги по команде
2. Топ-3 лучших показателя
3. Проблемные зоны
4. Рекомендации для улучшения

Ответь на русском языке, деловым стилем."""
        
        return prompt
    
    async def _make_request(self, prompt: str) -> str:
        """Make request to YandexGPT API."""
        headers = {
            "Authorization": f"Api-Key {self.api_key}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "modelUri": f"gpt://{self.folder_id}/yandexgpt-lite",
            "completionOptions": {
                "stream": False,
                "temperature": 0.3,
                "maxTokens": 1000
            },
            "messages": [
                {
                    "role": "user",
                    "text": prompt
                }
            ]
        }
        
        response = requests.post(
            self.base_url,
            headers=headers,
            data=json.dumps(payload),
            timeout=30
        )
        
        if response.status_code != 200:
            raise Exception(f"API error {response.status_code}: {response.text}")
        
        result = response.json()
        
        if "result" not in result or "alternatives" not in result["result"]:
            raise Exception(f"Unexpected API response format: {result}")
        
        return result["result"]["alternatives"][0]["message"]["text"]
