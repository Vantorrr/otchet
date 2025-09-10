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

    async def rank_top3(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Rank TOP-3 best and worst managers based on provided KPIs via YandexGPT.

        Returns dict: {"best": [names...], "worst": [names...], "reasons": {name: reason}}
        """
        if not self.api_key or not self.folder_id:
            return {"best": [], "worst": [], "reasons": {}, "error": "YANDEXGPT_NOT_CONFIGURED"}

        # Instruct the model to return STRICT JSON, no markdown, no commentary
        prompt = (
            "Ты аналитик продаж. На основе KPI менеджеров оцени эффективность и верни строго JSON.\n"
            "Правила ранжирования: основной вес — выполнение плана по перезвонам (calls_fact / calls_plan) и по объёму заявок (leads_volume_fact / leads_volume_plan).\n"
            "Вторичные факторы (для тай-брейка): issued_volume, approved_volume.\n"
            "Верни ТОЛЬКО JSON без пояснений в формате:\n"
            "{\n"
            "  \"best\": [\"Имя1\", \"Имя2\", \"Имя3\"],\n"
            "  \"worst\": [\"ИмяA\", \"ИмяB\", \"ИмяC\"],\n"
            "  \"reasons\": {\"Имя1\": \"краткая причина\", ...}\n"
            "}\n\n"
            "Данные:\n"
        )

        for manager, stats in data.items():
            prompt += (
                f"{manager}: calls {stats.get('calls_fact', 0)}/{stats.get('calls_plan', 0)}, "
                f"units {stats.get('leads_units_fact', 0)}/{stats.get('leads_units_plan', 0)}, "
                f"volume {stats.get('leads_volume_fact', 0)}/{stats.get('leads_volume_plan', 0)}, "
                f"approved {stats.get('approved_volume', 0)}, issued {stats.get('issued_volume', 0)}\n"
            )

        try:
            raw = await self._make_request(prompt)
            # Try to extract JSON block
            import json as _json
            text = raw.strip()
            start = text.find('{')
            end = text.rfind('}')
            if start != -1 and end != -1 and end > start:
                text = text[start:end+1]
            result = _json.loads(text)
            best = result.get("best") or []
            worst = result.get("worst") or []
            reasons = result.get("reasons") or {}
            # Ensure lists
            if not isinstance(best, list):
                best = []
            if not isinstance(worst, list):
                worst = []
            if not isinstance(reasons, dict):
                reasons = {}
            return {"best": best[:3], "worst": worst[:3], "reasons": reasons}
        except Exception as e:
            return {"best": [], "worst": [], "reasons": {}, "error": str(e)}

    async def generate_manager_comment(
        self,
        manager_name: str,
        previous: Dict[str, Any],
        current: Dict[str, Any],
        period_name: str,
    ) -> str:
        """Generate a brief per-manager comment (progress, risks, advice)."""
        if not self.api_key or not self.folder_id:
            return "❌ YandexGPT не настроен. Добавьте YANDEX_API_KEY и YANDEX_FOLDER_ID в .env"

        def line(label: str, key_plan: str, key_fact: str, data: Dict[str, Any]) -> str:
            plan = data.get(key_plan, 0)
            fact = data.get(key_fact, 0)
            return f"{label}: {fact}/{plan}"

        prev_lines = [
            line("Перезвоны", 'calls_plan', 'calls_fact', previous),
            line("Заявки шт", 'leads_units_plan', 'leads_units_fact', previous),
            line("Заявки млн", 'leads_volume_plan', 'leads_volume_fact', previous),
            f"Одобрено: {previous.get('approved_volume', 0)} млн",
            f"Выдано: {previous.get('issued_volume', 0)} млн",
            f"Новые звонки: {previous.get('new_calls', 0)}",
        ]

        cur_lines = [
            line("Перезвоны", 'calls_plan', 'calls_fact', current),
            line("Заявки шт", 'leads_units_plan', 'leads_units_fact', current),
            line("Заявки млн", 'leads_volume_plan', 'leads_volume_fact', current),
            f"Одобрено: {current.get('approved_volume', 0)} млн",
            f"Выдано: {current.get('issued_volume', 0)} млн",
            f"Новые звонки: {current.get('new_calls', 0)}",
        ]

        prompt = (
            "Ты руководитель отдела продаж. Дай короткий комментарий по менеджеру (60-90 слов), строго по делу, без воды.\n"
            "Структура: 1) Итоги и динамика vs прошлый период; 2) Где отстаёт/лидирует; 3) 2-3 конкретных рекомендации.\n"
            f"Менеджер: {manager_name}. Период: {period_name}.\n\n"
            "Прошлый период:\n" + "\n".join(prev_lines) + "\n\n"
            "Текущий период:\n" + "\n".join(cur_lines) + "\n\n"
            "Ответь кратко, деловым стилем, по-русски. Не используй markdown, только простой текст."
        )

        try:
            return await self._make_request(prompt)
        except Exception as e:
            return f"Комментарий недоступен: {str(e)}"

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
