from __future__ import annotations

import json
import re
import time
from typing import TypeVar

from pydantic import BaseModel, ValidationError

from errors import BridgeError
from orchestrator.config import AgentConfig

T = TypeVar("T", bound=BaseModel)

_JSON_FENCE_RE = re.compile(r"```(?:json)?\s*(.*?)\s*```", re.DOTALL)


def _extract_json(text: str) -> str:
    match = _JSON_FENCE_RE.search(text)
    if match:
        return match.group(1)
    start = text.find("{")
    if start != -1:
        return text[start:]
    return text


class LLMClient:
    def __init__(self, config: AgentConfig) -> None:
        self.config = config
        if config.llm_provider == "anthropic":
            try:
                import anthropic

                self._anthropic = anthropic.Anthropic(api_key=config.api_key)
            except ImportError as exc:
                raise BridgeError(
                    code="dependency_missing",
                    message="anthropic package not installed. Run: pip install anthropic",
                ) from exc
        elif config.llm_provider == "azure_openai":
            if not config.azure_endpoint:
                raise BridgeError(
                    code="configuration_error",
                    message="AZURE_OPENAI_ENDPOINT is required when PPTX_LLM_PROVIDER=azure_openai.",
                )
            if not config.azure_deployment:
                raise BridgeError(
                    code="configuration_error",
                    message="AZURE_OPENAI_DEPLOYMENT is required when PPTX_LLM_PROVIDER=azure_openai.",
                )
            try:
                import openai

                self._openai = openai.AzureOpenAI(
                    api_key=config.api_key,
                    azure_endpoint=config.azure_endpoint,
                    api_version=config.azure_api_version or "2024-02-01",
                )
            except ImportError as exc:
                raise BridgeError(
                    code="dependency_missing",
                    message="openai package not installed. Run: pip install openai",
                ) from exc
        else:
            try:
                import openai

                self._openai = openai.OpenAI(api_key=config.api_key)
            except ImportError as exc:
                raise BridgeError(
                    code="dependency_missing",
                    message="openai package not installed. Run: pip install openai",
                ) from exc

    def call_raw(self, system: str, user: str) -> str:
        for attempt in range(3):
            try:
                if self.config.llm_provider == "anthropic":
                    resp = self._anthropic.messages.create(
                        model=self.config.model,
                        max_tokens=self.config.max_tokens,
                        temperature=self.config.temperature,
                        system=system,
                        messages=[{"role": "user", "content": user}],
                    )
                    return resp.content[0].text
                else:
                    # Both "openai" and "azure_openai" use the same chat completions API.
                    # For Azure, azure_deployment is used as the model/deployment name.
                    model = (
                        self.config.azure_deployment
                        if self.config.llm_provider == "azure_openai"
                        else self.config.model
                    )
                    resp = self._openai.chat.completions.create(
                        model=model,
                        max_tokens=self.config.max_tokens,
                        temperature=self.config.temperature,
                        messages=[
                            {"role": "system", "content": system},
                            {"role": "user", "content": user},
                        ],
                    )
                    return resp.choices[0].message.content or ""
            except Exception as exc:
                if attempt == 2:
                    raise BridgeError(
                        code="internal_error",
                        message=f"LLM API call failed after 3 attempts: {exc}",
                    ) from exc
                time.sleep(2**attempt)
        return ""

    def call_structured(self, system: str, user: str, model_class: type[T]) -> T:
        current_user = user
        for attempt in range(3):
            raw = self.call_raw(system, current_user)
            try:
                json_str = _extract_json(raw)
                return model_class.model_validate_json(json_str)
            except (json.JSONDecodeError, ValidationError) as exc:
                if attempt == 2:
                    raise BridgeError(
                        code="internal_error",
                        message=f"LLM returned unparseable JSON after 3 attempts. Last error: {exc}",
                        details={"raw_response": raw[:500]},
                    ) from exc
                current_user = (
                    f"{user}\n\n"
                    f"IMPORTANT: Your previous response was not valid JSON. Error: {exc}\n"
                    f"Respond with raw JSON ONLY - no markdown, no explanation, no code fences."
                )
        raise BridgeError(code="internal_error", message="Unreachable")
