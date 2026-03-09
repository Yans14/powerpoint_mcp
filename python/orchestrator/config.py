from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass
class AgentConfig:
    llm_provider: str
    api_key: str
    model: str
    max_tokens: int = 4096
    temperature: float = 0.1
    max_steps: int = 100
    execution_timeout_sec: int = 600
    # Azure OpenAI-specific fields (only required when llm_provider == "azure_openai")
    azure_endpoint: str | None = None
    azure_api_version: str | None = None
    azure_deployment: str | None = None

    @classmethod
    def from_env(cls) -> AgentConfig | None:
        provider = os.getenv("PPTX_LLM_PROVIDER", "anthropic").lower()
        api_key = (
            os.getenv("PPTX_LLM_API_KEY")
            or (os.getenv("ANTHROPIC_API_KEY") if provider == "anthropic" else None)
            or (os.getenv("OPENAI_API_KEY") if provider == "openai" else None)
            or (os.getenv("AZURE_OPENAI_API_KEY") if provider == "azure_openai" else None)
        )
        if not api_key:
            return None

        default_models = {"anthropic": "claude-sonnet-4-20250514", "openai": "gpt-4o"}
        model = os.getenv("PPTX_LLM_MODEL", default_models.get(provider, "claude-sonnet-4-20250514"))

        azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT") if provider == "azure_openai" else None
        azure_api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-01") if provider == "azure_openai" else None
        azure_deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT") if provider == "azure_openai" else None

        return cls(
            llm_provider=provider,
            api_key=api_key,
            model=model,
            max_steps=int(os.getenv("PPTX_AGENT_MAX_STEPS", "100")),
            execution_timeout_sec=int(os.getenv("PPTX_AGENT_TIMEOUT", "600")),
            azure_endpoint=azure_endpoint,
            azure_api_version=azure_api_version,
            azure_deployment=azure_deployment,
        )
