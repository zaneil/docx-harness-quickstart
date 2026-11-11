"""Layer 3: the agent harness. No Word knowledge lives here."""

from .harness import Agent, AgentEvent, AgentResult, Plan
from .llm import DEFAULT_MODEL, LLMClient, OpenAICompatibleClient, OpenRouterClient
from .prompts import GENERIC_SYSTEM_PROMPT, PLAN_MODE_PROMPT, WORD_SYSTEM_PROMPT

__all__ = [
    "Agent", "AgentEvent", "AgentResult", "Plan",
    "LLMClient", "OpenAICompatibleClient", "OpenRouterClient", "DEFAULT_MODEL",
    "GENERIC_SYSTEM_PROMPT", "PLAN_MODE_PROMPT", "WORD_SYSTEM_PROMPT",
]
