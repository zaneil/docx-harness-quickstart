"""LLM clients.

The harness talks to models through one method, in OpenAI chat-completions
format: ``acomplete(messages, tools)`` takes a list of message dicts and a list
of tool definitions and returns the assistant message as a dict:

    {"role": "assistant", "content": "...",
     "tool_calls": [{"id": "...", "type": "function",
                     "function": {"name": "...", "arguments": "{...json...}"}}]}

Anything that implements that method is a valid client.
"""

from __future__ import annotations

import asyncio
import os
from typing import Any, Protocol, runtime_checkable

OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "openai/gpt-4o-mini"


@runtime_checkable
class LLMClient(Protocol):
    async def acomplete(self, messages: list[dict], tools: list[dict] | None = None) -> dict: ...


class OpenAICompatibleClient:
    """Client for any endpoint that speaks the OpenAI chat-completions API
    (OpenRouter, OpenAI, LiteLLM proxy, Ollama, vLLM...)."""

    def __init__(
        self,
        model: str,
        api_key: str | None = None,
        base_url: str | None = None,
        *,
        parallel_tool_calls: bool = True,
        **request_options: Any,
    ):
        self.model = model
        self.api_key = api_key
        self.base_url = base_url
        self.parallel_tool_calls = parallel_tool_calls
        self.request_options = request_options
        self.usage = {"prompt_tokens": 0, "completion_tokens": 0, "requests": 0}
        self._client = None
        self._loop = None

    def _get_client(self):
        # An AsyncOpenAI client is tied to the event loop it was first used on.
        # The sync wrappers start a new loop per call, so rebuild when it changes.
        loop = asyncio.get_running_loop()
        if self._client is None or self._loop is not loop:
            from openai import AsyncOpenAI

            self._client = AsyncOpenAI(api_key=self.api_key, base_url=self.base_url)
            self._loop = loop
        return self._client

    async def acomplete(self, messages: list[dict], tools: list[dict] | None = None) -> dict:
        request: dict[str, Any] = {"model": self.model, "messages": messages, **self.request_options}
        if tools:
            request["tools"] = tools
            if self.parallel_tool_calls:
                request["parallel_tool_calls"] = True
        response = await self._get_client().chat.completions.create(**request)
        if response.usage:
            self.usage["prompt_tokens"] += response.usage.prompt_tokens or 0
            self.usage["completion_tokens"] += response.usage.completion_tokens or 0
        self.usage["requests"] += 1
        message = response.choices[0].message
        out: dict[str, Any] = {"role": "assistant", "content": message.content}
        if message.tool_calls:
            out["tool_calls"] = [
                {
                    "id": call.id,
                    "type": "function",
                    "function": {"name": call.function.name, "arguments": call.function.arguments},
                }
                for call in message.tool_calls
            ]
        return out


class OpenRouterClient(OpenAICompatibleClient):
    """OpenRouter through the ``openai`` SDK.

    The key comes from ``api_key`` or ``OPENROUTER_API_KEY``; the model from
    ``model``, ``WORD_AGENT_MODEL`` or the default ``openai/gpt-4o-mini``.
    """

    def __init__(self, model: str | None = None, api_key: str | None = None, **kwargs: Any):
        api_key = api_key or os.environ.get("OPENROUTER_API_KEY")
        if not api_key:
            raise RuntimeError(
                "No OpenRouter API key. Set OPENROUTER_API_KEY (for example in a .env file) "
                "or pass api_key=..."
            )
        super().__init__(
            model or os.environ.get("WORD_AGENT_MODEL") or DEFAULT_MODEL,
            api_key=api_key,
            base_url=kwargs.pop("base_url", OPENROUTER_BASE_URL),
            **kwargs,
        )
