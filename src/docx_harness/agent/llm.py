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
import logging
import os
import random
from typing import Any, Protocol, runtime_checkable

OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "openai/gpt-6-luna"

log = logging.getLogger(__name__)

#: Error codes in a 200 reply that another attempt will not fix
#: (bad request, bad key, out of credits, refused by moderation).
_PERMANENT_CODES = {400, 401, 402, 403}


class LLMError(RuntimeError):
    """The endpoint answered, but with an error instead of a reply."""

    def __init__(self, message: str, code: int | None = None):
        super().__init__(message)
        self.code = code


def _reply_error(response: Any) -> LLMError:
    """Describe a response that has no choices. OpenRouter puts the upstream
    provider's failure in an ``error`` field of an HTTP 200 body."""
    error = getattr(response, "error", None)
    if isinstance(error, dict):
        code, message = error.get("code"), error.get("message")
    else:
        code, message = getattr(error, "code", None), getattr(error, "message", None)
    try:
        code = int(code) if code is not None else None
    except (TypeError, ValueError):
        code = None
    text = f"Model returned no reply: {message or 'empty response'}"
    if code is not None:
        text += f" (code {code})"
    return LLMError(text, code)


def _tokens(count: int) -> str:
    return f"{count / 1000:.1f}k" if count >= 1000 else str(count)


def format_usage(usage: dict | None, steps: int | None = None) -> str:
    """One line for people, such as ``12 steps · 48.2k in / 3.1k out · $0.0042``.

    ``usage`` is the dict an :class:`OpenAICompatibleClient` keeps on ``.usage``
    (``cost`` only when the provider reports it). Returns ``""`` without usage.
    """
    if not usage or not usage.get("requests"):
        return ""
    parts = [f"{steps} step{'s' if steps != 1 else ''}"] if steps is not None else []
    parts.append(f"{_tokens(usage.get('prompt_tokens', 0))} in / {_tokens(usage.get('completion_tokens', 0))} out")
    cost = usage.get("cost")
    if isinstance(cost, (int, float)):
        parts.append(f"${cost:.4f}" if cost < 1 else f"${cost:.2f}")
    return " · ".join(parts)


@runtime_checkable
class LLMClient(Protocol):
    async def acomplete(self, messages: list[dict], tools: list[dict] | None = None) -> dict: ...


class OpenAICompatibleClient:
    """Client for any endpoint that speaks the OpenAI chat-completions API
    (OpenRouter, OpenAI, LiteLLM proxy, Ollama, vLLM...).

    Failed calls are retried up to ``max_retries`` times with exponential
    backoff: connection errors, timeouts, 408, 409, 429 and 5xx by the
    ``openai`` SDK (which honours ``Retry-After``), and replies that arrive
    with no choices (an upstream failure reported in an HTTP 200) here.
    ``timeout`` is the limit per request, in seconds.
    """

    def __init__(
        self,
        model: str,
        api_key: str | None = None,
        base_url: str | None = None,
        *,
        parallel_tool_calls: bool = True,
        max_retries: int = 4,
        timeout: float | None = 120.0,
        **request_options: Any,
    ):
        self.model = model
        self.api_key = api_key
        self.base_url = base_url
        self.parallel_tool_calls = parallel_tool_calls
        self.max_retries = max_retries
        self.timeout = timeout
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

            self._client = AsyncOpenAI(
                api_key=self.api_key,
                base_url=self.base_url,
                max_retries=self.max_retries,
                timeout=self.timeout,
            )
            self._loop = loop
        return self._client

    async def _create(self, request: dict[str, Any]) -> Any:
        """Send the request; retry replies that came back without choices."""
        attempt = 0
        while True:
            response = await self._get_client().chat.completions.create(**request)
            if getattr(response, "choices", None):
                return response
            error = _reply_error(response)
            if error.code in _PERMANENT_CODES or attempt >= self.max_retries:
                raise error
            delay = min(0.5 * 2**attempt, 8.0) * random.uniform(0.8, 1.2)
            log.warning("%s; retrying in %.1fs", error, delay)
            await asyncio.sleep(delay)
            attempt += 1

    async def acomplete(self, messages: list[dict], tools: list[dict] | None = None) -> dict:
        request: dict[str, Any] = {"model": self.model, "messages": messages, **self.request_options}
        if tools:
            request["tools"] = tools
            if self.parallel_tool_calls:
                request["parallel_tool_calls"] = True
        response = await self._create(request)
        if response.usage:
            self.usage["prompt_tokens"] += response.usage.prompt_tokens or 0
            self.usage["completion_tokens"] += response.usage.completion_tokens or 0
            # Some providers (OpenRouter) also report what the request cost, in US dollars.
            cost = getattr(response.usage, "cost", None)
            if isinstance(cost, (int, float)):
                self.usage["cost"] = self.usage.get("cost", 0.0) + cost
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
    ``model``, ``DOCX_HARNESS_MODEL`` or the default ``openai/gpt-6-luna``.
    """

    def __init__(self, model: str | None = None, api_key: str | None = None, **kwargs: Any):
        api_key = api_key or os.environ.get("OPENROUTER_API_KEY")
        if not api_key:
            raise RuntimeError(
                "No OpenRouter API key. Set OPENROUTER_API_KEY (for example in a .env file) "
                "or pass api_key=..."
            )
        super().__init__(
            model or os.environ.get("DOCX_HARNESS_MODEL") or DEFAULT_MODEL,
            api_key=api_key,
            base_url=kwargs.pop("base_url", OPENROUTER_BASE_URL),
            **kwargs,
        )
