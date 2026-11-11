"""High-level helpers that wire the three layers together."""

from __future__ import annotations

import asyncio
import os
from typing import Any, Callable, Iterable

from .agent.harness import Agent, AgentResult, ApproveCallback, EventCallback, _run_sync
from .agent.llm import LLMClient, OpenRouterClient
from .agent.prompts import WORD_SYSTEM_PROMPT
from .document import WordDocument
from .tools import Tool, ToolRegistry, default_tools


def create_agent(
    doc: WordDocument,
    *,
    llm: LLMClient | None = None,
    model: str | None = None,
    extra_tools: Iterable[Tool | Callable] = (),
    include_raw: bool = False,
    system_prompt: str = WORD_SYSTEM_PROMPT,
    max_steps: int = 40,
    on_event: EventCallback | None = None,
) -> Agent:
    """Build an Agent with the default Word tools bound to ``doc``.

    ``extra_tools`` are registered after the defaults, so a tool with the same
    name as a built-in replaces it.
    """
    registry = ToolRegistry(default_tools(doc, include_raw=include_raw))
    for item in extra_tools:
        registry.register(item)
    return Agent(
        llm or OpenRouterClient(model=model),
        registry,
        system_prompt=system_prompt,
        max_steps=max_steps,
        on_event=on_event,
    )


async def adraft(
    instruction: str,
    input_path: str | os.PathLike | None = None,
    output_path: str | os.PathLike = "output.docx",
    *,
    plan: bool = False,
    approve: ApproveCallback | None = None,
    **agent_options: Any,
) -> AgentResult:
    """Draft a new document (no ``input_path``: starts from the bundled
    template) or edit an existing one, then save it to ``output_path``.

    With ``plan=True`` the agent plans first; nothing is saved if the plan is
    rejected. ``agent_options`` are passed to :func:`create_agent`.
    """
    doc = WordDocument.open(input_path)
    agent = create_agent(doc, **agent_options)
    result = await agent.arun(instruction, plan_mode=plan, approve=approve)
    if result.status != "rejected":
        doc.save(output_path)
    return result


def draft(instruction: str, input_path=None, output_path="output.docx", **kwargs: Any) -> AgentResult:
    """Synchronous :func:`adraft`."""
    return _run_sync(adraft(instruction, input_path, output_path, **kwargs))


async def adraft_many(jobs: Iterable[dict], concurrency: int = 4, **common: Any) -> list[AgentResult | Exception]:
    """Run several drafting jobs concurrently.

    Each job is a dict of :func:`adraft` arguments, for example
    ``{"instruction": "...", "output_path": "a.docx"}``. Results come back in
    job order; a job that fails yields its exception instead of a result.
    """
    semaphore = asyncio.Semaphore(max(concurrency, 1))

    async def run(job: dict):
        async with semaphore:
            return await adraft(**{**common, **job})

    return await asyncio.gather(*(run(job) for job in jobs), return_exceptions=True)
