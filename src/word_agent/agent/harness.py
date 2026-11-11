"""Agent: a generic, async tool-calling loop with an optional plan mode.

This module knows nothing about Word. It needs an ``LLMClient`` and a
``ToolRegistry`` and nothing else, so it can be reused, subclassed or replaced.

    agent = Agent(llm, registry, system_prompt="...")
    result = await agent.arun("do the thing")                  # act directly
    result = await agent.arun("do the thing", plan_mode=True)  # plan, approve, act
"""

from __future__ import annotations

import asyncio
import inspect
from dataclasses import dataclass, field
from typing import Any, Awaitable, Callable, Union

from ..tools.registry import ToolRegistry
from .llm import LLMClient
from .prompts import EXECUTE_PLAN_PROMPT, GENERIC_SYSTEM_PROMPT, PLAN_MODE_PROMPT, REVISE_PLAN_PROMPT


@dataclass
class AgentEvent:
    """Something that happened in the loop. ``type`` is one of:
    ``assistant`` (model reply), ``tool_call``, ``tool_result``, ``plan``."""

    type: str
    data: dict


@dataclass
class Plan:
    """The outcome of the planning phase."""

    instruction: str
    text: str
    messages: list[dict] = field(repr=False, default_factory=list)
    steps: int = 0


@dataclass
class AgentResult:
    final_text: str
    #: Number of model calls made.
    steps: int
    messages: list[dict] = field(repr=False, default_factory=list)
    #: "completed", "max_steps" (ran out of steps) or "rejected" (plan not approved).
    status: str = "completed"
    plan: Plan | None = None


#: ``approve(plan)`` returns True to execute, False to cancel, or a string of
#: feedback to request a revised plan. It may be sync or async.
ApproveCallback = Callable[[Plan], Union[bool, str, None, Awaitable[Union[bool, str, None]]]]
EventCallback = Callable[[AgentEvent], Any]


class Agent:
    def __init__(
        self,
        llm: LLMClient,
        tools: ToolRegistry,
        *,
        system_prompt: str = GENERIC_SYSTEM_PROMPT,
        plan_prompt: str = PLAN_MODE_PROMPT,
        max_steps: int = 40,
        max_plan_revisions: int = 5,
        on_event: EventCallback | None = None,
    ):
        self.llm = llm
        self.tools = tools
        self.system_prompt = system_prompt
        self.plan_prompt = plan_prompt
        self.max_steps = max_steps
        self.max_plan_revisions = max_plan_revisions
        self.on_event = on_event

    # ------------------------------------------------------------ public API

    async def arun(
        self,
        instruction: str,
        *,
        plan_mode: bool = False,
        approve: ApproveCallback | None = None,
    ) -> AgentResult:
        """Complete ``instruction``.

        With ``plan_mode=True`` the agent first plans with read-only tools. The
        plan goes to ``approve`` (auto-approved when no callback is given) and is
        executed only once approved.
        """
        if not plan_mode:
            messages = [
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": instruction},
            ]
            text, steps, status = await self._loop(messages, read_only=False)
            return AgentResult(text, steps, messages, status)

        plan = await self.aplan(instruction)
        revisions = 0
        while True:
            decision = True if approve is None else await _maybe_await(approve(plan))
            if decision is True:
                return await self.aexecute(plan)
            if isinstance(decision, str) and decision.strip() and revisions < self.max_plan_revisions:
                revisions += 1
                plan = await self.aplan(instruction, revise=plan, feedback=decision)
                continue
            return AgentResult(plan.text, plan.steps, plan.messages, "rejected", plan)

    async def aplan(self, instruction: str, *, revise: Plan | None = None, feedback: str | None = None) -> Plan:
        """Produce a plan using read-only tools only. Pass ``revise`` and
        ``feedback`` to refine an earlier plan."""
        if revise is not None:
            messages = list(revise.messages)
            messages.append({"role": "user", "content": REVISE_PLAN_PROMPT.format(feedback=feedback or "")})
            prior_steps = revise.steps
        else:
            messages = [
                {"role": "system", "content": f"{self.system_prompt}\n\n{self.plan_prompt}"},
                {"role": "user", "content": instruction},
            ]
            prior_steps = 0
        text, steps, _ = await self._loop(messages, read_only=True)
        plan = Plan(instruction, text, messages, prior_steps + steps)
        await self._emit("plan", {"text": text})
        return plan

    async def aexecute(self, plan: Plan) -> AgentResult:
        """Execute an approved plan with the full tool set. The planning
        conversation is kept, so what the model already read is not re-read."""
        messages = [{"role": "system", "content": self.system_prompt}, *plan.messages[1:]]
        messages.append({"role": "user", "content": EXECUTE_PLAN_PROMPT})
        text, steps, status = await self._loop(messages, read_only=False)
        return AgentResult(text, plan.steps + steps, messages, status, plan)

    # Synchronous conveniences. Inside a running event loop (Jupyter, FastAPI)
    # use the a* coroutines instead.
    def run(self, instruction: str, **kwargs) -> AgentResult:
        return _run_sync(self.arun(instruction, **kwargs))

    def plan(self, instruction: str, **kwargs) -> Plan:
        return _run_sync(self.aplan(instruction, **kwargs))

    def execute(self, plan: Plan) -> AgentResult:
        return _run_sync(self.aexecute(plan))

    # ------------------------------------------------------------------ loop

    async def _loop(self, messages: list[dict], *, read_only: bool) -> tuple[str, int, str]:
        """Call the model, run its tool calls, repeat until it answers in text.
        Appends to ``messages`` in place. Returns (final text, steps, status)."""
        schema = self.tools.to_openai_schema(read_only=read_only) or None
        last_text = ""
        for step in range(1, self.max_steps + 1):
            reply = await self.llm.acomplete(messages, schema)
            messages.append(reply)
            last_text = reply.get("content") or last_text
            calls = reply.get("tool_calls") or []
            await self._emit("assistant", {"content": reply.get("content"), "tool_calls": len(calls)})
            if not calls:
                return reply.get("content") or "", step, "completed"
            results = await self._run_tool_calls(calls, read_only)
            for call, result in zip(calls, results):
                messages.append({"role": "tool", "tool_call_id": call["id"], "content": result})
        return last_text, self.max_steps, "max_steps"

    async def _run_tool_calls(self, calls: list[dict], read_only: bool) -> list[str]:
        """Run one turn's tool calls. Consecutive read-only calls run
        concurrently; anything that mutates state runs alone, in order."""
        results: list[str] = [""] * len(calls)
        i = 0
        while i < len(calls):
            j = i
            while j < len(calls) and self._is_read_only(calls[j]):
                j += 1
            if j > i:
                results[i:j] = await asyncio.gather(*(self._call(c, read_only) for c in calls[i:j]))
                i = j
            else:
                results[i] = await self._call(calls[i], read_only)
                i += 1
        return results

    def _is_read_only(self, call: dict) -> bool:
        registered = self.tools.get(call["function"]["name"])
        return registered is not None and registered.read_only

    async def _call(self, call: dict, read_only: bool) -> str:
        name = call["function"]["name"]
        arguments = call["function"].get("arguments")
        await self._emit("tool_call", {"id": call["id"], "name": name, "arguments": arguments})
        result = await self.tools.aexecute(name, arguments, read_only=read_only)
        await self._emit("tool_result", {"id": call["id"], "name": name, "result": result})
        return result

    async def _emit(self, type_: str, data: dict) -> None:
        if self.on_event is not None:
            await _maybe_await(self.on_event(AgentEvent(type_, data)))


async def _maybe_await(value: Any) -> Any:
    return await value if inspect.isawaitable(value) else value


def _run_sync(coro):
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return asyncio.run(coro)
    coro.close()
    raise RuntimeError(
        "An event loop is already running. Use the async methods instead: "
        "await agent.arun(...), agent.aplan(...), agent.aexecute(...)."
    )
