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
import json
from dataclasses import dataclass, field
from typing import Any, Awaitable, Callable, Protocol, Union, runtime_checkable

from ..tools.registry import ToolRegistry
from .llm import LLMClient
from .prompts import (
    EXECUTE_PLAN_PROMPT,
    GENERIC_SYSTEM_PROMPT,
    PLAN_MODE_PROMPT,
    REVIEW_FEEDBACK_PROMPT,
    REVISE_PLAN_PROMPT,
)


@runtime_checkable
class Reviewer(Protocol):
    """Checks the agent's work once it has finished and returns feedback for
    it to act on, or None when there is nothing to fix.

    ``begin`` is called before the agent changes anything, so a reviewer can
    remember the starting state. The harness reviews only runs that changed
    something (called at least one tool that is not read-only).
    """

    def begin(self, instruction: str) -> Any: ...

    async def review(self, instruction: str, summary: str) -> str | None: ...


@dataclass
class AgentEvent:
    """Something that happened in the loop. ``type`` is one of:
    ``assistant`` (model reply), ``tool_call``, ``tool_result``, ``plan``,
    ``review`` (``{"round", "feedback"}``; feedback None means no problems,
    and ``error`` is set when the reviewer itself failed)."""

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
    #: The reviewer's feedback for each review round, None for a round that
    #: found no problems. Empty when there was no review.
    reviews: list[str | None] = field(default_factory=list)


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
        max_plan_steps: int = 8,
        max_plan_revisions: int = 5,
        on_event: EventCallback | None = None,
        reviewer: Reviewer | None = None,
        max_reviews: int = 2,
    ):
        self.llm = llm
        self.tools = tools
        self.system_prompt = system_prompt
        self.plan_prompt = plan_prompt
        self.max_steps = max_steps
        #: Model calls allowed while planning. On the last one the model gets
        #: no tools, so it has to answer with the plan.
        self.max_plan_steps = max_plan_steps
        self.max_plan_revisions = max_plan_revisions
        self.on_event = on_event
        #: Checks the finished work; its feedback goes back to the model for
        #: at most ``max_reviews`` rounds of fixes.
        self.reviewer = reviewer
        self.max_reviews = max_reviews
        self._changed = False

    # ------------------------------------------------------------ public API

    async def arun(
        self,
        instruction: str,
        *,
        plan_mode: bool = False,
        approve: ApproveCallback | None = None,
        history: list[dict] | None = None,
    ) -> AgentResult:
        """Complete ``instruction``.

        Pass ``history`` (the ``messages`` of an earlier result) to continue a
        conversation instead of starting a new one.

        With ``plan_mode=True`` the agent first plans with read-only tools. The
        plan goes to ``approve`` (auto-approved when no callback is given) and is
        executed only once approved.
        """
        if not plan_mode:
            messages = _start(self.system_prompt, history, instruction)
            return await self._act(messages, instruction)

        plan = await self.aplan(instruction, history=history)
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

    async def aplan(
        self,
        instruction: str,
        *,
        revise: Plan | None = None,
        feedback: str | None = None,
        history: list[dict] | None = None,
    ) -> Plan:
        """Produce a plan using read-only tools only. Pass ``revise`` and
        ``feedback`` to refine an earlier plan, or ``history`` to plan the next
        turn of an existing conversation."""
        if revise is not None:
            messages = list(revise.messages)
            messages.append({"role": "user", "content": REVISE_PLAN_PROMPT.format(feedback=feedback or "")})
            prior_steps = revise.steps
        else:
            messages = _start(f"{self.system_prompt}\n\n{self.plan_prompt}", history, instruction)
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
        result = await self._act(messages, plan.instruction)
        result.steps += plan.steps
        result.plan = plan
        return result

    # Synchronous conveniences. Inside a running event loop (Jupyter, FastAPI)
    # use the a* coroutines instead.
    def run(self, instruction: str, **kwargs) -> AgentResult:
        return _run_sync(self.arun(instruction, **kwargs))

    def plan(self, instruction: str, **kwargs) -> Plan:
        return _run_sync(self.aplan(instruction, **kwargs))

    def execute(self, plan: Plan) -> AgentResult:
        return _run_sync(self.aexecute(plan))

    # ------------------------------------------------------------------ loop

    async def _act(self, messages: list[dict], instruction: str) -> AgentResult:
        """Run the loop with the full tool set, then let the reviewer check
        the work and the model fix what it reports. All rounds share one
        ``max_steps`` budget."""
        reviewer = self.reviewer
        if reviewer is not None:
            try:
                await _maybe_await(reviewer.begin(instruction))
            except Exception as exc:  # noqa: BLE001 - a broken reviewer must not stop the work
                await self._emit("review", {"round": 0, "feedback": None, "error": f"{type(exc).__name__}: {exc}"})
                reviewer = None
        self._changed = False
        text, steps, status = await self._loop(messages, read_only=False)
        reviews: list[str | None] = []
        while (reviewer is not None and status == "completed" and self._changed
               and len(reviews) < self.max_reviews and steps < self.max_steps):
            round_ = len(reviews) + 1
            try:
                feedback = await _maybe_await(reviewer.review(instruction, text))
            except Exception as exc:  # noqa: BLE001 - keep the work, report the failed review
                await self._emit("review", {"round": round_, "feedback": None, "error": f"{type(exc).__name__}: {exc}"})
                break
            feedback = str(feedback).strip() if feedback else None
            reviews.append(feedback)
            await self._emit("review", {"round": round_, "feedback": feedback})
            if feedback is None:
                break
            messages.append({"role": "user", "content": REVIEW_FEEDBACK_PROMPT.format(feedback=feedback)})
            self._changed = False  # review the next round only if the model changed something
            more_text, more_steps, status = await self._loop(messages, read_only=False, limit=self.max_steps - steps)
            text = more_text or text
            steps += more_steps
        return AgentResult(text, steps, messages, status, reviews=reviews)

    async def _loop(self, messages: list[dict], *, read_only: bool, limit: int | None = None) -> tuple[str, int, str]:
        """Call the model, run its tool calls, repeat until it answers in text.
        Appends to ``messages`` in place. Returns (final text, steps, status)."""
        schema = self.tools.to_openai_schema(read_only=read_only) or None
        last_text = ""
        if limit is None:
            limit = min(self.max_steps, self.max_plan_steps) if read_only else self.max_steps
        self._seen_reads: dict[tuple[str, str], str] = {}
        for step in range(1, limit + 1):
            # Planning must end in a plan: on the final step offer no tools.
            offered = None if read_only and step == limit and limit > 1 else schema
            reply = await self.llm.acomplete(messages, offered)
            messages.append(reply)
            last_text = reply.get("content") or last_text
            calls = reply.get("tool_calls") or []
            await self._emit("assistant", {"content": reply.get("content"), "tool_calls": len(calls)})
            if not calls:
                return reply.get("content") or "", step, "completed"
            results = await self._run_tool_calls(calls, read_only)
            for call, result in zip(calls, results):
                messages.append({"role": "tool", "tool_call_id": call["id"], "content": result})
        return last_text, limit, "max_steps"

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
        key = (name, _canonical(arguments))
        if self._is_read_only(call):
            if key in self._seen_reads:
                # Weaker models can loop on identical reads. Say so instead of
                # paying for the same result again.
                result = (
                    "You already made this exact call and nothing has changed since; its result is "
                    "above. Do not repeat it. Continue with the next step"
                    + (" and write your plan." if read_only else ".")
                )
            else:
                result = await self.tools.aexecute(name, arguments, read_only=read_only)
                self._seen_reads[key] = result
        else:
            result = await self.tools.aexecute(name, arguments, read_only=read_only)
            if not read_only and not result.startswith("Error:"):
                self._seen_reads.clear()  # state changed; reads are fresh again
                self._changed = True
        await self._emit("tool_result", {"id": call["id"], "name": name, "result": result})
        return result

    async def _emit(self, type_: str, data: dict) -> None:
        if self.on_event is not None:
            await _maybe_await(self.on_event(AgentEvent(type_, data)))


def _start(system_prompt: str, history: list[dict] | None, instruction: str) -> list[dict]:
    """Messages for a new turn: system prompt, any earlier turns, the new request."""
    earlier = list(history or [])
    if earlier and earlier[0].get("role") == "system":
        earlier = earlier[1:]
    return [{"role": "system", "content": system_prompt}, *earlier, {"role": "user", "content": instruction}]


def _canonical(arguments: Any) -> str:
    """Arguments in a comparable form (key order and spacing ignored)."""
    if isinstance(arguments, str):
        try:
            arguments = json.loads(arguments or "{}")
        except ValueError:
            return arguments
    return json.dumps(arguments, sort_keys=True, default=str)


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
