"""Command line interface.

    word-agent "Draft a project proposal" -o proposal.docx
    word-agent "Make every heading blue" -i proposal.docx -o proposal2.docx
    word-agent "Tighten the introduction" -i report.docx -o report2.docx --plan
"""

from __future__ import annotations

import argparse
import sys

from .agent import AgentEvent, Plan
from .api import draft


def _printer(verbose: bool):
    def on_event(event: AgentEvent) -> None:
        if event.type == "tool_call":
            arguments = event.data.get("arguments") or ""
            if not verbose and len(arguments) > 100:
                arguments = arguments[:100] + "…"
            print(f"  → {event.data['name']} {arguments}", file=sys.stderr)
        elif event.type == "tool_result" and verbose:
            print(f"    {event.data['result'][:500]}", file=sys.stderr)

    return on_event


def _ask_approval(plan: Plan):
    print("\nProposed plan:\n")
    print(plan.text)
    print()
    try:
        answer = input("[a]pprove, [c]ancel, or type feedback to revise: ").strip()
    except EOFError:
        return False
    if answer.lower() in ("a", "approve", "y", "yes"):
        return True
    if answer.lower() in ("", "c", "cancel", "n", "no"):
        return False
    return answer


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="word-agent", description="Draft or edit a Word document with an LLM agent.")
    parser.add_argument("instruction", help="What to write or change.")
    parser.add_argument("-i", "--input", help="Existing .docx to edit. Omit to start from the built-in template.")
    parser.add_argument("-o", "--output", default="output.docx", help="Where to save the result (default: output.docx).")
    parser.add_argument("--model", help="OpenRouter model id (default: openai/gpt-4.1-mini or $WORD_AGENT_MODEL).")
    parser.add_argument("--plan", action="store_true", help="Plan first and ask for approval before changing anything.")
    parser.add_argument("--yes", action="store_true", help="With --plan: approve the plan automatically.")
    parser.add_argument("--raw", action="store_true", help="Also give the agent the raw XML write tool.")
    parser.add_argument("--max-steps", type=int, default=40, help="Maximum model calls (default: 40).")
    parser.add_argument("-v", "--verbose", action="store_true", help="Show full tool calls and results.")
    args = parser.parse_args(argv)

    try:
        from dotenv import load_dotenv

        load_dotenv()
    except ImportError:
        pass

    try:
        result = draft(
            args.instruction,
            args.input,
            args.output,
            plan=args.plan,
            approve=None if args.yes or not args.plan else _ask_approval,
            model=args.model,
            include_raw=args.raw,
            max_steps=args.max_steps,
            on_event=_printer(args.verbose),
        )
    except Exception as exc:  # noqa: BLE001 - top-level CLI error report
        print(f"error: {exc}", file=sys.stderr)
        return 1

    if result.status == "rejected":
        print("Plan not approved; nothing was changed.", file=sys.stderr)
        return 1
    if args.plan and args.yes and result.plan:
        print(f"Plan:\n{result.plan.text}\n")
    print(result.final_text)
    if result.status == "max_steps":
        print(f"warning: stopped after {result.steps} steps; the document may be incomplete.", file=sys.stderr)
    print(f"Saved {args.output}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
