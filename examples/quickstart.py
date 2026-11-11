"""Draft a new document, then edit it with plan mode.

Run: uv run python examples/quickstart.py   (needs OPENROUTER_API_KEY in .env)
"""

from dotenv import load_dotenv

from docx_harness import draft

load_dotenv()


def show(event):
    if event.type == "tool_call":
        print(f"  → {event.data['name']}")


# 1. New document from the bundled template.
result = draft(
    "Draft a one page project proposal for a team offsite: a title, three short "
    "sections with headings, a bulleted list of goals and a small budget table.",
    output_path="proposal.docx",
    on_event=show,
)
print(result.final_text, "\n")


# 2. Edit it. The agent plans with read-only tools; approve() decides what happens next.
def approve(plan):
    print("Plan:\n" + plan.text + "\n")
    return True  # or False to cancel, or a string of feedback to get a revised plan


result = draft(
    "Add an executive summary paragraph under the title and right-align the budget numbers.",
    input_path="proposal.docx",
    output_path="proposal-v2.docx",
    plan=True,
    approve=approve,
    on_event=show,
)
print(result.final_text)
