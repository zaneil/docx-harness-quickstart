"""Add your own tools next to the built-in ones.

Run: uv run python examples/custom_tool.py   (needs OPENROUTER_API_KEY in .env)
"""

import asyncio
import datetime

from dotenv import load_dotenv

from docx_harness import WordDocument, create_agent, tool

load_dotenv()
doc = WordDocument.open()


@tool(read_only=True)
def todays_date() -> str:
    """Get today's date, for dating letters and memos."""
    return datetime.date.today().strftime("%B %d, %Y")


@tool(read_only=True)
async def lookup_client(name: str) -> dict:
    """Look up a client's address in the CRM.

    Args:
        name: Client name.
    """
    await asyncio.sleep(0.1)  # stands in for a real async API call
    return {"name": name, "address": "1 Example Street, Springfield", "contact": "Alex Doe"}


# Read-only tools called in the same turn run concurrently.
agent = create_agent(doc, extra_tools=[todays_date, lookup_client])
result = agent.run("Write a short, dated cover letter to our client Acme Corp about the attached Q3 report.")
print(result.final_text)
doc.save("letter.docx")
