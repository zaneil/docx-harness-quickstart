"""Bring your own harness: drive the Word tools from a hand-written loop.

Nothing here uses word_agent.agent. The same pattern works for plugging the
tools into any agent framework that accepts OpenAI-style tool definitions.

Run: uv run python examples/custom_harness.py   (needs OPENROUTER_API_KEY in .env)
"""

import os

from dotenv import load_dotenv
from openai import OpenAI

from word_agent import ToolRegistry, WordDocument, default_tools

load_dotenv()

doc = WordDocument.open()
registry = ToolRegistry(default_tools(doc))
client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=os.environ["OPENROUTER_API_KEY"])

messages = [
    {"role": "system", "content": "You draft Word documents using tools. Call get_outline first. Reply 'done' when finished."},
    {"role": "user", "content": "Write a three line haiku about spreadsheets, with a title."},
]

for _ in range(10):
    reply = client.chat.completions.create(
        model="openai/gpt-4.1-mini", messages=messages, tools=registry.to_openai_schema()
    ).choices[0].message
    messages.append(reply)
    if not reply.tool_calls:
        break
    for tool_call in reply.tool_calls:
        result = registry.execute(tool_call.function.name, tool_call.function.arguments)
        print(f"{tool_call.function.name} -> {result[:80]}")
        messages.append({"role": "tool", "tool_call_id": tool_call.id, "content": result})

doc.save("haiku.docx")
print("saved haiku.docx")
