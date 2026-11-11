# Build a Python AI agent that creates and edits Word documents

`docx-harness-quickstart` is a source starter kit for DOCX document automation with LLM tool calling. It provides a document abstraction, reusable document tools, and an optional async agent loop. Choose the smallest layer that fits your task.

## Install from source

```bash
git clone https://github.com/zaneil/docx-harness-quickstart.git
cd docx-harness-quickstart
uv sync
uv run pytest
```

Installing dependencies requires package access on the first run. The tests use a fake model and need no API key. For the agent examples, copy `.env.example` to `.env` and set `OPENROUTER_API_KEY`. Direct document operations and tool calls work without a model or credentials.

## Create and edit a DOCX without an LLM

Use the document layer when your changes are known in advance: create a report, insert a table, replace text, or apply formatting. Save the following as `example_document.py` in the cloned repository and run `uv run python example_document.py`.

```python
from docx_harness import WordDocument

doc = WordDocument.open()  # Start from the bundled blank Word template.
doc.insert_paragraph("Quarterly report", "Title")
summary = doc.insert_paragraph("Revenue grew 12% this quarter.")
doc.insert_table([["Region", "Revenue"], ["EMEA", "4.2M"]], after=summary)
doc.replace_text("12%", "14%")
doc.format_text(summary, match="14%", bold=True)
doc.save("report.docx")

reopened = WordDocument.open("report.docx")
assert "Revenue grew 14% this quarter." in reopened.text()
print(reopened.outline())
```

Expected result: `report.docx` contains a title, a paragraph with **14%** in bold, and a two-row table. To edit an existing document, pass its path to `WordDocument.open("input.docx")` and save to a separate output path.

## Edit a Word document with an agent

Use the full agent when the changes arrive as natural-language instructions. After setting up `.env`, run:

```bash
uv run docx-harness "Draft a project proposal with a title, goals, and a budget table" -o proposal.docx
uv run docx-harness "Add an executive summary and make the headings dark blue" -i proposal.docx -o proposal-v2.docx
```

Expected result: the first command writes `proposal.docx`; the second reads it and writes an edited copy as `proposal-v2.docx`. Model-generated content varies. These commands call the configured model provider and incur its usage charges.

The [Python quickstart](../examples/quickstart.py) shows the same workflow with `draft`, tool-call events, and an approval callback.

## Add DOCX tools to an existing agent

Use the tools layer if you already have a tool-calling loop. This example constructs the tool schemas and executes a document edit locally, without calling a model:

```python
from docx_harness import ToolRegistry, WordDocument, default_tools

doc = WordDocument.open()
registry = ToolRegistry(default_tools(doc))
schemas = registry.to_openai_schema()

# In your loop, give schemas to the model and execute its selected tool calls.
result = registry.execute(
    "insert_paragraphs",
    {"paragraphs": [{"text": "Executive summary", "style": "Heading 1"}]},
)
print(result)
doc.save("agent-output.docx")
assert "Executive summary" in WordDocument.open("agent-output.docx").text()
```

Expected result: `agent-output.docx` contains an Executive summary heading. Your loop must pass tool results back to the model and save the document when edits are complete. Tool results are strings; a failed call returns an `Error: ...` message for the agent to handle.

See the [complete custom harness](../examples/custom_harness.py) for a model-driven loop, and the [custom tool example](../examples/custom_tool.py) to extend the available operations. No ready-made MCP server is included.

## Plan and approve document edits

Use plan mode to review a proposed edit before allowing document changes:

```bash
uv run docx-harness "Restructure this into a summary and an appendix" -i proposal.docx -o proposal-reviewed.docx --plan
```

The agent inspects the document using read-only tools, proposes a plan, and asks for approval. In Python, pass `plan=True` and an `approve` callback to `draft`; returning `False` rejects the plan, and returning a string provides feedback for revision. See the [approval callback example](../examples/quickstart.py).

## Check the capabilities before integrating

The built-in semantic tools edit body paragraphs, tables, and formatting. Other document parts, including headers, images, and comments, are preserved on round trip but have no semantic editing tools. Advanced raw XML tools are opt-in.

Writing tracked changes is not implemented. This project has no PDF renderer or page-layout engine and cannot verify pagination. Legacy `.doc` files are unsupported. See the full [scope and limitations](../README.md#scope-and-limitations).

For extending the source, read [AGENTS.md](../AGENTS.md) for setup commands, implementation entry points, and document-validity rules.
