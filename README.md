# word-agent

An open source Python SDK for building LLM agents that draft and edit Microsoft Word (`.docx`) documents.

A `.docx` file is a zip archive of XML. word-agent unzips it into an in-memory abstraction, exposes that abstraction to a model as tools (read, insert, update, delete, style), and ships a small agent harness to run the loop. Each of the three layers is usable on its own, so you can bring your own harness, add your own tools, or both.

> Not affiliated with or endorsed by Microsoft. "Microsoft Word" is a trademark of Microsoft Corporation.

## Install

```bash
git clone https://github.com/zaneil/word-agent && cd word-agent
uv sync                      # or: pip install -e .
cp .env.example .env         # then add your OpenRouter key
```

Models are reached through [OpenRouter](https://openrouter.ai) using the `openai` SDK. The default model is `openai/gpt-4o-mini`; set `WORD_AGENT_MODEL` or pass `model=` to use any other.

## Quickstart

Command line:

```bash
# New document, started from the bundled template
uv run word-agent "Draft a one page project proposal with a title, three sections and a budget table" -o proposal.docx

# Edit an existing document
uv run word-agent "Tighten the introduction and make every heading dark blue" -i proposal.docx -o proposal-v2.docx

# Plan mode: the agent inspects the document, proposes a plan, and changes nothing until you approve
uv run word-agent "Restructure this into an executive summary plus appendix" -i report.docx -o report-v2.docx --plan
```

Python:

```python
from dotenv import load_dotenv
from word_agent import draft

load_dotenv()
result = draft("Write a short memo announcing the office move", output_path="memo.docx")
print(result.final_text)

# Edit an existing file, planning first
draft("Fix the typos and add a summary table", "memo.docx", "memo-v2.docx", plan=True)
```

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│ agent/      Agent harness            generic, knows no Word   │
│             Agent · LLMClient · plan mode · async loop        │
├──────────────────────────────────────────────────────────────┤
│ tools/      LLM-callable tools       knows no agent loop      │
│             Tool · @tool · ToolRegistry · default_tools(doc)  │
├──────────────────────────────────────────────────────────────┤
│ document/   Abstraction layer        knows no LLM             │
│             WordPackage (unzipped parts) · WordDocument       │
│             (blocks) · styling · EditContext                  │
└──────────────────────────────────────────────────────────────┘
```

### Layer 1: the document abstraction

`WordPackage` is the unzipped file: every part in memory, XML parsed on demand, unmodified parts written back byte for byte. `WordDocument` sits on top and presents the body as an ordered list of **blocks** (paragraphs and tables) with stable ids.

```python
from word_agent import WordDocument

doc = WordDocument.open()                       # bundled template; or WordDocument.open("in.docx")
title = doc.insert_paragraph("Q3 Report", "Title")
intro = doc.insert_paragraph("Revenue grew 12% year over year.")
table = doc.insert_table([["Region", "Revenue"], ["EMEA", "4.2M"]], after=intro)

doc.replace_text("12%", "14%")                  # works across differently formatted runs
doc.format_text(intro, match="14%", bold=True, color="1F4E79")
doc.set_paragraph_format(title, alignment="center")
doc.set_cell_text(table, 1, 1, "4.5M")          # cells are also addressable as f"{table}[1,1]"

print(doc.outline())
doc.save("report.docx")
```

Block ids (`b1`, `b2`...) are assigned on load and on insert, and never shift when other blocks are added or removed. The raw package is always reachable: `doc.list_parts()`, `doc.read_part_xml("word/styles.xml")`, `doc.write_part_xml(...)`.

### Layer 2: tools

`default_tools(doc)` returns the built-in tools bound to a document:

| Kind | Tools |
| --- | --- |
| Read | `get_outline`, `read_blocks`, `search_text`, `list_styles` |
| Write | `insert_paragraphs`, `insert_table`, `insert_page_break`, `insert_table_row` |
| Update | `update_paragraph`, `replace_text`, `set_table_cell` |
| Delete | `delete_blocks`, `delete_table_row` |
| Style | `set_paragraph_style`, `format_text`, `set_paragraph_format` |
| Raw package | `list_parts`, `read_part_xml`, and `write_part_xml` (only with `include_raw=True`) |

Tool schemas are in OpenAI function-calling format. Tool errors are returned to the model as text rather than raised, so it can correct itself.

### Layer 3: the harness

`Agent` is a plain async tool-calling loop over any `LLMClient` and `ToolRegistry`.

- **Plan mode.** `agent.arun(instruction, plan_mode=True, approve=callback)` first runs with read-only tools only, returns a plan, and executes it only after `approve(plan)` returns `True`. Returning a string sends feedback for a revised plan; returning `False` cancels. Mutating tools are blocked by the registry during planning, not just by the prompt. `aplan()` and `aexecute()` are available separately.
- **Async.** The loop is async (`arun`, `aplan`, `aexecute`), with sync wrappers (`run`, `plan`, `execute`) for scripts. When the model issues several tool calls in one turn, read-only calls run concurrently and mutating calls run in order. `adraft_many` processes several documents concurrently.
- **Events.** Pass `on_event` to observe model replies, tool calls and results.

## Extending

### Add your own tool

```python
from word_agent import WordDocument, create_agent, tool

doc = WordDocument.open("contract.docx")

@tool(read_only=True)                 # read-only tools stay available in plan mode
def word_count() -> int:
    """Count the words in the document."""
    return len(doc.text().split())

agent = create_agent(doc, extra_tools=[word_count])
agent.run("Cut the document down to under 500 words")
doc.save("contract-short.docx")
```

The schema comes from the type hints and the docstring's `Args:` section. Tools can be `async def`; use `@tool(blocking=True)` for sync functions that do slow I/O. Registering a tool with the name of a built-in replaces it.

### Use your own harness

The tools are plain objects, so any loop or framework can drive them:

```python
from word_agent import ToolRegistry, WordDocument, default_tools

doc = WordDocument.open()
registry = ToolRegistry(default_tools(doc))

schemas = registry.to_openai_schema()          # hand these to your model
result = registry.execute("insert_paragraphs", {"paragraphs": [{"text": "Hi", "style": "Title"}]})
```

### Use a different model or provider

Anything with this method is an `LLMClient`; messages and tool calls are OpenAI chat-completions dicts:

```python
class MyClient:
    async def acomplete(self, messages: list[dict], tools: list[dict] | None = None) -> dict:
        ...  # return {"role": "assistant", "content": ..., "tool_calls": [...]}
```

For any OpenAI-compatible endpoint (OpenAI, a LiteLLM proxy, Ollama, vLLM) use `OpenAICompatibleClient(model, api_key=..., base_url=...)`.

Runnable versions of these are in [`examples/`](examples).

## Scope and limitations

- Editable through the semantic tools: body paragraphs and tables, paragraph styles, character and paragraph formatting.
- Preserved on round trip but not editable through the semantic tools: headers and footers, footnotes, images, comments, content controls. They can be reached through the raw part tools.
- Existing tracked changes are respected when reading (deleted text is not shown to the agent) and left intact on save. Writing edits *as* tracked changes is not implemented yet; all mutations already flow through one place (`document/edits.py`) so it can be added without API changes.
- There is no layout engine: the agent cannot know page counts or where lines break.
- A document without the standard style names (for example no `Heading 1`) gets a clear error listing the styles it does have; styles are not created automatically.
- Legacy `.doc` files are not supported.

## Security notes

- The loader refuses archives with path traversal entries, oversized or over-populated archives, and XML containing DTDs or entity declarations.
- Text inside a document is untrusted input. If you add tools with side effects (email, web, shell), a malicious document could try to instruct the model to misuse them. Keep such tools behind plan mode or your own approval step.
- `write_part_xml` is off by default because a bad write can produce a file Word cannot open.

## Development

```bash
uv sync
uv run pytest
uv run python scripts/build_template.py   # regenerate the bundled blank.docx
```

## License

MIT
