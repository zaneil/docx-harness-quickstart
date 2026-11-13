# DOCX agent harness quickstart: create and edit Word documents in Python

A starter kit for building LLM agents that draft and edit Microsoft Word (`.docx`) documents in Python. Clone it, point your coding agent at it, and adapt it.

A `.docx` file is a zip archive of XML. docx-harness unzips it into an in-memory abstraction, exposes that abstraction to a model as tools (read, insert, update, delete, style), and ships a small agent harness to run the loop. Each of the three layers is usable on its own, so you can bring your own harness, add your own tools, or both.

> Not affiliated with or endorsed by Microsoft. "Microsoft Word" is a trademark of Microsoft Corporation.

## When to use this starter kit

Use `docx-harness-quickstart` to **build a Python AI agent that creates or edits Word documents**, or to **add DOCX editing tools to an existing agent**. Clone and adapt the source; the Python import is `docx_harness`.

| What you want to build | Start with |
| --- | --- |
| An agent that drafts reports, proposals, or memos as `.docx` | [`draft` and the command-line quickstart](#quickstart) |
| An agent that edits paragraphs, tables, and formatting in an existing Word file | [Edit a Word document with an agent](docs/use-cases.md#edit-a-word-document-with-an-agent) |
| Word-editing tools for your own tool-calling loop | [Add DOCX tools to an existing agent](docs/use-cases.md#add-docx-tools-to-an-existing-agent) |
| Direct Python document editing without an LLM or API key | [Create and edit a DOCX without an LLM](docs/use-cases.md#create-and-edit-a-docx-without-an-llm) |
| A review step before an agent changes a document | [Plan and approve document edits](docs/use-cases.md#plan-and-approve-document-edits) |

The built-in semantic tools handle body paragraphs, tables, and formatting. Generating tracked changes, PDF rendering, page-layout inspection, and editing legacy `.doc` files are outside the current scope. See [scope and limitations](#scope-and-limitations) before choosing it for a task.

## For coding agents

If you are a coding agent (or are handing this repo to one), read [`AGENTS.md`](AGENTS.md) first. It has the commands, a task-to-file map, a minimal usage snippet for each layer, and the rules that keep generated documents valid. `CLAUDE.md` points to the same file.

For choosing an integration, see the [task guide](docs/use-cases.md). A compact [documentation index](llms.txt) links to the Markdown documentation and runnable examples.

- Everything is verifiable offline: `uv sync && uv run pytest` takes about 3 seconds and needs no API key.
- The package is `docx_harness` under `src/`; the layers depend one way only: `agent` → `tools` → `document`.
- Prompts that work well as a starting point:
  - "Read AGENTS.md, then add a tool that inserts an image after a given block, with tests."
  - "Read AGENTS.md, then replace the built-in harness with my LangGraph agent, keeping the default tools."
  - "Read AGENTS.md, then add tools for reading and replying to Word comments."

## Install

```bash
git clone https://github.com/zaneil/docx-harness-quickstart && cd docx-harness-quickstart
uv sync                      # or: pip install -e .
cp .env.example .env         # then add your OpenRouter key
```

Models are reached through [OpenRouter](https://openrouter.ai) using the `openai` SDK. The default model is `openai/gpt-4.1-mini`; set `DOCX_HARNESS_MODEL` or pass `model=` to use any other.

## Quickstart

Command line:

```bash
# New document, started from the bundled template
uv run docx-harness "Draft a one page project proposal with a title, three sections and a budget table" -o proposal.docx

# Edit an existing document
uv run docx-harness "Tighten the introduction and make every heading dark blue" -i proposal.docx -o proposal-v2.docx

# Plan mode: the agent inspects the document, proposes a plan, and changes nothing until you approve
uv run docx-harness "Restructure this into an executive summary plus appendix" -i report.docx -o report-v2.docx --plan

# Tracked changes: every edit is saved as a redline to accept or reject in Word
uv run docx-harness "Shorten the payment terms to 30 days and fix the typos" -i contract.docx -o contract-redline.docx --track-changes --author "Review Agent"
```

Python:

```python
from dotenv import load_dotenv
from docx_harness import draft

load_dotenv()
result = draft("Write a short memo announcing the office move", output_path="memo.docx")
print(result.final_text)

# Edit an existing file, planning first
draft("Fix the typos and add a summary table", "memo.docx", "memo-v2.docx", plan=True)

# Save the edits as tracked changes instead of applying them directly
draft("Make the tone more formal", "memo.docx", "memo-redline.docx", track_changes=True, author="Review Agent")
```

### Chat UI for testing

```bash
uv run docx-harness-ui        # opens http://127.0.0.1:8765
```

A small local page with a chat on the left and a live preview of the document on the right. Start from the blank template or upload a `.docx`, keep chatting to refine it (the conversation and block ids carry over between messages), tick **Plan first** to review a plan before anything changes, and download the result. It is a test tool: one in-memory session, standard library only, bound to localhost. The preview is approximate; open the downloaded file in Word for the real rendering.

To continue a conversation in your own code, pass the previous result's messages back in: `agent.arun("now make it shorter", history=result.messages)`.

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
from docx_harness import WordDocument

doc = WordDocument.open()                       # bundled template; or WordDocument.open("in.docx")
title = doc.insert_paragraph("Q3 Report", "Title")
intro = doc.insert_paragraph("Revenue grew 12% year over year.")
table = doc.insert_table([["Region", "Revenue"], ["EMEA", "4.2M"]], after=intro)

doc.replace_text("12%", "14%")                  # works across differently formatted runs
doc.format_text(intro, match="14%", bold=True, color="1F4E79")
doc.set_paragraph_format(title, alignment="center")
doc.set_cell_text(table, 1, 1, "4.5M")          # cells are also addressable as f"{table}[1,1]"

doc.insert_table_column(table, cells=["Growth", "8%"])
doc.insert_table_row(table, ["FY25 results", "", ""], index=0)
doc.merge_cells(table, 0, 0, 0, 2)              # one heading cell across the three columns
doc.format_cells(f"{table}[0,*]", shading="D9E2F3")          # a whole row; f"{table}[*,2]" is a column
doc.set_table_format(table, column_widths_in=[2.5, 2, 2], header_rows=2)

print(doc.outline())
doc.save("report.docx")
```

Block ids (`b1`, `b2`...) are assigned on load and on insert, and never shift when other blocks are added or removed.

Tables are read and addressed on their column grid, so files with merged cells behave: `[row,col]` means the same column in every row, a merged cell answers to every position it covers, and reading a table lists its `merged` cells. A newline in cell text is a paragraph break, and rewriting a cell keeps its paragraphs' formatting. A table nested inside a cell gets its own block id (listed under `nested` when you read the outer table) and is edited like any other table. The raw package is always reachable: `doc.list_parts()`, `doc.read_part_xml("word/styles.xml")`, `doc.write_part_xml(...)`.

### Layer 2: tools

`default_tools(doc)` returns the built-in tools bound to a document:

| Kind | Tools |
| --- | --- |
| Read | `get_outline`, `read_blocks`, `search_text`, `list_styles`, `read_headers_footnotes`, `list_revisions` |
| Write | `insert_paragraphs`, `insert_table`, `insert_page_break`, `insert_table_row`, `insert_table_column` |
| Update | `update_paragraph`, `replace_text`, `set_table_cell`, `merge_table_cells`, `unmerge_table_cells`, `move_blocks` |
| Delete | `delete_blocks`, `delete_table_row`, `delete_table_column` |
| Style | `set_paragraph_style`, `format_text`, `set_paragraph_format`, `format_table_cells`, `set_table_format` |
| Review | `accept_revisions`, `reject_revisions` |
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
from docx_harness import WordDocument, create_agent, tool

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
from docx_harness import ToolRegistry, WordDocument, default_tools

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

## Tracked changes

Open a document with `track_changes=True` and every edit made through the document API or the tools is recorded as a Word revision instead of being applied directly:

```python
doc = WordDocument.open("contract.docx", track_changes=True, author="Review Agent")
doc.replace_text("60 days", "30 days")            # old text struck out, new text inserted
doc.set_paragraph_text("b4", "Payment is due within 30 days of a valid invoice.")   # only the changed words are revised
doc.delete_block("b9")                            # stays in the file as a deletion, leaves the outline
doc.format_text("b2", match="Supplier", bold=True)   # formatting change, previous formatting kept
doc.move_block("b6", after="b2")                  # a tracked move: the paragraph keeps its id
print(doc.revisions())                            # what is pending, including other authors' changes
doc.save("contract-redline.docx")

doc.reject_revisions(author="Review Agent")       # undo them again; or by ids=[...] or block="b4"
doc.accept_revisions()                            # no arguments: everything that is pending
```

- Text, paragraphs, tables, table rows and table columns are tracked as insertions and deletions; character formatting, paragraph formatting, style changes and table formatting (cell shading and alignment, widths, header rows) are tracked as formatting changes.
- An inserted or deleted column is recorded cell by cell (`w:cellIns` / `w:cellDel`). Merging and unmerging cells cannot be recorded and are refused while tracking is on.
- The agent (and `doc.text()`, the outline and block ids) always sees the document as it would read with all changes accepted.
- Revisions already in the file from other authors are preserved. Undoing something inserted in the same session removes it rather than recording a deletion.
- With tracking on, the agent is told to make minimal, targeted edits so the redline stays reviewable.
- The saved file has Word's own Track Changes switched on (`w:trackRevisions` in `word/settings.xml`), so the reader's edits are tracked too. `doc.track_changes = False` turns both off.
- `replace_text` and `search` also cover headers, footers, footnotes and endnotes, and those edits are tracked the same way. `doc.stories()` returns their text.
- A moved paragraph is recorded as a move. Word has no move markup for tables, so a moved table is a deletion plus an insertion.
- `accept_revisions` and `reject_revisions` resolve insertions, deletions, moves and formatting changes from any author. The agent is told not to use them unless asked. The chat UI has Accept all and Reject all buttons.

Not covered: edits made through `write_part_xml`; accepting or rejecting tracked cell merges (`w:cellMerge`), section and table-grid changes and numbering changes, which are left in place (so rejecting a tracked column-width change restores the cell widths but not the `w:tblGrid` values); and pending revisions inside a paragraph you move (your own earlier edits included), which read as plain text in the moved copy. They are kept in the half left behind and come back if the move is rejected.

## Scope and limitations

- Editable through the semantic tools: body paragraphs and tables, paragraph styles, character and paragraph formatting.
- Tables: rows, columns, merged cells, nested tables, column widths, cell shading and vertical alignment, repeating header rows. Not covered: cell borders and margins, text direction, and the legacy `w:hMerge` merge markup (such cells read as separate cells).
- Headers, footers, footnotes and endnotes: their text can be read, searched and replaced, but not restyled or restructured.
- Preserved on round trip but not editable through the semantic tools: images, comments, content controls. They can be reached through the raw part tools.
- Tracked output has been checked with python-docx and macOS `textutil`, not yet in Word itself. The move markup and the tracked column markup are the parts most likely to need adjusting.
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
