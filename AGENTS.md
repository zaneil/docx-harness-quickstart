# AGENTS.md

Guidance for coding agents (Claude Code, Cursor, Codex, Copilot and similar) working in or building on this repository.

## What this repo is

A starter kit for LLM agents that draft and edit Microsoft Word `.docx` files. It is meant to be cloned and adapted: keep the parts you need, replace the rest. Python 3.10+, three dependencies (`lxml`, `openai`, `python-dotenv`).

Three layers with a strict one-way dependency: `agent` → `tools` → `document`.

| Layer | Path | Responsibility | Must not import |
| --- | --- | --- | --- |
| Document | `src/docx_harness/document/` | The unzipped `.docx` and block-level operations | `tools`, `agent`, any LLM code |
| Tools | `src/docx_harness/tools/` | LLM-callable functions with JSON schemas | `agent` |
| Agent | `src/docx_harness/agent/` | Generic async tool-calling loop, plan mode, LLM client | `document` (it knows nothing about Word) |

`src/docx_harness/api.py` wires the three together (`create_agent`, `draft`, `adraft`, `adraft_many`). `cli.py` and `ui/` are thin front ends over `api.py`.

## Commands

```bash
uv sync                                   # install (creates .venv)
uv run pytest                             # full test suite, about 3 seconds, no network or API key needed
uv run pytest tests/test_document.py -k replace   # one area
uv run docx-harness "Draft a memo about X" -o memo.docx          # needs OPENROUTER_API_KEY in .env
uv run docx-harness "Fix the typos" -i in.docx -o out.docx --plan
uv run docx-harness-ui                    # local chat UI at http://127.0.0.1:8765
uv run python scripts/build_template.py   # regenerate the bundled blank.docx after editing the script
```

Run `uv run pytest` before finishing any change. Tests use a scripted fake model (`ScriptedLLM` in `tests/conftest.py`), so they never call a real API.

## Where to make a change

| Task | Edit | Then test in |
| --- | --- | --- |
| Add a tool the model can call | `tools/builtin.py` (inside `default_tools`), or pass `extra_tools=[...]` to `create_agent` without touching the repo | `tests/test_tools.py` |
| Add a document capability (e.g. headers, images, comments) | `document/document.py`, helpers in `document/ooxml.py` or `document/styling.py`; then expose it as a tool | `tests/test_document.py` |
| Change how the loop runs, plan mode, concurrency | `agent/harness.py` | `tests/test_harness.py` |
| Use another model or provider | Implement `LLMClient.acomplete` (see `agent/llm.py`), or use `OpenAICompatibleClient(model, api_key=..., base_url=...)` | `tests/test_harness.py` |
| Change what the model is told | `agent/prompts.py` | live run |
| Add styles to new documents | `scripts/build_template.py`, then rerun it | `tests/test_document.py` |
| Change how tracked changes (redlines) are recorded, accepted or rejected | `document/edits.py` (`EditContext`); selection in `WordDocument._resolve`; reading helpers in `document/ooxml.py` | `tests/test_tracked.py` |
| Change the chat UI | `ui/server.py`, `ui/chat.html` | `tests/test_ui.py` |

## Minimal usage, for writing code against it

```python
from docx_harness import WordDocument, ToolRegistry, default_tools, create_agent, tool

doc = WordDocument.open("in.docx")          # or WordDocument.open() for the blank template

# 1. Direct, no LLM
block = doc.insert_paragraph("Summary", "Heading 1", after="end")   # returns a block id like "b7"
doc.replace_text("2024", "2025")
doc.format_text(block, bold=True, color="1F4E79")
doc.save("out.docx")

# 2. Tools only, for your own loop or framework
registry = ToolRegistry(default_tools(doc))
schemas = registry.to_openai_schema()                       # OpenAI function-calling format
text = registry.execute("get_outline", {})                  # always returns a string, never raises

# 3. Full agent
@tool(read_only=True)
def word_count() -> int:
    """Count the words in the document."""
    return len(doc.text().split())

agent = create_agent(doc, extra_tools=[word_count])         # needs OPENROUTER_API_KEY
result = agent.run("Shorten this to 300 words")             # or: await agent.arun(...)
result = agent.run("Now add a table", history=result.messages)   # continue the conversation
doc.save("out.docx")
```

## Rules that keep the code working

1. **Respect the layer boundaries** in the table above. The harness must stay free of Word-specific code, and the document layer free of LLM code.
2. **All mutations go through `EditContext`** (`document/edits.py`): `insert_run`, `delete_run`, `insert_block`, `delete_block`, `move_block`, and `resolve` for accepting or rejecting a revision. Do not call `parent.insert`, `parent.remove` or `append` on document content anywhere else, and call `edits.before_run_format(run)` / `edits.before_paragraph_format(p)` before changing properties. This single path is what records tracked changes when `track_changes=True`. Use `with doc.edits.untracked():` only while preparing content that is not in the document yet.
3. **Tools return strings and never raise to the caller.** `ToolRegistry.aexecute` converts exceptions to `"Error: ..."` so the model can recover. Write error messages that tell the model what to do next.
4. **Mark tools `@tool(read_only=True)` only if they change nothing.** Read-only tools are the only ones offered in plan mode, and they may run concurrently. Anything that mutates must not carry the flag.
5. **Block ids are in memory only.** `b1`, `b2`... are assigned on load and on insert, stay stable for the life of a `WordDocument`, and are not written to the file. Do not persist them across processes. Table cells are addressed as `"b7[row,col]"`, zero-based.
6. **Property elements must be in schema order.** When adding children to `w:rPr` or `w:pPr`, use `ensure_child(parent, tag, RPR_ORDER | PPR_ORDER)` from `document/ooxml.py`. Out-of-order properties can make Word reject the file.
7. **Read text through `iter_runs` / `paragraph_text`.** They skip tracked deletions (`w:del`) and include runs inside hyperlinks and tracked insertions. Iterating `w:r` directly gets this wrong.
8. **Parse XML only with `parse_xml`** in `document/package.py`. It disables entity expansion and network access and rejects DTDs. Do not call `etree.fromstring` on package content directly.
9. **The message format is OpenAI chat-completions dicts** end to end (`messages`, `tools`, `tool_calls`). Do not introduce a second message representation.
10. **Do not commit `.env`** or generated `.docx` files in the repo root (both are gitignored). `src/docx_harness/document/templates/blank.docx` is the one binary that is tracked.

## Verifying a change

- Unit level: add or extend a test next to the ones for that layer. For document changes, also reopen the saved file with `python-docx` (see `reopen_with_python_docx` in `tests/conftest.py`) as an independent check that the output is valid.
- Harness level: script the model with `ScriptedLLM([assistant(call("tool_name", arg=...)), assistant(content="done")])`.
- Live: `uv run docx-harness "..." -o out.docx -v` prints every tool call and result.

## Known limits (do not assume these work)

- Editable through the built-in tools: body paragraphs, tables, paragraph styles, character and paragraph formatting. Headers, footers, footnotes and endnotes are text only: `doc.stories()`, `search` and `replace_text` reach them (a part name such as `word/header1.xml` works as `ref`), nothing else does, and they have no block ids. Images, comments and content controls are preserved on save but have no semantic tools; they are reachable only through `list_parts` / `read_part_xml` / `write_part_xml` (the last requires `include_raw=True`).
- `WordDocument.open(path, track_changes=True, author=...)` records edits as revisions. A tracked-deleted block or table row stays in the XML but has no block id and is skipped by `_block_elements`, `_live_rows` and `_live_paragraphs`; do not iterate `w:p` / `w:tr` directly when reading. The same holds for the source half of a tracked move (`w:moveFrom`); a tracked move re-binds the block id to the copy at the destination. Setting `doc.track_changes` also writes `w:trackRevisions` to the settings part (skipped when the file has none).
- `accept_revisions` / `reject_revisions` handle `w:ins`, `w:del`, `w:moveFrom` / `w:moveTo` and the `w:rPrChange`, `w:pPrChange`, `w:trPrChange`, `w:tcPrChange`, `w:tblPrChange` family. `w:cellIns`, `w:cellDel`, `w:cellMerge`, `w:sectPrChange`, `w:tblGridChange` and `w:numberingChange` are left in place. Block ids can change after either call (a rejected deletion comes back under a new id).
- Tracked output is verified with python-docx and `textutil`, not in Word.
- No layout engine: page counts and line breaks are unknown.
- Styles are looked up, not created. Asking for a style the document lacks raises `StyleError` listing the styles it has.
- Legacy `.doc` is not supported.
