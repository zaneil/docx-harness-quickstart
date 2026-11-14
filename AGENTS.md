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
| Change how tables are read, addressed or laid out (merged cells, columns, widths, cell formatting) | `document/tables.py` (the grid and layout helpers), table methods in `document/document.py` | `tests/test_tables.py` |
| Change how the loop runs, plan mode, concurrency | `agent/harness.py` | `tests/test_harness.py` |
| Use another model or provider | Implement `LLMClient.acomplete` (see `agent/llm.py`), or use `OpenAICompatibleClient(model, api_key=..., base_url=...)` | `tests/test_harness.py` |
| Change what the model is told | `agent/prompts.py` | live run |
| Add styles to new documents, or well-known styles added to files that lack them | `document/builtin_styles.py` (`TEMPLATE` goes into `blank.docx`, `EXTRA` is created on demand), then rerun `scripts/build_template.py` | `tests/test_content.py` |
| Change lists (levels, restarts, the list definitions) | `document/numbering.py`; `_place_in_list`, `set_list_level`, `restart_numbering` in `document/document.py` | `tests/test_content.py` |
| Change hyperlinks, pictures, headers and footers | `document/document.py` (`add_hyperlink`, `insert_image`, `_set_header_footer`), `document/images.py`; new parts and relationships go through `WordPackage.add_part` / `relate` | `tests/test_content.py`, `tests/test_package.py` |
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
2. **All mutations go through `EditContext`** (`document/edits.py`): `insert_run`, `delete_run`, `insert_block`, `delete_block`, `insert_cell`, `delete_cell`, `move_block`, `wrap_runs` / `unwrap_runs` (hyperlinks), and `resolve` for accepting or rejecting a revision. Do not call `parent.insert`, `parent.remove` or `append` on document content anywhere else, and call the matching hook before changing properties: `edits.before_run_format(run)`, `before_paragraph_format(p)`, `before_cell_format(tc)`, `before_row_format(tr)`, `before_table_format(tbl)`. (The `w:tblGrid` columns are layout, written by the helpers in `document/tables.py`. Style and list definitions, relationships and `w:sectPr` references are not content either and are written directly.) This single path is what records tracked changes when `track_changes=True`. Use `with doc.edits.untracked():` only while preparing content that is not in the document yet.
3. **Tools return strings and never raise to the caller.** `ToolRegistry.aexecute` converts exceptions to `"Error: ..."` so the model can recover. Write error messages that tell the model what to do next.
4. **Mark tools `@tool(read_only=True)` only if they change nothing.** Read-only tools are the only ones offered in plan mode, and they may run concurrently. Anything that mutates must not carry the flag.
5. **Block ids are in memory only.** `b1`, `b2`... are assigned on load and on insert, stay stable for the life of a `WordDocument`, and are not written to the file. Do not persist them across processes. Table cells are addressed as `"b7[row,col]"`, zero-based, on the table's column grid: a merged cell answers to every position it covers and its content is in its top-left `w:tc`. `"b7[0,*]"` is a row and `"b7[*,1]"` a column. A table nested in a cell gets a block id when it is first read; it works with the table methods and `delete_block`, but not as a position or with `move_block`. Go through `tables.build_grid` / `tables.live_rows` / `tables.row_cells` rather than indexing `w:tr` / `w:tc`, which miss merges, content-control wrappers and tracked-deleted cells.
6. **Property elements must be in schema order.** When adding children to `w:rPr`, `w:pPr`, `w:tcPr`, `w:tblPr`, `w:trPr` or `w:sectPr`, use `ensure_child(parent, tag, RPR_ORDER | PPR_ORDER | TCPR_ORDER | TBLPR_ORDER | TRPR_ORDER | SECTPR_ORDER)` from `document/ooxml.py`. Out-of-order properties can make Word reject the file.
7. **Read text through `iter_runs` / `paragraph_text`.** They skip tracked deletions (`w:del`) and include runs inside hyperlinks and tracked insertions. Iterating `w:r` directly gets this wrong. A run that holds a picture, a field or a note reference has no text; do not delete runs just because they are in a paragraph being rewritten (see `_set_paragraph_text`).
8. **Add parts and relationships only through `WordPackage.add_part` and `relate`.** They keep `[Content_Types].xml` and the `.rels` parts consistent; a part without a content type makes Word reject the file. Relationships belong to the part that uses them: a link in a header goes in that header's `.rels` (`WordDocument._part_of`).
9. **Parse XML only with `parse_xml`** in `document/package.py`. It disables entity expansion and network access and rejects DTDs. Do not call `etree.fromstring` on package content directly.
10. **The message format is OpenAI chat-completions dicts** end to end (`messages`, `tools`, `tool_calls`). Do not introduce a second message representation.
11. **Do not commit `.env`** or generated `.docx` files in the repo root (both are gitignored). `src/docx_harness/document/templates/blank.docx` is the one binary that is tracked.

## Verifying a change

- Unit level: add or extend a test next to the ones for that layer. For document changes, also reopen the saved file with `python-docx` (see `reopen_with_python_docx` in `tests/conftest.py`) as an independent check that the output is valid.
- Harness level: script the model with `ScriptedLLM([assistant(call("tool_name", arg=...)), assistant(content="done")])`.
- Live: `uv run docx-harness "..." -o out.docx -v` prints every tool call and result.

## Known limits (do not assume these work)

- Editable through the built-in tools: body paragraphs, tables, paragraph styles, character and paragraph formatting, lists, hyperlinks, inline pictures, headers and footers.
- Tables: rows, columns, merging and unmerging, nested tables, column widths, cell shading, borders, padding, text direction and vertical alignment, repeating header rows. Cell text is one paragraph per line. A border between two cells is written on both (`_draw_borders`). Legacy `w:hMerge` is read by `build_grid` as one cell and rewritten as `w:gridSpan` by `_modernise_merges` before a structural edit; a cell that combines it with `w:vMerge` reads row by row. Not handled: merging or unmerging while tracking (raises `DocumentError`).
- Lists: a list item is a paragraph whose own `w:numPr` or whose style gives it a numbering id; read that through `WordDocument._list_item`, never from the style name. A numbered paragraph inserted away from other numbered items gets a new `w:num` (it starts at 1); `restart_numbering` re-points every later item of the same list. New definitions are found again by `w:name` (`docx-harness bullet` / `docx-harness number`). Number formats and indents of existing lists are not editable.
- Hyperlinks: external addresses only (`https`, `http`, `mailto`); bookmark links are read as `#name`, not created. Pictures: PNG, JPEG or GIF, inline, in a body paragraph of their own; an existing picture can be resized or deleted, not replaced, and its media part is not removed with it.
- Headers and footers: `set_header` / `set_footer` write the default, first-page or even-page one (text, `{PAGE}` / `{NUMPAGES}` as `w:fldSimple`), on the existing part when a section already references one, otherwise on a new part referenced from the first section. Per-section headers are not created. Header, footer, footnote and endnote text has no block ids: a part name such as `word/header1.xml` works as `ref` in `search`, `replace_text`, `format_text`, `set_paragraph_format`, `set_paragraph_style` and `add_hyperlink`. Footnotes and endnotes cannot be created.
- Comments and content controls are preserved on save but have no semantic tools; they are reachable only through `list_parts` / `read_part_xml` / `write_part_xml` (the last requires `include_raw=True`).
- Not recorded while tracking: style and list definitions, a newly created header or footer part, re-pointing an existing link, `set_image`. Linking or unlinking text is recorded as a deletion plus an insertion of the same text.
- `WordDocument.open(path, track_changes=True, author=...)` records edits as revisions. A tracked-deleted block, table row or table cell (`w:cellDel`) stays in the XML but has no block id and is skipped by `_block_elements`, `tables.live_rows`, `tables.row_cells` and `_live_paragraphs`; do not iterate `w:p` / `w:tr` / `w:tc` directly when reading. While a tracked column deletion is pending, the `w:tblGrid` still has its column: map a grid column to it with `Grid.physical_of`. The same holds for the source half of a tracked move (`w:moveFrom`); a tracked move re-binds the block id to the copy at the destination. Setting `doc.track_changes` also writes `w:trackRevisions` to the settings part (skipped when the file has none).
- `accept_revisions` / `reject_revisions` handle `w:ins`, `w:del`, `w:moveFrom` / `w:moveTo`, `w:cellIns` / `w:cellDel` and the `w:rPrChange`, `w:pPrChange`, `w:trPrChange`, `w:tcPrChange`, `w:tblPrChange` family. `w:cellMerge`, `w:sectPrChange`, `w:tblGridChange` and `w:numberingChange` are left in place, and `w:tblGridChange` is never written, so rejecting a tracked width change restores `w:tcW` but not the grid column values. Block ids can change after either call (a rejected deletion comes back under a new id).
- Output is verified with python-docx and `textutil`, not in Word. python-docx does not see pictures or fields inside revision marks, so tracked tests check those through `WordDocument.read`.
- No layout engine: page counts and line breaks are unknown.
- A style the document lacks is created only when it is one of the well-known ones in `document/builtin_styles.py`; any other name raises `StyleError` listing the styles it has. `create_style` defines a new one.
- Legacy `.doc` is not supported.
