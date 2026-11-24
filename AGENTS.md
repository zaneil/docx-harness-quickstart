# AGENTS.md

Guidance for coding agents (Claude Code, Cursor, Codex, Copilot and similar) working in or building on this repository.

## What this repo is

A starter kit for LLM agents that draft and edit Microsoft Word `.docx` files. It is meant to be cloned and adapted: keep the parts you need, replace the rest. Python 3.10+, three dependencies (`lxml`, `openai`, `python-dotenv`).

Three layers with a strict one-way dependency: `agent` → `tools` → `document`.

| Layer | Path | Responsibility | Must not import |
| --- | --- | --- | --- |
| Document | `src/docx_harness/document/` | The unzipped `.docx` and block-level operations | `tools`, `agent`, any LLM code |
| Tools | `src/docx_harness/tools/` | LLM-callable functions with JSON schemas, in groups (`default_tools(doc, groups=...)`) | `agent` |
| Agent | `src/docx_harness/agent/` | Generic async tool-calling loop, plan mode, LLM client | `document` (it knows nothing about Word) |

`src/docx_harness/api.py` wires the three together (`create_agent`, `draft`, `adraft`, `adraft_many`). `cli.py` is a thin front end over `api.py`. `demo/` (repo root) is a demo chat UI over the public API; it is not part of the library or the wheel.

## Commands

```bash
uv sync                                   # install (creates .venv)
uv run pytest                             # full test suite, about 3 seconds, no network or API key needed
uv run pytest tests/test_document.py -k replace   # one area
uv run docx-harness "Draft a memo about X" -o memo.docx          # needs OPENROUTER_API_KEY in .env
uv run docx-harness "Fix the typos" -i in.docx -o out.docx --plan
uv run python -m demo                     # demo chat UI (not shipped) at http://127.0.0.1:8765
uv run python -m demo --trace-dir traces  # also write each conversation to a JSONL trace (CLI: --trace PATH)
uv run python scripts/build_template.py   # regenerate the bundled blank.docx after editing the script
uv run python -m evals check              # offline: every eval task fails untouched and passes with its solution
uv run python -m evals run -k fix-typos   # live eval on a real model; needs OPENROUTER_API_KEY and costs tokens
```

Run `uv run pytest` before finishing any change. Tests use a scripted fake model (`ScriptedLLM` in `tests/conftest.py`), so they never call a real API.

## Where to make a change

| Task | Edit | Then test in |
| --- | --- | --- |
| Add a tool the model can call | The group module in `tools/builtin/` (`read.py`, `write.py`, `tables.py`, `formatting.py`, `revisions.py`, `raw.py`; shared helpers in `common.py`), or pass `extra_tools=[...]` to `create_agent` without touching the repo. `@tool(available=check)` offers it only while `check()` is true | `tests/test_tools.py` |
| Add a document capability (e.g. headers, images, comments) | `document/document.py` holds the core (block ids, reading, inserting, moving, text edits, styles); each other area is a mixin `WordDocument` inherits: `stories.py` (headers, footers, sections, reading notes), `notes.py` (creating and deleting footnotes and endnotes), `lists.py`, `table_ops.py`, `media.py` (pictures, links), `revisions.py`. Shared names are in `document/common.py`, XML helpers in `ooxml.py` or `styling.py`. Then expose it as a tool | `tests/test_document.py` |
| Change how tables are read, addressed or laid out (merged cells, columns, widths, cell formatting) | `document/tables.py` (the grid and layout helpers), table methods in `document/table_ops.py` | `tests/test_tables.py` |
| Change how the loop runs, plan mode, concurrency | `agent/harness.py` | `tests/test_harness.py` |
| Use another model or provider | Implement `LLMClient.acomplete` (see `agent/llm.py`), or use `OpenAICompatibleClient(model, api_key=..., base_url=..., max_retries=4, timeout=120)` (retries with backoff are in `_create` and the `openai` SDK) | `tests/test_harness.py` |
| Change what the model is told | `agent/prompts.py` | live run |
| Add styles to new documents, or well-known styles added to files that lack them | `document/builtin_styles.py` (`TEMPLATE` goes into `blank.docx`, `EXTRA` is created on demand), then rerun `scripts/build_template.py` | `tests/test_content.py` |
| Change lists (levels, restarts, the list definitions) | `document/numbering.py`; `_place_in_list`, `set_list_level`, `restart_numbering` in `document/lists.py` | `tests/test_content.py` |
| Change hyperlinks, pictures, headers and footers, sections, footnotes | `document/media.py` (`add_hyperlink`, `insert_image`), `document/stories.py` (`_set_header_footer`, `_set_section_header_footer`, `insert_section_break`), `document/notes.py` (`add_note`, `delete_note`), `document/images.py`; new parts and relationships go through `WordPackage.add_part` / `relate` | `tests/test_content.py`, `tests/test_package.py` |
| Change how tracked changes (redlines) are recorded, accepted or rejected | `document/edits.py` (`EditContext`); listing and selection in `document/revisions.py` (`_resolve`); reading helpers in `document/ooxml.py` | `tests/test_tracked.py` |
| Change the demo UI (not part of the library) | `demo/server.py`, `demo/chat.html` | `tests/test_ui.py` |
| Add an eval task | One YAML file in `evals/tasks/<category>/` (format in `evals/README.md`); find block ids for its `solution` with `uv run python -m evals show <id>` | `uv run python -m evals check <id>`, `tests/test_evals.py` |
| Add an eval check or input step | `evals/checks.py` (`@check`), `evals/builder.py` (`@step`); add a case to `CASES` in the test and a row in `evals/README.md` | `tests/test_evals.py` |

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
- Model behaviour: after changing a prompt in `agent/prompts.py` or a tool's description or arguments, run `uv run python -m evals run` and compare with the table in the README. `evals/` sits above `api.py`, imports only the public `docx_harness` API, and is not shipped in the wheel. A live run is never part of `uv run pytest`.
- Demo UI: `demo/` likewise imports only the public `docx_harness` API and is not in the wheel; keep library code out of it.

## Known limits (do not assume these work)

- Editable through the built-in tools: body paragraphs, tables, paragraph styles, character and paragraph formatting, lists, hyperlinks, inline pictures, headers and footers.
- Tables: rows, columns, merging and unmerging, nested tables, column widths, cell shading, borders, padding, text direction and vertical alignment, repeating header rows. Cell text is one paragraph per line. A border between two cells is written on both (`_draw_borders`). Legacy `w:hMerge` is read by `build_grid` as one cell and rewritten as `w:gridSpan` by `_modernise_merges` before a structural edit; a cell that combines it with `w:vMerge` reads row by row. Not handled: merging or unmerging while tracking (raises `DocumentError`).
- Lists: a list item is a paragraph whose own `w:numPr` or whose style gives it a numbering id; read that through `WordDocument._list_item`, never from the style name. A numbered paragraph inserted away from other numbered items gets a new `w:num` (it starts at 1); `restart_numbering` re-points every later item of the same list. New definitions are found again by `w:name` (`docx-harness bullet` / `docx-harness number`). Number formats and indents of existing lists are not editable.
- Hyperlinks: external addresses only (`https`, `http`, `mailto`); bookmark links are read as `#name`, not created. Pictures: PNG, JPEG or GIF, inline, in a body paragraph of their own; an existing picture can be resized or deleted, not replaced, and its media part is not removed with it.
- Headers and footers: `set_header` / `set_footer` write the default, first-page or even-page one (text, `{PAGE}` / `{NUMPAGES}` as `w:fldSimple`). Without `section`, on the existing parts of every section that references one, otherwise on a new part referenced from the first section. With `section=N` (numbered from 1), only section N changes: a part it shares with another reference is copied first (`WordPackage.duplicate_part`); when N inherits (Word's "link to previous"), section N+1 is first pinned to what it showed, a new blank part if nothing. Find sections through `_sections()` / `_section_spans()` / `_effective_reference`, never by iterating `w:sectPr`: a section ends at a body paragraph whose `w:pPr` holds a `w:sectPr`, and the body's own `w:sectPr` ends the last. `insert_section_break` adds an empty paragraph carrying a copy of the split section's `w:sectPr`; the original keeps ending the second half and gets the new `w:type`. Page size, margins and columns are not editable. Header, footer, footnote and endnote text has no block ids: a part name such as `word/header1.xml` works as `ref` in `search`, `replace_text`, `format_text`, `set_paragraph_format`, `set_paragraph_style` and `add_hyperlink`.
- Footnotes and endnotes: `add_note` creates the notes part when missing (separators `-1` and `0`, `w:footnotePr` / `w:endnotePr` in settings), appends a note with the next free id, and inserts the reference run after `match` or at the end of the paragraph; body paragraphs only. `delete_note` removes the mark and, once nothing refers to it, the note. Note marks have no text, so `paragraph_text` and offsets ignore them; `read` reports them under `notes`. Numbering format and position settings are not editable.
- Comments and content controls are preserved on save but have no semantic tools; they are reachable only through `list_parts` / `read_part_xml` / `write_part_xml` (the last requires `include_raw=True`).
- Not recorded while tracking: style and list definitions, a newly created header or footer part (including a section's copy of a shared one and a pinned blank one), section properties (`w:sectPrChange` is not written), re-pointing an existing link, `set_image`. A section break is a tracked paragraph insertion; rejecting it removes the empty paragraph with its break. A new note is tracked as its mark plus its inserted text; rejecting it leaves the empty note in the notes part, unreferenced. Linking or unlinking text is recorded as a deletion plus an insertion of the same text.
- `WordDocument.open(path, track_changes=True, author=...)` records edits as revisions. A tracked-deleted block, table row or table cell (`w:cellDel`) stays in the XML but has no block id and is skipped by `_block_elements`, `tables.live_rows`, `tables.row_cells` and `_live_paragraphs`; do not iterate `w:p` / `w:tr` / `w:tc` directly when reading. While a tracked column deletion is pending, the `w:tblGrid` still has its column: map a grid column to it with `Grid.physical_of`. The same holds for the source half of a tracked move (`w:moveFrom`); a tracked move re-binds the block id to the copy at the destination. Setting `doc.track_changes` also writes `w:trackRevisions` to the settings part (skipped when the file has none).
- `accept_revisions` / `reject_revisions` handle `w:ins`, `w:del`, `w:moveFrom` / `w:moveTo`, `w:cellIns` / `w:cellDel` and the `w:rPrChange`, `w:pPrChange`, `w:trPrChange`, `w:tcPrChange`, `w:tblPrChange` family. `w:cellMerge`, `w:sectPrChange`, `w:tblGridChange` and `w:numberingChange` are left in place, and `w:tblGridChange` is never written, so rejecting a tracked width change restores `w:tcW` but not the grid column values. Block ids can change after either call (a rejected deletion comes back under a new id).
- Output is verified with python-docx and `textutil`, not in Word. python-docx does not see pictures or fields inside revision marks, so tracked tests check those through `WordDocument.read`.
- No layout engine: page counts and line breaks are unknown.
- A style the document lacks is created only when it is one of the well-known ones in `document/builtin_styles.py`; any other name raises `StyleError` listing the styles it has. `create_style` defines a new one.
- Legacy `.doc` is not supported.
