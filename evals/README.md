# Evals

Tasks a real model is asked to do with the docx-harness tools, each graded by automatic checks on the document it produces.

```bash
uv run python -m evals list                         # the tasks
uv run python -m evals check                        # offline: prove every task with its reference solution
uv run python -m evals run                          # live, default model (needs OPENROUTER_API_KEY)
uv run python -m evals run --model openai/gpt-6.1-terra --model anthropic/claude-haiku-5.5
uv run python -m evals run -k typos --no-judge      # only tasks whose id contains "typos", no judge model
uv run python -m evals report evals/results/<run> --markdown    # the README table
```

A live run writes `evals/results/<time>_<model>/` (gitignored): `results.json`, and for each task `input.docx`, `output.docx` and `trace.json` (every message and tool call). Open the output and the trace when a task fails, to tell a model mistake from a wrong check.

A task passes when every check passes. The judge score (1 to 5) is reported beside it and never decides a pass.

## Add a task

Create `evals/tasks/<category>/<id>.yaml`. The file name is the task id and the folder is its category.

```yaml
instruction: Fix the spelling mistakes. Do not change anything else.
input:
  - heading: Status report
  - paragraph: Teh report is redy.
checks:
  - text_contains: The report is ready.
  - unchanged: [Status report]
  - block_count: 2
judge: The corrected sentence reads naturally.      # optional
solution:                                           # the tool calls a perfect model would make
  - replace_text: {find: Teh, replace: The}
  - replace_text: {find: redy, replace: ready}
```

Then:

```bash
uv run python -m evals show <id>     # the starting document with block ids (b1, b2...) for the solution
uv run python -m evals check <id>    # checks must fail on the untouched input and pass after the solution
uv run python -m evals run -k <id>   # try it on a real model
```

`uv run pytest` runs the same `check` on every task file, so a task with a check that cannot pass, or cannot fail, does not get merged. A misspelt key, input step, check or tool name is an error on load that lists the valid names.

### Keys

| Key | Meaning |
| --- | --- |
| `instruction` | What the model is asked to do. Required. |
| `checks` | List of checks, each `- name: argument`. Required. |
| `input` | Steps that build the starting document (below), or `{file: fixtures/name.docx}` for a file in `evals/fixtures/`. Omit for a blank document. |
| `solution` | Tool calls, each `- tool_name: {arguments}`, that solve the task. Optional but expected. |
| `judge` | A rubric for the judge model. Use it for what checks cannot see, such as prose quality. |
| `track_changes`, `author` | Run with tracked changes on, as that author. |
| `plan` | Run in plan mode (the plan is approved automatically). |
| `max_steps` | Model calls allowed (default 40). |
| `id`, `category` | Override the file and folder names. |

### Input steps

| Step | Builds |
| --- | --- |
| `title: text` | A paragraph in the Title style |
| `heading: text` (option `level: 2`) | A heading, level 1 unless given |
| `paragraph: text` (options `style`, `level`) | A paragraph |
| `bullets: [a, b]` / `numbered: [a, b]` | List items |
| `table: [[a, b], [c, d]]` (options `header: false`, `merge: [[row, col, end_row, end_col]]`) | A table; quote numbers (`"42"`) |
| `page_break: true` | A page break |
| `header: text` / `footer: text` | The page header or footer; `{PAGE}` and `{NUMPAGES}` are fields |
| `section_break: text` | A new section starts after the paragraph containing the text |
| `section_header: {section, text}` | That section's own page header |
| `footnote: {after, text}` | A footnote after text already in the document |
| `link: {text, url}` | Links text already in the document |
| `format: {text, bold: true, color: FF0000}` | Formats text already in the document |
| `tracked: {author, steps}` | Pending revisions by another author. Steps: `replace: [old, new]`, `insert: text` (with `after: text of the paragraph above`), `delete: text of the paragraph` |

### Checks

Text is compared with curly quotes, long dashes and runs of spaces normalised. Where a count is expected, give a number or `{min, max}`.

| Check | Passes when |
| --- | --- |
| `text_contains: [a, b]` | Every string is in the body. `{text: [...], case: false}` ignores case |
| `text_contains_any: [a, b]` | At least one is (case ignored) |
| `text_absent: [a, b]` | None is. `{text: [...], case: false}` ignores case |
| `regex: pattern` | The pattern matches (`^` and `$` are a paragraph's start and end). `{pattern, min, max}` bounds the matches; `max: 0` forbids it |
| `word_count: {min, max}` | The body has that many words |
| `section_words: {heading, min, max}` | The text under that heading, up to the next heading, has that many words |
| `unchanged: [a, b]` | Each string is still the exact text of a paragraph or table cell |
| `unchanged_except: [a]` | Every input paragraph is still there word for word, except those containing a listed string. `[]` means nothing may change |
| `order: [a, b, c]` | The strings appear in that order |
| `heading: text` | A heading contains it. `{text, level}` fixes the level; level 0 is the Title style |
| `heading_count: 3` | That many headings (Title not counted). `{level, min, max}` counts one level |
| `style_of: {text, style}` | The paragraph with that text uses the style (a name or a list of acceptable names) |
| `block_count: 4` | That many non-empty paragraphs and tables in the body |
| `list_items: {kind, items, min, max, level}` | Each of `items` is a list item of that `kind` (bullet or number); `min` / `max` bound how many there are |
| `paragraph_format: {text, alignment: center}` | The paragraph has that direct layout |
| `table_count: 1` | That many tables |
| `table: {...}` | One table, picked by `with_text` or `index` (default 0), has `rows` / `cols`, `cells: [{at: [row, col], text}]`, `has_row: [[...]]`, `contains` / `absent`, `shaded: [[row, col]]`, `shaded_row: 0` |
| `formatting: {text, bold: true, color: FF0000}` | Every occurrence of the text is formatted so. `only: true` also requires the rest of the paragraph not to be |
| `link: {text, url}` | Exactly those words link to the address. `exact: false` accepts a longer link containing them |
| `header_contains: a` / `footer_contains: a` | The header or footer has the text |
| `note: {text, after}` | A footnote containing `text` has its mark in the paragraph containing `after`. `kind: endnote` for an endnote |
| `section_header: {section, contains, absent, starts_with}` | Section `section` (from 1) shows a header with the `contains` strings and none of `absent`, and its first paragraph starts with `starts_with`. `footer: true` checks the footer |
| `revisions: 0` | That many pending tracked changes. `{author, type, min, max}` counts some (type: insert, delete, format, move) |
| `reject_all_restores_input: true` | Rejecting the tracked changes gives back the input, so every edit was recorded. `{author: name}` rejects only theirs |

Two checks run on every task without being listed: the agent finished (did not run out of steps), and the output opens with python-docx.

## Add a check or an input step

- A check is a function in `checks.py` decorated with `@check("name")`. It receives a `Context` (`ctx.out` and `ctx.inp` are the output and input as `WordDocument`s; `ctx.blocks`, `ctx.paragraphs`, `ctx.tables`, `ctx.lines` and `ctx.text` are read from the output) and the argument from the YAML, and returns `None` to pass or a short reason. Add a passing and a failing case to `CASES` in `tests/test_evals.py` (a test fails until you do) and a row to the table above.
- An input step is a function in `builder.py` decorated with `@step("name", "option", ...)`.

## Files

| File | Holds |
| --- | --- |
| `loader.py` | YAML to `Task`, validation |
| `builder.py` | `input:` steps to a document |
| `checks.py` | The checks |
| `judge.py` | The judge prompt and reply parsing |
| `runner.py` | `run_suite` (live), `solve` and `verify` (offline) |
| `report.py` | Console table, `results.json`, markdown table |
| `__main__.py` | The command line |
