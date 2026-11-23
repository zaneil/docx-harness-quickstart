# Demo chat UI

A local page for trying the docx-harness agent by hand. **It is a demo and test tool, not part of the library:** it lives at the repo root, is not installed with the package, and uses only the public `docx_harness` API (the same way `evals/` does). Use it as an example of wiring the library into an app, not as something to deploy.

```bash
uv run python -m demo                       # opens http://127.0.0.1:8765
uv run python -m demo --trace-dir traces    # also write each conversation to traces/<time>_<id>.jsonl
uv run python -m demo --port 8768 --no-browser --model openai/gpt-6-luna
```

It needs `OPENROUTER_API_KEY` in `.env` (or the environment), like the CLI.

Chat on the left, a live preview of the document on the right. Start from the blank template or upload a `.docx`, keep chatting to refine it (the conversation and block ids carry over between messages), tick **Plan first** to review a plan before anything changes, switch on **Track changes** to record edits as revisions (with Accept all / Reject all buttons), and download the result.

Limits:

- One in-memory session for the whole server; a reset or an upload starts over.
- Bound to `127.0.0.1`, with host and header checks; uploads up to 50 MB.
- The preview is approximate (no layout engine, pictures shown as sized placeholders). Open the downloaded file in Word for the real rendering.
- `insert_image` reads pictures only from the folder the server was started in.

Files: `server.py` (standard-library HTTP server over `create_agent`), `chat.html` (the page). Tests: `tests/test_ui.py`.
