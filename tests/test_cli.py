from __future__ import annotations

from docx_harness import WordDocument, cli

from conftest import ScriptedLLM, assistant, call


def test_cli_saves_and_prints_usage(tmp_path, monkeypatch, capsys):
    llm = ScriptedLLM([
        assistant(call("insert_paragraphs", paragraphs=[{"text": "Hello"}])),
        assistant(content="Added a greeting."),
    ])
    llm.usage = {"prompt_tokens": 2000, "completion_tokens": 50, "requests": 2}
    monkeypatch.setattr(cli, "OpenRouterClient", lambda model=None: llm)
    out = tmp_path / "out.docx"

    assert cli.main(["say hello", "-o", str(out)]) == 0
    captured = capsys.readouterr()
    assert captured.out.strip() == "Added a greeting."
    assert "→ insert_paragraphs" in captured.err
    assert captured.err.strip().endswith("2 steps · 2.0k in / 50 out")
    assert WordDocument.open(out).text() == "Hello"
