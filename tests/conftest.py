from __future__ import annotations

import copy
import io
import json
import zipfile

import pytest

from word_agent import WordDocument, WordPackage

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def template_parts() -> dict[str, bytes]:
    package = WordPackage.open()
    return {name: package.part_bytes(name) for name in package.list_parts()}


def make_docx(body_xml: str, extra_parts: dict[str, bytes] | None = None) -> bytes:
    """A .docx built from the template with a custom document body."""
    parts = template_parts()
    parts["word/document.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {W}><w:body>{body_xml}<w:sectPr/></w:body></w:document>"
    ).encode()
    parts.update(extra_parts or {})
    return zip_parts(parts)


def zip_parts(parts: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buffer.getvalue()


def doc_from_body(body_xml: str) -> WordDocument:
    return WordDocument(WordPackage.from_bytes(make_docx(body_xml)))


@pytest.fixture
def doc() -> WordDocument:
    return WordDocument.open()


def reopen_with_python_docx(document: WordDocument, tmp_path):
    """Save and reload with python-docx as an independent check of validity."""
    import docx

    path = tmp_path / "out.docx"
    document.save(path)
    return docx.Document(str(path))


# ---------------------------------------------------------------- fake LLM

def call(name: str, _id: str | None = None, **arguments) -> dict:
    return {
        "id": _id or f"call_{name}_{abs(hash(json.dumps(arguments, sort_keys=True))) % 10_000}",
        "type": "function",
        "function": {"name": name, "arguments": json.dumps(arguments)},
    }


def assistant(*calls: dict, content: str | None = None) -> dict:
    message: dict = {"role": "assistant", "content": content}
    if calls:
        message["tool_calls"] = list(calls)
    return message


class ScriptedLLM:
    """An LLMClient that replays a fixed list of assistant messages."""

    def __init__(self, turns: list[dict]):
        self.turns = list(turns)
        self.requests: list[dict] = []

    async def acomplete(self, messages, tools=None):
        self.requests.append({"messages": copy.deepcopy(messages), "tools": tools})
        if not self.turns:
            raise AssertionError("ScriptedLLM ran out of turns")
        return self.turns.pop(0)
