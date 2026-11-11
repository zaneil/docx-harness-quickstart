"""EditContext: the single path through which the document is mutated.

Every insert and delete of a run or a block goes through these four
primitives. That is what makes tracked changes (redlines) possible to add
without touching ``WordDocument`` or the tools: a tracked implementation only
has to wrap inserted content in ``w:ins`` and turn deletions into ``w:del``
here.

v1 implements direct edits only.
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from .ooxml import INS


@dataclass
class EditContext:
    track_changes: bool = False
    author: str = "word-agent"

    def __post_init__(self) -> None:
        if self.track_changes:
            raise NotImplementedError(
                "Tracked changes are not implemented yet. The hook is here: "
                "implement the w:ins / w:del path in EditContext."
            )

    # ------------------------------------------------------------------ runs

    def insert_run(self, parent: etree._Element, index: int, run: etree._Element) -> None:
        parent.insert(index, run)

    def delete_run(self, run: etree._Element) -> None:
        parent = run.getparent()
        if parent is None:
            return
        parent.remove(run)
        # Do not leave an empty tracked-insertion wrapper behind.
        if parent.tag == INS and len(parent) == 0 and parent.getparent() is not None:
            parent.getparent().remove(parent)

    # ---------------------------------------------------------------- blocks

    def insert_block(self, parent: etree._Element, index: int, block: etree._Element) -> None:
        parent.insert(index, block)

    def delete_block(self, block: etree._Element) -> None:
        parent = block.getparent()
        if parent is not None:
            parent.remove(block)
