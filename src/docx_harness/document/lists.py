"""List items: placing paragraphs in lists, nesting levels and numbering restarts.

Mixed into ``WordDocument`` (``document.py``), which provides the block-id
helpers it relies on (``_get``, ``_body_block``, ``_table``, ``_paragraphs_for``,
``_describe``...) and ``self.edits``, the single path for every mutation."""

from __future__ import annotations

from lxml import etree

from . import numbering, styling
from .common import (
    BLOCK_TAGS,
    DocumentError,
    _live_paragraphs,
)
from .ooxml import (
    P,
    is_deleted_block,
)


class ListsMixin:
    # ----------------------------------------------------------------- lists

    def _list_item(self, p: etree._Element) -> tuple[int, int, str] | None:
        """``(numbering id, level, kind)`` when the paragraph is a list item,
        whether its own ``w:numPr`` or its style makes it one."""
        if p.tag != P:
            return None
        num_id, level = numbering.num_pr(p)
        if num_id is None or level is None:
            style_num, style_level = styling.style_num_pr(self._styles, styling.paragraph_style_id(p))
            num_id = style_num if num_id is None else num_id
            level = style_level if level is None else level
        kind = numbering.kind_of(self._numbering, num_id, level or 0)
        return (num_id, level or 0, kind) if num_id and kind else None

    def _check_level(self, num_id: int, level: int) -> int:
        try:
            level = int(level)
        except (TypeError, ValueError):
            raise DocumentError("level must be a whole number: 0 for the outer level, 1 for the first nested level.") from None
        levels = numbering.level_count(numbering.abstract_of(self._numbering, num_id))
        if not 0 <= level < max(levels, 1):
            raise DocumentError(f"level must be between 0 and {max(levels, 1) - 1} for this list.")
        return level

    def _place_in_list(self, p: etree._Element, index: int, level: int | None) -> None:
        """Before a new paragraph goes into the body at ``index``: set its
        list level, and decide which numbered list it belongs to."""
        item = self._list_item(p)
        if item is None:
            if level:
                raise DocumentError(
                    "level only applies to list items. Give the paragraph a list style such as "
                    "'List Bullet' or 'List Number'."
                )
            return
        num_id, own_level, kind = item
        level = own_level if level is None else self._check_level(num_id, level)
        target = num_id
        if kind == "number":
            target = self._neighbouring_list(index)
            if target is None:
                root = self._numbering_root()
                target = numbering.new_num(root, numbering.abstract_id_of(root, num_id), start=1)
        if (target, level) != (num_id, own_level):
            numbering.set_num_pr(p, target, level)

    def _neighbouring_list(self, index: int) -> int | None:
        """The numbered list a paragraph placed at body position ``index``
        joins: that of the nearest numbered item in the run of list items
        just before it, or failing that just after it."""
        live = [child for child in self._body if child.tag in BLOCK_TAGS and not is_deleted_block(child)]
        before = [child for child in live if self._body.index(child) < index]
        after = [child for child in live if self._body.index(child) >= index]
        for run in (reversed(before), after):
            for block in run:
                item = self._list_item(block)
                if item is None:
                    break
                if item[2] == "number":
                    return item[0]
        return None

    def set_list_level(self, ref: str, level: int) -> int:
        """Nest list items: ``level`` 0 is the outer level, 1 the first
        nested level, up to 8. ``ref`` is a block id or a table cell. Returns
        the number of paragraphs changed."""
        changed = 0
        for p in self._paragraphs_for(ref):
            item = self._list_item(p)
            if item is None:
                raise DocumentError(
                    f"{ref} is not a list item. Give it a list style first with set_paragraph_style "
                    "('List Bullet' or 'List Number')."
                )
            wanted = self._check_level(item[0], level)
            if wanted != item[1]:
                self.edits.before_paragraph_format(p)
                numbering.set_num_pr(p, item[0], wanted)
                changed += 1
        return changed

    def restart_numbering(self, block_id: str, start: int = 1) -> int:
        """Make a numbered list item begin a new count at ``start``. The
        items of the same list that come after it follow on from there.
        Returns the number of paragraphs renumbered."""
        first = self._paragraph(block_id)
        item = self._list_item(first)
        if item is None or item[2] != "number":
            raise DocumentError(f"Block {block_id} is not a numbered list item, so there is no numbering to restart.")
        if int(start) < 0:
            raise DocumentError("start must be 0 or more.")
        num_id, level, _ = item
        root = self._numbering_root()
        fresh = numbering.new_num(root, numbering.abstract_id_of(root, num_id), start=int(start), level=level)
        changed, reached = 0, False
        for p in _live_paragraphs(self._body):
            reached = reached or p is first
            current = self._list_item(p) if reached else None
            if current is not None and current[0] == num_id:
                self.edits.before_paragraph_format(p)
                numbering.set_num_pr(p, fresh, current[1])
                changed += 1
        return changed
