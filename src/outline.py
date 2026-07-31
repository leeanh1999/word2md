"""Heading outline of a Markdown document - the equivalent of Word's
Navigation Pane - plus extraction of the selected sections.

A section owns every line from its heading up to (but excluding) the next
heading of the same or a higher rank, so selecting a node always pulls in its
whole subtree.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable, Iterator, Sequence

PREAMBLE_ID = "0"
PREAMBLE_TITLE = "(Phần mở đầu)"
UNTITLED = "(không có tiêu đề)"

_HEADING_RE = re.compile(r"^(#{1,6})[ \t]+(.*)$")
_FENCE_RE = re.compile(r"^\s*(```|~~~)")
_INVALID_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_BLANK_LINES_RE = re.compile(r"\n{3,}")


@dataclass
class OutlineNode:
    node_id: str
    level: int
    title: str
    start: int
    end: int
    children: list["OutlineNode"] = field(default_factory=list)

    @property
    def is_preamble(self) -> bool:
        return self.level == 0

    @property
    def line_count(self) -> int:
        return max(0, self.end - self.start)

    def walk(self) -> Iterator["OutlineNode"]:
        yield self
        for child in self.children:
            yield from child.walk()


def iter_nodes(roots: Sequence[OutlineNode]) -> Iterator[OutlineNode]:
    for root in roots:
        yield from root.walk()


def index_nodes(roots: Sequence[OutlineNode]) -> dict[str, OutlineNode]:
    return {node.node_id: node for node in iter_nodes(roots)}


# ------------------------------------------------------------------ parsing


def heading_positions(lines: Sequence[str]) -> list[tuple[int, int, str]]:
    """Return (line index, level, title) for every ATX heading outside code fences."""
    headings: list[tuple[int, int, str]] = []
    fence: str | None = None

    for index, line in enumerate(lines):
        fence_match = _FENCE_RE.match(line)
        if fence_match:
            token = fence_match.group(1)
            if fence is None:
                fence = token
            elif fence == token:
                fence = None
            continue
        if fence is not None:
            continue
        match = _HEADING_RE.match(line)
        if match:
            headings.append((index, len(match.group(1)), match.group(2).strip()))

    return headings


def build_outline(markdown: str) -> list[OutlineNode]:
    """Build the heading tree for a Markdown document."""
    lines = markdown.split("\n")
    headings = heading_positions(lines)
    roots: list[OutlineNode] = []
    stack: list[OutlineNode] = []

    first_heading_line = headings[0][0] if headings else len(lines)
    if any(line.strip() for line in lines[:first_heading_line]):
        roots.append(
            OutlineNode(
                node_id="",
                level=0,
                title=PREAMBLE_TITLE,
                start=0,
                end=first_heading_line,
            )
        )

    for position, (line_index, level, title) in enumerate(headings):
        end = len(lines)
        for following in headings[position + 1 :]:
            if following[1] <= level:
                end = following[0]
                break

        node = OutlineNode(
            node_id="",
            level=level,
            title=title or UNTITLED,
            start=line_index,
            end=end,
        )

        while stack and stack[-1].level >= level:
            stack.pop()
        if stack:
            stack[-1].children.append(node)
        else:
            roots.append(node)
        stack.append(node)

    _assign_ids(roots)
    return roots


def _assign_ids(nodes: Sequence[OutlineNode], prefix: str = "") -> None:
    for position, node in enumerate(nodes, start=1):
        node.node_id = f"{prefix}.{position}" if prefix else str(position)
        _assign_ids(node.children, node.node_id)


# --------------------------------------------------------------- extraction


def _is_descendant(node_id: str, ancestor_id: str) -> bool:
    return node_id.startswith(f"{ancestor_id}.")


def top_level_selection(
    roots: Sequence[OutlineNode], node_ids: Iterable[str]
) -> list[OutlineNode]:
    """Drop ids already covered by a selected ancestor, then order by position."""
    index = index_nodes(roots)
    wanted = [node_id for node_id in dict.fromkeys(node_ids) if node_id in index]
    selected = set(wanted)
    kept = [
        index[node_id]
        for node_id in wanted
        if not any(_is_descendant(node_id, other) for other in selected)
    ]
    return sorted(kept, key=lambda node: node.start)


def shift_headings(text: str, shift: int) -> str:
    """Promote every heading by `shift` ranks, never above h1."""
    if shift <= 0:
        return text

    output: list[str] = []
    fence: str | None = None
    for line in text.split("\n"):
        fence_match = _FENCE_RE.match(line)
        if fence_match:
            token = fence_match.group(1)
            if fence is None:
                fence = token
            elif fence == token:
                fence = None
            output.append(line)
            continue
        match = _HEADING_RE.match(line) if fence is None else None
        if match:
            level = max(1, len(match.group(1)) - shift)
            output.append(f"{'#' * level} {match.group(2)}")
        else:
            output.append(line)
    return "\n".join(output)


def section_text(
    markdown: str, node: OutlineNode, promote: bool = False, shift: int | None = None
) -> str:
    """Return the Markdown of a single section (heading plus everything under it)."""
    lines = markdown.split("\n")
    text = "\n".join(lines[node.start : node.end]).strip("\n")
    if promote:
        if shift is None:
            shift = max(0, node.level - 1)
        text = shift_headings(text, shift)
    return _BLANK_LINES_RE.sub("\n\n", text).strip()


def extract_sections(
    markdown: str,
    node_ids: Iterable[str],
    roots: Sequence[OutlineNode] | None = None,
    promote: bool = False,
) -> str:
    """Concatenate the selected sections into one Markdown document."""
    roots = build_outline(markdown) if roots is None else roots
    selected = top_level_selection(roots, node_ids)
    if not selected:
        return ""

    shift = 0
    if promote:
        levels = [node.level for node in selected if node.level > 0]
        shift = min(levels) - 1 if levels else 0

    blocks = [
        section_text(markdown, node, promote=promote, shift=shift) for node in selected
    ]
    body = "\n\n".join(block for block in blocks if block)
    return _BLANK_LINES_RE.sub("\n\n", body).strip() + "\n"


# ------------------------------------------------------------------- naming


def slugify_title(title: str, fallback: str = "section") -> str:
    """Turn a heading into a safe file name fragment."""
    text = _INVALID_FILENAME_RE.sub("", title)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:80].strip(" .") or fallback


def outline_to_text(roots: Sequence[OutlineNode], indent: str = "  ") -> str:
    """Render the outline for the command line."""
    lines = []
    for node in iter_nodes(roots):
        depth = node.node_id.count(".")
        prefix = indent * depth
        rank = "—" if node.is_preamble else f"H{node.level}"
        lines.append(
            f"{node.node_id:<10}{prefix}{node.title}  [{rank}, {node.line_count} dòng]"
        )
    return "\n".join(lines)
