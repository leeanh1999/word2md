"""Convert the small HTML subset produced by mammoth into Markdown.

Mammoth ships a Markdown writer, but it is deprecated upstream and silently
drops tables, so we render its HTML output ourselves instead.
"""

from __future__ import annotations

import re
from html import unescape
from html.parser import HTMLParser

VOID_TAGS = {"br", "img", "hr", "meta", "link", "input", "col"}

HEADING_TAGS = {"h1": 1, "h2": 2, "h3": 3, "h4": 4, "h5": 5, "h6": 6}

BLOCK_TAGS = {
    "p",
    "div",
    "section",
    "article",
    "ul",
    "ol",
    "li",
    "table",
    "thead",
    "tbody",
    "tfoot",
    "tr",
    "td",
    "th",
    "blockquote",
    "pre",
    "hr",
    *HEADING_TAGS,
}

_WHITESPACE_RE = re.compile(r"[ \t\r\n\f\v]+")
_BLANK_LINES_RE = re.compile(r"\n{3,}")


class Node:
    __slots__ = ("tag", "attrs", "children", "text")

    def __init__(self, tag: str, attrs: dict[str, str] | None = None, text: str = ""):
        self.tag = tag
        self.attrs = attrs or {}
        self.children: list[Node] = []
        self.text = text

    @property
    def is_text(self) -> bool:
        return self.tag == "#text"

    def find_all(self, tag: str) -> list["Node"]:
        found = []
        for child in self.children:
            if child.tag == tag:
                found.append(child)
            found.extend(child.find_all(tag))
        return found


class _TreeBuilder(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = Node("#root")
        self._stack = [self.root]

    def handle_starttag(self, tag, attrs):
        node = Node(tag, {k: (v or "") for k, v in attrs})
        self._stack[-1].children.append(node)
        if tag not in VOID_TAGS:
            self._stack.append(node)

    def handle_startendtag(self, tag, attrs):
        self._stack[-1].children.append(Node(tag, {k: (v or "") for k, v in attrs}))

    def handle_endtag(self, tag):
        for index in range(len(self._stack) - 1, 0, -1):
            if self._stack[index].tag == tag:
                del self._stack[index:]
                return

    def handle_data(self, data):
        self._stack[-1].children.append(Node("#text", text=data))


def parse_html(html: str) -> Node:
    builder = _TreeBuilder()
    builder.feed(html)
    builder.close()
    return builder.root


class MarkdownRenderer:
    """Render a parsed HTML tree as Markdown."""

    def __init__(self, first_row_as_header: bool = True):
        self.first_row_as_header = first_row_as_header
        # Inside a table cell a real newline would break the row, so hard
        # breaks degrade to <br> which every Markdown renderer understands.
        self._in_table_cell = False

    def render(self, html: str) -> str:
        root = parse_html(html)
        body = "\n\n".join(self._render_children_as_blocks(root, indent=""))
        return _BLANK_LINES_RE.sub("\n\n", body).strip() + "\n"

    # ---------------------------------------------------------------- blocks

    def _render_children_as_blocks(self, node: Node, indent: str) -> list[str]:
        blocks: list[str] = []
        inline_run: list[Node] = []

        def flush_inline() -> None:
            if not inline_run:
                return
            text = self._render_inline(inline_run).strip()
            inline_run.clear()
            if text:
                blocks.append(_prefix_lines(text, indent))

        for child in node.children:
            if child.is_text:
                if child.text.strip():
                    inline_run.append(child)
                continue
            if child.tag in BLOCK_TAGS:
                flush_inline()
                blocks.extend(self._render_block(child, indent))
            else:
                inline_run.append(child)

        flush_inline()
        return [block for block in blocks if block.strip()]

    def _render_block(self, node: Node, indent: str) -> list[str]:
        tag = node.tag

        if tag in HEADING_TAGS:
            text = self._render_inline(node.children).strip()
            if not text:
                return []
            return [f"{indent}{'#' * HEADING_TAGS[tag]} {text}"]

        if tag == "p":
            text = self._render_inline(node.children).strip()
            return [_prefix_lines(text, indent)] if text else []

        if tag in ("ul", "ol"):
            return self._render_list(node, indent)

        if tag == "table":
            table = self._render_table(node)
            return [_prefix_lines(table, indent)] if table else []

        if tag == "blockquote":
            inner = "\n\n".join(self._render_children_as_blocks(node, indent=""))
            if not inner.strip():
                return []
            quoted = "\n".join(
                f"> {line}" if line else ">" for line in inner.split("\n")
            )
            return [_prefix_lines(quoted, indent)]

        if tag == "pre":
            code = _collect_raw_text(node).strip("\n")
            if not code.strip():
                return []
            return [_prefix_lines(f"```\n{code}\n```", indent)]

        if tag == "hr":
            return [f"{indent}---"]

        # div / section / stray li / table parts encountered out of context
        return self._render_children_as_blocks(node, indent)

    def _render_list(self, node: Node, indent: str) -> list[str]:
        ordered = node.tag == "ol"
        lines: list[str] = []
        counter = 0

        for item in node.children:
            if item.tag != "li":
                continue
            counter += 1
            marker = f"{counter}. " if ordered else "- "
            child_indent = " " * len(marker)

            item_blocks = self._render_children_as_blocks(item, indent="")
            if not item_blocks:
                continue

            first, *rest = item_blocks
            first_lines = first.split("\n")
            lines.append(f"{indent}{marker}{first_lines[0]}")
            for line in first_lines[1:]:
                lines.append(f"{indent}{child_indent}{line}" if line else "")
            for block in rest:
                is_nested_list = bool(re.match(r"\s*(?:[-*]|\d+\.)\s", block))
                if not is_nested_list:
                    lines.append("")
                for line in block.split("\n"):
                    lines.append(f"{indent}{child_indent}{line}" if line else "")

        return ["\n".join(lines)] if lines else []

    def _render_table(self, node: Node) -> str:
        rows: list[tuple[list[str], bool]] = []
        for tr in node.find_all("tr"):
            cells: list[str] = []
            has_header_cell = False
            for cell in tr.children:
                if cell.tag not in ("td", "th"):
                    continue
                has_header_cell = has_header_cell or cell.tag == "th"
                cells.append(self._render_cell(cell))
                span = _to_int(cell.attrs.get("colspan"), 1)
                cells.extend([""] * max(0, span - 1))
            if cells:
                rows.append((cells, has_header_cell))

        if not rows:
            return ""

        width = max(len(cells) for cells, _ in rows)
        grid = [cells + [""] * (width - len(cells)) for cells, _ in rows]

        if rows[0][1] or self.first_row_as_header:
            header, body = grid[0], grid[1:]
        else:
            header, body = [""] * width, grid

        lines = [_table_row(header), "| " + " | ".join(["---"] * width) + " |"]
        lines.extend(_table_row(row) for row in body)
        return "\n".join(lines)

    def _render_cell(self, cell: Node) -> str:
        previous = self._in_table_cell
        self._in_table_cell = True
        try:
            blocks = self._render_children_as_blocks(cell, indent="")
        finally:
            self._in_table_cell = previous
        text = "<br>".join(block.replace("\n", "<br>") for block in blocks)
        return text.replace("|", "\\|").strip()

    # ---------------------------------------------------------------- inline

    def _render_inline(self, nodes: list[Node]) -> str:
        return "".join(self._render_inline_node(node) for node in nodes)

    def _render_inline_node(self, node: Node) -> str:
        if node.is_text:
            return _WHITESPACE_RE.sub(" ", node.text)

        tag = node.tag

        if tag == "br":
            return "<br>" if self._in_table_cell else "  \n"

        if tag == "img":
            alt = node.attrs.get("alt", "").strip()
            src = node.attrs.get("src", "").strip()
            return f"![{alt}]({src})" if src else ""

        if tag in BLOCK_TAGS:
            blocks = self._render_children_as_blocks(node, indent="")
            return "\n\n".join(blocks)

        inner = self._render_inline(node.children)

        if tag in ("strong", "b"):
            return _wrap(inner, "**")
        if tag in ("em", "i"):
            return _wrap(inner, "*")
        if tag in ("s", "del", "strike"):
            return _wrap(inner, "~~")
        if tag == "code":
            return _wrap(inner, "`")
        if tag in ("sup", "sub"):
            return f"<{tag}>{inner}</{tag}>" if inner.strip() else ""
        if tag == "a":
            href = node.attrs.get("href", "").strip()
            label = inner.strip()
            if not href:
                return inner
            if not label:
                return f"<{href}>"
            return f"[{label}]({href})"

        return inner


def html_to_markdown(html: str, first_row_as_header: bool = True) -> str:
    return MarkdownRenderer(first_row_as_header=first_row_as_header).render(html)


# --------------------------------------------------------------------- utils


def _wrap(text: str, marker: str) -> str:
    """Apply an emphasis marker, keeping surrounding spaces outside of it."""
    if not text.strip():
        return text
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()) :]
    return f"{lead}{marker}{text.strip()}{marker}{trail}"


def _prefix_lines(text: str, indent: str) -> str:
    if not indent:
        return text
    return "\n".join(f"{indent}{line}" if line else "" for line in text.split("\n"))


def _table_row(cells: list[str]) -> str:
    return "| " + " | ".join(cell or " " for cell in cells) + " |"


def _collect_raw_text(node: Node) -> str:
    if node.is_text:
        return node.text
    if node.tag == "br":
        return "\n"
    return "".join(_collect_raw_text(child) for child in node.children)


def _to_int(value: str | None, default: int) -> int:
    try:
        return int(str(value))
    except (TypeError, ValueError):
        return default


__all__ = ["html_to_markdown", "MarkdownRenderer", "parse_html", "unescape"]
