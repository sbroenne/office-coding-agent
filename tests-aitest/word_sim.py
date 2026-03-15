"""In-memory Word simulator for manifest-driven AI tool tests."""

from __future__ import annotations

import re
from typing import Any

from tool_result import ToolResult


def _strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", text).strip()


class WordSimulator:
    """Minimal stateful document simulator."""

    def __init__(self) -> None:
        self.html = "<h1>Project Update</h1><p>Initial draft content.</p>"
        self.selection_text = "Initial draft content."

    def _ok(self, value: Any) -> ToolResult:
        return ToolResult(success=True, value=value)

    def get_document_overview(self) -> ToolResult:
        paragraph_count = len(re.findall(r"<p\b", self.html))
        table_count = len(re.findall(r"<table\b", self.html))
        headings = re.findall(r"<(h[1-6])[^>]*>(.*?)</\1>", self.html, flags=re.IGNORECASE | re.DOTALL)
        heading_lines = [f"{tag.upper()}: {_strip_html(text)}" for tag, text in headings]
        overview = [
            "Document Overview",
            "=" * 40,
            f"Paragraphs: {paragraph_count}",
            f"Tables: {table_count}",
            "",
            "Headings:",
            *heading_lines,
        ]
        return self._ok("\n".join(overview))

    def get_document_content(self) -> ToolResult:
        return self._ok(self.html)

    def set_document_content(self, html: str) -> ToolResult:
        self.html = html
        self.selection_text = _strip_html(html)
        return self._ok("Document content replaced successfully.")

    def insert_content_at_selection(self, html: str, location: str | None = None) -> ToolResult:
        del location
        self.html += html
        self.selection_text = _strip_html(html)
        return self._ok("Content inserted at selection (location: Replace).")

    def apply_style_to_selection(
        self,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikeThrough: bool | None = None,  # noqa: N803
        fontSize: float | None = None,  # noqa: N803
        fontName: str | None = None,  # noqa: N803
        fontColor: str | None = None,  # noqa: N803
        highlightColor: str | None = None,  # noqa: N803
    ) -> ToolResult:
        if self.selection_text:
            styled = self.selection_text
            if bold:
                styled = f"<strong>{styled}</strong>"
            if italic:
                styled = f"<em>{styled}</em>"
            styles: list[str] = []
            if underline:
                styles.append("text-decoration: underline")
            if strikeThrough:
                styles.append("text-decoration: line-through")
            if fontSize is not None:
                styles.append(f"font-size: {fontSize}pt")
            if fontName:
                styles.append(f"font-family: {fontName}")
            if fontColor:
                styles.append(f"color: {fontColor}")
            if highlightColor:
                styles.append(f"background-color: {highlightColor}")
            if styles:
                styled = f'<span style="{"; ".join(styles)}">{styled}</span>'
            self.html += styled
        return self._ok("Applied font formatting to the current selection.")

    def insert_table(
        self,
        rows: float,
        columns: float,
        data: list[list[str]] | None = None,
        style: str | None = None,
        hasHeaderRow: bool | None = None,  # noqa: N803
    ) -> ToolResult:
        del style, hasHeaderRow
        row_count = int(rows)
        column_count = int(columns)
        body_rows: list[str] = []
        for r in range(row_count):
            cells: list[str] = []
            for c in range(column_count):
                value = data[r][c] if data and r < len(data) and c < len(data[r]) else ""
                cell_tag = "th" if r == 0 else "td"
                cells.append(f"<{cell_tag}>{value}</{cell_tag}>")
            body_rows.append(f"<tr>{''.join(cells)}</tr>")
        self.html += f"<table>{''.join(body_rows)}</table>"
        self.selection_text = ""
        return self._ok(f'Inserted a {row_count}×{column_count} table with "grid" style.')

    def find_and_replace(
        self,
        find: str,
        replace: str,
        matchCase: bool | None = None,  # noqa: N803
        matchWholeWord: bool | None = None,  # noqa: N803
    ) -> ToolResult:
        del matchWholeWord
        if matchCase:
            count = self.html.count(find)
            self.html = self.html.replace(find, replace)
        else:
            pattern = re.compile(re.escape(find), flags=re.IGNORECASE)
            self.html, count = pattern.subn(replace, self.html)
        return self._ok(f'Replaced {count} occurrence(s) of "{find}" with "{replace}".')
