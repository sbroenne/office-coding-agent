"""In-memory spreadsheet simulator for testing Excel tool schemas.

Provides a minimal spreadsheet engine that the ExcelMCP server uses
to back tool calls. This does NOT aim to replicate Excel's full feature
set — just enough to validate that an LLM can read/write ranges, create
tables/charts/comments, and apply formatting/validation via the tools.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any


# ---------------------------------------------------------------------------
# ToolResult — inlined from pytest_aitest.testing.types to keep this
# module self-contained within the excel-ai-addin project.
# ---------------------------------------------------------------------------


@dataclass
class ToolResult:
    """Result from a tool call."""

    success: bool
    value: Any = None
    error: str | None = None

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {"success": self.success, "value": self.value, "error": self.error}


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------


@dataclass
class Comment:
    """A cell comment."""

    cell_address: str
    text: str
    sheet_name: str


@dataclass
class Table:
    """A named table on a sheet."""

    name: str
    address: str
    sheet_name: str
    has_headers: bool = True

    @property
    def id(self) -> str:
        return f"tbl_{self.name}"


@dataclass
class Chart:
    """A chart on a sheet."""

    name: str
    chart_type: str
    data_range: str
    sheet_name: str
    title: str = ""


@dataclass
class ConditionalFormat:
    """A conditional format rule."""

    rule_type: str
    address: str
    sheet_name: str
    params: dict[str, Any] = field(default_factory=dict)


@dataclass
class DataValidation:
    """A data validation rule."""

    validation_type: str
    address: str
    sheet_name: str
    params: dict[str, Any] = field(default_factory=dict)


@dataclass
class NamedRange:
    """A workbook-scoped named range."""

    name: str
    address: str
    sheet_name: str
    comment: str = ""


@dataclass
class Sheet:
    """A worksheet with cells."""

    name: str
    cells: dict[str, Any] = field(default_factory=dict)
    formulas: dict[str, str] = field(default_factory=dict)
    formats: dict[str, dict[str, Any]] = field(default_factory=dict)
    number_formats: dict[str, str] = field(default_factory=dict)
    merged: list[str] = field(default_factory=list)
    position: int = 0
    frozen_at: str | None = None
    is_protected: bool = False
    visibility: str = "Visible"
    tab_color: str = ""
    hyperlinks: dict[str, dict[str, str]] = field(default_factory=dict)
    hidden_rows: set[int] = field(default_factory=set)
    hidden_columns: set[int] = field(default_factory=set)
    grouped_rows: list[str] = field(default_factory=list)
    grouped_columns: list[str] = field(default_factory=list)
    page_layout: dict[str, Any] = field(default_factory=lambda: {
        "orientation": "Portrait",
        "paperSize": "Letter",
        "margins": {"left": 0.75, "right": 0.75, "top": 1, "bottom": 1},
    })


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _col_to_index(col: str) -> int:
    """Convert column letter(s) to 0-based index. A=0, B=1, ..., Z=25, AA=26."""
    result = 0
    for c in col.upper():
        result = result * 26 + (ord(c) - ord("A") + 1)
    return result - 1


def _index_to_col(idx: int) -> str:
    """Convert 0-based index to column letter(s)."""
    result = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _parse_cell(ref: str) -> tuple[str, int]:
    """Parse 'A1' into ('A', 1)."""
    m = re.match(r"([A-Za-z]+)(\d+)", ref)
    if not m:
        msg = f"Invalid cell reference: {ref}"
        raise ValueError(msg)
    return m.group(1).upper(), int(m.group(2))


def _parse_range(address: str) -> tuple[str | None, str, str]:
    """Parse 'Sheet1!A1:B5' into (sheet_name, start_cell, end_cell).

    Also handles 'A1:B5' (no sheet) and 'A1' (single cell).
    """
    sheet_name = None
    cell_part = address

    if "!" in address:
        sheet_name, cell_part = address.split("!", 1)
        sheet_name = sheet_name.strip("'\"")

    if ":" in cell_part:
        start, end = cell_part.split(":", 1)
    else:
        start = end = cell_part

    return sheet_name, start.upper(), end.upper()


# ---------------------------------------------------------------------------
# Simulator
# ---------------------------------------------------------------------------


class ExcelSimulator:
    """Stateful in-memory spreadsheet simulator."""

    def __init__(self) -> None:
        self.sheets: dict[str, Sheet] = {"Sheet1": Sheet(name="Sheet1", position=0)}
        self.active_sheet: str = "Sheet1"
        self.tables: dict[str, Table] = {}
        self.charts: dict[str, Chart] = {}
        self.comments: list[Comment] = []
        self.conditional_formats: list[ConditionalFormat] = []
        self.data_validations: list[DataValidation] = []
        self.named_ranges: dict[str, NamedRange] = {}
        self._chart_counter: int = 0

    # ─── Sheet resolution ────────────────────────────────────────────

    def _resolve_sheet(self, sheet_name: str | None = None) -> Sheet:
        name = sheet_name or self.active_sheet
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")  # type: ignore[return-value]
        return self.sheets[name]

    def _error_result(self, msg: str) -> ToolResult:
        return ToolResult(success=False, error=msg)

    def _ok(self, value: Any = None) -> ToolResult:
        return ToolResult(success=True, value=value)

    # ─── Range Operations ────────────────────────────────────────────

    def get_range_values(self, address: str, sheet_name: str | None = None) -> ToolResult:
        sheet_ref, start, end = _parse_range(address)
        sheet = self._resolve_sheet(sheet_ref or sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet

        start_col, start_row = _parse_cell(start)
        end_col, end_row = _parse_cell(end)

        rows = []
        for r in range(start_row, end_row + 1):
            row = []
            for c in range(_col_to_index(start_col), _col_to_index(end_col) + 1):
                cell_ref = f"{_index_to_col(c)}{r}"
                row.append(sheet.cells.get(cell_ref, ""))
            rows.append(row)

        return self._ok(rows)

    def set_range_values(self, address: str, values: list[list[Any]], sheet_name: str | None = None) -> ToolResult:
        sheet_ref, start, _end = _parse_range(address)
        sheet = self._resolve_sheet(sheet_ref or sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet

        start_col, start_row = _parse_cell(start)
        base_col = _col_to_index(start_col)

        for ri, row in enumerate(values):
            for ci, val in enumerate(row):
                cell_ref = f"{_index_to_col(base_col + ci)}{start_row + ri}"
                sheet.cells[cell_ref] = val

        num_rows = len(values)
        num_cols = max((len(r) for r in values), default=0)
        end_cell = f"{_index_to_col(base_col + num_cols - 1)}{start_row + num_rows - 1}"
        return self._ok({"address": f"{start}:{end_cell}", "rowsWritten": num_rows, "columnsWritten": num_cols})

    def get_used_range(self, sheet_name: str | None = None, max_rows: int | None = None) -> ToolResult:
        sheet = self._resolve_sheet(sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet

        if not sheet.cells:
            result: dict[str, Any] = {"address": "A1", "rowCount": 1, "columnCount": 1}
            if max_rows is not None:
                result["values"] = [[""]]
            return self._ok(result)

        min_col = min_row = float("inf")
        max_col = max_row = 0
        for ref in sheet.cells:
            col_str, row_num = _parse_cell(ref)
            ci = _col_to_index(col_str)
            min_col = min(min_col, ci)
            max_col = max(max_col, ci)
            min_row = min(min_row, row_num)
            max_row = max(max_row, row_num)

        total_rows = int(max_row) - int(min_row) + 1
        total_cols = int(max_col) - int(min_col) + 1
        addr = f"{_index_to_col(int(min_col))}{int(min_row)}:{_index_to_col(int(max_col))}{int(max_row)}"

        result = {
            "address": addr,
            "rowCount": total_rows,
            "columnCount": total_cols,
        }

        # Only include values when maxRows is explicitly set
        if max_rows is not None:
            rows_to_read = min(max_rows, total_rows)
            rows = []
            for r in range(int(min_row), int(min_row) + rows_to_read):
                row = []
                for c in range(int(min_col), int(max_col) + 1):
                    cell_ref = f"{_index_to_col(c)}{r}"
                    row.append(sheet.cells.get(cell_ref, ""))
                rows.append(row)
            result["values"] = rows
            if max_rows < total_rows:
                result["truncated"] = True
                result["rowsReturned"] = rows_to_read

        return self._ok(result)

    def clear_range(self, address: str, sheet_name: str | None = None) -> ToolResult:
        sheet_ref, start, end = _parse_range(address)
        sheet = self._resolve_sheet(sheet_ref or sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet

        start_col, start_row = _parse_cell(start)
        end_col, end_row = _parse_cell(end)

        cleared = 0
        for r in range(start_row, end_row + 1):
            for c in range(_col_to_index(start_col), _col_to_index(end_col) + 1):
                cell_ref = f"{_index_to_col(c)}{r}"
                if cell_ref in sheet.cells:
                    del sheet.cells[cell_ref]
                    cleared += 1
                sheet.formulas.pop(cell_ref, None)
                sheet.formats.pop(cell_ref, None)

        return self._ok({"address": address, "cellsCleared": cleared})

    def format_range(self, address: str, sheet_name: str | None = None, **fmt: Any) -> ToolResult:
        sheet_ref, start, end = _parse_range(address)
        sheet = self._resolve_sheet(sheet_ref or sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet

        start_col, start_row = _parse_cell(start)
        end_col, end_row = _parse_cell(end)

        for r in range(start_row, end_row + 1):
            for c in range(_col_to_index(start_col), _col_to_index(end_col) + 1):
                cell_ref = f"{_index_to_col(c)}{r}"
                if cell_ref not in sheet.formats:
                    sheet.formats[cell_ref] = {}
                sheet.formats[cell_ref].update(fmt)

        return self._ok({"address": address, "formatsApplied": list(fmt.keys())})

    def set_number_format(self, address: str, format_code: str, sheet_name: str | None = None) -> ToolResult:
        sheet_ref, start, end = _parse_range(address)
        sheet = self._resolve_sheet(sheet_ref or sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet

        start_col, start_row = _parse_cell(start)
        end_col, end_row = _parse_cell(end)

        for r in range(start_row, end_row + 1):
            for c in range(_col_to_index(start_col), _col_to_index(end_col) + 1):
                cell_ref = f"{_index_to_col(c)}{r}"
                sheet.number_formats[cell_ref] = format_code

        return self._ok({"address": address, "numberFormat": format_code})

    def auto_fit_columns(self, address: str, sheet_name: str | None = None) -> ToolResult:
        return self._ok({"address": address, "autoFitColumns": True})

    def auto_fit_rows(self, address: str, sheet_name: str | None = None) -> ToolResult:
        return self._ok({"address": address, "autoFitRows": True})

    def set_range_formulas(self, address: str, formulas: list[list[str]], sheet_name: str | None = None) -> ToolResult:
        sheet_ref, start, _end = _parse_range(address)
        sheet = self._resolve_sheet(sheet_ref or sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet

        start_col, start_row = _parse_cell(start)
        base_col = _col_to_index(start_col)

        for ri, row in enumerate(formulas):
            for ci, formula in enumerate(row):
                cell_ref = f"{_index_to_col(base_col + ci)}{start_row + ri}"
                sheet.formulas[cell_ref] = formula
                sheet.cells[cell_ref] = f"[formula:{formula}]"

        return self._ok({"address": address, "formulasSet": len(formulas)})

    def get_range_formulas(self, address: str, sheet_name: str | None = None) -> ToolResult:
        sheet_ref, start, end = _parse_range(address)
        sheet = self._resolve_sheet(sheet_ref or sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet

        start_col, start_row = _parse_cell(start)
        end_col, end_row = _parse_cell(end)

        rows = []
        for r in range(start_row, end_row + 1):
            row = []
            for c in range(_col_to_index(start_col), _col_to_index(end_col) + 1):
                cell_ref = f"{_index_to_col(c)}{r}"
                row.append(sheet.formulas.get(cell_ref, ""))
            rows.append(row)

        return self._ok(rows)

    def sort_range(self, address: str, column: int = 0, ascending: bool = True,
                   has_headers: bool = False, sheet_name: str | None = None) -> ToolResult:
        return self._ok({"address": address, "sortedByColumn": column, "ascending": ascending})

    def copy_range(self, source_address: str, destination_address: str,
                   source_sheet_name: str | None = None,
                   destination_sheet_name: str | None = None) -> ToolResult:
        src_result = self.get_range_values(source_address, source_sheet_name)
        if not src_result.success:
            return src_result
        return self.set_range_values(destination_address, src_result.value, destination_sheet_name)

    def find_values(self, searchValue: str, address: str | None = None,
                    sheet_name: str | None = None, match_case: bool = False) -> ToolResult:
        sheet = self._resolve_sheet(sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet

        matches = []
        search = searchValue if match_case else searchValue.lower()

        for ref, val in sheet.cells.items():
            cell_val = str(val) if match_case else str(val).lower()
            if search in cell_val:
                matches.append({"address": ref, "value": val})

        return self._ok(matches)

    def insert_range(self, address: str, shift: str, sheet_name: str | None = None) -> ToolResult:
        return self._ok({"address": address, "shift": shift, "inserted": True})

    def delete_range(self, address: str, shift: str, sheet_name: str | None = None) -> ToolResult:
        return self._ok({"address": address, "shift": shift, "deleted": True})

    def merge_cells(self, address: str, across: bool = False, sheet_name: str | None = None) -> ToolResult:
        sheet = self._resolve_sheet(sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet
        sheet.merged.append(address)
        return self._ok({"address": address, "merged": True})

    def unmerge_cells(self, address: str, sheet_name: str | None = None) -> ToolResult:
        sheet = self._resolve_sheet(sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet
        sheet.merged = [m for m in sheet.merged if m != address]
        return self._ok({"address": address, "unmerged": True})

    def replace_values(
        self, find: str, replace: str, address: str | None = None, sheet_name: str | None = None
    ) -> ToolResult:
        sheet = self._resolve_sheet(sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet
        count = 0
        for ref, val in list(sheet.cells.items()):
            cell_str = str(val)
            if find.lower() in cell_str.lower():
                sheet.cells[ref] = cell_str.replace(find, replace)
                count += 1
        return self._ok({"find": find, "replace": replace, "replacements": count})

    def remove_duplicates(
        self, address: str, columns: list[str], sheet_name: str | None = None
    ) -> ToolResult:
        # Simplified: just acknowledge the operation
        return self._ok({"address": address, "rowsRemoved": 0, "rowsRemaining": 0})

    def set_hyperlink(
        self,
        address: str,
        url: str,
        text_to_display: str | None = None,
        tooltip: str | None = None,
        sheet_name: str | None = None,
    ) -> ToolResult:
        sheet = self._resolve_sheet(sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet
        if url == "":
            sheet.hyperlinks.pop(address, None)
        else:
            sheet.hyperlinks[address] = {
                "address": url,
                "textToDisplay": text_to_display or url,
                "screenTip": tooltip or "",
            }
        return self._ok({"address": address, "url": url or None})

    def toggle_row_column_visibility(
        self, address: str, hidden: bool, target: str, sheet_name: str | None = None
    ) -> ToolResult:
        sheet = self._resolve_sheet(sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet
        return self._ok({"address": address, "target": target, "hidden": hidden})

    def group_rows_columns(self, address: str, sheet_name: str | None = None) -> ToolResult:
        sheet = self._resolve_sheet(sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet
        sheet.grouped_rows.append(address)
        return self._ok({"address": address, "grouped": True})

    def ungroup_rows_columns(self, address: str, sheet_name: str | None = None) -> ToolResult:
        sheet = self._resolve_sheet(sheet_name)
        if isinstance(sheet, ToolResult):
            return sheet
        sheet.grouped_rows = [g for g in sheet.grouped_rows if g != address]
        return self._ok({"address": address, "ungrouped": True})

    def set_cell_borders(
        self,
        address: str,
        border_style: str,
        border_color: str | None = None,
        side: str | None = None,
        sheet_name: str | None = None,
    ) -> ToolResult:
        return self._ok({"address": address, "borderStyle": border_style, "side": side, "color": border_color or "000000"})

    # ─── Sheet Operations ────────────────────────────────────────────

    def list_sheets(self) -> ToolResult:
        sheets = [{"name": s.name, "position": s.position, "isActive": s.name == self.active_sheet}
                  for s in sorted(self.sheets.values(), key=lambda s: s.position)]
        return self._ok(sheets)

    def create_sheet(self, name: str) -> ToolResult:
        if name in self.sheets:
            return self._error_result(f"Sheet '{name}' already exists")
        pos = len(self.sheets)
        self.sheets[name] = Sheet(name=name, position=pos)
        return self._ok({"name": name, "id": f"sheet_{name}", "position": pos})

    def rename_sheet(self, current_name: str, new_name: str) -> ToolResult:
        if current_name not in self.sheets:
            return self._error_result(f"Sheet '{current_name}' not found")
        if new_name in self.sheets:
            return self._error_result(f"Sheet '{new_name}' already exists")
        sheet = self.sheets.pop(current_name)
        sheet.name = new_name
        self.sheets[new_name] = sheet
        if self.active_sheet == current_name:
            self.active_sheet = new_name
        return self._ok({"previousName": current_name, "newName": new_name})

    def delete_sheet(self, name: str) -> ToolResult:
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")
        if len(self.sheets) <= 1:
            return self._error_result("Cannot delete the last sheet")
        del self.sheets[name]
        if self.active_sheet == name:
            self.active_sheet = next(iter(self.sheets))
        return self._ok({"deleted": name})

    def activate_sheet(self, name: str) -> ToolResult:
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")
        self.active_sheet = name
        return self._ok({"activated": name})

    def freeze_panes(self, name: str, freeze_at: str | None = None) -> ToolResult:
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")
        sheet = self.sheets[name]
        sheet.frozen_at = freeze_at
        return self._ok({"sheet": name, "frozenAt": freeze_at, "unfrozen": freeze_at is None})

    def protect_sheet(self, name: str, password: str | None = None) -> ToolResult:
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")
        self.sheets[name].is_protected = True
        return self._ok({"sheet": name, "protected": True})

    def unprotect_sheet(self, name: str, password: str | None = None) -> ToolResult:
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")
        self.sheets[name].is_protected = False
        return self._ok({"sheet": name, "protected": False})

    def set_sheet_visibility(
        self, name: str, visibility: str | None = None, tab_color: str | None = None
    ) -> ToolResult:
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")
        sheet = self.sheets[name]
        if visibility is not None:
            sheet.visibility = visibility
        if tab_color is not None:
            sheet.tab_color = tab_color
        return self._ok({"name": name, "visibility": sheet.visibility, "tabColor": sheet.tab_color})

    def copy_sheet(self, name: str, new_name: str | None = None) -> ToolResult:
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")
        source = self.sheets[name]
        copied_name = new_name or f"{name} (2)"
        if copied_name in self.sheets:
            return self._error_result(f"Sheet '{copied_name}' already exists")
        pos = source.position + 1
        new_sheet = Sheet(
            name=copied_name,
            cells=dict(source.cells),
            formulas=dict(source.formulas),
            formats={k: dict(v) for k, v in source.formats.items()},
            number_formats=dict(source.number_formats),
            merged=list(source.merged),
            position=pos,
        )
        self.sheets[copied_name] = new_sheet
        return self._ok({"sourceSheet": name, "copiedSheet": copied_name, "position": pos})

    def move_sheet(self, name: str, position: int) -> ToolResult:
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")
        self.sheets[name].position = position
        return self._ok({"name": name, "position": position})

    def set_page_layout(
        self,
        name: str,
        orientation: str | None = None,
        paper_size: str | None = None,
        left_margin: float | None = None,
        right_margin: float | None = None,
        top_margin: float | None = None,
        bottom_margin: float | None = None,
    ) -> ToolResult:
        if name not in self.sheets:
            return self._error_result(f"Sheet '{name}' not found")
        sheet = self.sheets[name]
        if orientation:
            sheet.page_layout["orientation"] = orientation
        if paper_size:
            sheet.page_layout["paperSize"] = paper_size
        if left_margin is not None:
            sheet.page_layout["margins"]["left"] = left_margin
        if right_margin is not None:
            sheet.page_layout["margins"]["right"] = right_margin
        if top_margin is not None:
            sheet.page_layout["margins"]["top"] = top_margin
        if bottom_margin is not None:
            sheet.page_layout["margins"]["bottom"] = bottom_margin
        return self._ok({
            "sheet": name,
            "orientation": sheet.page_layout["orientation"],
            "paperSize": sheet.page_layout["paperSize"],
            "margins": sheet.page_layout["margins"],
        })

    # ─── Table Operations ────────────────────────────────────────────

    def list_tables(self, sheet_name: str | None = None) -> ToolResult:
        tables = list(self.tables.values())
        if sheet_name:
            tables = [t for t in tables if t.sheet_name == sheet_name]
        return self._ok([{"name": t.name, "id": t.id, "address": t.address,
                         "worksheetName": t.sheet_name} for t in tables])

    def create_table(self, address: str, name: str | None = None,
                     has_headers: bool = True, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        tbl_name = name or f"Table{len(self.tables) + 1}"
        if tbl_name in self.tables:
            return self._error_result(f"Table '{tbl_name}' already exists")
        table = Table(name=tbl_name, address=address, sheet_name=sn, has_headers=has_headers)
        self.tables[tbl_name] = table
        return self._ok({"name": tbl_name, "id": table.id, "address": address})

    def add_table_rows(self, table_name: str, values: list[list[Any]]) -> ToolResult:
        if table_name not in self.tables:
            return self._error_result(f"Table '{table_name}' not found")
        return self._ok({"tableName": table_name, "rowsAdded": len(values)})

    def get_table_data(self, table_name: str) -> ToolResult:
        if table_name not in self.tables:
            return self._error_result(f"Table '{table_name}' not found")
        table = self.tables[table_name]
        result = self.get_range_values(table.address, table.sheet_name)
        if not result.success:
            return result
        return self._ok({"tableName": table_name, "address": table.address, "values": result.value})

    def delete_table(self, table_name: str) -> ToolResult:
        if table_name not in self.tables:
            return self._error_result(f"Table '{table_name}' not found")
        del self.tables[table_name]
        return self._ok({"deleted": table_name})

    def sort_table(self, table_name: str, column: int, ascending: bool = True) -> ToolResult:
        if table_name not in self.tables:
            return self._error_result(f"Table '{table_name}' not found")
        return self._ok({"tableName": table_name, "sortedByColumn": column, "ascending": ascending})

    def filter_table(self, table_name: str, column: str, values: list[str]) -> ToolResult:
        if table_name not in self.tables:
            return self._error_result(f"Table '{table_name}' not found")
        return self._ok({"tableName": table_name, "filteredColumn": column, "filterValues": values})

    def clear_table_filters(self, table_name: str) -> ToolResult:
        if table_name not in self.tables:
            return self._error_result(f"Table '{table_name}' not found")
        return self._ok({"tableName": table_name, "filtersCleared": True})

    def add_table_column(
        self, table_name: str, column_name: str | None = None, column_data: list[str] | None = None
    ) -> ToolResult:
        if table_name not in self.tables:
            return self._error_result(f"Table '{table_name}' not found")
        col_name = column_name or f"Column{len(self.tables[table_name].columns or []) + 1}"
        return self._ok({"tableName": table_name, "columnName": col_name, "added": True})

    def delete_table_column(self, table_name: str, column_name: str) -> ToolResult:
        if table_name not in self.tables:
            return self._error_result(f"Table '{table_name}' not found")
        return self._ok({"tableName": table_name, "columnName": column_name, "deleted": True})

    def convert_table_to_range(self, table_name: str) -> ToolResult:
        if table_name not in self.tables:
            return self._error_result(f"Table '{table_name}' not found")
        table = self.tables[table_name]
        del self.tables[table_name]
        return self._ok({"tableName": table_name, "rangeAddress": table.address, "converted": True})

    # ─── Chart Operations ────────────────────────────────────────────

    def list_charts(self, sheet_name: str | None = None) -> ToolResult:
        charts = list(self.charts.values())
        if sheet_name:
            charts = [c for c in charts if c.sheet_name == sheet_name]
        return self._ok([{"name": c.name, "chartType": c.chart_type,
                         "dataRange": c.data_range, "title": c.title} for c in charts])

    def create_chart(self, data_range: str, chart_type: str,
                     title: str | None = None, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        self._chart_counter += 1
        name = f"Chart {self._chart_counter}"
        chart = Chart(name=name, chart_type=chart_type, data_range=data_range,
                      sheet_name=sn, title=title or "")
        self.charts[name] = chart
        return self._ok({"name": name, "chartType": chart_type, "title": title or "", "dataRange": data_range})

    def delete_chart(self, chart_name: str, sheet_name: str | None = None) -> ToolResult:
        if chart_name not in self.charts:
            return self._error_result(f"Chart '{chart_name}' not found")
        del self.charts[chart_name]
        return self._ok({"deleted": chart_name})

    def set_chart_title(self, chart_name: str, title: str, sheet_name: str | None = None) -> ToolResult:
        if chart_name not in self.charts:
            return self._error_result(f"Chart '{chart_name}' not found")
        self.charts[chart_name].title = title
        return self._ok({"chartName": chart_name, "title": title})

    def set_chart_type(self, chart_name: str, chart_type: str, sheet_name: str | None = None) -> ToolResult:
        if chart_name not in self.charts:
            return self._error_result(f"Chart '{chart_name}' not found")
        self.charts[chart_name].chart_type = chart_type
        return self._ok({"chartName": chart_name, "chartType": chart_type})

    def set_chart_data_source(
        self, chart_name: str, data_range: str, sheet_name: str | None = None
    ) -> ToolResult:
        if chart_name not in self.charts:
            return self._error_result(f"Chart '{chart_name}' not found")
        self.charts[chart_name].data_range = data_range
        return self._ok({"chartName": chart_name, "dataRange": data_range, "updated": True})

    # ─── Workbook Operations ─────────────────────────────────────────

    def get_workbook_info(self) -> ToolResult:
        sheets_info = [{"name": s.name, "position": s.position}
                       for s in sorted(self.sheets.values(), key=lambda s: s.position)]
        return self._ok({
            "activeSheet": self.active_sheet,
            "sheets": sheets_info,
            "tableCount": len(self.tables),
            "chartCount": len(self.charts),
            "namedRangeCount": len(self.named_ranges),
        })

    def recalculate_workbook(self, recalc_type: str | None = None) -> ToolResult:
        recalc = recalc_type or "Full"
        return self._ok({"recalculated": True, "type": recalc})

    def get_selected_range(self) -> ToolResult:
        return self._ok({"address": "A1", "values": [[""]], "sheetName": self.active_sheet})

    def define_named_range(self, name: str, address: str, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        self.named_ranges[name] = NamedRange(name=name, address=address, sheet_name=sn)
        return self._ok({"name": name, "address": address, "comment": ""})

    def list_named_ranges(self) -> ToolResult:
        return self._ok([{"name": n.name, "address": n.address, "sheetName": n.sheet_name}
                        for n in self.named_ranges.values()])

    # ─── Comment Operations ──────────────────────────────────────────

    def add_comment(self, cell_address: str, text: str, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        self.comments.append(Comment(cell_address=cell_address, text=text, sheet_name=sn))
        return self._ok({"cellAddress": cell_address, "text": text, "sheetName": sn})

    def list_comments(self, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        filtered = [c for c in self.comments if c.sheet_name == sn]
        return self._ok([{"cellAddress": c.cell_address, "text": c.text} for c in filtered])

    def edit_comment(self, cell_address: str, new_text: str, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        for c in self.comments:
            if c.cell_address == cell_address and c.sheet_name == sn:
                c.text = new_text
                return self._ok({"cellAddress": cell_address, "newText": new_text})
        return self._error_result(f"No comment at {cell_address}")

    def delete_comment(self, cell_address: str, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        for i, c in enumerate(self.comments):
            if c.cell_address == cell_address and c.sheet_name == sn:
                self.comments.pop(i)
                return self._ok({"deleted": cell_address})
        return self._error_result(f"No comment at {cell_address}")

    # ─── Conditional Format Operations ───────────────────────────────

    def add_conditional_format(self, rule_type: str, address: str,
                               sheet_name: str | None = None, **params: Any) -> ToolResult:
        sn = sheet_name or self.active_sheet
        cf = ConditionalFormat(rule_type=rule_type, address=address, sheet_name=sn, params=params)
        self.conditional_formats.append(cf)
        return self._ok({"ruleType": rule_type, "address": address, "applied": True})

    def list_conditional_formats(self, address: str, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        filtered = [cf for cf in self.conditional_formats if cf.sheet_name == sn]
        return self._ok([{"ruleType": cf.rule_type, "address": cf.address, "params": cf.params}
                        for cf in filtered])

    def clear_conditional_formats(self, address: str | None = None,
                                  sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        before = len(self.conditional_formats)
        self.conditional_formats = [cf for cf in self.conditional_formats if cf.sheet_name != sn]
        cleared = before - len(self.conditional_formats)
        return self._ok({"cleared": cleared})

    # ─── Data Validation Operations ──────────────────────────────────

    def set_data_validation(self, validation_type: str, address: str,
                            sheet_name: str | None = None, **params: Any) -> ToolResult:
        sn = sheet_name or self.active_sheet
        dv = DataValidation(validation_type=validation_type, address=address, sheet_name=sn, params=params)
        self.data_validations = [v for v in self.data_validations if not (v.address == address and v.sheet_name == sn)]
        self.data_validations.append(dv)
        return self._ok({"validationType": validation_type, "address": address, "applied": True})

    def get_data_validation(self, address: str, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        for dv in self.data_validations:
            if dv.address == address and dv.sheet_name == sn:
                return self._ok({"validationType": dv.validation_type, "address": address, "params": dv.params})
        return self._ok({"address": address, "validationType": None, "message": "No validation set"})

    def clear_data_validation(self, address: str, sheet_name: str | None = None) -> ToolResult:
        sn = sheet_name or self.active_sheet
        before = len(self.data_validations)
        self.data_validations = [v for v in self.data_validations if not (v.address == address and v.sheet_name == sn)]
        cleared = before - len(self.data_validations)
        return self._ok({"address": address, "cleared": cleared})

    # ─── Pivot Table Operations ──────────────────────────────────────

    def list_pivot_tables(self, sheet_name: str | None = None) -> ToolResult:
        return self._ok([])

    def refresh_pivot_table(self, pivot_table_name: str, sheet_name: str | None = None) -> ToolResult:
        re

    def add_pivot_field(
        self, pivot_table_name: str, field_name: str, field_type: str, sheet_name: str | None = None
    ) -> ToolResult:
        return self._ok({"pivotTableName": pivot_table_name, "fieldName": field_name, "fieldType": field_type, "added": True})

    def remove_pivot_field(
        self, pivot_table_name: str, field_name: str, field_type: str, sheet_name: str | None = None
    ) -> ToolResult:
        return self._ok({"pivotTableName": pivot_table_name, "fieldName": field_name, "fieldType": field_type, "removed": True})

    def delete_pivot_table(self, pivot_table_name: str, sheet_name: str | None = None) -> ToolResult:
        return self._ok({"pivotTableName": pivot_table_name, "deleted": True})

    def create_pivot_table(
        self,
        name: str,
        source_address: str,
        destination_address: str,
        row_fields: list[str],
        value_fields: list[str],
        source_sheet_name: str | None = None,
        destination_sheet_name: str | None = None,
    ) -> ToolResult:
        return self._ok({
            "pivotTableName": name,
            "rowFields": row_fields,
            "valueFields": value_fields,
            "created": True,
        })
