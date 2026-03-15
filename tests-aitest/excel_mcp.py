"""Excel MCP server for integration testing with pytest-skill-engineering.

Registers individual decomposed tools (e.g. ``get_range_values``,
``set_range_values``) backed by an in-memory :class:`ExcelSimulator`.

The aggregate manifest (10 config-group tools) is loaded only for
reference descriptions.  Tool schemas are derived from the simulator
method signatures + the routing table's camelCase ↔ snake_case map.

Run as::

    python tests-aitest/excel_mcp.py --manifest tests-aitest/manifests/excel-tools-manifest.json
"""

from __future__ import annotations

import argparse
import inspect
import json
import sys
from pathlib import Path
from typing import Annotated, Any, get_args, get_origin

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from excel_sim import ExcelSimulator
from tool_result import ToolResult

# ---------------------------------------------------------------------------
# Server & simulator
# ---------------------------------------------------------------------------

mcp = FastMCP("excel-ai-addin-test-server")
_sim = ExcelSimulator()

# ---------------------------------------------------------------------------
# Tool routing — individual tool name → (simulator_method, camelCase→snake remap)
# ---------------------------------------------------------------------------

_TOOL_ROUTES: dict[str, tuple[str, dict[str, str] | None]] = {
    # Range
    "get_range_values": ("get_range_values", {"sheetName": "sheet_name", "maxRows": "max_rows", "maxColumns": "max_columns", "startRow": "start_row", "startColumn": "start_column"}),
    "set_range_values": ("set_range_values", None),
    "get_used_range": ("get_used_range", {"sheetName": "sheet_name", "maxRows": "max_rows", "maxColumns": "max_columns", "startRow": "start_row", "startColumn": "start_column"}),
    "clear_range": ("clear_range", None),
    "format_range": ("format_range", None),
    "set_number_format": ("set_number_format", {"formatCode": "format_code"}),
    "auto_fit_columns": ("auto_fit_columns", None),
    "auto_fit_rows": ("auto_fit_rows", None),
    "set_range_formulas": ("set_range_formulas", None),
    "get_range_formulas": ("get_range_formulas", None),
    "sort_range": ("sort_range", None),
    "copy_range": ("copy_range", {"sourceAddress": "source_address", "destinationAddress": "destination_address",
                                   "sourceSheetName": "source_sheet_name", "destinationSheetName": "destination_sheet_name"}),
    "find_values": ("find_values", {"matchCase": "match_case"}),
    "insert_range": ("insert_range", None),
    "delete_range": ("delete_range", None),
    "merge_cells": ("merge_cells", None),
    "unmerge_cells": ("unmerge_cells", None),
    "replace_values": ("replace_values", None),
    "remove_duplicates": ("remove_duplicates", None),
    "set_hyperlink": ("set_hyperlink", {"textToDisplay": "text_to_display"}),
    "toggle_row_column_visibility": ("toggle_row_column_visibility", None),
    "group_rows_columns": ("group_rows_columns", None),
    "ungroup_rows_columns": ("ungroup_rows_columns", None),
    "set_cell_borders": ("set_cell_borders", {"borderStyle": "border_style", "borderColor": "border_color"}),
    # Sheet
    "list_sheets": ("list_sheets", None),
    "create_sheet": ("create_sheet", {"name": "name"}),
    "rename_sheet": ("rename_sheet", {"currentName": "current_name", "newName": "new_name"}),
    "delete_sheet": ("delete_sheet", {"name": "name"}),
    "activate_sheet": ("activate_sheet", {"name": "name"}),
    "freeze_panes": ("freeze_panes", {"name": "name", "freezeAt": "freeze_at"}),
    "protect_sheet": ("protect_sheet", {"name": "name"}),
    "unprotect_sheet": ("unprotect_sheet", {"name": "name"}),
    "set_sheet_visibility": ("set_sheet_visibility", {"name": "name", "tabColor": "tab_color"}),
    "copy_sheet": ("copy_sheet", {"name": "name", "newName": "new_name"}),
    "move_sheet": ("move_sheet", {"name": "name"}),
    "set_page_layout": ("set_page_layout", {"name": "name", "paperSize": "paper_size", "leftMargin": "left_margin", "rightMargin": "right_margin", "topMargin": "top_margin", "bottomMargin": "bottom_margin"}),
    # Table
    "list_tables": ("list_tables", {"sheetName": "sheet_name"}),
    "create_table": ("create_table", {"hasHeaders": "has_headers", "sheetName": "sheet_name"}),
    "add_table_rows": ("add_table_rows", {"tableName": "table_name"}),
    "get_table_data": ("get_table_data", {"tableName": "table_name"}),
    "delete_table": ("delete_table", {"tableName": "table_name"}),
    "sort_table": ("sort_table", {"tableName": "table_name"}),
    "filter_table": ("filter_table", {"tableName": "table_name"}),
    "add_table_column": ("add_table_column", {"tableName": "table_name", "columnName": "column_name", "columnData": "column_data"}),
    "delete_table_column": ("delete_table_column", {"tableName": "table_name", "columnName": "column_name"}),
    "convert_table_to_range": ("convert_table_to_range", {"tableName": "table_name"}),
    "clear_table_filters": ("clear_table_filters", {"tableName": "table_name"}),
    "set_chart_title": ("set_chart_title", {"chartName": "chart_name"}),
    "set_chart_type": ("set_chart_type", {"chartName": "chart_name", "chartType": "chart_type"}),
    "set_chart_data_source": ("set_chart_data_source", {"chartName": "chart_name", "dataRange": "data_range"}),
    # Chart
    "list_charts": ("list_charts", {"sheetName": "sheet_name"}),
    "create_chart": ("create_chart", {"dataRange": "data_range", "chartType": "chart_type", "sheetName": "sheet_name"}),
    "recalculate_workbook": ("recalculate_workbook", {"recalcType": "recalc_type"}),
    "delete_chart": ("delete_chart", {"chartName": "chart_name", "sheetName": "sheet_name"}),
    # Workbook
    "get_workbook_info": ("get_workbook_info", None),
    "get_selected_range": ("get_selected_range", None),
    "define_named_range": ("define_named_range", {"sheetName": "sheet_name"}),
    "list_named_ranges": ("list_named_ranges", None),
    # Comment
    "add_comment": ("add_comment", {"cellAddress": "cell_address", "sheetName": "sheet_name"}),
    "list_comments": ("list_comments", {"sheetName": "sheet_name"}),
    "edit_comment": ("edit_comment", {"cellAddress": "cell_address", "newText": "new_text", "sheetName": "sheet_name"}),
    "delete_comment": ("delete_comment", {"cellAddress": "cell_address", "sheetName": "sheet_name"}),
    # Conditional Format (decomposed → generic add_conditional_format)
    "add_color_scale": ("add_conditional_format", None),
    "add_data_bar": ("add_conditional_format", None),
    "add_cell_value_format": ("add_conditional_format", None),
    "add_top_bottom_format": ("add_conditional_format", None),
    "add_contains_text_format": ("add_conditional_format", None),
    "add_custom_format": ("add_conditional_format", None),
    "list_conditional_formats": ("list_conditional_formats", None),
    "clear_conditional_formats": ("clear_conditional_formats", None),
    # Data Validation (decomposed → generic set_data_validation)
    "set_list_validation": ("set_data_validation", None),
    "set_number_validation": ("set_data_validation", None),
    "set_date_validation": ("set_data_validation", None),
    "set_text_length_validation": ("set_data_validation", None),
    "set_custom_validation": ("set_data_validation", None),
    "get_data_validation": ("get_data_validation", None),
    "clear_data_validation": ("clear_data_validation", None),
    # Pivot Table
    "list_pivot_tables": ("list_pivot_tables", {"sheetName": "sheet_name"}),
    "refresh_pivot_table": ("refresh_pivot_table", {"pivotTableName": "pivot_table_name", "sheetName": "sheet_name"}),
    "delete_pivot_table": ("delete_pivot_table", {"pivotTableName": "pivot_table_name", "sheetName": "sheet_name"}),
    "create_pivot_table": ("create_pivot_table", {
        "sourceAddress": "source_address",
        "destinationAddress": "destination_address",
        "rowFields": "row_fields",
        "valueFields": "value_fields",
        "sourceSheetName": "source_sheet_name",
        "destinationSheetName": "destination_sheet_name",
    }),
    "add_pivot_field": ("add_pivot_field", {"pivotTableName": "pivot_table_name", "fieldName": "field_name", "fieldType": "field_type"}),
    "remove_pivot_field": ("remove_pivot_field", {"pivotTableName": "pivot_table_name", "fieldName": "field_name", "fieldType": "field_type"}),
}

# Params that the dispatch layer synthesises — exclude from the MCP schema.
_DISPATCH_ONLY_PARAMS = {"rule_type", "validation_type"}


# ---------------------------------------------------------------------------
# camelCase ↔ snake_case helpers
# ---------------------------------------------------------------------------


def _remap_params(params: dict[str, Any], remap: dict[str, str] | None) -> dict[str, Any]:
    """Remap camelCase param names to snake_case for the simulator."""
    if not remap:
        result = {}
        for k, v in params.items():
            if v is None:
                continue
            snake = ""
            for i, c in enumerate(k):
                if c.isupper() and i > 0:
                    snake += "_"
                snake += c.lower()
            result[snake] = v
        return result

    result = {}
    for k, v in params.items():
        if v is None:
            continue
        result[remap.get(k, k)] = v
    return result


# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------


def _dispatch(tool_name: str, params: dict[str, Any]) -> str:
    """Route a tool call to the appropriate simulator method."""
    route = _TOOL_ROUTES.get(tool_name)
    if not route:
        return json.dumps({"error": f"Unknown tool: {tool_name}"})

    method_name, remap = route
    method = getattr(_sim, method_name, None)
    if not method:
        return json.dumps({"error": f"Simulator has no method: {method_name}"})

    # Decomposed conditional-format tools → synthesise rule_type
    if tool_name.startswith("add_") and method_name == "add_conditional_format":
        rule_type = tool_name.replace("add_", "")
        py_params = _remap_params(params, remap)
        address = py_params.pop("address", "")
        sheet_name = py_params.pop("sheet_name", None)
        result = method(rule_type=rule_type, address=address, sheet_name=sheet_name, **py_params)
    # Decomposed data-validation tools → synthesise validation_type
    elif tool_name.startswith("set_") and method_name == "set_data_validation":
        validation_type = tool_name.replace("set_", "").replace("_validation", "")
        py_params = _remap_params(params, remap)
        address = py_params.pop("address", "")
        sheet_name = py_params.pop("sheet_name", None)
        result = method(validation_type=validation_type, address=address, sheet_name=sheet_name, **py_params)
    else:
        py_params = _remap_params(params, remap)
        result = method(**py_params)

    if isinstance(result, ToolResult):
        if result.success:
            return json.dumps(result.value, default=str)
        return json.dumps({"error": result.error})
    return json.dumps(result, default=str)


# ---------------------------------------------------------------------------
# Schema introspection — derive MCP tool schemas from the simulator methods
# ---------------------------------------------------------------------------


def _annotation_to_json_type(ann: Any) -> str:  # noqa: PLR0911
    """Map a Python type annotation to a JSON Schema type string."""
    if ann is str:
        return "string"
    if ann is int:
        return "integer"
    if ann is float:
        return "number"
    if ann is bool:
        return "boolean"

    origin = get_origin(ann)
    if origin is list:
        return "array"

    # Handle Union / Optional (e.g. str | None)
    if origin is type(str | None):
        args = [a for a in get_args(ann) if a is not type(None)]
        if args:
            return _annotation_to_json_type(args[0])
    return "string"


def _annotation_to_schema(ann: Any) -> dict[str, Any]:
    """Build a mini JSON Schema for a single parameter annotation."""
    origin = get_origin(ann)

    # Handle Optional / Union types (e.g. str | None, list[str] | None)
    if origin is type(str | None):
        args = [a for a in get_args(ann) if a is not type(None)]
        if args:
            return _annotation_to_schema(args[0])

    if origin is list:
        inner = get_args(ann)
        if inner:
            inner_origin = get_origin(inner[0])
            if inner_origin is list:
                return {"type": "array", "items": {"type": "array"}}
            return {"type": "array", "items": {"type": _annotation_to_json_type(inner[0])}}
        return {"type": "array"}

    return {"type": _annotation_to_json_type(ann)}


def _humanize(name: str) -> str:
    """``get_range_values`` → ``Get range values``."""
    return name.replace("_", " ").capitalize()


def register_tools_from_routes() -> None:
    """Register individual tools from the routing table.

    Schemas are derived by inspecting the simulator method signatures and
    reversing the camelCase→snake_case remap in each routing entry.
    """
    for tool_name, (method_name, remap) in _TOOL_ROUTES.items():
        method = getattr(_sim, method_name, None)
        if method is None:
            continue

        # Reverse remap: snake_case → camelCase
        reverse: dict[str, str] = {}
        if remap:
            reverse = {v: k for k, v in remap.items()}

        sig = inspect.signature(method)
        sig_params: list[inspect.Parameter] = []
        annotations: dict[str, Any] = {}

        for pname, param in sig.parameters.items():
            if pname == "self" or pname in _DISPATCH_ONLY_PARAMS:
                continue
            # Skip **kwargs
            if param.kind == inspect.Parameter.VAR_KEYWORD:
                continue

            camel_name = reverse.get(pname, pname)
            ann = param.annotation if param.annotation is not inspect.Parameter.empty else str
            schema = _annotation_to_schema(ann)
            base_type = {"string": str, "integer": int, "number": float, "boolean": bool, "array": list}.get(schema["type"], str)
            desc = f"The {pname.replace('_', ' ')}"
            extra = {"items": schema["items"]} if "items" in schema else None

            if param.default is not inspect.Parameter.empty:
                pydantic_ann = Annotated[base_type, Field(default=None, description=desc, json_schema_extra=extra)]
                sig_params.append(inspect.Parameter(
                    camel_name, inspect.Parameter.POSITIONAL_OR_KEYWORD,
                    default=None, annotation=pydantic_ann,
                ))
            else:
                pydantic_ann = Annotated[base_type, Field(description=desc, json_schema_extra=extra)]
                sig_params.append(inspect.Parameter(
                    camel_name, inspect.Parameter.POSITIONAL_OR_KEYWORD,
                    annotation=pydantic_ann,
                ))
            annotations[camel_name] = pydantic_ann

        annotations["return"] = str
        required_params = [p for p in sig_params if p.default is inspect.Parameter.empty]
        optional_params = [p for p in sig_params if p.default is not inspect.Parameter.empty]

        # Build handler closure
        def make_handler(tn: str = tool_name) -> Any:
            def handler(**kwargs: Any) -> str:
                return _dispatch(tn, kwargs)
            handler.__name__ = tn
            handler.__doc__ = _humanize(tn)
            handler.__signature__ = inspect.Signature(required_params + optional_params, return_annotation=str)
            handler.__annotations__ = dict(annotations)
            return handler

        mcp.tool()(make_handler())


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Parse args and run the server."""
    parser = argparse.ArgumentParser(description="Excel MCP server for testing")
    parser.add_argument(
        "--manifest",
        type=Path,
        default=Path(__file__).parent / "manifests" / "excel-tools-manifest.json",
        help="Path to Excel aggregate manifest (kept for reference; tools are registered from routing table).",
    )
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse", "streamable-http"],
        default="stdio",
    )
    parser.add_argument("--port", type=int, default=8080)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()

    # Register individual decomposed tools from routing table
    register_tools_from_routes()
    print(f"Registered {len(mcp._tool_manager._tools)} tools", file=sys.stderr)

    mcp.settings.host = args.host
    mcp.settings.port = args.port

    if args.transport == "streamable-http":
        mcp.settings.stateless_http = True
        mcp.settings.json_response = True

    mcp.run(transport=args.transport)


if __name__ == "__main__":
    main()
