"""Excel MCP server for integration testing with pytest-aitest.

Reads tools-manifest.json to dynamically register MCP tools backed
by an in-memory ExcelSimulator. This lets pytest-aitest tests exercise
the same tool schemas the real Excel add-in exposes.

Run as: python tests-aitest/excel_mcp.py --manifest src/tools/tools-manifest.json
"""

from __future__ import annotations

import argparse
import inspect
import json
import sys
from pathlib import Path
from typing import Annotated, Any

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from excel_sim import ExcelSimulator

# ---------------------------------------------------------------------------
# Server & simulator
# ---------------------------------------------------------------------------

mcp = FastMCP("excel-ai-addin-test-server")
_sim = ExcelSimulator()

# ---------------------------------------------------------------------------
# Tool routing — maps manifest tool names to simulator methods
# ---------------------------------------------------------------------------

# Mapping from manifest tool name → (simulator_method_name, param_remapping)
# Most tools map 1:1. Decomposed tools merge back into generic methods.
_TOOL_ROUTES: dict[str, tuple[str, dict[str, str] | None]] = {
    # Range
    "get_range_values": ("get_range_values", None),
    "set_range_values": ("set_range_values", None),
    "get_used_range": ("get_used_range", {"maxRows": "max_rows"}),
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
    # Conditional Format (decomposed → generic)
    "add_color_scale": ("add_conditional_format", None),
    "add_data_bar": ("add_conditional_format", None),
    "add_cell_value_format": ("add_conditional_format", None),
    "add_top_bottom_format": ("add_conditional_format", None),
    "add_contains_text_format": ("add_conditional_format", None),
    "add_custom_format": ("add_conditional_format", None),
    "list_conditional_formats": ("list_conditional_formats", None),
    "clear_conditional_formats": ("clear_conditional_formats", None),
    # Data Validation (decomposed → generic)
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


def _remap_params(params: dict[str, Any], remap: dict[str, str] | None) -> dict[str, Any]:
    """Remap camelCase param names from manifest to snake_case for Python methods."""
    if not remap:
        # Default: convert common camelCase patterns
        result = {}
        for k, v in params.items():
            if v is None:
                continue
            # Convert camelCase to snake_case
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
        key = remap.get(k, k)
        result[key] = v
    return result


def _dispatch(tool_name: str, params: dict[str, Any]) -> str:
    """Route a tool call to the appropriate simulator method."""
    route = _TOOL_ROUTES.get(tool_name)
    if not route:
        return json.dumps({"error": f"Unknown tool: {tool_name}"})

    method_name, remap = route
    method = getattr(_sim, method_name, None)
    if not method:
        return json.dumps({"error": f"Simulator has no method: {method_name}"})

    # Special handling for decomposed conditional format tools
    if tool_name.startswith("add_") and method_name == "add_conditional_format":
        rule_type = tool_name.replace("add_", "")
        py_params = _remap_params(params, remap)
        address = py_params.pop("address", "")
        sheet_name = py_params.pop("sheet_name", None)
        result = method(rule_type=rule_type, address=address, sheet_name=sheet_name, **py_params)
    elif tool_name.startswith("set_") and method_name == "set_data_validation":
        validation_type = tool_name.replace("set_", "").replace("_validation", "")
        py_params = _remap_params(params, remap)
        address = py_params.pop("address", "")
        sheet_name = py_params.pop("sheet_name", None)
        result = method(validation_type=validation_type, address=address, sheet_name=sheet_name, **py_params)
    else:
        py_params = _remap_params(params, remap)
        result = method(**py_params)

    if result.success:
        return json.dumps(result.value, default=str)
    return json.dumps({"error": result.error})


def _load_manifest(manifest_path: Path) -> list[dict[str, Any]]:
    """Load and return tools from the manifest file."""
    with manifest_path.open() as f:
        manifest = json.load(f)
    return manifest["tools"]


def _build_tool_docstring(tool_def: dict[str, Any]) -> str:
    """Build a Google-style docstring from manifest tool definition."""
    doc = tool_def["description"] + "\n"
    params = tool_def.get("params", {})
    if params:
        doc += "\nArgs:\n"
        for name, param in params.items():
            req = " (required)" if param.get("required", True) else " (optional)"
            doc += f"    {name}: {param['description']}{req}\n"
    return doc


def register_tools(manifest_path: Path) -> None:
    """Register all manifest tools as MCP tools backed by the simulator."""
    tools = _load_manifest(manifest_path)

    for tool_def in tools:
        tool_name = tool_def["name"]

        if tool_name not in _TOOL_ROUTES:
            continue

        # Build the param signature for the tool
        params_meta = tool_def.get("params", {})

        # Create a closure-based handler
        def make_handler(tn: str, pm: dict[str, Any]) -> Any:
            """Create a tool handler closure."""
            def handler(**kwargs: Any) -> str:
                return _dispatch(tn, kwargs)

            # Set function metadata for FastMCP
            handler.__name__ = tn
            handler.__doc__ = _build_tool_docstring({"description": tool_def["description"], "params": pm})

            # Build proper inspect.Signature so FastMCP exposes individual
            # parameters with descriptions and enum constraints matching
            # what the production Zod schemas generate.
            sig_params: list[inspect.Parameter] = []
            annotations: dict[str, Any] = {}
            for pname, pdef in pm.items():
                ptype = pdef.get("type", "string")
                required = pdef.get("required", True)
                desc = pdef.get("description", "")
                enum_values = pdef.get("enum")

                if ptype == "string":
                    base = str
                elif ptype == "number":
                    base = float
                elif ptype == "boolean":
                    base = bool
                elif ptype in ("string[]",):
                    base = list[str]
                elif ptype in ("any[][]", "string[][]"):
                    base = list[list[Any]]
                else:
                    base = str

                # Build Pydantic Field with description + optional enum.
                # Use base type directly (not base | None) so FastMCP
                # generates {"type": "string"} instead of {"anyOf": [{string}, {null}]},
                # matching how Zod .optional() serializes in production.
                extra = {"enum": enum_values} if enum_values else None

                if not required:
                    ann = Annotated[base, Field(default=None, description=desc, json_schema_extra=extra)]
                    sig_params.append(inspect.Parameter(
                        pname,
                        inspect.Parameter.POSITIONAL_OR_KEYWORD,
                        default=None,
                        annotation=ann,
                    ))
                else:
                    ann = Annotated[base, Field(description=desc, json_schema_extra=extra)]
                    sig_params.append(inspect.Parameter(
                        pname,
                        inspect.Parameter.POSITIONAL_OR_KEYWORD,
                        annotation=ann,
                    ))
                annotations[pname] = ann

            annotations["return"] = str
            # Required params must come before optional ones in the signature
            required_params = [p for p in sig_params if p.default is inspect.Parameter.empty]
            optional_params = [p for p in sig_params if p.default is not inspect.Parameter.empty]
            handler.__signature__ = inspect.Signature(required_params + optional_params, return_annotation=str)
            handler.__annotations__ = annotations

            return handler

        handler = make_handler(tool_name, params_meta)
        mcp.tool()(handler)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Parse args and run the server."""
    parser = argparse.ArgumentParser(description="Excel MCP server for testing")
    parser.add_argument(
        "--manifest",
        type=Path,
        default=Path(__file__).parent / "src" / "tools" / "tools-manifest.json",
        help="Path to tools-manifest.json",
    )
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse", "streamable-http"],
        default="stdio",
    )
    parser.add_argument("--port", type=int, default=8080)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()

    if not args.manifest.exists():
        print(f"Manifest not found: {args.manifest}", file=sys.stderr)
        print("Run 'npm run manifest' first.", file=sys.stderr)
        raise SystemExit(1)

    register_tools(args.manifest)
    print(f"Registered {len(mcp._tool_manager._tools)} tools from {args.manifest}", file=sys.stderr)

    mcp.settings.host = args.host
    mcp.settings.port = args.port

    if args.transport == "stdio":
        mcp.run(transport="stdio")
    elif args.transport == "sse":
        mcp.run(transport="sse")
    elif args.transport == "streamable-http":
        mcp.settings.stateless_http = True
        mcp.settings.json_response = True
        mcp.run(transport="streamable-http")


if __name__ == "__main__":
    main()
