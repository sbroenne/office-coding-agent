"""Excel MCP server for integration testing with pytest-skill-engineering.

Registers individual decomposed tools (e.g. ``get_range_values``,
``set_range_values``) backed by an in-memory :class:`ExcelSimulator`.

The aggregate manifest (10 config-group tools) is loaded at startup to
provide rich tool and parameter descriptions. Tool schemas are still
derived from the simulator method signatures + the routing table's
camelCase ↔ snake_case map.

Run as::

    python tests-aitest/excel_mcp.py --manifest tests-aitest/manifests/excel-tools-manifest.json
"""

from __future__ import annotations

import argparse
import inspect
import json
import re
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

# Individual tool name → (manifest tool, aggregate action, param aliases, fixed type)
_TOOL_MANIFEST_ROUTE_MAP: dict[str, tuple[str, str | None, dict[str, str], str | None]] = {}


def _add_manifest_routes(
    manifest_tool: str,
    route_actions: dict[str, str],
    *,
    param_aliases: dict[str, dict[str, str]] | None = None,
    fixed_variants: dict[str, str] | None = None,
) -> None:
    """Register manifest metadata for decomposed MCP tools."""
    alias_lookup = param_aliases or {}
    variant_lookup = fixed_variants or {}
    for tool_name, action in route_actions.items():
        _TOOL_MANIFEST_ROUTE_MAP[tool_name] = (
            manifest_tool,
            action,
            dict(alias_lookup.get(tool_name, {})),
            variant_lookup.get(tool_name),
        )


_add_manifest_routes(
    "range",
    {
        "get_range_values": "get_values",
        "set_range_values": "set_values",
        "get_used_range": "get_used",
        "clear_range": "clear",
        "set_range_formulas": "set_formulas",
        "get_range_formulas": "get_formulas",
        "sort_range": "sort",
        "copy_range": "copy",
        "find_values": "find",
        "insert_range": "insert",
        "delete_range": "delete",
        "merge_cells": "merge",
        "unmerge_cells": "unmerge",
        "replace_values": "replace",
        "remove_duplicates": "remove_duplicates",
        "group_rows_columns": "group",
        "ungroup_rows_columns": "ungroup",
    },
    param_aliases={
        "copy_range": {
            "sourceSheetName": "sourceSheet",
            "destinationSheetName": "destinationSheet",
        },
        "find_values": {"searchValue": "searchText"},
    },
)

_add_manifest_routes(
    "range_format",
    {
        "format_range": "format",
        "set_number_format": "set_number_format",
        "auto_fit_columns": "auto_fit",
        "auto_fit_rows": "auto_fit",
        "set_hyperlink": "set_hyperlink",
        "toggle_row_column_visibility": "toggle_visibility",
        "set_cell_borders": "set_borders",
    },
    param_aliases={"set_number_format": {"formatCode": "format"}},
)

_add_manifest_routes(
    "sheet",
    {
        "list_sheets": "list",
        "create_sheet": "create",
        "rename_sheet": "rename",
        "delete_sheet": "delete",
        "activate_sheet": "activate",
        "freeze_panes": "freeze",
        "protect_sheet": "protect",
        "unprotect_sheet": "unprotect",
        "set_sheet_visibility": "set_visibility",
        "copy_sheet": "copy",
        "move_sheet": "move",
        "set_page_layout": "set_page_layout",
    },
)

_add_manifest_routes(
    "table",
    {
        "list_tables": "list",
        "create_table": "create",
        "add_table_rows": "add_rows",
        "get_table_data": "get_data",
        "delete_table": "delete",
        "sort_table": "sort",
        "filter_table": "filter",
        "add_table_column": "add_column",
        "delete_table_column": "delete_column",
        "convert_table_to_range": "convert_to_range",
        "clear_table_filters": "clear_filters",
    },
    param_aliases={"filter_table": {"values": "filterValues"}},
)

_add_manifest_routes(
    "chart",
    {
        "list_charts": "list",
        "create_chart": "create",
        "delete_chart": "delete",
        "set_chart_title": "configure",
        "set_chart_type": "configure",
        "set_chart_data_source": "configure",
    },
)

_add_manifest_routes(
    "workbook",
    {
        "get_workbook_info": "get_info",
        "recalculate_workbook": "recalculate",
        "get_selected_range": "get_selected_range",
        "define_named_range": "define_named_range",
        "list_named_ranges": "list_named_ranges",
    },
)

_add_manifest_routes(
    "comment",
    {
        "add_comment": "add",
        "list_comments": "list",
        "edit_comment": "edit",
        "delete_comment": "delete",
    },
    param_aliases={"edit_comment": {"newText": "text"}},
)

_add_manifest_routes(
    "conditional_format",
    {
        "add_color_scale": "add",
        "add_data_bar": "add",
        "add_cell_value_format": "add",
        "add_top_bottom_format": "add",
        "add_contains_text_format": "add",
        "add_custom_format": "add",
        "list_conditional_formats": "list",
        "clear_conditional_formats": "clear",
    },
    fixed_variants={
        "add_color_scale": "colorScale",
        "add_data_bar": "dataBar",
        "add_cell_value_format": "cellValue",
        "add_top_bottom_format": "topBottom",
        "add_contains_text_format": "containsText",
        "add_custom_format": "custom",
    },
)

_add_manifest_routes(
    "data_validation",
    {
        "set_list_validation": "set",
        "set_number_validation": "set",
        "set_date_validation": "set",
        "set_text_length_validation": "set",
        "set_custom_validation": "set",
        "get_data_validation": "get",
        "clear_data_validation": "clear",
    },
    fixed_variants={
        "set_list_validation": "list",
        "set_number_validation": "number",
        "set_date_validation": "date",
        "set_text_length_validation": "textLength",
        "set_custom_validation": "custom",
    },
)

_add_manifest_routes(
    "pivot",
    {
        "list_pivot_tables": "list",
        "refresh_pivot_table": "refresh",
        "delete_pivot_table": "delete",
        "create_pivot_table": "create",
        "add_pivot_field": "add_field",
        "remove_pivot_field": "remove_field",
    },
)


# ---------------------------------------------------------------------------
# Manifest description helpers
# ---------------------------------------------------------------------------


def _normalize_manifest_name(name: str) -> str:
    """Normalize manifest tool names for lookup."""
    return re.sub(r"[_\s]+", "", name).lower()


def _extract_manifest_params(tool: dict[str, Any]) -> dict[str, Any]:
    """Return the parameter schema mapping from a manifest tool entry."""
    params = tool.get("params")
    if isinstance(params, dict):
        return params

    input_schema = tool.get("inputSchema")
    if isinstance(input_schema, dict):
        properties = input_schema.get("properties")
        if isinstance(properties, dict):
            return properties

    return {}


def _load_manifest_lookup(manifest_path: Path) -> dict[str, dict[str, Any]]:
    """Load the aggregate manifest into a normalized lookup table."""
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    lookup: dict[str, dict[str, Any]] = {}
    for tool in manifest.get("tools", []):
        name = str(tool.get("name", "")).strip()
        if not name:
            continue

        params = _extract_manifest_params(tool)
        lookup[_normalize_manifest_name(name)] = {
            "name": name,
            "description": str(tool.get("description", "")).strip(),
            "params": {
                param_name: str(param_schema.get("description", "")).strip()
                for param_name, param_schema in params.items()
                if isinstance(param_schema, dict)
            },
        }
    return lookup


def _description_prefix(description: str) -> str:
    """Return the shared prefix before action-specific manifest details."""
    first_quote = description.find('"')
    if first_quote == -1:
        return description.strip()

    prefix = description[:first_quote].strip()
    prefix = re.sub(r"(?:Actions?:|Use action)\s*$", "", prefix, flags=re.IGNORECASE).strip()
    return prefix


def _extract_action_detail(description: str, action: str | None) -> str | None:
    """Extract the action-specific detail text from an aggregate description."""
    if not action:
        return None

    matches = list(re.finditer(r'"([^"]+)"', description))
    for index, match in enumerate(matches):
        if match.group(1) != action:
            continue

        end = matches[index + 1].start() if index + 1 < len(matches) else len(description)
        detail = description[match.end():end].strip()
        detail = re.sub(r"^[,;:\s]+", "", detail)
        detail = re.sub(r"^or\s+", "", detail, flags=re.IGNORECASE)

        if detail.startswith("("):
            close = detail.find(")")
            if close != -1:
                detail = detail[1:close].strip()
            else:
                detail = detail.strip("() ")
        else:
            detail = re.sub(r"^to\s+", "", detail, flags=re.IGNORECASE)
            detail = re.sub(r"\s+(?:or|and)$", "", detail, flags=re.IGNORECASE)
            detail = detail.rstrip(" ,.;")

        return detail or None

    return None


def _build_tool_description(
    tool_name: str,
    manifest_tool: str,
    manifest_description: str,
    action: str | None,
    fixed_variant: str | None,
) -> str:
    """Build a per-tool description from the aggregate manifest metadata."""
    if not manifest_description:
        return _humanize(tool_name)

    prefix = _description_prefix(manifest_description)
    detail = _extract_action_detail(manifest_description, action)
    parts: list[str] = []

    if prefix:
        parts.append(prefix.rstrip("."))

    if action:
        if detail:
            parts.append(f'Action "{action}": {detail.rstrip(".")}')
        else:
            parts.append(f'Action "{action}"')

    if fixed_variant:
        if manifest_tool == "conditional_format":
            parts.append(f'Rule type is fixed to "{fixed_variant}"')
        elif manifest_tool == "data_validation":
            parts.append(f'Validation type is fixed to "{fixed_variant}"')

    if not parts:
        return manifest_description

    return ". ".join(part.rstrip(".") for part in parts if part).strip() + "."


def _resolve_manifest_param_description(
    manifest_entry: dict[str, Any] | None,
    param_name: str,
    param_aliases: dict[str, str],
) -> str | None:
    """Resolve a parameter description from the manifest, honoring aliases."""
    if not manifest_entry:
        return None

    manifest_param_name = param_aliases.get(param_name, param_name)
    return manifest_entry.get("params", {}).get(manifest_param_name)


def _humanize_param(name: str) -> str:
    """Convert parameter names like ``sheetName`` into readable words."""
    spaced = re.sub(r"(?<!^)(?=[A-Z])", " ", name).replace("_", " ")
    return spaced.lower()


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


def register_tools_from_routes(manifest_path: Path) -> None:
    """Register individual tools from the routing table.

    Schemas are derived by inspecting the simulator method signatures and
    enriched with descriptions loaded from the aggregate tool manifest.
    """
    manifest_lookup = _load_manifest_lookup(manifest_path)

    for tool_name, (method_name, remap) in _TOOL_ROUTES.items():
        method = getattr(_sim, method_name, None)
        if method is None:
            continue

        manifest_entry: dict[str, Any] | None = None
        manifest_param_aliases: dict[str, str] = {}
        tool_description = _humanize(tool_name)
        manifest_route = _TOOL_MANIFEST_ROUTE_MAP.get(tool_name)
        if manifest_route:
            manifest_tool, action, manifest_param_aliases, fixed_variant = manifest_route
            manifest_entry = manifest_lookup.get(_normalize_manifest_name(manifest_tool))
            if manifest_entry:
                tool_description = _build_tool_description(
                    tool_name,
                    manifest_tool,
                    manifest_entry.get("description", ""),
                    action,
                    fixed_variant,
                )

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
            desc = _resolve_manifest_param_description(manifest_entry, camel_name, manifest_param_aliases)
            if not desc:
                desc = f"The {_humanize_param(camel_name)}"
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
        def make_handler(tn: str = tool_name, doc: str = tool_description) -> Any:
            def handler(**kwargs: Any) -> str:
                return _dispatch(tn, kwargs)
            handler.__name__ = tn
            handler.__doc__ = doc
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
        help="Path to the Excel aggregate manifest used for tool and parameter descriptions.",
    )
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse", "streamable-http"],
        default="stdio",
    )
    parser.add_argument("--port", type=int, default=8080)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()

    # Register individual decomposed tools from routing table + manifest descriptions
    register_tools_from_routes(args.manifest)
    print(f"Registered {len(mcp._tool_manager._tools)} tools", file=sys.stderr)

    mcp.settings.host = args.host
    mcp.settings.port = args.port

    if args.transport == "streamable-http":
        mcp.settings.stateless_http = True
        mcp.settings.json_response = True

    mcp.run(transport=args.transport)


if __name__ == "__main__":
    main()
