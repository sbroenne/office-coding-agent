"""Shared helpers for manifest-driven MCP servers used in AI tool tests."""

from __future__ import annotations

import argparse
import inspect
import json
import sys
from pathlib import Path
from typing import Annotated, Any

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from tool_result import ToolResult


def _legacy_params_to_schema(tool_def: dict[str, Any]) -> dict[str, Any]:
    params = tool_def.get("params", {})
    properties: dict[str, Any] = {}
    required: list[str] = []

    for name, meta in params.items():
        property_schema: dict[str, Any] = {
            "type": meta.get("type", "string"),
            "description": meta.get("description", ""),
        }
        if meta.get("enum"):
            property_schema["enum"] = meta["enum"]
        if meta.get("default") is not None:
            property_schema["default"] = meta["default"]
        properties[name] = property_schema
        if meta.get("required", True):
            required.append(name)

    return {"type": "object", "properties": properties, "required": required}


def get_input_schema(tool_def: dict[str, Any]) -> dict[str, Any]:
    """Return a top-level JSON schema for a manifest tool."""
    schema = tool_def.get("inputSchema")
    if isinstance(schema, dict):
        required = schema.get("required")
        return {
            **schema,
            "type": "object",
            "properties": schema.get("properties", {}),
            "required": required if isinstance(required, list) else [],
        }
    return _legacy_params_to_schema(tool_def)


def _schema_to_type(schema: dict[str, Any]) -> Any:
    schema_type = schema.get("type", "string")

    if schema_type == "string":
        return str
    if schema_type == "number":
        return float
    if schema_type == "integer":
        return int
    if schema_type == "boolean":
        return bool
    if schema_type == "object":
        return dict[str, Any]
    if schema_type == "array":
        items = schema.get("items", {})
        if not isinstance(items, dict):
            return list[Any]
        item_type = items.get("type")
        if item_type == "string":
            return list[str]
        if item_type == "number":
            return list[float]
        if item_type == "integer":
            return list[int]
        if item_type == "boolean":
            return list[bool]
        if item_type == "object":
            return list[dict[str, Any]]
        if item_type == "array":
            return list[list[Any]]
        return list[Any]
    return Any


def _build_tool_docstring(tool_def: dict[str, Any], input_schema: dict[str, Any]) -> str:
    doc = tool_def.get("description", "") + "\n"
    properties = input_schema.get("properties", {})
    required = set(input_schema.get("required", []))
    if properties:
        doc += "\nArgs:\n"
        for name, schema in properties.items():
            req = " (required)" if name in required else " (optional)"
            desc = schema.get("description", "") if isinstance(schema, dict) else ""
            doc += f"    {name}: {desc}{req}\n"
    return doc


def _build_signature(input_schema: dict[str, Any]) -> tuple[inspect.Signature, dict[str, Any]]:
    properties = input_schema.get("properties", {})
    required = set(input_schema.get("required", []))
    sig_params: list[inspect.Parameter] = []
    annotations: dict[str, Any] = {}

    for name, schema in properties.items():
        schema = schema if isinstance(schema, dict) else {}
        base = _schema_to_type(schema)
        desc = schema.get("description", "")
        enum_values = schema.get("enum")
        extra = {"enum": enum_values} if enum_values else None

        if name in required:
            annotation = Annotated[base, Field(description=desc, json_schema_extra=extra)]
            sig_params.append(
                inspect.Parameter(name, inspect.Parameter.POSITIONAL_OR_KEYWORD, annotation=annotation)
            )
        else:
            annotation = Annotated[base, Field(default=None, description=desc, json_schema_extra=extra)]
            sig_params.append(
                inspect.Parameter(
                    name,
                    inspect.Parameter.POSITIONAL_OR_KEYWORD,
                    default=None,
                    annotation=annotation,
                )
            )
        annotations[name] = annotation

    required_params = [p for p in sig_params if p.default is inspect.Parameter.empty]
    optional_params = [p for p in sig_params if p.default is not inspect.Parameter.empty]
    annotations["return"] = str
    return inspect.Signature(required_params + optional_params, return_annotation=str), annotations


def _serialize_result(value: Any) -> str:
    if isinstance(value, str):
        return value
    return json.dumps(value, default=str)


def create_manifest_server(
    server_name: str,
    simulator: Any,
    *,
    routes: dict[str, str] | None = None,
) -> tuple[FastMCP, callable]:
    """Create a FastMCP server that exposes manifest-defined tools."""

    mcp = FastMCP(server_name)
    route_map = routes or {}

    def dispatch(tool_name: str, params: dict[str, Any]) -> str:
        method_name = route_map.get(tool_name, tool_name)
        method = getattr(simulator, method_name, None)
        if method is None:
            return json.dumps({"error": f"Simulator has no method for tool: {tool_name}"})

        result = method(**params)
        if isinstance(result, ToolResult):
            if result.success:
                return _serialize_result(result.value)
            return json.dumps({"error": result.error or f"Tool failed: {tool_name}"})

        return _serialize_result(result)

    def register_tools(manifest_path: Path) -> None:
        with manifest_path.open(encoding="utf-8") as handle:
            manifest = json.load(handle)

        for tool_def in manifest["tools"]:
            tool_name = tool_def["name"]
            input_schema = get_input_schema(tool_def)

            def make_handler(name: str = tool_name, schema: dict[str, Any] = input_schema, definition: dict[str, Any] = tool_def):
                def handler(**kwargs: Any) -> str:
                    return dispatch(name, kwargs)

                handler.__name__ = name
                handler.__doc__ = _build_tool_docstring(definition, schema)
                handler.__signature__, handler.__annotations__ = _build_signature(schema)
                return handler

            mcp.tool()(make_handler())

    return mcp, register_tools


def run_manifest_server(
    server_name: str,
    simulator: Any,
    *,
    routes: dict[str, str] | None = None,
    default_manifest: Path | None = None,
) -> None:
    """CLI entry point used by per-host MCP wrapper scripts."""

    mcp, register_tools = create_manifest_server(server_name, simulator, routes=routes)

    parser = argparse.ArgumentParser(description=f"{server_name} for testing")
    parser.add_argument(
        "--manifest",
        type=Path,
        default=default_manifest,
        help="Path to the JSON tool manifest.",
    )
    parser.add_argument("--transport", choices=["stdio", "sse", "streamable-http"], default="stdio")
    parser.add_argument("--port", type=int, default=8080)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()

    if args.manifest is None or not args.manifest.exists():
        print(f"Manifest not found: {args.manifest}", file=sys.stderr)
        print("Run 'npm run manifest' first.", file=sys.stderr)
        raise SystemExit(1)

    register_tools(args.manifest)
    print(f"Registered {len(mcp._tool_manager._tools)} tools from {args.manifest}", file=sys.stderr)

    mcp.settings.host = args.host
    mcp.settings.port = args.port

    if args.transport == "streamable-http":
        mcp.settings.stateless_http = True
        mcp.settings.json_response = True

    mcp.run(transport=args.transport)
