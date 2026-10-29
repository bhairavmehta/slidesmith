"""MCP server exposing the toolchain to Kiro (or any MCP client).

Run:  python -m slidesmith.mcp_server     (stdio transport)

Tools: get_schema, list_themes, validate_spec, build, view_pages.
`view_pages` returns the rendered PNGs as images so the agent can look at them
directly — that is the "inspect" step of the render → inspect → fix loop.
"""
from __future__ import annotations

import json
from pathlib import Path

try:  # MCP Python SDK 2.x
    from mcp.server.mcpserver import Image, MCPServer as _Server
except ImportError:  # SDK 1.x
    from mcp.server.fastmcp import FastMCP as _Server, Image  # type: ignore

from pydantic import ValidationError

from .build import build as _build, format_validation_error
from .schema import json_schema, load_spec
from .theme import list_themes as _list_themes

server = _Server(
    "slidesmith",
    instructions=(
        "Deterministic renderer for slide decks (.pptx) and documents (.docx). "
        "Write a JSON spec that follows get_schema, call build, then call view_pages to "
        "look at the result and fix the spec. Never hand-edit the output files."
    ),
)


@server.tool()
def get_schema(kind: str = "deck") -> str:
    """JSON Schema for a 'deck' or 'doc' spec. Read it before writing a spec."""
    return json.dumps(json_schema("doc" if kind == "doc" else "deck"))


@server.tool()
def list_themes() -> list[str]:
    """Theme names usable in a spec's "theme" field (a path to a theme .json also works)."""
    return _list_themes()


@server.tool()
def validate_spec(spec_path: str) -> dict:
    """Fast schema check of a spec file. Returns {valid, errors}."""
    try:
        load_spec(json.loads(Path(spec_path).read_text(encoding="utf-8-sig")))
        return {"valid": True, "errors": []}
    except json.JSONDecodeError as e:
        return {"valid": False, "errors": [f"JSON syntax: {e}"]}
    except ValidationError as e:
        return {"valid": False, "errors": [f"{x['where']}: {x['message']}" for x in format_validation_error(e)]}


@server.tool()
def build(spec_path: str, out_dir: str | None = None, theme: str | None = None,
          rasterize: bool = True, renderer: str = "auto") -> dict:
    """Validate, render (.pptx/.docx), rasterize to PNGs and lint. Returns paths, score and issues.

    renderer: 'auto' (Microsoft Office if available, else LibreOffice, else built-in preview),
    or force 'office' / 'libreoffice' / 'preview'. The result's 'renderer' says which drew the pages.
    Fix every 'error', then the 'warning's that matter, then call view_pages to review visually.
    """
    return _build(spec_path, out_dir, theme, rasterize=rasterize, renderer=renderer).to_dict()


@server.tool()
def view_pages(build_dir: str, pages: list[int] | None = None, contact_sheet: bool = False) -> list:
    """Return rendered pages as images for visual review.

    pages: 1-based page numbers (default: all, max 12). contact_sheet=True returns one
    overview image instead — cheapest way to judge rhythm and consistency across a deck.
    """
    d = Path(build_dir)
    if contact_sheet:
        return [Image(path=str(d / "contact_sheet.png"))]
    files = sorted((d / "pages").glob("page-*.png"), key=lambda p: int(p.stem.split("-")[-1]))
    if pages:
        files = [f for f in files if int(f.stem.split("-")[-1]) in set(pages)]
    return [Image(path=str(f)) for f in files[:12]]


def main():
    server.run()


if __name__ == "__main__":
    main()
