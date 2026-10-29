"""`slidesmith doctor` — is this machine ready? Plain-ASCII output (safe in any Windows console).

Checks Python, packages, Microsoft Office (via COM on Windows), fonts, Kiro wiring,
and finishes with a real test build so you know which renderer will draw the pages.
"""
from __future__ import annotations

import importlib
import json
import os
import platform
import struct
import sys
import tempfile
import time
from pathlib import Path

OK, WARN, FAIL = "[ OK ]", "[WARN]", "[FAIL]"


class Report:
    def __init__(self):
        self.failed = False

    def line(self, status: str, what: str, detail: str = ""):
        if status == FAIL:
            self.failed = True
        print(f"{status} {what}" + (f"  -  {detail}" if detail else ""), flush=True)


def _version(mod: str) -> str | None:
    try:
        from importlib.metadata import version

        dist = {"pptx": "python-pptx", "docx": "python-docx", "PIL": "pillow",
                "win32com": "pywin32"}.get(mod, mod)
        return version(dist)
    except Exception:  # noqa: BLE001
        try:
            m = importlib.import_module(mod)
            return getattr(m, "__version__", "installed")
        except Exception:  # noqa: BLE001
            return None


def _font_found(family: str) -> str | None:
    from matplotlib import font_manager

    from .textfit import ALIASES

    for fam in ALIASES.get(family, [family]):
        try:
            return font_manager.findfont(font_manager.FontProperties(family=fam),
                                         fallback_to_default=False)
        except ValueError:
            continue
    return None


def _kiro_check(r: Report, workspace: Path | None):
    home = Path.home() / ".kiro"
    places = [("user", home)]
    if workspace:
        places.append(("workspace", workspace / ".kiro"))
    found_server = False
    for label, base in places:
        cfg = base / "settings" / "mcp.json"
        if cfg.exists():
            try:
                data = json.loads(cfg.read_text(encoding="utf-8-sig"))
                entry = data.get("mcpServers", {}).get("slidesmith")
            except Exception as e:  # noqa: BLE001
                r.line(WARN, f"Kiro {label} MCP config unreadable", f"{cfg}: {e}")
                continue
            if entry:
                found_server = True
                cmd = entry.get("command", "")
                if cmd and (Path(cmd).exists() or not (":" in cmd or "/" in cmd or "\\" in cmd)):
                    r.line(OK, f"Kiro {label} MCP server 'slidesmith'", cmd)
                else:
                    r.line(FAIL, f"Kiro {label} MCP server points at a missing Python", cmd)
        steer = base / "steering" / "slidesmith-workflow.md"
        if steer.exists():
            r.line(OK, f"Kiro {label} steering files", str(steer.parent))
    if not found_server:
        r.line(WARN, "Kiro MCP server 'slidesmith' not registered",
               "run install.cmd (or add .kiro/settings/mcp.json); the CLI still works")


def run(workspace: str | None = None, skip_build: bool = False) -> int:
    r = Report()
    print(f"slidesmith doctor  -  {platform.system()} {platform.release()} ({platform.version()})\n")

    # Python ----------------------------------------------------------------
    v = sys.version_info
    bits = struct.calcsize("P") * 8
    if v >= (3, 10):
        r.line(OK, f"Python {v.major}.{v.minor}.{v.micro} ({bits}-bit)", sys.executable)
    else:
        r.line(FAIL, f"Python {v.major}.{v.minor} is too old", "need 3.10+")

    # Packages --------------------------------------------------------------
    for mod, need in [("pptx", True), ("docx", True), ("pydantic", True), ("PIL", True),
                      ("matplotlib", True), ("lxml", True), ("pypdfium2", False), ("mcp", False)]:
        ver = _version(mod)
        if ver:
            r.line(OK, f"{mod} {ver}")
        else:
            why = {"pypdfium2": "needed to rasterize Word's PDF export: pip install pypdfium2",
                   "mcp": "needed for the Kiro MCP server: pip install mcp"}.get(mod, f"pip install {mod}")
            r.line(FAIL if need else WARN, f"{mod} missing", why)

    # Office ----------------------------------------------------------------
    system = platform.system()
    if system == "Windows":
        if _version("win32com") is None:
            r.line(WARN, "pywin32 missing", "pip install pywin32 (needed to drive PowerPoint/Word)")
        else:
            r.line(OK, f"pywin32 {_version('win32com')}")
            from .rasterize import RasterizeError, run_office_child

            try:
                res = run_office_child(["probe"], timeout=90)
                for app in ("powerpoint", "word"):
                    info = res.get(app, {})
                    label = "PowerPoint" if app == "powerpoint" else "Word"
                    if "version" in info:
                        r.line(OK, f"Microsoft {label} {info['version']} reachable over COM")
                    else:
                        r.line(WARN, f"Microsoft {label} not reachable",
                               info.get("error", "unknown error") + " - the built-in preview will be used")
            except RasterizeError as e:
                r.line(WARN, "Microsoft Office probe failed", str(e))
    elif system == "Darwin":
        for app in ("PowerPoint", "Word"):
            p = f"/Applications/Microsoft {app}.app"
            r.line(OK if os.path.exists(p) else WARN, f"Microsoft {app}", p)
    else:
        r.line(WARN, "Microsoft Office", "not available on this OS; built-in preview will be used")

    # Fonts -----------------------------------------------------------------
    for fam, use in (("Calibri", "default theme"), ("Inter", "warm theme")):
        p = _font_found(fam)
        if p:
            r.line(OK, f"Font {fam} ({use})", p)
        else:
            r.line(WARN, f"Font {fam} not found ({use})",
                   "install it for your user (right-click the .ttf > Install) or avoid that theme")

    # Kiro ------------------------------------------------------------------
    _kiro_check(r, Path(workspace) if workspace else None)

    # Test build ------------------------------------------------------------
    if not skip_build:
        from .build import build

        spec = {"kind": "deck", "title": "Doctor check", "slides": [
            {"layout": "title", "title": "slidesmith doctor"},
            {"layout": "kpis", "title": "A test slide renders end to end without problems",
             "items": [{"value": "1", "label": "deck"}, {"value": "2", "label": "slides"}]}]}
        with tempfile.TemporaryDirectory() as tmp:
            sp = Path(tmp) / "doctor.deck.json"
            sp.write_text(json.dumps(spec), encoding="utf-8")
            t0 = time.time()
            try:
                res = build(sp, Path(tmp) / "out")
                dt = time.time() - t0
                if res.output and res.pages:
                    status = OK if res.renderer == "office" else WARN
                    note = "" if res.renderer == "office" else " (Office not used - see warnings above)"
                    r.line(status, f"Test build rendered {len(res.pages)} pages with '{res.renderer}' "
                                   f"in {dt:.1f}s{note}")
                elif res.output:
                    r.line(WARN, "Test build wrote the file but no page images", res.report or "")
                else:
                    r.line(FAIL, "Test build failed", json.dumps(res.summary)[:300])
            except Exception as e:  # noqa: BLE001
                r.line(FAIL, "Test build crashed", f"{type(e).__name__}: {e}")

    print("\n" + ("Some checks FAILED." if r.failed else "Ready."))
    return 1 if r.failed else 0
