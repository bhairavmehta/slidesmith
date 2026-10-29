"""slidesmith CLI — the commands an agent (or a person) runs.

  slidesmith schema deck|doc          JSON Schema for writing a spec
  slidesmith themes                   available themes
  slidesmith validate SPEC            schema check only (fast)
  slidesmith build SPEC [-o DIR]      render + rasterize + lint; exit 2 if errors remain
                 [--renderer auto|office|libreoffice|preview]
  slidesmith doctor [--workspace DIR] check Python, packages, Office, fonts, Kiro wiring
"""
from __future__ import annotations

import argparse
import json
import sys

from pydantic import ValidationError

from .build import build, format_validation_error, read_spec
from .schema import json_schema, load_spec
from .theme import list_themes


def _utf8_console():
    """Windows pipes (e.g. when Kiro runs the CLI) default to cp1252; our output has · → κ."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main(argv=None) -> int:
    _utf8_console()
    ap = argparse.ArgumentParser(prog="slidesmith", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("schema")
    s.add_argument("kind", choices=["deck", "doc"])
    sub.add_parser("themes")
    v = sub.add_parser("validate")
    v.add_argument("spec")
    b = sub.add_parser("build")
    b.add_argument("spec")
    b.add_argument("-o", "--out", help="output dir (default: <spec dir>/build/<name>)")
    b.add_argument("--theme", help="theme name or path to a theme .json (overrides the spec)")
    b.add_argument("--no-raster", action="store_true", help="skip PDF/PNG rendering")
    b.add_argument("--dpi", type=int, default=90)
    b.add_argument("--renderer", default="auto", choices=["auto", "office", "libreoffice", "preview"],
                   help="who draws the review PNGs (auto: Microsoft Office, then LibreOffice, then built-in preview)")
    b.add_argument("--json", action="store_true", help="print the full result as JSON")
    d = sub.add_parser("doctor")
    d.add_argument("--workspace", help="Kiro workspace folder to check for .kiro/ config")
    d.add_argument("--skip-build", action="store_true")
    a = ap.parse_args(argv)

    if a.cmd == "doctor":
        from .doctor import run
        return run(a.workspace, a.skip_build)

    if a.cmd == "schema":
        print(json.dumps(json_schema(a.kind), indent=2))
        return 0
    if a.cmd == "themes":
        print("\n".join(list_themes()))
        return 0
    if a.cmd == "validate":
        try:
            load_spec(read_spec(a.spec))
            print("valid")
            return 0
        except ValidationError as e:
            for x in format_validation_error(e):
                print(f"{x['where']}: {x['message']}")
            return 2
    if a.cmd == "build":
        res = build(a.spec, a.out, a.theme, rasterize=not a.no_raster, dpi=a.dpi, renderer=a.renderer)
        if a.json:
            print(json.dumps(res.to_dict(), indent=2))
        else:
            s = res.summary
            print(f"score {s['score']}/100 · {s['errors']} errors · {s['warnings']} warnings · {s['info']} info")
            if res.output:
                print(f"output: {res.output}")
            if res.contact_sheet:
                print(f"contact sheet: {res.contact_sheet}  (rendered by {res.renderer})")
            print(f"report: {res.report}")
            for x in s["issues"]:
                if x["severity"] != "info":
                    print(f"  [{x['severity']}] {x['where']}: {x['message']}")
        return 0 if res.ok else 2
    return 1


if __name__ == "__main__":
    sys.exit(main())
