"""Install slidesmith for the current user and wire it into Kiro. No admin rights needed.

  install.cmd                          (Windows: double-click or run from a terminal)
  python install.py [options]          (any OS)

What it does
  1. creates a private virtual environment in .venv\\ next to this file
  2. pip-installs slidesmith with its Office (pywin32, pypdfium2) and MCP extras
  3. writes a `slidesmith.cmd` launcher next to this file
  4. registers the `slidesmith` MCP server and copies the steering files into Kiro:
       user level   %USERPROFILE%\\.kiro\\   (all workspaces; skip with --no-user-kiro)
       workspace    <DIR>\\.kiro\\           (with --workspace DIR; repeatable)
     An existing mcp.json is merged, never overwritten, and backed up to mcp.json.bak.
  5. runs `slidesmith doctor`

Options
  --workspace DIR     also configure this Kiro workspace (repeatable)
  --no-user-kiro      don't touch %USERPROFILE%\\.kiro
  --index-url URL     corporate PyPI mirror for pip
  --offline           don't download anything: reuse packages already installed in this
                      Python (venv created with --system-site-packages)
  --no-venv           install into this Python itself (pip --user) instead of a private .venv.
                      Use when AppLocker/WDAC blocks python.exe copies in your user folder.
  --skip-doctor       don't run the health check at the end
  --uninstall         remove the venv, the launcher and the Kiro entries this script added
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import venv
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VENV = ROOT / ".venv"
WINDOWS = os.name == "nt"
AUTO_APPROVE = ["get_schema", "list_themes", "validate_spec", "view_pages"]


def say(msg: str):
    print(f"==> {msg}", flush=True)


def die(msg: str, code: int = 1):
    print(f"\nERROR: {msg}", file=sys.stderr, flush=True)
    sys.exit(code)


def venv_python() -> Path:
    return VENV / ("Scripts/python.exe" if WINDOWS else "bin/python")


# --------------------------------------------------------------------------- #
def preflight():
    """Warn about Windows setups that are known to break Python tooling."""
    p = str(ROOT)
    if "onedrive" in p.lower():
        print(f"""
WARNING: slidesmith is inside a OneDrive-synced folder:
    {p}
OneDrive locks and re-uploads files while pip writes thousands of them, which makes the
install slow or fail. Recommended: move the folder to e.g.  C:\\Users\\<you>\\tools\\slidesmith
and run install.cmd from there. Continuing anyway in 10 seconds (Ctrl+C to stop)...""", flush=True)
        try:
            import time
            time.sleep(10)
        except KeyboardInterrupt:
            sys.exit(1)
    if WINDOWS and len(p) > 90:
        print(f"WARNING: install path is long ({len(p)} chars); Windows' 260-char limit can break "
              "pip. A short path like C:\\Users\\<you>\\tools\\slidesmith is safer.", flush=True)


def make_venv(offline: bool):
    py = venv_python()
    if py.exists():
        say(f"Using existing virtual environment {VENV}")
        return py
    say(f"Creating virtual environment {VENV}")
    venv.EnvBuilder(with_pip=True, system_site_packages=offline, clear=False).create(VENV)
    if not py.exists():
        die("venv creation failed (no python in .venv).")
    return py


def pip_install(py: Path, index_url: str | None, offline: bool, user: bool = False):
    target = f"{ROOT}[mcp,office]"
    cmd = [str(py), "-m", "pip", "install", "--disable-pip-version-check", "-e", target]
    if offline:
        cmd = [str(py), "-m", "pip", "install", "--disable-pip-version-check", "--no-deps",
               "--no-build-isolation", "--no-index", "-e", str(ROOT)]
    elif index_url:
        cmd += ["--index-url", index_url]
    if user:
        cmd.insert(cmd.index("install") + 1, "--user")
    say("Installing packages (python-pptx, python-docx, matplotlib, pywin32, pypdfium2, mcp ...)")
    r = subprocess.run(cmd)
    if r.returncode != 0:
        die("pip install failed.\n"
            "  Behind a corporate proxy/mirror?  re-run with  --index-url https://<your-mirror>/simple\n"
            "  or set it once:  python -m pip config set global.index-url https://<your-mirror>/simple\n"
            "  Packages already installed by IT?  re-run with  --offline")


def write_launcher(py: Path) -> Path:
    if WINDOWS:
        p = ROOT / "slidesmith.cmd"
        try:  # relative to the .cmd when possible: survives renamed/odd user folders
            target = '"%~dp0' + str(py.relative_to(ROOT)) + '"'
        except ValueError:
            target = f'"{py}"'
        lines = ["@echo off", f"{target} -m slidesmith %*"]
        if not target.isascii():
            lines.insert(1, "chcp 65001 >nul")  # let cmd read a non-ASCII path
        p.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
    else:
        p = ROOT / "bin" / "slidesmith"
        p.parent.mkdir(exist_ok=True)
        p.write_text(f'#!/bin/sh\nexec "{py}" -m slidesmith "$@"\n', encoding="utf-8")
        p.chmod(0o755)
    say(f"Launcher: {p}")
    return p


# --------------------------------------------------------------------------- #
def _load_json(p: Path) -> dict:
    if not p.exists() or not p.read_text(encoding="utf-8-sig").strip():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as e:
        die(f"{p} is not valid JSON ({e}); fix or remove it and run again.")
    return data if isinstance(data, dict) else {}


def configure_kiro(base: Path, py: Path):
    """base = <something>/.kiro"""
    settings, steering = base / "settings", base / "steering"
    settings.mkdir(parents=True, exist_ok=True)
    steering.mkdir(parents=True, exist_ok=True)
    for f in sorted((ROOT / ".kiro" / "steering").glob("slidesmith-*.md")):
        shutil.copy2(f, steering / f.name)
    cfg_path = settings / "mcp.json"
    cfg = _load_json(cfg_path)
    if cfg_path.exists():
        shutil.copy2(cfg_path, cfg_path.with_suffix(".json.bak"))
    cfg.setdefault("mcpServers", {})["slidesmith"] = {
        "command": str(py),
        "args": ["-m", "slidesmith.mcp_server"],
        "env": {"PYTHONIOENCODING": "utf-8"},
        "disabled": False,
        "autoApprove": AUTO_APPROVE,
    }
    cfg_path.write_text(json.dumps(cfg, indent=2) + "\n", encoding="utf-8")  # no BOM
    say(f"Kiro configured: {base}")


def unconfigure_kiro(base: Path):
    cfg_path = base / "settings" / "mcp.json"
    if cfg_path.exists():
        cfg = _load_json(cfg_path)
        if cfg.get("mcpServers", {}).pop("slidesmith", None) is not None:
            cfg_path.write_text(json.dumps(cfg, indent=2) + "\n", encoding="utf-8")
            say(f"Removed MCP server from {cfg_path}")
    for f in (base / "steering").glob("slidesmith-*.md"):
        f.unlink()
        say(f"Removed {f}")


# --------------------------------------------------------------------------- #
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workspace", action="append", default=[])
    ap.add_argument("--no-user-kiro", action="store_true")
    ap.add_argument("--index-url")
    ap.add_argument("--offline", action="store_true")
    ap.add_argument("--no-venv", action="store_true")
    ap.add_argument("--skip-doctor", action="store_true")
    ap.add_argument("--uninstall", action="store_true")
    a = ap.parse_args(argv)

    if sys.version_info < (3, 10):
        die(f"Python 3.10+ is required (this is {sys.version.split()[0]}).")

    kiro_bases = [] if a.no_user_kiro else [Path.home() / ".kiro"]
    kiro_bases += [Path(w).resolve() / ".kiro" for w in a.workspace]

    if a.uninstall:
        for base in kiro_bases:
            unconfigure_kiro(base)
        for p in (ROOT / "slidesmith.cmd", ROOT / "bin" / "slidesmith"):
            if p.is_file():
                p.unlink()
        if (ROOT / "bin").is_dir() and not any((ROOT / "bin").iterdir()):
            (ROOT / "bin").rmdir()
        if VENV.exists():
            shutil.rmtree(VENV)
            say(f"Removed {VENV}")
        else:  # --no-venv install: remove the package from this Python
            subprocess.run([sys.executable, "-m", "pip", "uninstall", "-y", "slidesmith"])
        say("Uninstalled.")
        return 0

    preflight()
    if a.no_venv:
        py = Path(sys.executable)
        say(f"Installing into this Python (no venv): {py}")
    else:
        py = make_venv(a.offline)
    pip_install(py, a.index_url, a.offline, user=a.no_venv and sys.prefix == sys.base_prefix)
    launcher = write_launcher(py)
    for base in kiro_bases:
        configure_kiro(base, py)

    if not a.skip_doctor:
        print()
        cmd = [str(py), "-m", "slidesmith", "doctor"]
        if a.workspace:
            cmd += ["--workspace", a.workspace[0]]
        subprocess.run(cmd, env={**os.environ, "PYTHONIOENCODING": "utf-8"})

    print(f"""
Done.
  Try it:      "{launcher}" build "{ROOT / 'examples' / 'ai-eval-program.deck.json'}"
  In Kiro:     restart Kiro (or reconnect MCP servers), then ask e.g.
               #slidesmith-workflow Build a 10-slide deck from notes.md for the leadership review
  Uninstall:   install.cmd --uninstall
""")
    return 0


if __name__ == "__main__":
    sys.exit(main())
