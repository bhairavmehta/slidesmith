"""Turn the generated .pptx/.docx into page images so the agent can look at them.

Backends, tried in this order with --renderer auto:

  office       Microsoft PowerPoint / Word do the export themselves — exactly what the
               audience will see. Windows via COM (pip: pywin32); macOS via AppleScript.
               PowerPoint writes PNGs directly; Word writes a PDF that pypdfium2 (pip,
               no system install) turns into PNGs.
  libreoffice  soffice headless -> PDF -> PNG (pypdfium2 or poppler's pdftoppm).
  preview      pure Python (Pillow + matplotlib), always available. Same geometry and
               font metrics as the renderer; approximate pagination for documents.

The backend actually used is recorded in the build report as `renderer`.
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from PIL import Image, ImageDraw, ImageFont

BACKENDS = ("auto", "office", "libreoffice", "preview")


class RasterizeError(RuntimeError):
    pass


@dataclass
class RasterResult:
    renderer: str
    pages: list[Path]
    pdf: Optional[Path] = None
    notes: list[str] = field(default_factory=list)


def _clear_pages(outdir: Path):
    outdir.mkdir(parents=True, exist_ok=True)
    for old in outdir.glob("page-*.png"):
        old.unlink()


def _sorted_pages(outdir: Path) -> list[Path]:
    return sorted(outdir.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[-1]))


# =========================================================================== #
# PDF -> PNG (pypdfium2 first: pip-installable, no system dependency)
# =========================================================================== #
def pdf_to_pngs(pdf: Path, outdir: Path, dpi: int = 90) -> list[Path]:
    _clear_pages(outdir)
    try:
        import pypdfium2 as pdfium

        doc = pdfium.PdfDocument(str(pdf))
        try:
            for i in range(len(doc)):
                img = doc[i].render(scale=dpi / 72).to_pil()
                img.convert("RGB").save(outdir / f"page-{i + 1:02d}.png")
        finally:
            doc.close()
        return _sorted_pages(outdir)
    except ImportError:
        pass
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(outdir / "page")],
                       check=True, capture_output=True)
        return _sorted_pages(outdir)
    raise RasterizeError("No PDF rasterizer: `pip install pypdfium2` (no admin rights needed).")


# =========================================================================== #
# Microsoft Office
# =========================================================================== #
def office_available() -> tuple[bool, str]:
    system = platform.system()
    if system == "Windows":
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            return False, "pywin32 not installed (`pip install pywin32`)."
        return True, ""
    if system == "Darwin":
        apps = ["/Applications/Microsoft PowerPoint.app", "/Applications/Microsoft Word.app"]
        if all(os.path.exists(a) for a in apps):
            return True, ""
        return False, "Microsoft PowerPoint/Word not found in /Applications."
    return False, "Microsoft Office export needs Windows or macOS."


OFFICE_TIMEOUT_S = int(os.environ.get("SLIDESMITH_OFFICE_TIMEOUT", "240"))


def run_office_child(args: list[str], timeout: int = OFFICE_TIMEOUT_S) -> dict:
    """Run slidesmith.office_export in a child process (or in-process when
    SLIDESMITH_OFFICE_INPROC=1, for debugging). Returns its JSON result."""
    import json

    if os.environ.get("SLIDESMITH_OFFICE_INPROC") == "1":
        from . import office_export
        import io
        import contextlib

        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            office_export.main(["office_export"] + args)
        res = json.loads(buf.getvalue().strip().splitlines()[-1])
    else:
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)  # no console flash on Windows
        env = os.environ.copy()
        pkg_root = str(Path(__file__).resolve().parent.parent)  # works without `pip install` too
        env["PYTHONPATH"] = pkg_root + (os.pathsep + env["PYTHONPATH"] if env.get("PYTHONPATH") else "")
        env["PYTHONIOENCODING"] = "utf-8"
        try:
            r = subprocess.run([sys.executable, "-m", "slidesmith.office_export"] + args,
                               capture_output=True, text=True, encoding="utf-8", errors="replace",
                               timeout=timeout, creationflags=flags, env=env)
        except subprocess.TimeoutExpired:
            raise RasterizeError(
                f"Office did not respond within {timeout}s. It is probably showing a dialog "
                "(sign-in, activation, document recovery). Open PowerPoint and Word once, dismiss "
                "any prompts, close them, and build again. A hidden WINWORD.EXE left over from this "
                "attempt can be ended in Task Manager.")
        lines = [ln for ln in (r.stdout or "").strip().splitlines() if ln.startswith("{")]
        if not lines:
            raise RasterizeError(f"Office export produced no result: {(r.stderr or '').strip()[-400:]}")
        res = json.loads(lines[-1])
    if "error" in res:
        raise RasterizeError(f"Office export failed: {res['error']}")
    return res


def _mac_export_pdf(src: Path, pdf: Path, app: str):
    """AppleScript export. Office for Mac is sandboxed, so we export inside its container
    (no permission prompt) and move the result out."""
    bundle = {"PowerPoint": "com.microsoft.Powerpoint", "Word": "com.microsoft.Word"}[app]
    box = Path.home() / "Library" / "Containers" / bundle / "Data" / "tmp" / "slidesmith"
    box.mkdir(parents=True, exist_ok=True)
    staged_src = box / src.name
    staged_pdf = box / pdf.name
    shutil.copy2(src, staged_src)
    if app == "PowerPoint":
        script = f'''
        tell application "Microsoft PowerPoint"
            set p to open POSIX file "{staged_src}"
            save p in POSIX file "{staged_pdf}" as save as PDF
            close p saving no
        end tell'''
    else:
        script = f'''
        tell application "Microsoft Word"
            open POSIX file "{staged_src}"
            set d to active document
            save as d file name POSIX path of "{staged_pdf}" file format format PDF
            close d saving no
        end tell'''
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=300)
    if r.returncode != 0 or not staged_pdf.exists():
        raise RasterizeError(f"{app} export failed: {r.stderr.strip()}")
    shutil.move(str(staged_pdf), pdf)
    staged_src.unlink(missing_ok=True)


def render_office(target: Path, outdir: Path, dpi: int) -> RasterResult:
    ok, why = office_available()
    if not ok:
        raise RasterizeError(why)
    pdf = target.with_suffix(".pdf")
    is_deck = target.suffix.lower() == ".pptx"
    if platform.system() == "Windows":
        if is_deck:
            res = run_office_child(["pptx", str(target), str(outdir), str(pdf), str(dpi)])
            return RasterResult("office", [Path(p) for p in res["pages"]], Path(res["pdf"]))
        res = run_office_child(["docx", str(target), str(pdf)])
        pdf = Path(res["pdf"])
    else:
        _mac_export_pdf(target, pdf, "PowerPoint" if is_deck else "Word")
    return RasterResult("office", pdf_to_pngs(pdf, outdir, dpi), pdf)


# =========================================================================== #
# LibreOffice
# =========================================================================== #
def _soffice() -> Optional[str]:
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return p
    return None


def render_libreoffice(target: Path, outdir: Path, dpi: int, timeout: int = 180) -> RasterResult:
    exe = _soffice()
    if not exe:
        raise RasterizeError("LibreOffice not found.")
    with tempfile.TemporaryDirectory() as prof:
        cmd = [exe, f"-env:UserInstallation={Path(prof).as_uri()}", "--headless",
               "--convert-to", "pdf", "--outdir", str(target.parent), str(target)]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    pdf = target.with_suffix(".pdf")
    if not pdf.exists():
        raise RasterizeError(f"LibreOffice PDF conversion failed: {r.stderr.strip() or r.stdout.strip()}")
    return RasterResult("libreoffice", pdf_to_pngs(pdf, outdir, dpi), pdf)


# =========================================================================== #
# Pure-Python preview
# =========================================================================== #
def render_preview(target: Path, outdir: Path, dpi: int, spec=None, theme=None) -> RasterResult:
    from .preview import preview_doc, preview_pptx
    from .theme import load_theme

    theme = theme or load_theme("default")
    if target.suffix.lower() == ".pptx":
        pages = preview_pptx(target, outdir, theme, dpi)
    else:
        if spec is None:
            raise RasterizeError("Document preview needs the spec.")
        pages = preview_doc(spec, theme, outdir, dpi)
    note = ("Pages drawn by the built-in preview, not by Office: layout and text fit are "
            "faithful; document pagination is approximate.")
    return RasterResult("preview", pages, None, [note])


# =========================================================================== #
def render_pages(target: Path, outdir: Path, dpi: int = 90, renderer: str = "auto",
                 spec=None, theme=None) -> RasterResult:
    """Render with the requested backend; 'auto' falls through office → libreoffice → preview."""
    if renderer not in BACKENDS:
        raise RasterizeError(f"Unknown renderer '{renderer}'. Use one of {', '.join(BACKENDS)}.")
    order = ["office", "libreoffice", "preview"] if renderer == "auto" else [renderer]
    skipped = []
    for name in order:
        try:
            if name == "office":
                res = render_office(target, outdir, dpi)
            elif name == "libreoffice":
                res = render_libreoffice(target, outdir, dpi)
            else:
                res = render_preview(target, outdir, dpi, spec, theme)
            res.notes = [f"{s} unavailable: {why}" for s, why in skipped] + res.notes
            return res
        except RasterizeError as e:
            skipped.append((name, str(e)))
        except Exception as e:  # noqa: BLE001 — COM/AppleScript failures surface as many types
            skipped.append((name, f"{type(e).__name__}: {e}"))
    raise RasterizeError("; ".join(f"{s}: {why}" for s, why in skipped))


def contact_sheet(pages: list[Path], out: Path, cols: int = 3, thumb_w: int = 560) -> Path:
    """One image with every page labelled — cheap for a vision model to review at once."""
    if not pages:
        raise RasterizeError("no pages to put on a contact sheet")
    thumbs = []
    for p in pages:
        im = Image.open(p).convert("RGB")
        r = thumb_w / im.width
        thumbs.append(im.resize((thumb_w, int(im.height * r)), Image.LANCZOS))
    th = max(t.height for t in thumbs)
    cols = min(cols, len(thumbs))
    gap, label_h = 24, 30
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * thumb_w + (cols + 1) * gap,
                              rows * (th + label_h) + (rows + 1) * gap), (232, 234, 238))
    d = ImageDraw.Draw(sheet)
    try:
        from .textfit import font_path
        f = ImageFont.truetype(font_path("DejaVu Sans", bold=True), 18)
    except OSError:
        f = ImageFont.load_default()
    for k, im in enumerate(thumbs):
        r, c = divmod(k, cols)
        x = gap + c * (thumb_w + gap)
        y = gap + r * (th + label_h + gap)
        d.text((x, y), f"{k + 1}", fill=(60, 66, 76), font=f)
        sheet.paste(im, (x, y + label_h))
        d.rectangle([x - 1, y + label_h - 1, x + im.width, y + label_h + im.height], outline=(190, 196, 204))
    sheet.save(out)
    return out
