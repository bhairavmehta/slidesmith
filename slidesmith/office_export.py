"""Microsoft Office export over COM (Windows).

Runs as a child process (`python -m slidesmith.office_export ...`) so the caller
can time it out: if Office is sitting on a modal dialog (sign-in, activation,
"recover documents?"), a COM call blocks forever, and inside the MCP server that
would hang Kiro's tool call. The child prints one JSON line with the result.

Commands:
  pptx  <file.pptx> <png_dir> <out.pdf> <dpi>   slides -> page-NN.png + PDF (PowerPoint)
  docx  <file.docx> <out.pdf>                   document -> PDF (Word)
  probe                                         PowerPoint/Word versions, for `slidesmith doctor`
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PP_SAVE_AS_PDF = 32
PP_ALERTS_NONE = 1
WD_EXPORT_PDF = 17
MSO_TRUE, MSO_FALSE = -1, 0


def _com():
    import pythoncom
    import win32com.client

    return pythoncom, win32com.client


def _attach_or_start(client, progid: str):
    """Reuse a running instance (and leave it running) or start our own."""
    try:
        return client.GetActiveObject(progid), False
    except Exception:  # noqa: BLE001 — pywintypes.com_error when not running
        return client.DispatchEx(progid), True


def free_path(p: Path) -> Path:
    """p, or a sibling name if p is locked (e.g. the PDF is open in a viewer)."""
    if not p.exists():
        return p
    try:
        p.unlink()
        return p
    except PermissionError:
        for n in range(1, 100):
            alt = p.with_name(f"{p.stem}-{n}{p.suffix}")
            if not alt.exists():
                return alt
        raise


def export_pptx(pptx: Path, png_dir: Path, pdf: Path, dpi: int) -> dict:
    pythoncom, client = _com()
    pythoncom.CoInitialize()
    try:
        app, started = _attach_or_start(client, "PowerPoint.Application")
        old_alerts = None
        try:
            old_alerts = app.DisplayAlerts
            app.DisplayAlerts = PP_ALERTS_NONE
        except Exception:  # noqa: BLE001
            pass
        # Open(FileName, ReadOnly, Untitled, WithWindow)
        pres = app.Presentations.Open(str(pptx.resolve()), MSO_TRUE, MSO_FALSE, MSO_FALSE)
        try:
            w_px = int(pres.PageSetup.SlideWidth / 72 * dpi)
            h_px = int(pres.PageSetup.SlideHeight / 72 * dpi)
            png_dir.mkdir(parents=True, exist_ok=True)
            for old in png_dir.glob("page-*.png"):
                old.unlink()
            pages = []
            for i in range(1, pres.Slides.Count + 1):
                out = (png_dir / f"page-{i:02d}.png").resolve()
                pres.Slides(i).Export(str(out), "PNG", w_px, h_px)
                pages.append(str(out))
            pdf = free_path(pdf)
            pres.SaveAs(str(pdf.resolve()), PP_SAVE_AS_PDF)
        finally:
            pres.Close()
            if old_alerts is not None:
                try:
                    app.DisplayAlerts = old_alerts
                except Exception:  # noqa: BLE001
                    pass
            if started and app.Presentations.Count == 0:
                app.Quit()
        return {"pages": pages, "pdf": str(pdf)}
    finally:
        pythoncom.CoUninitialize()


def export_docx_pdf(docx: Path, pdf: Path) -> dict:
    pythoncom, client = _com()
    pythoncom.CoInitialize()
    try:
        word = client.DispatchEx("Word.Application")  # private instance: never touches the user's Word
        try:
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone
            doc = word.Documents.Open(str(docx.resolve()), ConfirmConversions=False, ReadOnly=True,
                                      AddToRecentFiles=False, Visible=False)
            try:
                pdf = free_path(pdf)
                doc.ExportAsFixedFormat(OutputFileName=str(pdf.resolve()), ExportFormat=WD_EXPORT_PDF)
            finally:
                doc.Close(SaveChanges=0)
        finally:
            word.Quit()
        return {"pdf": str(pdf)}
    finally:
        pythoncom.CoUninitialize()


def probe() -> dict:
    pythoncom, client = _com()
    pythoncom.CoInitialize()
    out = {}
    try:
        for name, progid in (("powerpoint", "PowerPoint.Application"), ("word", "Word.Application")):
            try:
                if name == "word":
                    app, started = client.DispatchEx(progid), True
                else:
                    app, started = _attach_or_start(client, progid)
                out[name] = {"version": str(app.Version)}
                if started and (name == "word" or app.Presentations.Count == 0):
                    app.Quit()
            except Exception as e:  # noqa: BLE001
                out[name] = {"error": f"{type(e).__name__}: {e}"}
    finally:
        pythoncom.CoUninitialize()
    return out


def main(argv: list[str]) -> int:
    try:
        cmd = argv[1]
        if cmd == "pptx":
            res = export_pptx(Path(argv[2]), Path(argv[3]), Path(argv[4]), int(argv[5]))
        elif cmd == "docx":
            res = export_docx_pdf(Path(argv[2]), Path(argv[3]))
        elif cmd == "probe":
            res = probe()
        else:
            raise ValueError(f"unknown command {cmd!r}")
        print(json.dumps(res))
        return 0
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"error": f"{type(e).__name__}: {e}"}))
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
