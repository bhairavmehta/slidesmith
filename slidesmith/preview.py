"""Pure-Python page previews — no LibreOffice, Office or poppler required.

* Slides: reads the generated .pptx with python-pptx and draws every shape it
  contains (rectangles, ovals, arrows, text boxes with bullets, tables, charts).
  Because slidesmith only emits that small vocabulary of shapes, the drawing is
  close to what PowerPoint shows, and it uses the same font metrics the renderer
  used to size the text — so overflow and balance problems show up faithfully.
* Documents: lays the doc spec out on pages with the same theme tokens the .docx
  uses. Pagination is approximate (Word's line breaking differs slightly), which
  is fine for judging hierarchy, density and rhythm.

Fidelity is labelled "preview" in the build report so the reviewing agent knows.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Optional

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn

from . import markup
from .charts_mpl import render_chart_png
from .schema import ChartSpec, Document, Series, bullet_subs, bullet_text
from .textfit import LINE_HEIGHT, font_path
from .theme import Theme, hex_to_rgb, mix

EMU_PER_IN = 914400


# =========================================================================== #
# Rich-text layout shared by slides and docs
# =========================================================================== #
@dataclass(frozen=True)
class Style:
    font: str
    size_pt: float
    bold: bool = False
    italic: bool = False
    color: str = "000000"


@lru_cache(maxsize=1024)
def _pil_font(family: str, bold: bool, italic: bool, px: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(font_path(family, bold, italic), size=max(1, px))


class Typesetter:
    def __init__(self, dpi: int):
        self.dpi = dpi

    def px(self, inches: float) -> int:
        return int(round(inches * self.dpi))

    def font(self, st: Style) -> ImageFont.FreeTypeFont:
        return _pil_font(st.font, st.bold, st.italic, int(round(st.size_pt * self.dpi / 72)))

    def layout(self, tokens: list[tuple[str, Style]], width_px: float) -> list[list[tuple[str, Style, float]]]:
        """Greedy word wrap over styled runs. Returns lines of (text, style, x_offset)."""
        words: list[tuple[str, Style, bool]] = []  # (word, style, space_before)
        pending_space = False
        for text, st in tokens:
            parts = re.split(r"(\s+)", text)
            for part in parts:
                if not part:
                    continue
                if part.isspace():
                    pending_space = True
                    continue
                words.append((part, st, pending_space))
                pending_space = False
        lines: list[list[tuple[str, Style, float]]] = [[]]
        x = 0.0
        for w, st, sp in words:
            f = self.font(st)
            space = f.getlength(" ") if sp and lines[-1] else 0.0
            ww = f.getlength(w)
            if lines[-1] and x + space + ww > width_px:
                lines.append([])
                x, space = 0.0, 0.0
            lines[-1].append((w, st, x + space))
            x += space + ww
        return lines

    def line_width(self, line) -> float:
        if not line:
            return 0.0
        w, st, x = line[-1]
        return x + self.font(st).getlength(w)

    def line_height(self, line, default: Style, spacing: float) -> float:
        size = max([st.size_pt for _, st, _ in line] or [default.size_pt])
        return size * self.dpi / 72 * LINE_HEIGHT * spacing

    def draw_lines(self, draw: ImageDraw.ImageDraw, lines, x: float, y: float, width: float,
                   default: Style, spacing: float, align: str = "l") -> float:
        """Draw wrapped lines; returns the height used."""
        top = y
        for line in lines:
            lh = self.line_height(line, default, spacing)
            big = max([st for _, st, _ in line] or [default], key=lambda s: s.size_pt)
            asc, desc = self.font(big).getmetrics()
            baseline = y + (lh - (asc + desc)) / 2 + asc
            lw = self.line_width(line)
            dx = {"l": 0, "c": (width - lw) / 2, "r": width - lw}[align]
            for w, st, ox in line:
                draw.text((x + dx + ox, baseline), w, font=self.font(st), fill="#" + st.color, anchor="ls")
            y += lh
        return y - top


# =========================================================================== #
# PPTX -> PNG
# =========================================================================== #
def _hex(color_format, default: str) -> str:
    try:
        rgb = color_format.rgb
        return str(rgb) if rgb is not None else default
    except (AttributeError, TypeError, ValueError):
        return default


def _solid_fill_hex(el) -> Optional[str]:
    """First a:solidFill/a:srgbClr under el (spPr or tcPr)."""
    if el is None:
        return None
    sf = el.find(qn("a:solidFill"))
    if sf is None:
        return None
    clr = sf.find(qn("a:srgbClr"))
    return clr.get("val") if clr is not None else None


CHART_BACK = {
    XL_CHART_TYPE.COLUMN_CLUSTERED: "column", XL_CHART_TYPE.BAR_CLUSTERED: "bar",
    XL_CHART_TYPE.LINE_MARKERS: "line", XL_CHART_TYPE.LINE: "line", XL_CHART_TYPE.PIE: "pie",
    XL_CHART_TYPE.DOUGHNUT: "doughnut", XL_CHART_TYPE.COLUMN_STACKED: "stacked_column",
    XL_CHART_TYPE.BAR_STACKED: "stacked_bar",
}


class SlidePreview:
    def __init__(self, theme: Theme, dpi: int = 90):
        self.t = theme
        self.ts = Typesetter(dpi)
        self.dpi = dpi

    def e2p(self, emu) -> float:
        return int(emu or 0) / EMU_PER_IN * self.dpi

    def render(self, pptx_path: str | Path, outdir: Path) -> list[Path]:
        prs = Presentation(str(pptx_path))
        W, H = self.e2p(prs.slide_width), self.e2p(prs.slide_height)
        outdir.mkdir(parents=True, exist_ok=True)
        for old in outdir.glob("page-*.png"):
            old.unlink()
        pages = []
        for n, slide in enumerate(prs.slides, start=1):
            bg = "FFFFFF"
            try:
                if slide.background.fill.type is not None:
                    bg = _hex(slide.background.fill.fore_color, bg)
            except (AttributeError, TypeError):
                pass
            im = Image.new("RGB", (int(round(W)), int(round(H))), "#" + bg)
            draw = ImageDraw.Draw(im)
            for shp in slide.shapes:
                self._shape(im, draw, shp)
            p = outdir / f"page-{n:02d}.png"
            im.save(p)
            pages.append(p)
        return pages

    # ---------------------------------------------------------------- shapes
    def _shape(self, im, draw, shp):
        x, y = self.e2p(shp.left), self.e2p(shp.top)
        w, h = self.e2p(shp.width), self.e2p(shp.height)
        if shp.has_chart:
            self._chart(im, shp, x, y, w, h)
            return
        if shp.has_table:
            self._table(draw, shp.table, x, y)
            return
        if shp.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            fill = _solid_fill_hex(shp._element.spPr)
            if fill:
                kind = shp.auto_shape_type
                box = [x, y, x + w - 1, y + h - 1]
                if kind == MSO_SHAPE.OVAL:
                    draw.ellipse(box, fill="#" + fill)
                elif kind == MSO_SHAPE.RIGHT_ARROW:
                    shaft = h * 0.25
                    head = min(w * 0.5, h * 0.5)
                    draw.polygon([(x, y + shaft), (x + w - head, y + shaft), (x + w - head, y),
                                  (x + w, y + h / 2), (x + w - head, y + h), (x + w - head, y + h - shaft),
                                  (x, y + h - shaft)], fill="#" + fill)
                else:
                    draw.rectangle(box, fill="#" + fill)
        if shp.has_text_frame and shp.text_frame.text.strip():
            self._text_frame(draw, shp.text_frame, x, y, w, h)

    # ---------------------------------------------------------------- text
    def _paragraphs(self, tf, width_px: float):
        """Lay out each paragraph -> list of (lines, default_style, spacing, align, bullet, indent_px, space_before_px)."""
        out = []
        auto_no = 0
        for p in tf.paragraphs:
            runs = [r for r in p.runs]
            if not runs and not p.text:
                continue
            tokens = []
            first = None
            for r in runs:
                f = r.font
                st = Style(
                    font=f.name or self.t.body_font,
                    size_pt=f.size.pt if f.size else 18.0,
                    bold=bool(f.bold), italic=bool(f.italic),
                    color=_hex(f.color, self.t.c("ink")) if f.color and f.color.type is not None else self.t.c("ink"),
                )
                first = first or st
                tokens.append((r.text, st))
            first = first or Style(self.t.body_font, 18)
            ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
            sb = p.space_before.pt * self.dpi / 72 if p.space_before is not None else 0.0
            align = {PP_ALIGN.CENTER: "c", PP_ALIGN.RIGHT: "r"}.get(p.alignment, "l")
            pPr = p._p.pPr
            bullet, mar, ind = None, 0.0, 0.0
            if pPr is not None:
                mar = self.e2p(pPr.get("marL") or 0)
                ind = self.e2p(pPr.get("indent") or 0)
                bc = pPr.find(qn("a:buChar"))
                if bc is not None:
                    bullet = bc.get("char")
                elif pPr.find(qn("a:buAutoNum")) is not None:
                    auto_no += 1
                    bullet = f"{auto_no}."
                if bullet:
                    clr = pPr.find(qn("a:buClr"))
                    bcol = clr.find(qn("a:srgbClr")).get("val") if clr is not None and clr.find(qn("a:srgbClr")) is not None else first.color
                    bullet = (bullet, Style(self.t.body_font, first.size_pt, False, False, bcol))
            lines = self.ts.layout(tokens, width_px - mar)
            out.append((lines, first, ls, align, bullet, mar, ind, sb))
        return out

    def _text_frame(self, draw, tf, x, y, w, h, anchor=None):
        ml, mr = self.e2p(tf.margin_left), self.e2p(tf.margin_right)
        mt, mb = self.e2p(tf.margin_top), self.e2p(tf.margin_bottom)
        x0, y0, iw, ih = x + ml, y + mt, w - ml - mr, h - mt - mb
        paras = self._paragraphs(tf, iw)
        total = sum(sb + sum(self.ts.line_height(l, st, ls) for l in lines)
                    for lines, st, ls, _a, _b, _m, _i, sb in paras)
        va = anchor if anchor is not None else tf.vertical_anchor
        if va == MSO_ANCHOR.MIDDLE:
            y0 += (ih - total) / 2
        elif va == MSO_ANCHOR.BOTTOM:
            y0 += ih - total
        cy = y0
        for k, (lines, st, ls, align, bullet, mar, ind, sb) in enumerate(paras):
            if k:
                cy += sb
            if bullet:
                ch, bst = bullet
                lh = self.ts.line_height(lines[0] if lines else [], st, ls)
                self.ts.draw_lines(draw, [[(ch, bst, 0.0)]], x0 + mar + ind, cy, iw, st, ls)
            cy += self.ts.draw_lines(draw, lines, x0 + mar, cy, iw - mar, st, ls, align)

    # ---------------------------------------------------------------- tables
    def _table(self, draw, tbl, x, y):
        col_w = [self.e2p(c.width) for c in tbl.columns]
        cy = y
        for r_i, row in enumerate(tbl.rows):
            rh = self.e2p(row.height)
            cx = x
            for c_i, cell in enumerate(row.cells):
                cw = col_w[c_i]
                tcPr = cell._tc.tcPr
                fill = _solid_fill_hex(tcPr)
                if fill:
                    draw.rectangle([cx, cy, cx + cw, cy + rh], fill="#" + fill)
                if tcPr is not None:
                    lnB = tcPr.find(qn("a:lnB"))
                    if lnB is not None:
                        col = _solid_fill_hex(lnB) or self.t.c("line")
                        wpx = max(1, round(int(lnB.get("w", "9525")) / 12700 * self.dpi / 72))
                        draw.rectangle([cx, cy + rh - wpx, cx + cw, cy + rh], fill="#" + col)
                tf = cell.text_frame
                ml, mr = self.e2p(cell.margin_left), self.e2p(cell.margin_right)
                mt, mb = self.e2p(cell.margin_top), self.e2p(cell.margin_bottom)

                class _TF:  # adapter: cells keep their own margins
                    paragraphs = tf.paragraphs
                    margin_left = margin_right = margin_top = margin_bottom = 0
                    vertical_anchor = cell.vertical_anchor

                self._text_frame(draw, _TF, cx + ml, cy + mt, cw - ml - mr, rh - mt - mb)
                cx += cw
            cy += rh

    # ---------------------------------------------------------------- charts
    def _chart(self, im, shp, x, y, w, h):
        ch = shp.chart
        plot = ch.plots[0]
        kind = CHART_BACK.get(ch.chart_type, "column")
        cats = [str(c) for c in plot.categories]
        series = [Series(name=s.name or f"Series {i + 1}", values=[v or 0 for v in s.values])
                  for i, s in enumerate(plot.series)]
        nf = "General"
        try:
            if plot.has_data_labels:
                nf = plot.data_labels.number_format
            elif kind not in ("pie", "doughnut"):
                nf = ch.value_axis.tick_labels.number_format
        except (AttributeError, ValueError):
            pass
        highlight = None
        if len(series) >= 1 and kind not in ("line",):
            ser = list(plot.series)[0]._element
            cols = {}
            for dpt in ser.findall(qn("c:dPt")):
                idx = int(dpt.find(qn("c:idx")).get("val"))
                c = _solid_fill_hex(dpt.find(qn("c:spPr")))
                if c:
                    cols[idx] = c
            if cols and len(cols) == len(cats):
                counts = {}
                for v in cols.values():
                    counts[v] = counts.get(v, 0) + 1
                odd = [i for i, v in cols.items() if counts[v] == 1]
                if len(odd) == 1 and len(cats) > 2:
                    highlight = odd[0]
                elif kind in ("pie", "doughnut"):
                    highlight = None
        spec = ChartSpec(type=kind, categories=cats, series=series, number_format=nf or "General",
                         data_labels=bool(plot.has_data_labels), highlight=highlight)
        buf = io.BytesIO()
        render_chart_png(spec, self.t, buf, width_in=w / self.dpi, height_in=h / self.dpi,
                         dpi=self.dpi, base_pt=13)
        buf.seek(0)
        chart_im = Image.open(buf).convert("RGB")
        im.paste(chart_im.resize((int(w), int(h))), (int(x), int(y)))


def preview_pptx(pptx_path, outdir: Path, theme: Theme, dpi: int = 90) -> list[Path]:
    return SlidePreview(theme, dpi).render(pptx_path, Path(outdir))


# =========================================================================== #
# Document spec -> PNG pages
# =========================================================================== #
NUMERIC = re.compile(r"^[\s$€£¥+\-−~≈<>]*[\d.,]+\s*(%|x|pp|bps|[KMBkmb]|ms|s|pts?)?\s*$")


class DocPreview:
    """A small page-layout engine mirroring render_docx's styles."""

    def __init__(self, doc: Document, theme: Theme, dpi: int = 90):
        self.spec, self.t, self.dpi = doc, theme, dpi
        self.ts = Typesetter(dpi)
        self.sz = theme.doc["size"]
        self.ls = theme.doc.get("line_spacing", 1.2)
        a4 = theme.doc.get("page", "letter") == "a4"
        self.pw, self.ph = (8.27, 11.69) if a4 else (8.5, 11.0)
        self.m = theme.doc.get("margin_in", 1.0)
        self.pages: list[Image.Image] = []
        self.new_page()

    # ---------------- geometry ----------------
    def P(self, inches):
        return inches * self.dpi

    def pt(self, points):
        return points * self.dpi / 72

    @property
    def left(self): return self.P(self.m)

    @property
    def width(self): return self.P(self.pw - 2 * self.m)

    @property
    def bottom(self): return self.P(self.ph - self.m)

    def new_page(self):
        im = Image.new("RGB", (int(self.P(self.pw)), int(self.P(self.ph))), "#" + self.t.c("bg"))
        self.pages.append(im)
        self.draw = ImageDraw.Draw(im)
        self.y = self.P(self.m)

    def ensure(self, h):
        if self.y + h > self.bottom and self.y > self.P(self.m) + 1:
            self.new_page()

    # ---------------- text helpers ----------------
    def st(self, size, color=None, bold=False, italic=False, font=None) -> Style:
        return Style(font or self.t.body_font, size, bold, italic, (color or self.t.c("ink")))

    def tokens(self, text: str, base: Style):
        return [(seg, Style(base.font, base.size_pt, base.bold or b, base.italic or it, base.color))
                for seg, b, it in markup.runs(text)]

    def flow(self, text: str, base: Style, x=None, width=None, spacing=None, after_pt=0.0,
             align="l", before_pt=0.0):
        """Paragraph that may break across pages line by line."""
        x = self.left if x is None else x
        width = self.width if width is None else width
        spacing = spacing or self.ls
        self.y += self.pt(before_pt)
        lines = self.ts.layout(self.tokens(text, base), width)
        for line in lines:
            lh = self.ts.line_height(line, base, spacing)
            self.ensure(lh)
            self.y += self.ts.draw_lines(self.draw, [line], x, self.y, width, base, spacing, align)
        self.y += self.pt(after_pt)

    def measure(self, text: str, base: Style, width, spacing=None) -> float:
        spacing = spacing or self.ls
        lines = self.ts.layout(self.tokens(text, base), width)
        return sum(self.ts.line_height(l, base, spacing) for l in lines)

    # ---------------- render ----------------
    def render(self, outdir: Path) -> list[Path]:
        t, sz = self.t, self.sz
        self.flow(self.spec.title, self.st(sz["title"], bold=True, font=t.heading_font), spacing=1.05, after_pt=4)
        if self.spec.subtitle:
            self.flow(self.spec.subtitle, self.st(sz["subtitle"], t.c("muted"), font=t.heading_font),
                      spacing=1.05, after_pt=6)
        meta = " · ".join(v for v in [self.spec.author, self.spec.date] if v)
        if meta:
            self.flow(meta, self.st(sz["small"] + 1, t.c("muted")))
        self.y += self.pt(6)
        self.draw.rectangle([self.left, self.y, self.left + self.width, self.y + self.pt(1.5)],
                            fill="#" + t.c("accent"))
        self.y += self.pt(18)

        blocks = self.spec.blocks
        for i, b in enumerate(blocks):
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            getattr(self, f"_{b.type}")(b, nxt)

        # running header + page numbers
        for n, im in enumerate(self.pages, start=1):
            d = ImageDraw.Draw(im)
            small = self.st(sz["small"], t.c("muted"))
            if n > 1:
                txt = self.spec.header or self.spec.title
                self.ts.draw_lines(d, self.ts.layout([(txt, small)], self.width), self.left,
                                   self.P(0.5), self.width, small, 1.0, "r")
            self.ts.draw_lines(d, [[(str(n), small, 0.0)]], self.left, self.P(self.ph - 0.6),
                               self.width, small, 1.0, "r")
        outdir.mkdir(parents=True, exist_ok=True)
        for old in outdir.glob("page-*.png"):
            old.unlink()
        paths = []
        for n, im in enumerate(self.pages, start=1):
            p = outdir / f"page-{n:02d}.png"
            im.save(p)
            paths.append(p)
        return paths

    def _heading(self, b, nxt):
        t, sz = self.t, self.sz
        size, color, before, after = {
            1: (sz["h1"], t.c("primary"), 22, 8), 2: (sz["h2"], t.c("ink"), 16, 6),
            3: (sz["h3"], t.c("accent"), 12, 4)}[b.level]
        st = self.st(size, color, bold=True, font=t.heading_font)
        h = self.measure(b.text, st, self.width, 1.05)
        body = self.pt(sz["body"]) * LINE_HEIGHT * self.ls * 3  # keep with ~3 lines of what follows
        self.y += self.pt(before)
        self.ensure(h + self.pt(after) + body)
        self.flow(b.text, st, spacing=1.05, after_pt=after)

    def _paragraph(self, b, nxt):
        if nxt is not None and nxt.type == "chart":  # mirror Word's keep-with-next before figures
            self.ensure(self.measure(b.text, self.st(self.sz["body"]), self.width) + self.P(3.2))
        self.flow(b.text, self.st(self.sz["body"]), after_pt=self.t.doc.get("para_after_pt", 8))

    def _bullets(self, b, nxt):
        body = self.st(self.sz["body"])
        for k, item in enumerate(b.items, start=1):
            glyph = f"{k}." if b.ordered else "•"
            line_h = self.pt(self.sz["body"]) * LINE_HEIGHT * self.ls
            self.ensure(line_h)
            self.ts.draw_lines(self.draw, [[(glyph, body, 0.0)]], self.left + self.P(0.25), self.y,
                               self.P(0.25), body, self.ls)
            self.flow(bullet_text(item), body, x=self.left + self.P(0.5), width=self.width - self.P(0.5),
                      after_pt=4)
            for s in bullet_subs(item):
                self.ensure(line_h)
                self.ts.draw_lines(self.draw, [[("◦", body, 0.0)]], self.left + self.P(0.5), self.y,
                                   self.P(0.25), body, self.ls)
                self.flow(s, self.st(self.sz["body"], self.t.c("muted")), x=self.left + self.P(0.75),
                          width=self.width - self.P(0.75), after_pt=4)
        self.y += self.pt(self.t.doc.get("para_after_pt", 8) - 2)

    def _box_text(self, parts: list[tuple[str, Style, float]], x, width) -> float:
        """Height of stacked (text, style, after_pt) parts."""
        return sum(self.measure(tx, st, width) + self.pt(a) for tx, st, a in parts)

    def _callout(self, b, nxt):
        t = self.t
        stripe = {"key": t.c("accent"), "info": t.c("primary"), "warning": t.c("negative")}[b.kind]
        padx, pady = self.P(0.15), self.P(0.1)
        parts = []
        if b.title:
            parts.append((b.title, self.st(self.sz["body"], bold=True), 2))
        parts.append((b.text, self.st(self.sz["body"]), 0))
        inner = self.width - 2 * padx - self.pt(3)
        h = self._box_text(parts, 0, inner) + 2 * pady
        self.ensure(h)
        y0 = self.y
        self.draw.rectangle([self.left, y0, self.left + self.width, y0 + h], fill="#" + mix(stripe, t.c("bg"), 0.9))
        self.draw.rectangle([self.left, y0, self.left + self.pt(3), y0 + h], fill="#" + stripe)
        self.y = y0 + pady
        for tx, st, a in parts:
            self.flow(tx, st, x=self.left + self.pt(3) + padx, width=inner, after_pt=a)
        self.y = y0 + h + self.pt(12)

    def _table(self, b, nxt):
        t, sz = self.t, self.sz
        lens = [max([len(b.columns[c])] + [len(r[c]) for r in b.rows]) for c in range(len(b.columns))]
        weights = [max(6, min(n, 40)) ** 0.75 for n in lens]
        widths = [self.width * w / sum(weights) for w in weights]
        numeric = [sum(bool(NUMERIC.match(r[c])) for r in b.rows) >= max(1, len(b.rows) * 0.6)
                   for c in range(len(b.columns))]
        padx, pady = self.P(0.076), self.P(0.05)
        head_st = self.st(sz["body"] - 1, t.c("muted"), bold=True)
        cell_st = self.st(sz["body"] - 0.5)

        def row_h(cells, st):
            return max(self.measure(c, st, widths[i] - 2 * padx, 1.1) for i, c in enumerate(cells)) + 2 * pady

        def draw_row(cells, st, header):
            h = row_h(cells, st)
            x = self.left
            for i, c in enumerate(cells):
                lines = self.ts.layout(self.tokens(c, st), widths[i] - 2 * padx)
                th = sum(self.ts.line_height(l, st, 1.1) for l in lines)
                self.ts.draw_lines(self.draw, lines, x + padx, self.y + (h - th) / 2, widths[i] - 2 * padx,
                                   st, 1.1, "r" if numeric[i] and i > 0 else "l")
                x += widths[i]
            line_col = t.c("primary") if header else t.c("line")
            lw = self.pt(1.5 if header else 0.5)
            self.draw.rectangle([self.left, self.y + h - lw, self.left + self.width, self.y + h], fill="#" + line_col)
            self.y += h

        self.ensure(row_h(b.columns, head_st) + row_h(b.rows[0], cell_st))
        draw_row(b.columns, head_st, True)
        for r in b.rows:
            if self.y + row_h(r, cell_st) > self.bottom:
                self.new_page()
                draw_row(b.columns, head_st, True)  # repeated header row
            draw_row(r, cell_st, False)
        self._caption(b.caption)

    def _caption(self, text):
        if text:
            self.flow(text, self.st(self.sz["small"], self.t.c("muted"), italic=True), before_pt=4,
                      after_pt=self.t.doc.get("para_after_pt", 8))
        else:
            self.y += self.pt(10)

    def _chart(self, b, nxt):
        w_in, h_in = self.width / self.dpi, 3.2
        self.ensure(self.P(h_in))
        buf = io.BytesIO()
        render_chart_png(b.chart, self.t, buf, width_in=w_in, height_in=h_in, dpi=self.dpi)
        buf.seek(0)
        self.pages[-1].paste(Image.open(buf).convert("RGB"), (int(self.left), int(self.y)))
        self.y += self.P(h_in) + self.pt(2)
        self._caption(b.caption)

    def _kpis(self, b, nxt):
        t, sz = self.t, self.sz
        n = len(b.items)
        cw = self.width / n
        padx, pady = self.P(0.11), self.P(0.083)
        parts_all = []
        for k in b.items:
            parts = [(k.value, self.st(sz["kpi_value"], bold=True, font=t.heading_font), 2),
                     (k.label, self.st(sz["body"] - 1.5, t.c("muted")), 2)]
            if k.delta:
                sent = k.sentiment or ("negative" if k.delta.strip().startswith(("-", "−", "▼", "↓")) else "positive")
                col = {"positive": t.c("positive"), "negative": t.c("negative"), "neutral": t.c("muted")}[sent]
                parts.append((k.delta, self.st(sz["body"] - 2, col, bold=True), 0))
            parts_all.append(parts)
        h = max(self._box_text(p, 0, cw - 2 * padx) for p in parts_all) + 2 * pady
        self.ensure(h)
        y0 = self.y
        for j, parts in enumerate(parts_all):
            x = self.left + j * cw
            gap = self.pt(1.5)
            self.draw.rectangle([x + (gap if j else 0), y0, x + cw - (gap if j < n - 1 else 0), y0 + h],
                                fill="#" + t.c("surface"))
            self.draw.rectangle([x + (gap if j else 0), y0, x + cw - (gap if j < n - 1 else 0), y0 + self.pt(2.25)],
                                fill="#" + (t.c("accent") if j == 0 else t.c("primary")))
            self.y = y0 + pady
            for tx, st, a in parts:
                self.flow(tx, st, x=x + padx, width=cw - 2 * padx, after_pt=a, spacing=1.1)
        self.y = y0 + h + self.pt(14)

    def _quote(self, b, nxt):
        t = self.t
        x = self.left + self.P(0.25) + self.pt(12)
        w = self.width - self.P(0.25) - self.pt(12)
        qst = self.st(self.sz["body"] + 3, italic=True, font=t.heading_font)
        ast = self.st(self.sz["small"] + 1, t.c("muted"), bold=True)
        h = self.measure(b.text, qst, w) + (self.measure(f"— {b.attribution}", ast, w) + self.pt(2)
                                             if b.attribution else 0)
        self.y += self.pt(6)
        self.ensure(h)
        y0 = self.y
        self.draw.rectangle([self.left + self.P(0.25), y0, self.left + self.P(0.25) + self.pt(3), y0 + h],
                            fill="#" + t.c("accent"))
        self.flow(b.text, qst, x=x, width=w, after_pt=2 if b.attribution else 0)
        if b.attribution:
            self.flow(f"— {b.attribution}", ast, x=x, width=w)
        self.y += self.pt(12)

    def _page_break(self, b, nxt):
        self.new_page()


def preview_doc(doc: Document, theme: Theme, outdir: Path, dpi: int = 90) -> list[Path]:
    return DocPreview(doc, theme, dpi).render(Path(outdir))
