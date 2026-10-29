"""Measure text with real font metrics so the renderer can pick a size that fits.

PowerPoint's own autofit only kicks in when the file is opened in PowerPoint (and
never in LibreOffice or previews), so we size text ourselves at build time and
report anything that cannot fit even at the minimum size. That report is what
lets the agent fix overflow by *editing words*, not by shrinking type forever.
"""
from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache

from matplotlib import font_manager
from PIL import ImageFont

# Metric-compatible open substitutes, used only for measurement.
ALIASES = {
    "Calibri": ["Calibri", "Carlito"],
    "Cambria": ["Cambria", "Caladea"],
    "Arial": ["Arial", "Liberation Sans", "Arimo"],
    "Helvetica": ["Helvetica", "Liberation Sans", "Arimo"],
    "Times New Roman": ["Times New Roman", "Liberation Serif", "Tinos"],
    "Aptos": ["Aptos", "Carlito"],
}
_SCALE = 10  # measure at 10x for sub-point precision
LINE_HEIGHT = 1.2  # typical ascent+descent for sans fonts, as a multiple of size
# Text is fitted into this fraction of the real box width. PowerPoint, Word, LibreOffice and
# font substitutes never wrap exactly like our measurement; the slack keeps a line that
# "just fits" here from wrapping onto an extra line there.
WIDTH_SAFETY = 0.95


@lru_cache(maxsize=64)
def font_path(family: str, bold: bool = False, italic: bool = False) -> str:
    weight = "bold" if bold else "normal"
    style = "italic" if italic else "normal"
    for fam in ALIASES.get(family, [family]):
        try:
            return font_manager.findfont(
                font_manager.FontProperties(family=fam, weight=weight, style=style),
                fallback_to_default=False,
            )
        except ValueError:
            continue
    # DejaVu is wider than most UI fonts, so falling back errs on the safe side.
    return font_manager.findfont(font_manager.FontProperties(family="DejaVu Sans", weight=weight,
                                                             style=style))


@lru_cache(maxsize=512)
def _font(family: str, bold: bool, pt: float) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(font_path(family, bold), size=max(1, int(round(pt * _SCALE))))


def _is_wide(ch: str) -> bool:
    """Characters body fonts usually lack (CJK, emoji): Office draws them with a fallback
    font, so measuring them with the body font underestimates their width."""
    o = ord(ch)
    return (0x1100 <= o <= 0x115F or 0x2E80 <= o <= 0xA4CF or 0xAC00 <= o <= 0xD7A3
            or 0xF900 <= o <= 0xFAFF or 0xFE30 <= o <= 0xFE4F or 0xFF00 <= o <= 0xFF60
            or 0xFFE0 <= o <= 0xFFE6 or 0x2600 <= o <= 0x27BF or 0x1F000 <= o <= 0x1FAFF
            or 0x20000 <= o <= 0x3FFFD)


def text_width_in(text: str, family: str, pt: float, bold: bool = False) -> float:
    font = _font(family, bold, pt)
    wide = sum(1 for ch in text if _is_wide(ch))
    if not wide:
        return font.getlength(text) / _SCALE / 72.0
    narrow = "".join(ch for ch in text if not _is_wide(ch))
    # Fallback fonts draw these: CJK ~1 em, emoji up to ~1.3 em. Be conservative.
    return (font.getlength(narrow) / _SCALE + wide * pt * 1.25) / 72.0


def wrap(text: str, family: str, pt: float, width_in: float, bold: bool = False) -> list[str]:
    words = text.split()
    if not words:
        return [""]
    lines, cur = [], ""
    for w in words:
        trial = f"{cur} {w}" if cur else w
        if text_width_in(trial, family, pt, bold) <= width_in:
            cur = trial
            continue
        if cur:
            lines.append(cur)
        # break words that are wider than the line on their own
        while text_width_in(w, family, pt, bold) > width_in and len(w) > 1:
            cut = len(w)
            while cut > 1 and text_width_in(w[:cut], family, pt, bold) > width_in:
                cut -= 1
            lines.append(w[:cut])
            w = w[cut:]
        cur = w
    lines.append(cur)
    return lines


@dataclass
class Para:
    text: str
    scale: float = 1.0          # relative to the base size (e.g. 0.8 for sub-bullets)
    indent_in: float = 0.0      # left indent eats into the available width
    bold: bool = False
    space_before: float = 0.0   # in multiples of base pt


@dataclass
class Fit:
    pt: float
    height_in: float
    fits: bool
    lines: int


def measure(paras: list[Para], family: str, pt: float, width_in: float,
            line_spacing: float = 1.1) -> tuple[float, int]:
    h, n = 0.0, 0
    for p in paras:
        size = pt * p.scale
        lines = wrap(p.text, family, size, width_in - p.indent_in, p.bold)
        n += len(lines)
        h += len(lines) * size * LINE_HEIGHT * line_spacing / 72.0
        h += p.space_before * pt / 72.0
    return h, n


def fit(paras: list[Para], family: str, width_in: float, height_in: float,
        max_pt: float, min_pt: float, line_spacing: float = 1.1, max_lines: int | None = None) -> Fit:
    """Largest size in [min_pt, max_pt] (1pt steps) at which paras fit the box."""
    pt = float(max_pt)
    width_in = width_in * WIDTH_SAFETY
    while True:
        h, n = measure(paras, family, pt, width_in, line_spacing)
        ok = h <= height_in and (max_lines is None or n <= max_lines)
        if ok or pt - 1 < min_pt:
            return Fit(pt=pt, height_in=h, fits=ok, lines=n)
        pt -= 1
