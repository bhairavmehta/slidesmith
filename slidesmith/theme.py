"""Design tokens. A theme is a JSON file; brand themes are just more JSON files.

Taste lives here, not in the LLM: the model picks a layout, the theme decides
colour, type scale and spacing. To on-board a brand, copy themes/default.json,
change the tokens, and pass --theme path/to/brand.json (or set "theme" in the spec).
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

THEME_DIR = Path(__file__).parent / "themes"


@dataclass
class Theme:
    name: str
    fonts: dict
    colors: dict
    chart_palette: list[str]
    slide: dict
    doc: dict
    raw: dict = field(repr=False, default_factory=dict)

    # convenience accessors -------------------------------------------------
    def c(self, key: str) -> str:
        return self.colors[key].lstrip("#")

    @property
    def heading_font(self) -> str:
        return self.fonts["heading"]

    @property
    def body_font(self) -> str:
        return self.fonts["body"]


def list_themes() -> list[str]:
    return sorted(p.stem for p in THEME_DIR.glob("*.json"))


def load_theme(name_or_path: str = "default") -> Theme:
    p = Path(name_or_path)
    if not p.suffix:
        p = THEME_DIR / f"{name_or_path}.json"
    if not p.exists():
        raise FileNotFoundError(
            f"Theme '{name_or_path}' not found. Available: {', '.join(list_themes())}"
        )
    data = json.loads(p.read_text(encoding="utf-8-sig"))
    base = {}
    if data.get("extends"):
        base = load_theme(data["extends"]).raw
    merged = _deep_merge(base, data)
    return Theme(
        name=merged["name"],
        fonts=merged["fonts"],
        colors=merged["colors"],
        chart_palette=merged["chart_palette"],
        slide=merged["slide"],
        doc=merged["doc"],
        raw=merged,
    )


def _deep_merge(a: dict, b: dict) -> dict:
    out = dict(a)
    for k, v in b.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


# --------------------------------------------------------------------------- #
# Colour helpers (used by renderers and by the contrast lint rule)
# --------------------------------------------------------------------------- #
def hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _lum(h: str) -> float:
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4

    r, g, b = hex_to_rgb(h)
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(a: str, b: str) -> float:
    la, lb = sorted([_lum(a), _lum(b)], reverse=True)
    return (la + 0.05) / (lb + 0.05)


def mix(a: str, b: str, t: float) -> str:
    """Blend colour a toward b by t (0..1). Returns hex without '#'."""
    ra, ga, ba = hex_to_rgb(a)
    rb, gb, bb = hex_to_rgb(b)
    return "%02X%02X%02X" % (
        round(ra + (rb - ra) * t), round(ga + (gb - ga) * t), round(ba + (bb - ba) * t)
    )
