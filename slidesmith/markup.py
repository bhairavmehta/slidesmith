"""Tiny inline markup: **bold** and *italic*. Anything richer belongs in blocks."""
from __future__ import annotations

import re

_TOKEN = re.compile(r"(\*\*.+?\*\*|\*[^*\s][^*]*?\*)")


def runs(text: str) -> list[tuple[str, bool, bool]]:
    """Split text into (segment, bold, italic) runs."""
    out: list[tuple[str, bool, bool]] = []
    for part in _TOKEN.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            out.append((part[2:-2], True, False))
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            out.append((part[1:-1], False, True))
        else:
            out.append((part, False, False))
    return out or [("", False, False)]


def plain(text: str) -> str:
    return "".join(seg for seg, _, _ in runs(text))
