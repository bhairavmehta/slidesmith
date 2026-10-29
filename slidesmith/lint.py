"""Deterministic design QA — the cheap, reliable half of the critique loop.

The vision critique (done by the agent, see kiro/steering) catches what only eyes
can see. These rules catch everything that can be checked mechanically, so the
agent spends its review on judgement, not on counting words.

Each issue: {where, index, severity: error|warning|info, rule, message, hint}.
"""
from __future__ import annotations

from collections import Counter
from pathlib import Path

from PIL import Image

from . import markup
from .schema import (
    BulletsSlide, ChartSlide, ChartSpec, Deck, Document, TableSlide, bullet_subs, bullet_text,
)
from .theme import Theme, contrast

CONTENT_LAYOUTS = {"bullets", "two_column", "big_number", "kpis", "chart", "table", "timeline"}
SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}


def _words(s: str | None) -> int:
    return len(markup.plain(s or "").split())


class _Collector:
    def __init__(self, kind: str):
        self.kind = kind
        self.items: list[dict] = []

    def add(self, idx: int, severity: str, rule: str, message: str, hint: str = ""):
        where = f"{'slide' if self.kind == 'deck' else 'block'} {idx}" if idx else "global"
        self.items.append({"where": where, "index": idx, "severity": severity, "rule": rule,
                           "message": message, "hint": hint})


# --------------------------------------------------------------------------- #
def lint_theme(theme: Theme, c: _Collector):
    col = theme.colors
    checks = [
        ("ink", "bg", 7.0, "error", "body text"),
        ("muted", "bg", 4.5, "warning", "secondary text"),
        ("on_primary", "primary", 4.5, "error", "text on dark slides"),
        ("accent", "bg", 3.0, "warning", "accent text and big numbers"),
        ("ink", "surface", 7.0, "warning", "text on cards"),
    ]
    for fg, bg, need, sev, what in checks:
        r = contrast(col[fg], col[bg])
        if r < need:
            c.add(0, sev, "contrast",
                  f"theme '{theme.name}': {fg} on {bg} contrast {r:.1f}:1 (< {need}:1) for {what}.",
                  "Adjust the theme colours, not the spec.")


def lint_chart(ch: ChartSpec, idx: int, c: _Collector):
    n = len(ch.categories)
    if ch.type in ("pie", "doughnut") and n > 6:
        c.add(idx, "warning", "chart_type", f"{ch.type} with {n} slices is hard to read.",
              "Use a sorted bar chart, or group small slices into 'Other'.")
    if ch.type in ("column", "bar", "stacked_column", "stacked_bar") and n > 12:
        c.add(idx, "warning", "chart_density", f"{n} categories is too many to label.",
              "Show the top 8–10 and group the rest, or use a table.")
    if ch.type == "column" and max(len(x) for x in ch.categories) > 14 and n > 4:
        c.add(idx, "info", "chart_type", "Long category names on a column chart will wrap or tilt.",
              "Use type 'bar' (horizontal) for long labels.")
    if ch.type == "line" and n < 3:
        c.add(idx, "info", "chart_type", "Line chart with fewer than 3 points.",
              "Use a column chart or a KPI instead.")
    if "%" in ch.number_format:
        mx = max((abs(v) for s in ch.series for v in s.values), default=0)
        if mx > 1.5:
            c.add(idx, "error", "number_format",
                  f"number_format is a percent but values go up to {mx:g} (would show {mx*100:g}%).",
                  "Give percentages as fractions (0.42 for 42%) or drop the % format.")
    if ch.type in ("bar", "column") and len(ch.series) == 1 and n >= 4 and ch.highlight is None:
        vals = ch.series[0].values
        sorted_ok = vals == sorted(vals) or vals == sorted(vals, reverse=True)
        if not sorted_ok and not _looks_ordinal(ch.categories):
            c.add(idx, "info", "chart_sort", "Unordered categories in a single-series bar chart.",
                  "Sort bars by value unless the categories have a natural order (time, stages).")


def _looks_ordinal(cats: list[str]) -> bool:
    months = {"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"}
    low = [x.lower()[:3] for x in cats]
    return all(x in months for x in low) or any(ch.isdigit() for x in cats for ch in x)


# --------------------------------------------------------------------------- #
def lint_deck(deck: Deck, theme: Theme) -> list[dict]:
    c = _Collector("deck")
    lint_theme(theme, c)
    slides = deck.slides
    if slides[0].layout != "title":
        c.add(1, "info", "structure", "Deck does not open with a title slide.")
    if len(slides) > 25:
        c.add(0, "info", "length", f"{len(slides)} slides; most audiences lose the thread after ~20.",
              "Move detail to an appendix section.")

    titles = Counter()
    for i, s in enumerate(slides, start=1):
        title = getattr(s, "title", None)
        if title:
            titles[title.strip().lower()] += 1
        if s.layout in CONTENT_LAYOUTS and title:
            w = _words(title)
            if w < 4:
                c.add(i, "warning", "action_title",
                      f"Title '{title}' names a topic instead of stating the point.",
                      "Rewrite as a full-sentence takeaway, e.g. 'Billing drives most failed resolutions'.")
            elif w > 18 or len(title) > 110:
                c.add(i, "warning", "title_length", f"Title is {w} words; it will wrap to 2+ lines.",
                      "Keep action titles to ≤ 15 words.")

        if isinstance(s, BulletsSlide):
            n = len(s.bullets)
            if n > 6:
                c.add(i, "warning", "density", f"{n} bullets on one slide.",
                      "Keep to ≤ 5, split the slide, or move detail to notes.")
            total = _words(s.title) + _words(s.takeaway)
            for b in s.bullets:
                bw = _words(bullet_text(b))
                total += bw + sum(_words(x) for x in bullet_subs(b))
                if bw > 25:
                    c.add(i, "warning", "bullet_length",
                          f"Bullet of {bw} words: '{markup.plain(bullet_text(b))[:60]}…'",
                          "Bullets should be ≤ 20 words; lead with the bolded point.")
            if total > 80:
                c.add(i, "warning", "density", f"~{total} words on the slide.",
                      "Aim for ≤ 60 words; slides support the speaker, they are not the script.")
            if all(isinstance(b, str) for b in s.bullets) and n >= 3 and not any("**" in b for b in s.bullets):
                c.add(i, "info", "hierarchy", "Bullets have no emphasis; every line reads with equal weight.",
                      "Bold the key phrase of each bullet with **…**.")
        if isinstance(s, ChartSlide):
            lint_chart(s.chart, i, c)
        if isinstance(s, TableSlide) and len(s.rows) > 8:
            c.add(i, "warning", "density", f"Table with {len(s.rows)} rows.",
                  "Show the rows that support the title; put the full table in an appendix.")
        if s.layout == "big_number" and len(s.number) > 9:
            c.add(i, "warning", "big_number", f"'{s.number}' is long for a hero number.",
                  "Use a short figure ('$4.2M', '3×') and put precision in the label.")
        if s.layout == "section" and (i == len(slides) or slides[i].layout in ("section", "closing")):
            c.add(i, "warning", "structure", "Section divider with no content after it.")

    for t, n in titles.items():
        if n > 1:
            c.add(0, "warning", "duplicate_title", f"Title used {n} times: '{t[:60]}'.",
                  "Every slide should make a distinct point.")

    # visual rhythm
    run, prev = 1, None
    for i, s in enumerate(slides, start=1):
        run = run + 1 if s.layout == prev and s.layout in CONTENT_LAYOUTS else 1
        if run == 3:
            c.add(i, "warning", "monotony", f"3 '{s.layout}' slides in a row.",
                  "Vary layouts: turn one into a chart, kpis, two_column or big_number.")
        prev = s.layout
    content = [s for s in slides if s.layout in CONTENT_LAYOUTS]
    if len(content) >= 5 and sum(s.layout == "bullets" for s in content) / len(content) > 0.5:
        c.add(0, "warning", "variety", "More than half the content slides are bullet lists.",
              "Show, don't list: numbers → kpis/big_number, comparisons → two_column, trends → chart.")
    return c.items


def lint_doc(doc: Document, theme: Theme) -> list[dict]:
    c = _Collector("doc")
    lint_theme(theme, c)
    blocks = doc.blocks
    if blocks[0].type not in ("callout", "paragraph", "kpis"):
        c.add(1, "info", "bluf", "Document opens with a heading, not the answer.",
              "Lead with a 2–3 sentence summary or a 'key' callout (bottom line up front).")
    level, prose_run = 0, 0
    for i, b in enumerate(blocks, start=1):
        if b.type == "heading":
            if b.level > level + 1:
                c.add(i, "warning", "heading_levels", f"Heading jumps from H{level or '-'} to H{b.level}.",
                      "Don't skip heading levels.")
            nxt = blocks[i] if i < len(blocks) else None
            if nxt is None or (nxt.type == "heading" and nxt.level <= b.level):
                c.add(i, "warning", "empty_section", f"Heading '{b.text[:50]}' has no content.")
            if _words(b.text) > 12:
                c.add(i, "info", "heading_length", "Long heading; headings scan best at ≤ 10 words.")
            level = b.level
            prose_run = 0
        elif b.type == "paragraph":
            w = _words(b.text)
            if w > 150:
                c.add(i, "warning", "paragraph_length", f"Paragraph of {w} words.",
                      "Split it; one idea per paragraph, ≤ 120 words.")
            prose_run += w
            if prose_run > 450:
                c.add(i, "info", "wall_of_text", f"~{prose_run} words of unbroken prose.",
                      "Break it up with a subheading, bullets, a table or a callout.")
                prose_run = 0
        else:
            prose_run = 0
            if b.type == "chart":
                lint_chart(b.chart, i, c)
                if not b.caption:
                    c.add(i, "info", "caption", "Chart without a caption.",
                          "Add 'Figure N. what it shows (source)'.")
            if b.type == "table" and len(b.rows) > 15:
                c.add(i, "warning", "density", f"Table with {len(b.rows)} rows.",
                      "Summarise in the body; move the full table to an appendix.")
            if b.type == "bullets" and len(b.items) > 8:
                c.add(i, "info", "density", f"List of {len(b.items)} items.",
                      "Group into themes or use a table.")
    return c.items


# --------------------------------------------------------------------------- #
def lint_pages(kind: str, pages: list[Path], bg_hex: str) -> list[dict]:
    """Raster checks on the rendered pages."""
    c = _Collector(kind)
    if kind == "doc" and len(pages) > 1:
        im = Image.open(pages[-1]).convert("L").resize((200, 260))
        ink = sum(1 for p in im.getdata() if p < 235) / (200 * 260)
        if ink < 0.02:
            c.add(0, "info", "orphan_page",
                  f"Last page (p{len(pages)}) is nearly empty.",
                  "Trim a little text or remove a trailing spacer so it ends a page earlier.")
    return c.items


def summarize(issues: list[dict]) -> dict:
    issues.sort(key=lambda x: (SEVERITY_ORDER[x["severity"]], x["index"]))
    # notes about the build environment (which renderer ran) are not design issues
    counts = Counter(x["severity"] for x in issues
                     if x["rule"] not in ("renderer", "raster", "file_locked"))
    score = max(0, 100 - 15 * counts["error"] - 5 * counts["warning"] - 1 * counts["info"])
    return {"score": score, "errors": counts["error"], "warnings": counts["warning"],
            "info": counts["info"], "issues": issues}
