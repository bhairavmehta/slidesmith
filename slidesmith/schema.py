"""Spec schema for decks and documents.

The LLM (Kiro) writes JSON that conforms to these models; the renderers never
see free-form text. Keeping the contract small and opinionated is what makes the
output consistent: the model chooses *which* layout and *what* words, the code
owns every pixel decision.
"""
from __future__ import annotations

from typing import Annotated, Literal, Optional, Union

from pydantic import BaseModel, ConfigDict, Field, model_validator


class _M(BaseModel):
    model_config = ConfigDict(extra="forbid")


# --------------------------------------------------------------------------- #
# Shared building blocks
# --------------------------------------------------------------------------- #
class Series(_M):
    name: str
    values: list[float]


class ChartSpec(_M):
    type: Literal["column", "bar", "line", "pie", "doughnut", "stacked_column", "stacked_bar"]
    categories: list[str]
    series: list[Series] = Field(min_length=1)
    number_format: str = Field(
        default="General",
        description='Excel-style format for axis/labels, e.g. "0%", "$#,##0", "0.0".',
    )
    data_labels: bool = True
    highlight: Optional[int] = Field(
        default=None,
        description="Index of the category to emphasise; all other points are muted. "
        "Use it when the slide's point is about one bar/slice.",
    )
    y_title: Optional[str] = None

    @model_validator(mode="after")
    def _lengths(self):
        n = len(self.categories)
        for s in self.series:
            if len(s.values) != n:
                raise ValueError(
                    f"series '{s.name}' has {len(s.values)} values but there are {n} categories"
                )
        if self.type in ("pie", "doughnut") and len(self.series) != 1:
            raise ValueError("pie/doughnut charts take exactly one series")
        if self.highlight is not None and not (0 <= self.highlight < n):
            raise ValueError("highlight index out of range")
        return self


class KPI(_M):
    value: str = Field(description='Pre-formatted, e.g. "$4.2M", "38%", "1.7x".')
    label: str
    delta: Optional[str] = Field(default=None, description='e.g. "+12% YoY"; prefix - for a decline.')
    sentiment: Optional[Literal["positive", "negative", "neutral"]] = Field(
        default=None,
        description="Colour of the delta. Defaults to the sign; set it when down is good (e.g. cost -71%).",
    )


class Bullet(_M):
    text: str
    sub: list[str] = Field(default_factory=list, max_length=4)


BulletItem = Union[str, Bullet]


class Column(_M):
    heading: str
    bullets: list[BulletItem] = Field(default_factory=list)
    body: Optional[str] = None


class Step(_M):
    label: str = Field(description='Short marker, e.g. "Q1", "Week 2", "Phase 1".')
    title: str
    detail: Optional[str] = None


# --------------------------------------------------------------------------- #
# Slides — one model per layout archetype
# --------------------------------------------------------------------------- #
class _Slide(_M):
    notes: Optional[str] = Field(default=None, description="Speaker notes.")


class TitleSlide(_Slide):
    layout: Literal["title"]
    title: str
    subtitle: Optional[str] = None
    kicker: Optional[str] = Field(default=None, description="Small overline text, e.g. a team or date.")


class SectionSlide(_Slide):
    layout: Literal["section"]
    title: str
    number: Optional[str] = None
    subtitle: Optional[str] = None


class BulletsSlide(_Slide):
    layout: Literal["bullets"]
    title: str
    bullets: list[BulletItem] = Field(min_length=1)
    takeaway: Optional[str] = Field(default=None, description="One-line 'so what' shown in a callout bar.")


class TwoColumnSlide(_Slide):
    layout: Literal["two_column"]
    title: str
    left: Column
    right: Column
    style: Literal["plain", "compare", "before_after"] = "plain"


class BigNumberSlide(_Slide):
    layout: Literal["big_number"]
    title: str
    number: str
    label: str
    context: Optional[str] = None


class KPISlide(_Slide):
    layout: Literal["kpis"]
    title: str
    items: list[KPI] = Field(min_length=2, max_length=4)
    takeaway: Optional[str] = None


class ChartSlide(_Slide):
    layout: Literal["chart"]
    title: str
    chart: ChartSpec
    takeaway: Optional[str] = None
    source: Optional[str] = None


class TableSlide(_Slide):
    layout: Literal["table"]
    title: str
    columns: list[str] = Field(min_length=2, max_length=7)
    rows: list[list[str]] = Field(min_length=1, max_length=10)
    highlight_row: Optional[int] = None
    source: Optional[str] = None

    @model_validator(mode="after")
    def _shape(self):
        for i, r in enumerate(self.rows):
            if len(r) != len(self.columns):
                raise ValueError(f"row {i} has {len(r)} cells, expected {len(self.columns)}")
        return self


class TimelineSlide(_Slide):
    layout: Literal["timeline"]
    title: str
    steps: list[Step] = Field(min_length=3, max_length=6)


class QuoteSlide(_Slide):
    layout: Literal["quote"]
    quote: str
    attribution: Optional[str] = None


class ClosingSlide(_Slide):
    layout: Literal["closing"]
    title: str
    bullets: list[str] = Field(default_factory=list, max_length=5)
    contact: Optional[str] = None


Slide = Annotated[
    Union[
        TitleSlide, SectionSlide, BulletsSlide, TwoColumnSlide, BigNumberSlide, KPISlide,
        ChartSlide, TableSlide, TimelineSlide, QuoteSlide, ClosingSlide,
    ],
    Field(discriminator="layout"),
]


class Deck(_M):
    kind: Literal["deck"] = "deck"
    title: str
    sources: list[str] = Field(
        default_factory=list,
        description="Files the figures come from (paths relative to the spec). Every number in "
                    "the output is checked against them; unmatched figures are flagged.",
    )
    theme: str = "default"
    author: Optional[str] = None
    date: Optional[str] = None
    footer: Optional[str] = Field(default=None, description="Shown on content slides, e.g. 'Confidential'.")
    slides: list[Slide] = Field(min_length=1)


# --------------------------------------------------------------------------- #
# Documents — a flat list of blocks
# --------------------------------------------------------------------------- #
class HeadingBlock(_M):
    type: Literal["heading"]
    level: int = Field(ge=1, le=3)
    text: str


class ParagraphBlock(_M):
    type: Literal["paragraph"]
    text: str = Field(description="Supports **bold** and *italic* inline markup.")


class BulletsBlock(_M):
    type: Literal["bullets"]
    items: list[BulletItem] = Field(min_length=1)
    ordered: bool = False


class CalloutBlock(_M):
    type: Literal["callout"]
    kind: Literal["info", "key", "warning"] = "key"
    title: Optional[str] = None
    text: str


class TableBlock(_M):
    type: Literal["table"]
    columns: list[str] = Field(min_length=2)
    rows: list[list[str]] = Field(min_length=1)
    caption: Optional[str] = None

    @model_validator(mode="after")
    def _shape(self):
        for i, r in enumerate(self.rows):
            if len(r) != len(self.columns):
                raise ValueError(f"row {i} has {len(r)} cells, expected {len(self.columns)}")
        return self


class ChartBlock(_M):
    type: Literal["chart"]
    chart: ChartSpec
    caption: Optional[str] = None


class KPIBlock(_M):
    type: Literal["kpis"]
    items: list[KPI] = Field(min_length=2, max_length=4)


class QuoteBlock(_M):
    type: Literal["quote"]
    text: str
    attribution: Optional[str] = None


class PageBreakBlock(_M):
    type: Literal["page_break"]


Block = Annotated[
    Union[
        HeadingBlock, ParagraphBlock, BulletsBlock, CalloutBlock, TableBlock,
        ChartBlock, KPIBlock, QuoteBlock, PageBreakBlock,
    ],
    Field(discriminator="type"),
]


class Document(_M):
    kind: Literal["doc"] = "doc"
    title: str
    sources: list[str] = Field(
        default_factory=list,
        description="Files the figures come from (paths relative to the spec). Every number in "
                    "the output is checked against them; unmatched figures are flagged.",
    )
    subtitle: Optional[str] = None
    theme: str = "default"
    author: Optional[str] = None
    date: Optional[str] = None
    header: Optional[str] = Field(default=None, description="Running header text, e.g. 'Internal — Draft'.")
    blocks: list[Block] = Field(min_length=1)


def load_spec(data: dict) -> Deck | Document:
    kind = data.get("kind")
    if kind is None:
        kind = "deck" if "slides" in data else "doc"
    return Deck.model_validate(data) if kind == "deck" else Document.model_validate(data)


def json_schema(kind: str) -> dict:
    return (Deck if kind == "deck" else Document).model_json_schema()


def bullet_text(b: BulletItem) -> str:
    return b if isinstance(b, str) else b.text


def bullet_subs(b: BulletItem) -> list[str]:
    return [] if isinstance(b, str) else b.sub
