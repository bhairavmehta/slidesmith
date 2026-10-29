"""ChartSpec -> themed PNG with matplotlib (used for Word documents).

Slides use native PowerPoint charts instead (editable); this mirrors their styling.
"""
from __future__ import annotations

import re

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from matplotlib.ticker import FuncFormatter  # noqa: E402

from .schema import ChartSpec  # noqa: E402
from .textfit import font_path  # noqa: E402
from .theme import Theme, mix  # noqa: E402


def _fmt(number_format: str):
    """Approximate the Excel number formats the schema documents (%, $, decimals,
    thousands separators, ',' scaling and a quoted suffix such as ,,"M")."""
    nf = number_format or "General"
    m = re.match(r'^(?P<body>[^"]*?)(?P<scale>,*)(?:"(?P<suffix>[^"]*)")?$', nf)
    body, scale, suffix = (m.group("body"), m.group("scale"), m.group("suffix") or "") if m else (nf, "", "")
    div = 1000 ** len(scale)

    def f(v, _pos=None):
        if "%" in body:
            dec = body.split(".")[1].count("0") if "." in body else 0
            return f"{v * 100:.{dec}f}%"
        x = v / div
        dec = body.split(".")[1].count("0") if "." in body else (0 if body != "General" else None)
        prefix = "$" if "$" in body else "€" if "€" in body else "£" if "£" in body else ""
        if dec is None:
            s = f"{x:,.2f}".rstrip("0").rstrip(".") if abs(x) < 1000 else f"{x:,.0f}"
        else:
            s = f"{x:,.{dec}f}" if "," in body else f"{x:.{dec}f}"
        return prefix + s + suffix

    return f


def _family(name: str) -> str:
    path = font_path(name)
    font_manager.fontManager.addfont(path)
    return font_manager.FontProperties(fname=path).get_name()


def render_chart_png(spec: ChartSpec, theme: Theme, out: str, width_in: float = 6.5,
                     height_in: float = 3.4, dpi: int = 200, base_pt: float = 9) -> str:
    c = lambda k: "#" + theme.c(k)  # noqa: E731
    pal = theme.chart_palette
    fam = _family(theme.body_font)
    fmt = _fmt(spec.number_format)
    n_cat, n_ser = len(spec.categories), len(spec.series)
    k = base_pt / 9  # scale every hard-coded size below

    plt.rcParams.update({"font.family": fam, "font.size": base_pt, "text.color": c("ink"),
                         "axes.labelcolor": c("muted"), "xtick.color": c("ink"),
                         "ytick.color": c("muted")})
    fig, ax = plt.subplots(figsize=(width_in, height_in), dpi=dpi)
    fig.patch.set_facecolor(c("bg"))
    ax.set_facecolor(c("bg"))

    def colors_for(si: int) -> list[str]:
        base = pal[si % len(pal)]
        if spec.highlight is None:
            return [base] * n_cat
        if n_ser == 1:
            return [c("accent") if i == spec.highlight else c("chart_muted") for i in range(n_cat)]
        return [base if i == spec.highlight else "#" + mix(base, theme.c("bg"), 0.6) for i in range(n_cat)]

    t = spec.type
    if t in ("pie", "doughnut"):
        vals = spec.series[0].values
        cols = [pal[i % len(pal)] for i in range(n_cat)]
        if spec.highlight is not None:
            cols = [col if i == spec.highlight else "#" + mix(col, theme.c("bg"), 0.55)
                    for i, col in enumerate(cols)]
        wedges, _t, autot = ax.pie(
            vals, colors=cols, startangle=90, counterclock=False,
            wedgeprops={"width": 0.42 if t == "doughnut" else 1, "edgecolor": c("bg"), "linewidth": 1.5},
            autopct=(lambda p: f"{p:.0f}%") if spec.data_labels else None, pctdistance=0.79 if t == "doughnut" else 0.68,
            textprops={"color": "white", "fontsize": 9 * k, "fontweight": "bold"})
        ax.legend(wedges, spec.categories, loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False)
        ax.set_aspect("equal")
    else:
        horizontal = t in ("bar", "stacked_bar")
        stacked = t.startswith("stacked")
        idx = list(range(n_cat))
        width = 0.7 / (1 if stacked or t == "line" else n_ser)
        bottoms = [0.0] * n_cat
        for si, se in enumerate(spec.series):
            if t == "line":
                col = pal[si % len(pal)]
                ax.plot(idx, se.values, color=col, linewidth=2.2 * k, marker="o", markersize=5 * k,
                        markeredgecolor=c("bg"), label=se.name)
                if spec.data_labels and n_cat <= 8:
                    for i, v in enumerate(se.values):
                        ax.annotate(fmt(v), (i, v), textcoords="offset points", xytext=(0, 7),
                                    ha="center", fontsize=8 * k, fontweight="bold", color=c("ink"))
                continue
            offs = [i + (0 if stacked else (si - (n_ser - 1) / 2) * width) for i in idx]
            kw = dict(color=colors_for(si), label=se.name)
            if horizontal:
                bars = ax.barh(offs, se.values, height=width, left=bottoms if stacked else None, **kw)
            else:
                bars = ax.bar(offs, se.values, width=width, bottom=bottoms if stacked else None, **kw)
            if spec.data_labels:
                ax.bar_label(bars, labels=[fmt(v) for v in se.values], padding=0 if stacked else 3,
                             label_type="center" if stacked else "edge", fontsize=8 * k, fontweight="bold",
                             color="white" if stacked else c("ink"))
            if stacked:
                bottoms = [b + v for b, v in zip(bottoms, se.values)]
        if horizontal:
            ax.set_yticks(idx, spec.categories)
            ax.invert_yaxis()
            value_axis, cat_axis = ax.xaxis, ax.yaxis
        else:
            ax.set_xticks(idx, spec.categories)
            value_axis, cat_axis = ax.yaxis, ax.xaxis
        value_axis.set_major_formatter(FuncFormatter(fmt))
        for side in ("top", "right"):
            ax.spines[side].set_visible(False)
        ax.spines["left" if not horizontal else "bottom"].set_visible(False)
        ax.spines["bottom" if not horizontal else "left"].set_color(c("line"))
        ax.tick_params(length=0)
        clean = spec.data_labels and t != "line"
        if clean:
            (ax.set_xticks if horizontal else ax.set_yticks)([])
        else:
            ax.grid(axis="x" if horizontal else "y", color=c("line"), linewidth=0.6)
            ax.set_axisbelow(True)
        if spec.y_title and not clean:
            (ax.set_xlabel if horizontal else ax.set_ylabel)(spec.y_title)
        if n_ser > 1:
            ax.legend(frameon=False, loc="lower left", bbox_to_anchor=(0, 1.0), ncol=n_ser,
                      fontsize=8.5 * k)
        ax.margins(x=0.15 if horizontal else 0.03, y=0.12 if not horizontal else 0.03)
    fig.tight_layout()
    fig.savefig(out, facecolor=fig.get_facecolor())
    plt.close(fig)
    return out
