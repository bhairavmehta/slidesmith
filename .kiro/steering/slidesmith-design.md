---
inclusion: fileMatch
fileMatchPattern: ["**/*.deck.json", "**/*.doc.json"]
---
# Slidesmith design rules

The theme owns colour, type and spacing. Your job is **what to say** and **which layout
says it best**. These rules are what make the output look designed rather than generated.

## Titles
- Every content slide title is an **action title**: a full sentence stating the takeaway.
  - ✗ "Q3 results" ✗ "Channel quality"
  - ✓ "Billing drives most failed resolutions" ✓ "Costs fell 71% while coverage tripled"
- ≤ 15 words, one line if possible, two at most. No trailing period.
- The title and the visual must make the **same** claim. If the chart doesn't prove the
  title, change one of them.

## One idea per slide
- If you need "and" to describe the slide's point, it is two slides.
- Word budget per content slide: ≈ 60 words including the title. Bullets ≤ 5, each ≤ 20
  words, sub-bullets ≤ 2 per bullet.
- Bold the key phrase that starts each bullet: `"**Fix billing retrieval** before …"`.
- Anything longer goes in `notes`.

## Pick the layout from the content's shape

| Content | Layout |
|---|---|
| One striking figure | `big_number` (+ `context` explaining why it matters) |
| 2–4 headline metrics | `kpis` (+ `takeaway`) |
| Trend over time | `chart` type `line` (≤ 4 series) |
| Ranking / comparison across categories | `chart` type `bar`, sorted, `highlight` the one that matters |
| Composition (≤ 5 parts) | `chart` `doughnut`/`stacked_bar`; prefer bar if > 5 parts |
| Two options, before/after, pros/cons | `two_column` with `style` `compare` / `before_after` |
| Exact values people will look up | `table` (≤ 7 rows; use `highlight_row`) |
| Plan, phases, roadmap | `timeline` (3–6 steps) |
| Reasons, actions, recommendations | `bullets` (+ `takeaway`) |
| A voice of the customer / leader | `quote` |
| Change of topic in decks > 10 slides | `section` with `number` |

Rhythm: never three of the same content layout in a row; no more than half the deck as
bullets. Show, don't list — if a bullet slide contains numbers, it probably wants to be
`kpis`, `big_number` or a chart.

## Charts
- Percentages as fractions with `number_format` `"0%"` or `"0.0%"` (0.42 → 42%).
- Currency `"$#,##0"`, `"$#,##0.0,,\"M\""` for millions. Decimals `"0.00"`.
- Sort bars by value unless the categories have a natural order (time, funnel stage).
- Use `highlight` to point at the bar/slice the title is about; everything else mutes.
- Put the "so what" in `takeaway` — it renders as a side panel next to the chart.
- Always give `source` when the data comes from somewhere specific.
- Never invent data. If you only have qualitative info, use bullets or two_column.

## KPIs
- `value` pre-formatted and short ("$4.2M", "38%", "1.7×"). `label` ≤ 8 words.
- `delta` with an explicit comparison ("+12 pts vs. Q2"). Set `sentiment` when down is good
  (cost, latency, errors) so it renders green.

## Documents (`*.doc.json`)
- **Bottom line up front**: the first block is a `callout` (kind `key`) or a short paragraph
  with the answer/recommendation. Then optional `kpis`.
- H1 for main sections, H2 inside them; never skip levels. Headings ≤ 10 words and
  informative ("Billing drives most failures", not "Analysis").
- Paragraphs ≤ 120 words, one idea each. Break prose every ~400 words with a heading,
  list, table, chart or callout.
- `callout` kinds: `key` (the main point/recommendation), `info` (context, definitions),
  `warning` (risks, caveats). Max ~1 callout per page.
- Captions on every chart and table: "Figure N. What it shows (source)."
- Use `page_break` only before major sections in long documents.

## Voice
- Plain, specific, active. Numbers over adjectives ("cut costs 71%" not "dramatically
  reduced costs"). No filler ("In today's fast-paced world…", "leverage", "synergies").
- Consistent tense and terminology across the whole file.
