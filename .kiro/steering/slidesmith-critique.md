---
inclusion: fileMatch
fileMatchPattern: ["**/*.deck.json", "**/*.doc.json"]
---
# Slidesmith visual critique rubric

Use after every `build`, on the rendered images (`view_pages`). Look at the contact
sheet first, then open individual pages that look off. Judge what you **see**, not what
the spec says should be there.

## Score each page 1–5 on

| Criterion | 5 looks like | Typical fix in the spec |
|---|---|---|
| **Message** | Title states one clear point; the visual proves it at a glance | Rewrite title; change layout; add `highlight` |
| **Hierarchy** | Eye goes title → hero element → support, in that order | Bold key phrases; move detail to notes; use `takeaway` |
| **Density** | Breathing room; nothing crammed; ≤ ~60 words | Cut words; split slide; move to notes/appendix |
| **Fit & alignment** | No clipped, overlapping or wrapped-badly text; columns line up | Shorten the offending text (labels, titles, cells) |
| **Data integrity** | Labels readable, formats right (%, $), sorted, sources present | Fix `number_format`, sort, set `source` |
| **Consistency** | Same terms, tense, number formats, and capitalisation across pages | Normalise wording across the spec |
| **Balance** | Page doesn't feel empty or lopsided | Combine thin slides; add `context`/`takeaway`; switch layout |

Also check at deck level (contact sheet): visual rhythm (layouts vary), sections are
evenly sized, the titles read as a story when scanned in order, and the first and last
slides carry the answer and the ask.

## Output format (write this in chat before editing the spec)

```
Round N — score X/100 (build), visual verdict: ship | fix
p3  fix   Fit: bar labels "Account access management" wrap to 3 lines → shorten to "Account access"
p6  fix   Message: number "3 of 3" lacks the why → add context sentence
p9  nit   Balance: section slide fine
```

- `fix` = a reader would notice; must be addressed this round.
- `nit` = optional polish; address only if cheap.
- Verdict `ship` only when no `fix` items remain.

## Common defects to look for specifically
- Chart category labels truncated or wrapping onto 3+ lines.
- Data labels colliding (dense line charts) → drop `data_labels` or reduce series.
- A KPI value much longer than its siblings ("$1,234,567" vs "38%") → shorten format.
- Takeaway that repeats the title instead of adding the "so what".
- Table where only one column matters → turn it into a bar chart.
- Doc page that ends with a lone heading or starts with a stranded caption.
- Near-empty final page in a doc (lint reports `orphan_page`).
