---
inclusion: manual
---
# Slidesmith workflow — decks and documents

Use this when asked for a slide deck, presentation, memo, report, proposal or any
formatted document. You write a JSON **spec**; the `slidesmith` toolchain owns every
pixel. Never write python-pptx/python-docx code and never edit the output files by hand —
change the spec and rebuild.

Tools (MCP server `slidesmith`; the CLI equivalents work the same):

| Step | MCP tool | CLI |
|---|---|---|
| Read the contract | `get_schema(kind)` | `slidesmith schema deck\|doc` |
| Pick a theme | `list_themes()` | `slidesmith themes` |
| Quick check | `validate_spec(path)` | `slidesmith validate SPEC` |
| Render + lint | `build(path, out_dir)` | `slidesmith build SPEC -o DIR` |
| Look at it | `view_pages(dir, pages, contact_sheet)` | open `DIR/pages/*.png` |

## Non-negotiable: numbers
Every figure that reaches a slide or page must come from the user's material.
- Never invent, estimate, "round up for impact" or fill in a plausible number. If a number
  you need is missing, leave it out and say so, or ask.
- Put every source file you used in the spec's `sources` list (paths relative to the
  spec). The build then checks every number against those files.
- An `unverified_figure` warning is never ignored. Either fix the number, or — if it is
  derived (difference, ratio, sum) — re-check the arithmetic and write the calculation in
  that slide's `notes` (e.g. "0.74 − 0.61 = 0.13 → 13 points").
- Your final message must list every figure still flagged, so the user checks them.

## The loop

### 1. Understand before writing (no tools yet)
- Who is the audience, what decision or action do we want, how long is the slot?
- Gather the material first: the user's files, data, numbers. A polished file with thin
  content is a failure. If key facts are missing, ask once, briefly.
- Deck or doc? Presented live or read alone → deck. Read asynchronously, needs argument
  and detail → doc. If the request is ambiguous, ask.

### 2. Storyline (write it in chat, short)
- Write the **one-sentence answer** first.
- Then the action titles, one per slide (deck) or the H1/H2 outline (doc). Read the titles
  alone, top to bottom: they must tell the whole story. If they don't, fix the story now —
  it is 10× cheaper than after rendering.
- Deck default: 8–14 slides for a 20–30 min slot. Structure: title → the answer/key numbers
  → evidence sections → plan/ask → closing with the decisions requested.

### 3. Write the spec
- Save as `<name>.deck.json` or `<name>.doc.json` next to (or above) the source files (the
  design rules in `slidesmith-design.md` load automatically for these files — follow them).
- Fill `sources` with the files the content comes from.
- Choose layouts by the content's shape, not habit (see the layout guide there).
- Put detail the speaker needs in `notes`, not on the slide.

### 4. Build
- Call `build`. Fix **every `error`** (overflow, schema, wrong % format) by editing the spec —
  usually by cutting words, never by asking for smaller type.
- Resolve every `unverified_figure` warning as described above.
- Fix warnings that weaken the story (topic titles, density, monotony). An `info` is advice.
- `file_locked`: the user has the previous version open; tell them, don't loop.

### 5. Look at it (mandatory — do not skip)
- `view_pages(dir, contact_sheet=true)` for rhythm and consistency, then individual pages
  that look off. Score them with the rubric in `slidesmith-critique.md`.
- Check `renderer` in the build result. `office` means you are looking at exactly what
  PowerPoint/Word show. `preview` means the built-in drawing: trust layout and text fit,
  but treat document page breaks as approximate — mention that in your final message.
- Write down concrete defects → edit the spec → rebuild.

### 6. Stop rule
- At most **3** build → look → fix rounds. Stop earlier when: 0 errors, score ≥ 90, and the
  critique finds no defect rated "fix".
- If something can't be fixed in the spec (needs a new layout or theme change), say so in
  your final message instead of looping.

### 7. Deliver
- Report: the output path, slide/page count, the one-line storyline, which renderer drew
  the review images, and **every figure still flagged as unverified** (or "all figures
  verified against <sources>"). Don't narrate the steps.
- Remind the user to open the file and read it once before sending it on.
