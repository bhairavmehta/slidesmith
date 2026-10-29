# First run on your Windows 11 laptop — do this before using it for real work

Allow ~20 minutes. Every step tells you what "good" looks like. If a step doesn't look
like that, stop and send the output to whoever is supporting you (or back to Claude) —
don't use the tool for real work until every step passes.

## 1. Install (5 min)

1. Unzip to a **short, non-OneDrive** folder, e.g. `C:\Users\<you>\tools\slidesmith`.
   (Not Desktop/Documents if those sync to OneDrive.)
2. Double-click `install.cmd`.
   - Behind a corporate package mirror? Run it from a terminal instead:
     `install.cmd --index-url https://<your-mirror>/simple`
   - Got "Python was not found"? Install Python for your user, then rerun:
     `winget install -e --id Python.Python.3.12 --scope user`
   - Got an "Access denied / blocked by your administrator" error when the venv's
     python.exe runs (AppLocker)? Rerun with `install.cmd --no-venv`.

**Good looks like:** the doctor at the end prints mostly `[ OK ]`, including
`Microsoft PowerPoint 16.0 reachable over COM`, `Microsoft Word 16.0 reachable over COM`,
and `Test build rendered 2 pages with 'office'`. Then `Ready.`

If PowerPoint or Word shows `[WARN]`: open each app once by hand, dismiss any sign-in /
"what's new" / recovery dialogs, close them, and run `slidesmith.cmd doctor` again.

## 2. Prove Office opens everything (5 min)

In a terminal in the slidesmith folder:

```
.venv\Scripts\python.exe tests\stress_test.py --renderer office
.venv\Scripts\python.exe tests\test_provenance.py
```

(With `--no-venv`, use `python` instead of `.venv\Scripts\python.exe`.)

**Good looks like:** `4/4 passed` and `All provenance checks passed.` The stress test
makes PowerPoint and Word open ~40 slides/pages of deliberately hostile content — long
text, emoji, negative charts, 30-row tables. If Office can open those, it can open yours.
"design errors reported" in that output is expected (the test includes overflow on purpose).

## 3. Look at the samples yourself (3 min)

```
slidesmith.cmd build examples\ai-eval-program.deck.json
slidesmith.cmd build examples\release-gating-memo.doc.json
```

Open `examples\build\ai-eval-program\ai-eval-program.pptx` in PowerPoint and
`examples\build\release-gating-memo\release-gating-memo.docx` in Word.

**Good looks like:** no "repair" or "unreadable content" prompt; text fits its boxes;
charts are editable (right-click → Edit Data). Both builds say `score 100/100`.

## 4. Try it in Kiro on something that doesn't matter (5 min)

Restart Kiro. In a scratch folder, put a small `notes.md` with a few real-looking numbers.
In Kiro chat:

> #slidesmith-workflow Make a 5-slide deck from notes.md for my team

**Good looks like:** Kiro writes `*.deck.json` with `"sources": ["notes.md"]`, calls
`build` and `view_pages`, fixes things, and its final message says the renderer was
`office` and either "all figures verified" or lists the ones you need to check.

## Every time you use it for real

- **Check the numbers.** The build checks every figure against the files in `sources` and
  flags anything it can't find. Kiro must list those for you. Still read every number on
  the final slides once — you own what you present.
- **Open the final file yourself** before sending it. It takes one minute.
- Keep the previous version closed while Kiro rebuilds (or it writes `name (2).pptx`).
- If anything fails, the build report (`build\<name>\report.md`) says what and why;
  `slidesmith.cmd doctor` checks the machine.

## Undo everything

`install.cmd --uninstall` removes the private Python environment, the launcher, and the
`slidesmith` entries it added to Kiro (your other Kiro settings are untouched; the
original is also backed up as `mcp.json.bak`).
