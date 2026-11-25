# Talk notes scaffold: PowerPoint slides + LaTeX handouts

One deck and one notes file produce three things:

| Output | What it is |
|---|---|
| `output/<deck>-handout.pdf` | Participant handout: **one slide per page, your notes below it** |
| `output/<deck>-speaker.pdf` | Speaker copy: the same page plus a **3-inch shaded column** (full page height) holding your **script**. Whatever the script does not fill becomes a **NOTES** area with ruled lines |
| `output/<deck>-with-script.pptx` | A copy of your deck with the script in PowerPoint's own notes pane (presenter view). Your original `.pptx` is never modified |

You only ever edit three things: the **deck**, **`notes/notes.md`** and (rarely) **`config.yml`**.

## Quick start

```bash
pip install -r requirements.txt     # python-pptx, PyYAML, pymupdf
python build.py                     # builds everything into output/
```

Also needed on your machine: **pandoc**, and **XeLaTeX** (TeX Live or MiKTeX).
**LibreOffice** is needed only if you do not export a PDF from PowerPoint (see below).

1. Put your deck in `deck/` (replace the example `FineTuning.pdf`).
2. Run `python build.py`. On the first run it creates `notes/notes.md` with one empty section per slide
   (and copies in any speaker notes already stored in the `.pptx`).
3. Write your notes, run `python build.py` again.

`make`, `make sync`, `make handout`, `make speaker`, `make clean` do the same if you prefer make.

## The deck

* **Best fidelity:** save your `.pptx` in `deck/`, then in PowerPoint do *File > Export > PDF* into the same folder with
  the **same file name** (`Talk.pptx` + `Talk.pdf`). The build uses the PDF for the slide pictures, so they look exactly
  like PowerPoint. If the PDF is older than the `.pptx` the build tells you and falls back to LibreOffice.
* **No LibreOffice?** Export the PDF every time you change slides.
* **PDF only** also works (as with the included example), but then there is no `-with-script.pptx` output.
* Hidden slides are skipped.

## The notes file

```markdown
# 10 | Training loop steps

## Handout
What participants read under the slide. Normal Markdown: **bold**, *italic*, lists, `inline code`.

```python
for batch in loader:
    loss.backward()
```

## Script
What you say. This goes in the shaded column of the speaker copy.
```

* One section per slide: `# <slide number> | <slide title>`. The number is what counts; the title is a label.
* Under it, `## Handout` and `## Script`. Either can be empty.
* Use Markdown, but not `#` headings inside a section (use `**bold**` lines). See `notes/example-notes.md`.
* Code fences: `python`, `bash`, `json`, `yaml`, `sql`, `r`, `c`, `cpp`, `java`, `tex`, `html`, `xml`, `matlab` are highlighted.
  Add more in `latex/codeblocks.lua`.

### Keeping slides and notes aligned (the "git" part)

After you **add, delete or reorder slides**, run:

```bash
python build.py sync
```

It matches each notes section to its slide by **title**, renumbers sections that moved, adds empty sections for new
slides (seeded with that slide's existing PowerPoint notes), and keeps orphaned sections at the end. It only edits
headings and adds sections; your text is never rewritten. Review with `git diff notes/notes.md`.
Duplicate or empty slide titles cannot be matched, so those sections stay where they are and you get a warning.

A normal `python build.py` also warns you when the notes and the deck have drifted apart.

### Using git

Commit: `notes/`, `config.yml`, `latex/`, `build.py`, `deck/` (the `.pptx`, and ideally its PDF export).
Ignored (see `.gitignore`): `build/` and `output/`, which are rebuilt on demand.
`notes.md` is plain text, so `git diff`, `git blame` and branches work on your script. A `.pptx` is a binary,
so git stores each version whole; use Git LFS for it if it gets large.

## The speaker column

* Width is `script_column_in` (3 inches by default) and runs the full height of the page.
* The script is set in its own font (default DejaVu Serif) and **shrinks automatically** (11 pt down to 8.5 pt)
  so that at least `min_notes_mm` of ruled Notes space stays free.
* If a script is still too long at 8.5 pt, it is set at 9.5 pt and **continues in the column on extra pages**,
  and the build tells you which slides did. Nothing is cut off.
* Short or empty script: the rest of the column is ruled lines.
* `speaker_orientation: landscape` in `config.yml` gives a much larger slide with the same column.

## Customising

`config.yml`: title, subtitle, speaker, event, date, paper (`letter`/`a4`), cover page, column width,
fonts, and whether the speaker copy repeats the handout notes.
`latex/preamble.tex`: colours (your four palette colours are defined there), the code-block style
(`\lstdefinestyle{talk}`), the cover page, and all page layout.

**Fonts** default to Crimson Text (body), Fira Sans (headings), DejaVu Serif (script) and DejaVu Sans Mono (code).
All are free (Google Fonts / dejavu-fonts.github.io). A font that is not installed falls back to Latin Modern and the
build prints a warning. Change names under `fonts:` in `config.yml` to use others.

## Troubleshooting

| Message | Fix |
|---|---|
| `xelatex not found` | Install TeX Live or MiKTeX and make sure `xelatex` is on your PATH |
| `pandoc not found` | Install pandoc (pandoc.org/installing) |
| `The .pptx has N slides but the PDF has M pages` | Re-export the PDF from the current `.pptx` |
| `font '...' is not installed` | Install it, or pick another under `fonts:` |
| LaTeX error | The last lines of the log are printed; the full log is `build/handout.log` / `build/speaker.log` |
| Slide pictures look stale | `python build.py clean`, then build again |
