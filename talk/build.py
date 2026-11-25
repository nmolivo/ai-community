#!/usr/bin/env python3
"""
Talk builder: one deck + one notes file  ->  participant handout, speaker copy, and
(optionally) a copy of the .pptx with the script placed in PowerPoint's notes pane.

    python build.py             build everything
    python build.py sync        reconcile notes/notes.md with the deck (new / moved slides)
    python build.py slides      only render the slide images
    python build.py clean       delete build/ and output/

Inputs (the only files you edit):  deck/<your>.pptx (or .pdf),  notes/notes.md,  config.yml
Outputs:                           output/*.pdf, output/*-with-script.pptx
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    sys.exit("Missing PyYAML.  Run:  pip install -r requirements.txt")

ROOT = Path(__file__).resolve().parent
BUILD = ROOT / "build"
OUT = ROOT / "output"

DEFAULTS = {
    "title": "My Talk",
    "subtitle": "",
    "speaker": "",
    "event": "",
    "date": "",
    "deck_dir": "deck",
    "deck_file": "",            # optional: force a specific file inside deck_dir
    "notes_file": "notes/notes.md",
    "paper": "letter",          # letter | a4
    "speaker_orientation": "portrait",  # portrait | landscape (landscape gives a bigger slide)
    "dpi": 150,                 # slide image resolution
    "cover_page": True,
    "script_column_in": 3.0,    # width of the speaker-script column, inches
    "min_notes_mm": 24,         # keep at least this much ruled "Notes" space under the script
    "speaker_shows_handout_notes": True,
    "handout_blank_lines": False,   # fill the rest of each handout page with ruled lines
    "export_pptx_notes": True,      # write build of the .pptx with the script in its notes pane
    "fonts": {
        "body": "Crimson Text",
        "heading": "Fira Sans",
        "script": "DejaVu Serif",   # the speaker-script font (deliberately different)
        "code": "DejaVu Sans Mono",
    },
}


# ----------------------------------------------------------------------------- helpers
def say(msg: str = "") -> None:
    print(msg, flush=True)


def warn(msg: str) -> None:
    print(f"  ! {msg}", flush=True)


def die(msg: str) -> None:
    sys.exit(f"\nERROR: {msg}\n")


def load_config() -> dict:
    cfg = json.loads(json.dumps(DEFAULTS))
    path = ROOT / "config.yml"
    if path.exists():
        user = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        for k, v in user.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update(v)
            else:
                cfg[k] = v
    return cfg


def which(*names: str) -> str | None:
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None


def sha(data: bytes) -> str:
    return hashlib.sha1(data).hexdigest()


def tex_escape(s: str) -> str:
    table = {"\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
             "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}", "^": r"\textasciicircum{}"}
    return "".join(table.get(c, c) for c in s)


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", s.lower())


# ----------------------------------------------------------------------------- deck
class Deck:
    """Resolved deck: where the slide images come from and what the slides are called."""

    def __init__(self) -> None:
        self.pptx: Path | None = None
        self.pdf: Path | None = None      # PDF used for the slide images
        self.titles: list[str] = []
        self.pptx_notes: list[str] = []   # speaker notes already in the .pptx (to seed notes.md)
        self.slide_map: list[int] = []    # deck slide index (0-based) for each PDF page


def find_deck(cfg: dict) -> tuple[Path | None, Path | None]:
    d = ROOT / cfg["deck_dir"]
    if cfg["deck_file"]:
        f = d / cfg["deck_file"]
        if not f.exists():
            die(f"deck_file '{cfg['deck_file']}' not found in {d}/")
        return (f, None) if f.suffix.lower() == ".pptx" else (None, f)
    pptx = sorted((p for p in d.glob("*.pptx") if not p.name.startswith("~$")),
                  key=lambda p: p.stat().st_mtime, reverse=True)
    pdfs = sorted(d.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    if len(pptx) > 1:
        warn(f"several .pptx in {d}/; using the newest: {pptx[0].name} (set deck_file to choose)")
    if pptx:
        return pptx[0], None
    if pdfs:
        return None, pdfs[0]
    die(f"No deck found. Put your .pptx (or an exported .pdf) in {d}/")


def render_pptx_to_pdf(pptx: Path) -> Path:
    """LibreOffice render (fallback when there is no up-to-date PDF export from PowerPoint)."""
    soffice = which("soffice", "libreoffice")
    if not soffice:
        die("Need either an up-to-date PDF export next to the .pptx (same file name), "
            "or LibreOffice installed to render the .pptx.")
    cache = BUILD / "render"
    cache.mkdir(parents=True, exist_ok=True)
    key = cache / f"{pptx.stem}.key"
    target = cache / f"{pptx.stem}.pdf"
    h = sha(pptx.read_bytes())
    if target.exists() and key.exists() and key.read_text() == h:
        return target
    say(f"  rendering {pptx.name} with LibreOffice (slides may differ slightly from PowerPoint) ...")
    r = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(cache), str(pptx)],
                       capture_output=True, text=True, timeout=600)
    if r.returncode != 0 or not target.exists():
        die(f"LibreOffice failed:\n{r.stdout}\n{r.stderr}")
    key.write_text(h)
    return target


def pdf_titles(pdf: Path) -> list[str]:
    import pymupdf
    titles = []
    with pymupdf.open(pdf) as doc:
        for page in doc:
            best_size, best = 0.0, ""
            for b in page.get_text("dict")["blocks"]:
                for line in b.get("lines", []):
                    spans = [s for s in line["spans"] if s["text"].strip()]
                    if not spans:
                        continue
                    size = max(s["size"] for s in spans)
                    if size > best_size + 0.5:
                        best_size = size
                        best = " ".join(s["text"].strip() for s in spans)
            titles.append(re.sub(r"\s+", " ", best).strip())
    return titles


def pptx_slides(pptx: Path):
    """Returns list of (hidden, title, notes_text) for every slide in the .pptx."""
    try:
        from pptx import Presentation
    except ImportError:
        die("Missing python-pptx.  Run:  pip install -r requirements.txt")
    prs = Presentation(str(pptx))
    out = []
    for s in prs.slides:
        hidden = s._element.get("show") == "0"
        title = ""
        try:
            if s.shapes.title is not None and s.shapes.title.has_text_frame:
                title = s.shapes.title.text_frame.text
        except Exception:
            pass
        notes = ""
        if s.has_notes_slide:
            notes = s.notes_slide.notes_text_frame.text if s.notes_slide.notes_text_frame is not None else ""
        out.append((hidden, re.sub(r"\s+", " ", title).strip(), notes.strip()))
    return out


def resolve_deck(cfg: dict) -> Deck:
    deck = Deck()
    pptx, pdf = find_deck(cfg)
    if pptx:
        deck.pptx = pptx
        sibling = pptx.with_suffix(".pdf")
        if sibling.exists() and sibling.stat().st_mtime >= pptx.stat().st_mtime - 2:
            deck.pdf = sibling                       # exact PowerPoint export
            say(f"  deck: {pptx.name}  (slide images from {sibling.name})")
        else:
            if sibling.exists():
                warn(f"{sibling.name} is older than {pptx.name}; re-export it from PowerPoint "
                     "for exact slide images. Rendering with LibreOffice for now.")
            else:
                say(f"  deck: {pptx.name}  (no PDF export next to it)")
            deck.pdf = render_pptx_to_pdf(pptx)
    else:
        deck.pdf = pdf
        say(f"  deck: {pdf.name}  (PDF only: titles guessed from text; no PowerPoint notes export)")

    import pymupdf
    with pymupdf.open(deck.pdf) as doc:
        npages = doc.page_count
    ptitles = pdf_titles(deck.pdf)

    if deck.pptx:
        slides = pptx_slides(deck.pptx)
        visible = [i for i, s in enumerate(slides) if not s[0]]
        if len(slides) == npages:
            deck.slide_map = list(range(len(slides)))
        elif len(visible) == npages:
            deck.slide_map = visible
            say(f"  ({len(slides) - len(visible)} hidden slide(s) skipped, matching the PDF)")
        else:
            die(f"The .pptx has {len(slides)} slides ({len(visible)} visible) but the PDF has {npages} pages. "
                "Re-export the PDF from the current .pptx.")
        for pos, i in enumerate(deck.slide_map):
            _, t, n = slides[i]
            deck.titles.append(t or ptitles[pos] or f"Slide {pos + 1}")
            deck.pptx_notes.append(n)
    else:
        deck.titles = [t or f"Slide {i + 1}" for i, t in enumerate(ptitles)]
        deck.pptx_notes = [""] * npages
    return deck


def render_slides(deck: Deck, dpi: int) -> list[Path]:
    import pymupdf
    out = BUILD / "slides"
    out.mkdir(parents=True, exist_ok=True)
    n = len(deck.titles)
    width = max(2, len(str(n)))
    key_file = out / ".key"
    key = f"{sha(deck.pdf.read_bytes())}:{dpi}:{width}"
    files = [out / f"slide-{i + 1:0{width}d}.png" for i in range(n)]
    if key_file.exists() and key_file.read_text() == key and all(f.exists() for f in files):
        say(f"  slide images up to date ({n})")
        return files
    for old in out.glob("slide-*.png"):
        old.unlink()
    say(f"  rendering {n} slide images at {dpi} dpi ...")
    with pymupdf.open(deck.pdf) as doc:
        for i, page in enumerate(doc):
            page.get_pixmap(dpi=dpi, alpha=False).save(files[i])
    key_file.write_text(key)
    return files


# ----------------------------------------------------------------------------- notes
HEAD_RE = re.compile(r"^#\s+(\d+)\s*(?:\|\s*(.*?))?\s*$")
SUB_RE = re.compile(r"^##\s+(handout|script)\s*$", re.I)
FENCE_RE = re.compile(r"^\s*(```|~~~)")


class Section:
    def __init__(self, num: int, title: str, body: list[str]):
        self.num, self.title, self.body = num, title, body

    def part(self, name: str) -> str:
        """Markdown under '## Handout' or '## Script' (everything until the next '## ...')."""
        out, active, in_fence = [], False, False
        for line in self.body:
            if FENCE_RE.match(line):
                in_fence = not in_fence
            if not in_fence and line.startswith("## "):
                m = SUB_RE.match(line)
                active = bool(m and m.group(1).lower() == name)
                continue
            if active:
                out.append(line)
        return "\n".join(out).strip("\n")


def parse_notes(text: str) -> tuple[list[str], list[Section]]:
    pre: list[str] = []
    sections: list[Section] = []
    cur: Section | None = None
    in_fence = False
    for line in text.splitlines():
        if FENCE_RE.match(line):
            in_fence = not in_fence
        m = None if in_fence else HEAD_RE.match(line)
        if m:
            cur = Section(int(m.group(1)), (m.group(2) or "").strip(), [])
            sections.append(cur)
        elif cur is None:
            pre.append(line)
        else:
            cur.body.append(line)
    return pre, sections


def write_notes(path: Path, pre: list[str], sections: list[Section]) -> None:
    lines = list(pre)
    while lines and not lines[-1].strip():
        lines.pop()
    out = "\n".join(lines) + "\n" if lines else ""
    for s in sorted(sections, key=lambda s: s.num):
        body = list(s.body)
        while body and not body[-1].strip():
            body.pop()
        while body and not body[0].strip():
            body.pop(0)
        head = f"# {s.num} | {s.title}" if s.title else f"# {s.num}"
        out += f"\n{head}\n\n" + "\n".join(body) + "\n" if body else f"\n{head}\n"
    path.write_text(out.lstrip("\n"), encoding="utf-8")


NOTES_HEADER = """\
<!--
  One section per slide:   # <slide number> | <slide title>
  Under it:                ## Handout   (what participants see below the slide)
                           ## Script    (what you say; goes in the shaded speaker column)
  Use normal Markdown: **bold**, *italic*, lists, `inline code`, and fenced code blocks
  (```python ... ```). Do not use # headings inside a section; use **bold** lines instead.
  After you add, delete or reorder slides, run:  python build.py sync
-->
"""


def skeleton(num: int, title: str, seeded_script: str) -> Section:
    body = ["", "## Handout", "", "", "## Script", ""]
    if seeded_script:
        body += seeded_script.splitlines()
    return Section(num, title, body)


def sync_notes(cfg: dict, deck: Deck, write: bool) -> bool:
    """Make notes.md line up with the deck. Returns True when it is already aligned."""
    path = ROOT / cfg["notes_file"]
    path.parent.mkdir(parents=True, exist_ok=True)
    n = len(deck.titles)
    if not path.exists():
        pre, sections = NOTES_HEADER.splitlines(), []
    else:
        pre, sections = parse_notes(path.read_text(encoding="utf-8"))

    dtitles = [norm(t) for t in deck.titles]
    moves: dict[int, int] = {}
    for s in sections:
        t = norm(s.title)
        if not t or (1 <= s.num <= n and dtitles[s.num - 1] == t):
            continue
        hits = [i + 1 for i, d in enumerate(dtitles) if d == t]
        if len(hits) == 1:
            moves[s.num] = hits[0]
    targets: dict[int, list[int]] = {}
    for s in sections:
        targets.setdefault(moves.get(s.num, s.num), []).append(s.num)
    for tgt, srcs in targets.items():
        if len(srcs) > 1:
            warn(f"sections {srcs} would all land on slide {tgt}; leaving them where they are")
            for src in srcs:
                moves.pop(src, None)

    changes = []
    for s in sections:
        new = moves.get(s.num)
        if new and new != s.num:
            changes.append(f"section '{s.title}' moved: slide {s.num} -> {new}")
            s.num = new
    have = {s.num for s in sections}
    orphans = [s for s in sections if s.num > n or s.num < 1]
    for s in orphans:
        changes.append(f"section {s.num} ('{s.title}') has no slide (deck has {n}); kept at the end")
    for i in range(n):
        if (i + 1) not in have:
            changes.append(f"new slide {i + 1}: added empty section '{deck.titles[i]}'")
            sections.append(skeleton(i + 1, deck.titles[i], deck.pptx_notes[i]))
    for s in sections:
        if 1 <= s.num <= n and s.title != deck.titles[s.num - 1]:
            if s.title:
                changes.append(f"slide {s.num}: title in notes updated to '{deck.titles[s.num - 1]}'")
            s.title = deck.titles[s.num - 1]

    if not changes:
        return True
    for c in changes:
        say(f"  sync: {c}")
    if write:
        if not any("<!--" in l for l in pre):
            pre = NOTES_HEADER.splitlines() + [""] + pre
        write_notes(path, pre, sections)
        say(f"  sync: wrote {path.relative_to(ROOT)}  (review it with `git diff`)")
    else:
        warn("notes.md is out of step with the deck; run `python build.py sync`")
    return False


# ----------------------------------------------------------------------------- pandoc
def md_prepare(md: str) -> str:
    """Demote stray '###' headings to bold lines (headings inside notes would break the layout)."""
    out, in_fence = [], False
    for line in md.splitlines():
        if FENCE_RE.match(line):
            in_fence = not in_fence
        m = None if in_fence else re.match(r"^#{1,6}\s+(.*)$", line)
        out.append(f"\n**{m.group(1).strip()}**\n" if m else line)
    return "\n".join(out)


def pandoc(md: str, to: str) -> str:
    exe = which("pandoc")
    if not exe:
        die("pandoc not found. Install it from https://pandoc.org/installing.html")
    args = [exe, "-f", "markdown+smart", "-t", to, "--wrap=preserve"]
    if to == "latex":
        args += ["--lua-filter", str(ROOT / "latex" / "codeblocks.lua")]
    r = subprocess.run(args, input=md_prepare(md), capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        die(f"pandoc failed:\n{r.stderr}")
    return r.stdout


def convert_notes(cfg: dict, deck: Deck) -> dict[int, str]:
    """Writes build/tex/handout-NN.tex and script-NN.tex. Returns plain-text script per slide."""
    path = ROOT / cfg["notes_file"]
    _, sections = parse_notes(path.read_text(encoding="utf-8")) if path.exists() else ([], [])
    by_num: dict[int, Section] = {}
    for s in sections:
        if s.num in by_num:
            warn(f"slide {s.num} appears twice in notes.md; using the first")
        else:
            by_num[s.num] = s
    tex = BUILD / "tex"
    tex.mkdir(parents=True, exist_ok=True)
    cache_file = tex / ".cache.json"
    cache = json.loads(cache_file.read_text()) if cache_file.exists() else {}
    new_cache, plain = {}, {}
    n = len(deck.titles)
    width = max(2, len(str(n)))
    for i in range(1, n + 1):
        sec = by_num.get(i)
        for part in ("handout", "script"):
            md = sec.part(part) if sec else ""
            f = tex / f"{part}-{i:0{width}d}.tex"
            key = sha(("v3|" + md).encode())
            new_cache[f.name] = key
            if cache.get(f.name) == key and f.exists():
                pass
            else:
                f.write_text(pandoc(md, "latex") if md.strip() else "", encoding="utf-8")
            if part == "script":
                p = tex / f"script-{i:0{width}d}.txt"
                if not (cache.get(p.name) == key and p.exists()):
                    p.write_text(pandoc(md, "plain").strip() if md.strip() else "", encoding="utf-8")
                new_cache[p.name] = key
                plain[i] = p.read_text(encoding="utf-8")
    cache_file.write_text(json.dumps(new_cache))
    for num in by_num:
        if num > n or num < 1:
            warn(f"notes for slide {num} ignored (deck has {n} slides)")
    return plain


# ----------------------------------------------------------------------------- latex
def write_config_tex(cfg: dict, n: int) -> None:
    paper = "letterpaper" if str(cfg["paper"]).lower().startswith("l") else "a4paper"
    col = float(cfg["script_column_in"])
    right_mm = col * 25.4 + 8
    orient = ",landscape" if str(cfg["speaker_orientation"]).lower().startswith("l") else ""
    f = cfg["fonts"]
    lines = [
        "% generated by build.py; do not edit",
        rf"\newcommand{{\TalkTitle}}{{{tex_escape(str(cfg['title']))}}}",
        rf"\newcommand{{\TalkSubtitle}}{{{tex_escape(str(cfg['subtitle']))}}}",
        rf"\newcommand{{\TalkSpeaker}}{{{tex_escape(str(cfg['speaker']))}}}",
        rf"\newcommand{{\TalkEvent}}{{{tex_escape(str(cfg['event']))}}}",
        rf"\newcommand{{\TalkDate}}{{{tex_escape(str(cfg['date']))}}}",
        rf"\newcommand{{\TalkNumSlides}}{{{n}}}",
        rf"\newcommand{{\FontBody}}{{{f['body']}}}",
        rf"\newcommand{{\FontHeading}}{{{f['heading']}}}",
        rf"\newcommand{{\FontScript}}{{{f['script']}}}",
        rf"\newcommand{{\FontCode}}{{{f['code']}}}",
        rf"\newcommand{{\GeometryHandout}}{{\geometry{{{paper},top=16mm,bottom=16mm,left=18mm,right=18mm,footskip=9mm}}}}",
        rf"\newcommand{{\GeometrySpeaker}}{{\geometry{{{paper}{orient},top=14mm,bottom=14mm,left=14mm,right={right_mm:.2f}mm,footskip=8mm}}}}",
        rf"\newlength{{\ColW}}\setlength{{\ColW}}{{{col}in}}",
        rf"\newlength{{\MinNotes}}\setlength{{\MinNotes}}{{{cfg['min_notes_mm']}mm}}",
        r"\newif\ifSpeakerHandout" + ("\n" + (r"\SpeakerHandouttrue" if cfg["speaker_shows_handout_notes"] else r"\SpeakerHandoutfalse")),
        r"\newif\ifBlankLines" + ("\n" + (r"\BlankLinestrue" if cfg["handout_blank_lines"] else r"\BlankLinesfalse")),
        r"\newif\ifCover" + ("\n" + (r"\Covertrue" if cfg["cover_page"] else r"\Coverfalse")),
    ]
    (BUILD / "tex").mkdir(parents=True, exist_ok=True)
    (BUILD / "tex" / "config.tex").write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_bodies(cfg: dict, n: int) -> None:
    width = max(2, len(str(n)))
    for kind, macro in (("handout", "HandoutPage"), ("speaker", "SpeakerPage")):
        body = [rf"\{macro}{{{i:0{width}d}}}{{{i}}}" for i in range(1, n + 1)]
        (BUILD / "tex" / f"{kind}-body.tex").write_text("\n".join(body) + "\n", encoding="utf-8")


def compile_pdf(name: str, prefix: str) -> Path | None:
    exe = which("xelatex")
    if not exe:
        die("xelatex not found. Install TeX Live or MiKTeX (the build needs XeLaTeX).")
    src = ROOT / "latex" / f"{name}.tex"
    log = BUILD / f"{name}.log"
    cmd = [exe, "-interaction=nonstopmode", "-halt-on-error", "-file-line-error",
           f"-output-directory={BUILD}", str(src.relative_to(ROOT))]
    for run in (1, 2):                      # two passes: page-anchored overlays need the .aux
        r = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode != 0:
            tail = "\n".join((log.read_text(errors="replace") if log.exists() else r.stdout).splitlines()[-40:])
            die(f"LaTeX failed on {name}.tex (pass {run}). Last lines of the log:\n{tail}")
    OUT.mkdir(exist_ok=True)
    dest = OUT / f"{prefix}-{name}.pdf"
    shutil.copyfile(BUILD / f"{name}.pdf", dest)
    # surface things the author should know about
    text = log.read_text(errors="replace")
    cont = sorted({int(m) for m in re.findall(r"SCRIPT-CONTINUED slide (\d+)", text)})
    if cont:
        warn(f"{name}: script on slide(s) {cont} is longer than one column; it continues on extra page(s)")
    for m in sorted(set(re.findall(r"Missing character: There is no (.*?) in font (.*?)!", text)))[:5]:
        warn(f"{name}: glyph {m[0]} missing in font {m[1]}")
    for m in sorted(set(re.findall(r"talk: font '(.*?)' not installed", text))):
        warn(f"font '{m}' is not installed; using Latin Modern instead (see fonts: in config.yml)")
    return dest


# ----------------------------------------------------------------------------- pptx export
def export_pptx(deck: Deck, plain: dict[int, str], prefix: str) -> Path | None:
    if not deck.pptx:
        return None
    from pptx import Presentation
    prs = Presentation(str(deck.pptx))
    slides = list(prs.slides)
    n = 0
    for pos, idx in enumerate(deck.slide_map):
        text = plain.get(pos + 1, "")
        if not text:
            continue
        slides[idx].notes_slide.notes_text_frame.text = text
        n += 1
    OUT.mkdir(exist_ok=True)
    dest = OUT / f"{prefix}-with-script.pptx"
    prs.save(str(dest))
    say(f"  wrote {dest.relative_to(ROOT)}  (script in the notes pane of {n} slides; your original is untouched)")
    return dest


# ----------------------------------------------------------------------------- main
def clean() -> None:
    for d in (BUILD, OUT):
        if d.exists():
            shutil.rmtree(d)
    say("removed build/ and output/")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", nargs="?", default="build", choices=["build", "sync", "slides", "clean"])
    ap.add_argument("--only", choices=["handout", "speaker"], help="build just one PDF")
    ap.add_argument("--no-pptx", action="store_true", help="skip writing the .pptx with script notes")
    args = ap.parse_args()

    if args.command == "clean":
        return clean()
    cfg = load_config()
    BUILD.mkdir(exist_ok=True)
    say("Reading deck")
    deck = resolve_deck(cfg)
    say(f"  {len(deck.titles)} slides")

    if args.command == "sync":
        say("Syncing notes")
        if sync_notes(cfg, deck, write=True):
            say("  notes.md already matches the deck")
        return
    say("Rendering slides")
    render_slides(deck, int(cfg["dpi"]))
    if args.command == "slides":
        return

    say("Checking notes")
    path = ROOT / cfg["notes_file"]
    if not path.exists():
        say("  no notes file yet: creating one from the deck")
        sync_notes(cfg, deck, write=True)
    else:
        sync_notes(cfg, deck, write=False)
    say("Converting notes")
    plain = convert_notes(cfg, deck)

    n = len(deck.titles)
    write_config_tex(cfg, n)
    write_bodies(cfg, n)
    prefix = (deck.pptx or deck.pdf).stem
    say("Compiling LaTeX")
    made = []
    for name in ("handout", "speaker"):
        if args.only and args.only != name:
            continue
        made.append(compile_pdf(name, prefix))
        say(f"  wrote {made[-1].relative_to(ROOT)}")
    if cfg["export_pptx_notes"] and not args.no_pptx and not args.only:
        export_pptx(deck, plain, prefix)
    say("\nDone.")


if __name__ == "__main__":
    main()
