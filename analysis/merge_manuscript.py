"""
Merge all section documents into a single manuscript file.
Writes to: ../docs/Manuscript_Trends_2026.docx

Section order:
  Title page
  Abstract          <- Abstract_Trends_2026.docx
  1. Introduction   <- Introduction_Trends_2026.docx
  2. Methods        <- Method Section Trends.docx  (existing draft)
  3. Results        <- Results_Trends_2026.docx
  4. Discussion     <- Discussion_Trends_2026.docx
  5. Conclusions    <- Conclusions_Trends_2026.docx
"""

from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import lxml.etree as etree

BASE = Path(__file__).parent.parent
DOCS = BASE / "docs"
OUT  = DOCS / "Manuscript_Trends_2026.docx"

# Source files in manuscript order
SOURCES = [
    ("ABSTRACT",      DOCS / "Abstract_Trends_2026.docx"),
    ("INTRODUCTION",  DOCS / "Introduction_Trends_2026.docx"),
    ("METHODS",       DOCS / "Method Section Trends.docx"),
    ("RESULTS",       DOCS / "Results_Trends_2026.docx"),
    ("DISCUSSION",    DOCS / "Discussion_Trends_2026.docx"),
    ("CONCLUSIONS",   DOCS / "Conclusions_Trends_2026.docx"),
]


# ── Helpers ───────────────────────────────────────────────────

def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(__import__("docx.enum.text", fromlist=["WD_BREAK"]).WD_BREAK.PAGE)


def copy_paragraph(src_para, tgt_doc):
    """Deep-copy a paragraph's XML into the target document body."""
    new_para = deepcopy(src_para._element)
    tgt_doc.element.body.append(new_para)


def add_divider_heading(doc, label):
    """Insert a bold grey section-divider label (not part of the manuscript flow)."""
    p = doc.add_paragraph()
    run = p.add_run(f"── {label} ──────────────────────────────────────")
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    run.font.size = Pt(8)
    run.font.bold = False
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)


def append_document(master, src_path, label):
    """Append all non-empty paragraphs from src_path into master."""
    if not src_path.exists():
        print(f"  MISSING: {src_path.name} — skipping")
        return

    src = Document(src_path)
    add_divider_heading(master, label)

    para_count = 0
    for para in src.paragraphs:
        if para.text.strip() or para.style.name.startswith("Heading"):
            copy_paragraph(para, master)
            para_count += 1

    print(f"  Appended {para_count:>3} paragraphs from {src_path.name}")


# ── Build manuscript ──────────────────────────────────────────

master = Document()

# Page layout
for sec in master.sections:
    sec.top_margin    = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin   = Cm(3)
    sec.right_margin  = Cm(3)

# ── Title page ────────────────────────────────────────────────
p = master.add_paragraph()
r = p.add_run("Drivers, Trends, and Signals in Fisheries Research:\n"
              "A Structured Literature Review and Horizon Scan")
r.bold = True
r.font.name = "Times New Roman"
r.font.size = Pt(14)
p.paragraph_format.space_after = Pt(18)

p = master.add_paragraph()
r = p.add_run(
    "Timo Szczepanska\n"
    "Norwegian College of Fishery Science\n"
    "UiT The Arctic University of Norway, Tromsø, Norway\n"
    "tsz000@uit.no\n"
    "ORCID: 0000-0003-2442-8223"
)
r.font.name = "Times New Roman"
r.font.size = Pt(11)
p.paragraph_format.space_after = Pt(6)

p = master.add_paragraph()
r = p.add_run("Target journal: Reviews in Fish Biology and Fisheries")
r.italic = True
r.font.name = "Times New Roman"
r.font.size = Pt(10)
p.paragraph_format.space_after = Pt(24)

p = master.add_paragraph()
r = p.add_run(
    "Keywords: horizon scanning; megatrends; fisheries futures; STEEP framework; "
    "systematic literature review; weak signals; foresight"
)
r.font.name = "Times New Roman"
r.font.size = Pt(10)

add_page_break(master)

# ── Append all sections ───────────────────────────────────────
print("Merging sections:")
for label, path in SOURCES:
    append_document(master, path, label)

# ── Remove the initial empty paragraph docx always adds ───────
# (first element of body is often an empty <w:p>)
body = master.element.body
first = body[0]
if first.tag == qn("w:p") and not first.text_content if hasattr(first, "text_content") else True:
    pass  # leave it; cosmetic only

# ── Save ──────────────────────────────────────────────────────
master.save(OUT)
print(f"\nSaved: {OUT}")
print(f"Total paragraphs in manuscript: {len(master.paragraphs)}")
