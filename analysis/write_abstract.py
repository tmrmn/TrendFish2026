"""
Generate Abstract as a formatted Word document.
Writes to: ../docs/Abstract_Trends_2026.docx
Target: ~300 words (journal allows up to 400 for long articles).
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm

OUT = Path(__file__).parent.parent / "docs" / "Abstract_Trends_2026.docx"

doc = Document()

for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(3)
    section.right_margin  = Cm(3)


def body(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.runs[0]
    run.font.name = "Times New Roman"
    run.font.size = Pt(11)
    return p


def bold_label(para, label):
    run = para.add_run(label)
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(11)
    return run


# ── Title block ───────────────────────────────────────────────
p = doc.add_paragraph()
r = p.add_run("Drivers, Trends, and Signals in Fisheries Research: "
              "A Structured Literature Review and Horizon Scan")
r.bold = True
r.font.name = "Times New Roman"
r.font.size = Pt(12)
p.paragraph_format.space_after = Pt(12)

# ── Abstract ──────────────────────────────────────────────────
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
bold_label(p, "Background. ")
p.add_run(
    "Fisheries systems face an accelerating array of structural pressures, yet "
    "futures studies methodologies — including horizon scanning, megatrend analysis, "
    "and weak signal detection — remain comparatively underutilised in fisheries "
    "research relative to other sectors. Systematic knowledge of the drivers, "
    "trends, and signals shaping the future of fisheries is a prerequisite for "
    "proactive governance, adaptive management, and evidence-based education."
).font.name = "Times New Roman"

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
bold_label(p, "Methods. ")
p.add_run(
    "We conducted a three-step search strategy comprising: (1) a structured "
    "review of peer-reviewed fisheries futures literature (n = 60 sources after "
    "manual screening, with 720 additional candidate papers from an automated "
    "OpenAlex search pending full-text eligibility assessment); (2) a cross-sectoral "
    "scan of megatrend foresight publications (n = 145 sources); and (3) an initial "
    "horizon scan for weak and emerging signals (n = 1 formally documented entry). "
    "All sources were thematically classified using the STEEP framework (Social, "
    "Technological, Economic, Environmental, Political). The search strategy followed "
    "PRISMA 2020 reporting guidelines (Page et al., 2021)."
).font.name = "Times New Roman"

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
bold_label(p, "Results. ")
p.add_run(
    "Across the cross-sectoral megatrend database, environmental and technological "
    "drivers were equally prominent (28% each), followed by economic (23%) and "
    "social (12%) drivers. In the fisheries-specific futures literature, the "
    "environmental domain was strongly dominant (53% of classified entries), whilst "
    "no entry was classified under the technological domain. The resulting "
    "technology gap of +28 percentage points represents the largest disciplinary "
    "discrepancy identified. Social and economic drivers were also substantially "
    "underrepresented in the fisheries literature relative to the cross-sectoral "
    "baseline (+11 and +15 percentage points respectively)."
).font.name = "Times New Roman"

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
bold_label(p, "Conclusions. ")
p.add_run(
    "Fisheries futures research is characterised by a pronounced environmental "
    "framing that leaves technological, social, and economic drivers of change "
    "systematically underrepresented. This has implications for researchers — who "
    "should prioritise interdisciplinary foresight work on technological transformation "
    "in fisheries — for educators — who should incorporate futures literacy into "
    "fisheries science curricula — and for policymakers — who risk designing "
    "governance frameworks with an incomplete view of the forces shaping the sector. "
    "The structured database of drivers, trends, and signals introduced in this "
    "review is offered as a living resource for the research community."
).font.name = "Times New Roman"

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(12)
bold_label(p, "Keywords. ")
p.add_run(
    "horizon scanning; megatrends; fisheries futures; STEEP framework; "
    "systematic literature review; weak signals; foresight"
).font.name = "Times New Roman"

# ── Save ──────────────────────────────────────────────────────
OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(OUT)
print(f"Saved: {OUT}")
