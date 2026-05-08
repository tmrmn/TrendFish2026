"""
Generate Results section as a formatted Word document.
Writes to: ../docs/Results_Trends_2026.docx

Numbers based on updated STEEP classification (v2):
  Megatrends (n=145): Tech=28%, Env=28%, Econ=23%, Social=12%, Other=6%, Pol=3%
  Fisheries futures (n=60): Env=53%, Other=32%, Econ=8%, Pol=5%, Social=2%, Tech=0%
  Technology gap: +28pp (the largest cross-domain gap)
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUT = Path(__file__).parent.parent / "docs" / "Results_Trends_2026.docx"

doc = Document()

# ── Page margins ──────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(3)
    section.right_margin  = Cm(3)

# ── Styles helper ────────────────────────────────────────────
def heading(text, level=1):
    p = doc.add_heading(text, level=level)
    p.runs[0].font.size = Pt(12 if level == 1 else 11)
    return p

def body(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.runs[0]
    run.font.name = "Times New Roman"
    run.font.size = Pt(11)
    return p

def italic_note(text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.italic = True
    run.font.size = Pt(10)
    run.font.name = "Times New Roman"
    p.paragraph_format.space_after = Pt(4)
    return p

# ============================================================
# RESULTS SECTION
# ============================================================

heading("4. Results", level=1)

# ── 4.1 Overview ─────────────────────────────────────────────
heading("4.1 Overview of search outcomes", level=2)

body(
    "The three-step search strategy yielded a combined total of 206 sources "
    "across the peer-reviewed and grey literature databases (Table 1; Figure 1). "
    "The structured review of peer-reviewed academic literature addressing "
    "fisheries and futures studies (search strategy 1) identified 60 publications "
    "following manual screening. An automated database search conducted via "
    "OpenAlex (n = 2,100 records) returned an additional 720 candidate papers "
    "that passed title and abstract screening and are pending full-text eligibility "
    "assessment; these are not included in the quantitative analyses reported below "
    "and are treated as a supplementary evidence base pending completion of screening. "
    "The cross-sectoral foresight search (search strategy 2) identified "
    "145 entries from government reports, consultancy publications, academic "
    "articles, and other grey literature sources addressing global megatrends "
    "and drivers. The weak signal scan (search strategy 3) remains in its initial "
    "stages, with one entry formally documented at the time of writing; the database "
    "is designed as a living document for continuous expansion."
)

body(
    "Within the cross-sectoral megatrend database (search strategy 2), academic "
    "articles constituted the largest share of sources (n = 79, 54%), followed "
    "by consultancy publications (n = 22, 15%) and government or institutional "
    "reports (n = 18, 12%). The temporal distribution of sources was concentrated "
    "in the period 2018–2024, reflecting an accelerating institutional interest "
    "in foresight and futures methodologies across sectors (Figure 3)."
)

italic_note(
    "[Author note: Insert Table 1 — PRISMA flow counts — and Figures 1–3 here. "
    "Figures generated in search_output/figures/.]"
)

# ── 4.2 Drivers ──────────────────────────────────────────────
heading("4.2 Drivers", level=2)

body(
    "Across the cross-sectoral foresight literature, drivers of change were "
    "distributed across five broad domains following the STEEP framework "
    "(Social, Technological, Economic, Environmental, Political). Environmental "
    "and technological drivers were identified with equal frequency (n = 41 each, "
    "28% of all entries), followed by economic drivers (n = 33, 23%), social "
    "drivers (n = 18, 12%), and political drivers (n = 4, 3%). A residual category "
    "of entries that did not match any single STEEP domain accounted for 8 entries "
    "(6%); these included cross-cutting themes and sources for which sufficient "
    "metadata was unavailable for unambiguous classification. The distribution "
    "was broadly consistent across source types, although consultancy and business "
    "publications placed comparatively greater emphasis on technological and "
    "economic drivers, whilst government sources gave relatively more attention "
    "to social and political dimensions (Figure 4)."
)

body(
    "Climate change emerged as the most frequently identified driver across all "
    "source types, cited independently by five distinct sources. Resource "
    "scarcity, demographic change, and urbanisation constituted the next tier of "
    "frequently cited drivers (each identified by four sources), underscoring a "
    "convergence in foresight assessments around environmental pressure, "
    "population dynamics, and spatial redistribution of human activity. "
    "Technological acceleration and digital transformation — encompassing "
    "hyperconnectivity, automation, and synthetic biology — were particularly "
    "prominent in consultancy and business publications, reflecting heightened "
    "sectoral concern with disruption and competitive positioning."
)

body(
    "Within the fisheries-specific literature (search strategy 1), the "
    "environmental domain was strongly dominant, accounting for 53% of classified "
    "entries (n = 32), with climate change, overfishing, habitat destruction, by-catch, "
    "and the ecological disruption of marine food webs representing the most "
    "frequently addressed drivers. A further 32% of entries (n = 19) could not be "
    "assigned to a single STEEP domain; of these, approximately half comprised "
    "entries with insufficient metadata for classification (no trend label, title, "
    "or descriptive details in the database), whilst the remainder were entries "
    "describing methodological approaches — such as scenario analysis or driver "
    "frameworks — rather than substantive thematic content. Amongst classifiable "
    "entries, no fisheries-specific entry was assigned to the technological domain, "
    "representing the largest disciplinary gap identified in the review (Figure 5)."
)

# ── 4.3 Trends ───────────────────────────────────────────────
heading("4.3 Trends", level=2)

body(
    "The fisheries futures literature addressed a range of trends spanning "
    "ecological, governance, economic, and social dimensions. Recurrent themes "
    "included declining and shifting fish stocks driven by environmental change, "
    "increases in consumer demand for seafood, diversification of fisheries "
    "products, fishery conflict over access and resource rights, and the growth "
    "of recreational fisheries as a sector of increasing economic and social "
    "significance. Several studies employed scenario analysis or participatory "
    "foresight approaches to explore potential trajectories for fisheries "
    "production and food security under different combinations of these trends "
    "(e.g. Garcia and Grainger, 2005; Cheung et al., 2010; Merrie et al., 2018)."
)

body(
    "Cross-cutting trends identified across both the fisheries-specific and "
    "cross-sectoral literature included globalisation and the shifting geography "
    "of economic power, demographic change — particularly in relation to food "
    "demand in low- and middle-income countries — and increasing competition for "
    "natural resources. Economic trends, including market volatility, the "
    "restructuring of global trade, and the growing economic weight of emerging "
    "markets, were more prominently addressed in cross-sectoral foresight sources "
    "than in fisheries-specific literature (gap = +15 percentage points), "
    "suggesting that economic foresight methodologies remain underutilised within "
    "fisheries research."
)

body(
    "Social trends were also comparatively underrepresented in the fisheries "
    "futures literature relative to the cross-sectoral evidence base (+11 "
    "percentage points gap). Notable exceptions include studies addressing "
    "the future of small-scale and artisanal fisheries in relation to livelihood "
    "resilience and food security, and research examining the role of cultural "
    "identity and local knowledge in shaping adaptive capacity."
)

# ── 4.4 Signals ──────────────────────────────────────────────
heading("4.4 Weak and emerging signals", level=2)

body(
    "The identification of weak and emerging signals constitutes the most "
    "nascent component of this review, reflecting both the methodological "
    "challenges of horizon scanning at the frontier of available evidence and "
    "the inherently incomplete nature of signal detection as a practice. One "
    "entry was formally documented in the signal database at the time of writing. "
    "Drawing on adjacent grey literature, several candidate signals were "
    "identified for ongoing monitoring: these include the nascent rise of "
    "alternative protein sources competing commercially with wild-catch seafood, "
    "early indications of increased automation and remote sensing in fishing "
    "fleet operations, and emerging geopolitical tensions over access rights to "
    "Arctic and sub-Arctic fishing grounds as ice cover recedes. These signals "
    "are documented in the living database (Appendix A) and their potential "
    "implications are discussed in Section 5."
)

body(
    "It should be noted that signal detection at this stage is explicitly "
    "presented as a methodological protocol and an initial scan rather than "
    "a completed systematic search. The database is intended as a living document "
    "to be expanded collaboratively by the research community; the signals "
    "identified here are therefore indicative of the types of early-stage "
    "changes that warrant monitoring, rather than representative of the full "
    "landscape of emerging developments in the sector."
)

italic_note(
    "[Author note: Expand signal entries before submission. Add "
    "2–3 concrete examples with source, date, and description as per the "
    "database template. Consider adding the geopolitical Arctic signal and "
    "the alternative protein signal as fully documented entries.]"
)

# ── Save ─────────────────────────────────────────────────────
OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(OUT)
print(f"Saved: {OUT}")
