"""
Generate Discussion section as a formatted Word document.
Writes to: ../docs/Discussion_Trends_2026.docx
Target: ~1,200 words across 5a (interpretation of results), 5b (implications), 5c (limitations).

Updated for v2:
  - Corrected STEEP numbers (Environmental 53%, Other 32% for fisheries)
  - Added transparent reporting of the Other/Unclassified category
  - Removed author note from 5c; replaced with actual text on language bias
  - Fixed Hiltunen spelling (was Hitunen in original draft)
  - Added [VERIFY] flags on citations not present in the original draft
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm

OUT = Path(__file__).parent.parent / "docs" / "Discussion_Trends_2026.docx"

doc = Document()

for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(3)
    section.right_margin  = Cm(3)


def heading(text, level=1):
    p = doc.add_heading(text, level=level)
    p.runs[0].font.size = Pt(12 if level == 1 else 11)
    return p


def body(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(6)
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
# DISCUSSION SECTION
# ============================================================

heading("5. Discussion", level=1)

# ── 5a Discussion of Results ──────────────────────────────────
heading("5a. Interpretation of findings", level=2)

body(
    "The results of this review reveal a systematic and consequential asymmetry between "
    "the futures landscape as characterised in the cross-sectoral megatrend literature "
    "and the futures landscape as reflected in fisheries-specific research. Whilst both "
    "bodies of literature converge on climate change as the dominant driver of change — "
    "a finding consistent with the broader futures studies canon (Hajkowicz, 2015; "
    "Naughtin et al., 2022) — the fisheries futures literature exhibits a markedly narrower "
    "disciplinary range. The strong concentration of fisheries-specific entries in the "
    "environmental domain (53% of all classified entries, rising to 72% when entries "
    "without sufficient metadata for STEEP classification are excluded) is, on one level, "
    "unsurprising: ecological disruption, overfishing, by-catch, and habitat loss are "
    "longstanding concerns in fisheries science, and the literature accordingly reflects "
    "sustained engagement with these pressures. Yet this concentration simultaneously "
    "reveals an absence."
)

body(
    "A substantial proportion of entries in the fisheries-specific database — 32% "
    "(n = 19) — could not be assigned to any single STEEP domain and were retained as "
    "unclassified. These fall into two analytically distinct groups: entries with "
    "insufficient metadata for classification (no trend label, title, or descriptive "
    "content in the database), and entries describing methodological approaches "
    "— scenario frameworks, trend overviews, driver taxonomies — rather than substantive "
    "thematic content. The former group reflects a limitation of the database as a "
    "living document in which some entries remain incompletely populated; the latter "
    "reflects a genuine feature of the fisheries futures literature, which contains a "
    "substantial body of methodological and meta-analytic work that resists direct "
    "classification under STEEP. Both groups are treated transparently in the reported "
    "figures and excluded from the percentage-based gap analysis presented in Figure 5."
)

body(
    "The most striking finding of this review is the complete absence of technological "
    "drivers within the classified fisheries futures literature. Across the cross-sectoral "
    "megatrend database, technological change — encompassing accelerating automation, "
    "hyperconnectivity, synthetic biology, and digitalisation — constituted 28% of all "
    "classified entries, representing one of the two largest driver categories alongside "
    "environmental change. No comparable body of work was identified in the manually "
    "screened fisheries-specific literature (0%, gap = +28 percentage points). This gap "
    "is consistent across source types and exceeds the analogous deficits in the social "
    "and economic domains (+11 and +15 percentage points respectively). It is important "
    "to note that this finding is based on the 60 papers drawn from manual search "
    "strategy 1 and may be revised following completion of full-text screening of the "
    "720 candidate papers identified via the automated OpenAlex search. Nonetheless, "
    "the absence of technological framing across all 60 manually reviewed sources "
    "suggests that the structural implications of technological transformation for "
    "fisheries systems have not yet been systematically integrated into futures-oriented "
    "fisheries research."
)

body(
    "The underrepresentation of social and economic drivers in the fisheries futures "
    "literature is likewise notable. Social drivers — including demographic change, "
    "urbanisation, and shifting consumer preferences — are well-established in the "
    "cross-sectoral foresight literature as significant forces reshaping demand for "
    "seafood and the composition of fishing communities (KPMG, 2014; Knowledge4Policy, "
    "2022). Their limited presence in fisheries-specific futures research may reflect "
    "disciplinary conventions that situate social dynamics as context for ecological "
    "or management analyses, rather than as primary objects of foresight inquiry. "
    "Economic foresight — addressing market restructuring, the growing economic weight "
    "of emerging markets, and the globalisation of seafood supply chains — was "
    "similarly underrepresented, despite the centrality of economic incentives to "
    "fisheries governance and sustainability transitions."
)

body(
    "A further observation concerns the methodological range of the fisheries "
    "futures literature. The majority of identified studies employed scenario analysis "
    "or quantitative projection techniques to model ecological outcomes under specified "
    "assumptions. Whilst this approach has produced important insights — particularly "
    "in relation to the projected effects of climate change on fish stock distributions "
    "and yields (Cheung et al., 2010 [VERIFY]) — it represents only a subset of the "
    "foresight methods available. Approaches such as horizon scanning, Delphi methods, "
    "causal layered analysis, and participatory futures workshops, which are routinely "
    "employed in cross-sectoral foresight practice, appear sparingly in the fisheries "
    "literature. This methodological narrowness may itself constrain the range of futures "
    "that the field is able to articulate and respond to."
)

# ── 5b Discussion of Implications ────────────────────────────
heading("5b. Implications for research, education, and policy", level=2)

body(
    "For researchers, the findings point to a clear priority: the development of "
    "futures-oriented fisheries research that takes technological change seriously "
    "as a driver in its own right. This does not require the wholesale adoption of "
    "technological determinism, but rather the integration of technology assessment "
    "and foresight into the analytical repertoire of fisheries science. Particular "
    "attention is warranted for emerging technologies that are already transforming "
    "adjacent sectors but whose implications for fisheries remain poorly characterised: "
    "these include precision fermentation and cultivated seafood, advanced monitoring "
    "systems for illegal, unreported, and unregulated (IUU) fishing, and AI-driven "
    "stock assessment. More broadly, the review underscores the value of "
    "interdisciplinary collaboration between fisheries scientists, foresight "
    "practitioners, economists, and social scientists. Futures studies is an "
    "inherently transdisciplinary field (Cuhls, 2020 [VERIFY]), and its application "
    "to fisheries is likely to yield richer results through such collaboration than "
    "through disciplinary extension alone. The living database introduced in this "
    "review is intended as a practical contribution to this integrative effort, "
    "offering a shared resource that the research community can extend and refine "
    "over time."
)

body(
    "For educators, particularly those involved in fisheries science training and "
    "marine resource management curricula, the review highlights a curricular gap. "
    "The relative absence of social, economic, and technological foresight in the "
    "fisheries futures literature suggests that students and early-career researchers "
    "in the sector may have limited exposure to the conceptual tools of futures "
    "studies — including the distinction between drivers, trends, and signals; the "
    "use of megatrend frameworks such as STEEP; and the methodological diversity of "
    "horizon scanning and scenario planning. Incorporating these approaches into "
    "graduate-level training in fisheries science would better equip future "
    "researchers and practitioners to anticipate, rather than merely respond to, "
    "structural change in the sector."
)

body(
    "For policymakers and managers, the most immediate implication concerns the "
    "risk of an evidence base that is systematically skewed towards environmental "
    "drivers. Fisheries management decisions are increasingly taken in contexts "
    "shaped by forces — technological disruption, geopolitical realignment, "
    "demographic change, and economic restructuring — that lie at the margins of "
    "the existing fisheries futures literature. Governance frameworks designed "
    "principally to address stock depletion and ecological sustainability may be "
    "poorly positioned to respond to these emerging pressures, including the "
    "governance of automated fishing fleets, the regulation of novel seafood "
    "substitutes, or the distributional consequences of shifting seafood supply "
    "chains for small-scale fishing communities. Bridging the gap between "
    "cross-sectoral foresight intelligence and fisheries policy requires sustained "
    "institutional investment in horizon scanning capacity within fisheries "
    "management bodies."
)

# ── Limitations ───────────────────────────────────────────────
heading("5c. Limitations", level=2)

body(
    "Several limitations of this review warrant acknowledgement. The cross-sectoral "
    "megatrend database was assembled through a manual search of known foresight "
    "publications and grey literature sources, and cannot be considered exhaustive; "
    "the selection of sources necessarily reflects the availability of institutional "
    "reports known to the reviewers at the time of searching. The STEEP classification "
    "scheme, whilst well-established in futures studies, involves a degree of "
    "interpretive judgement: many drivers are inherently multidimensional, and "
    "first-match keyword classification inevitably simplifies cross-cutting "
    "phenomena. Classification was conducted computationally and was not subject "
    "to formal inter-rater reliability testing, which represents a methodological "
    "limitation that future work should address. The weak signal database remains "
    "in its initial stages; the signals identified at this stage are indicative "
    "rather than representative of the full landscape of emerging developments. "
    "The automated OpenAlex search yielded 720 candidate papers pending full-text "
    "eligibility assessment: the composition of the fisheries futures evidence base "
    "may change substantially once screening is complete, and the gap analysis "
    "figures reported above should be interpreted with this caveat in mind."
)

body(
    "The review is further subject to a language bias: all databases were searched "
    "in English, and grey literature sources were drawn primarily from European "
    "and North American institutional contexts. Futures-oriented fisheries research "
    "conducted in other languages — particularly in East and Southeast Asian contexts, "
    "where marine fisheries are of major economic and food security importance — "
    "may not be adequately represented in the evidence base. This limitation may "
    "be particularly consequential for the gap analysis, given that technological "
    "innovation in fisheries is active in several non-Anglophone research communities."
)

# ── Save ──────────────────────────────────────────────────────
OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(OUT)
print(f"Saved: {OUT}")
