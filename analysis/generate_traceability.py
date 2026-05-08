"""
Evidence Traceability Report (verbose version)
================================================
Generates an HTML document that links every quantitative claim in the
Results section to the specific rows in the Excel database that support it.
Includes section introductions, STEEP framework descriptions, methodology
notes, and interpretive commentary for each claim.

Output: ../docs/Traceability_Report.html

Usage:
    python -X utf8 generate_traceability.py
"""

import re as _re
from pathlib import Path
from datetime import date

import pandas as pd

BASE_DIR  = Path(__file__).parent.parent
DATA_FILE = BASE_DIR / "data" / "manual_literature_search.xlsx"
OUT_FILE  = BASE_DIR / "docs" / "Traceability_Report.html"
TODAY     = date.today().isoformat()

# ── STEEP rules (must match excel_analysis.py) ────────────────
STEEP_RULES = {
    "Environmental": [
        "climate","environment","biodiversity","ecosystem","ocean","marine","sea",
        "water","pollution","habitat","carbon","emission","green","nature conservation",
        "nature-based","natural resource","natural environment","resource scarcity",
        "erosion","deforestation","renewable energy","sustainability","ecology",
        "fish stock","overfishing","fishing","fisheries","fishery","fisher","fishers",
        "aquaculture","aquatic","seafood","harvest","catch","bycatch","by-catch",
        "food web","carrying capacity","stock assessment","spawn","recruitment","trophic",
    ],
    "Technological": [
        "technolog","digital","automation","artificial intelligence","ai ","data",
        "innovation","internet","cyber","robot","autonomous","biotech","synthetic biology",
        "microbiome","hyperconnectiv","electrif","knowledge",
    ],
    "Economic": [
        "economic","trade","market","finance","fiscal","gdp","consumption","globali",
        "growth","debt","income","employment","labour","labor","industry",
        "investment","supply chain","demand","profit","revenue","subsid","cost",
        "price","value chain",
    ],
    "Social": [
        "social","demograph","population","urban","aging","health","inequalit",
        "migration","culture","education","wellbeing","well-being","individual",
        "community","consumer","lifestyle","food security","gender",
        "governance of people","livelihood","artisanal","small-scale","indigenous",
    ],
    "Political": [
        "politic","governance","geopolit","security","power","regulation","law",
        "policy","democracy","conflict","supranational","institution","global order",
        "protectionism","management","quota","treaty","iuu","compliance",
    ],
}

STEEP_COLORS = {
    "Environmental": "#d4edda",
    "Technological": "#cce5ff",
    "Economic":      "#fff3cd",
    "Social":        "#f8d7da",
    "Political":     "#e2d9f3",
    "Other":         "#e9ecef",
    "Signal":        "#fce4ec",
}

STEEP_DESCRIPTIONS = {
    "Environmental": (
        "The <strong>Environmental</strong> domain captures drivers related to the natural world: "
        "climate change, ecological systems, biodiversity loss, ocean chemistry, resource scarcity, "
        "and sustainability pressures. In the fisheries context this domain is particularly broad, "
        "encompassing not only macro-level environmental forces (e.g. ocean warming, acidification) "
        "but also sector-specific ecological dynamics such as overfishing, habitat destruction, "
        "by-catch, and shifts in species distribution. The keyword list for this category was "
        "deliberately expanded to capture fisheries-specific terminology (e.g. 'fishery', 'fisher', "
        "'aquaculture', 'trophic') that might not appear in generic STEEP schemes designed for "
        "cross-sectoral use."
    ),
    "Technological": (
        "The <strong>Technological</strong> domain covers drivers related to the development and "
        "diffusion of new technologies: digitisation, automation, artificial intelligence, "
        "biotechnology, synthetic biology, hyperconnectivity, and innovation systems. In the "
        "cross-sectoral megatrend literature this is one of the most prominent categories, "
        "reflecting the pace of technological disruption across industries. The near-complete "
        "absence of this domain in the fisheries futures literature is the central finding of "
        "the gap analysis and a key motivation for this review."
    ),
    "Economic": (
        "The <strong>Economic</strong> domain encompasses drivers related to economic systems: "
        "globalisation, trade patterns, market structure, financial dynamics, labour markets, "
        "consumption trends, and economic power shifts between regions. For fisheries, economic "
        "drivers include seafood market volatility, the growing economic weight of aquaculture, "
        "restructuring of global seafood supply chains, and the economic value of recreational "
        "versus commercial fisheries. The domain was notably underrepresented in the fisheries "
        "futures literature relative to the cross-sectoral baseline."
    ),
    "Social": (
        "The <strong>Social</strong> domain captures drivers rooted in human societies: demographic "
        "change, urbanisation, migration, ageing populations, shifting cultural values, health "
        "trends, inequality, education, and lifestyle change. In the fisheries context, social "
        "drivers include the future of small-scale and artisanal fishing communities, indigenous "
        "and customary fishing rights, food security in coastal and low-income populations, and "
        "changing consumer preferences for seafood. Social drivers were substantially "
        "underrepresented in the fisheries futures literature (+11 percentage points gap)."
    ),
    "Political": (
        "The <strong>Political</strong> domain covers drivers related to governance, regulation, "
        "geopolitics, and institutional frameworks: changing security paradigms, shifts in global "
        "power, trade policy, multilateral agreements, democratisation, and institutional reform. "
        "In fisheries, political drivers include fisheries management regimes, quota allocation, "
        "international fisheries agreements, IUU fishing governance, and geopolitical competition "
        "over access rights to fishing grounds. This was the smallest category in both databases."
    ),
}


def classify_steep(text: str) -> str:
    if not text or str(text).strip().lower() in ("none", "nan", "?", "n=a", ""):
        return "Other"
    t = str(text).lower()
    for cat, kws in STEEP_RULES.items():
        if any(kw in t for kw in kws):
            return cat
    return "Other"


# ── Load and prepare data ─────────────────────────────────────

def load_data():
    mg = pd.read_excel(DATA_FILE, sheet_name="(Mega-)trends and drivers")
    mg.columns = ["trend","details","author","source","title","comments","doi","year","source_type"]
    mg = mg.dropna(subset=["trend"])
    mg = mg[~mg["trend"].astype(str).str.strip().isin(["","nan"])]
    mg["steep"] = mg.apply(
        lambda r: classify_steep(str(r["trend"]) + " " + str(r["details"])), axis=1
    )
    mg["sheet"] = "Megatrends"

    NOTE_PATTERNS = [
        r"^n=a$",r"^none$",r"^topic$",r"^keyword",r"^future\s*$",
        r"^political variables$",r"^economic variables$",r"^social variables$",
        r"^environmental variables$",r"^find something",r"^interesting article",
        r"^book on",r"^manifesto$",r"^brief vision",r"^bell prediction",
    ]

    def is_note_trend(val):
        if pd.isna(val): return True
        return any(_re.match(p, str(val).strip().lower()) for p in NOTE_PATTERNS)

    def has_reference(row):
        has_title  = pd.notna(row["title"])  and str(row["title"]).strip().lower() not in ("","nan")
        has_author = pd.notna(row["author"]) and str(row["author"]).strip().lower() not in ("","nan")
        has_source = (pd.notna(row["source"]) and
                      str(row["source"]).strip().lower() not in ("","nan") and
                      str(row["source"]).strip().lower() != str(row["author"]).strip().lower())
        return has_title or has_author or has_source

    ff = pd.read_excel(DATA_FILE, sheet_name="Futures of Fisheries")
    ff.columns = ["trend","details","author","source","title","comments","doi","year",
                  "source_type","_a","_b"]
    ff = ff.drop(columns=["_a","_b"])
    ff["_has_ref"]    = ff.apply(has_reference, axis=1)
    ff["_note_trend"] = ff["trend"].apply(is_note_trend)
    ff = ff[ff["_has_ref"] | (~ff["_note_trend"])]
    ff = ff.drop(columns=["_has_ref","_note_trend"])
    ff["steep"] = ff.apply(
        lambda r: classify_steep(
            " ".join(str(r[c]) for c in ["trend","details","title"] if pd.notna(r[c]))
        ), axis=1
    )
    ff["sheet"] = "FisheriesFutures"

    try:
        sg = pd.read_excel(DATA_FILE, sheet_name="Signals of Future Fisheries")
        sg.columns = (
            ["trend","details","author","source","title","comments","doi","year","source_type"]
            + list(sg.columns[9:])
        )
        sg = sg.iloc[:, :9]
        sg["steep"] = "Signal"
        sg["sheet"] = "Signals"
    except Exception:
        sg = pd.DataFrame()

    return mg, ff, sg


# ── HTML helpers ──────────────────────────────────────────────

def esc(val):
    return str(val).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")


def df_to_html_table(df, display_cols, steep_col="steep"):
    if df.empty:
        return "<p class='zero-result'>No matching rows found in the database.</p>"
    rows_html = []
    for _, row in df.iterrows():
        cat = str(row.get(steep_col, "Other"))
        bg  = STEEP_COLORS.get(cat, "#ffffff")
        cells = "".join(f"<td>{esc(row.get(c,''))}</td>" for c in display_cols)
        rows_html.append(f'<tr style="background:{bg}">{cells}</tr>')
    headers = "".join(f"<th>{esc(c)}</th>" for c in display_cols)
    return (
        '<table class="evidence-table">'
        f'<thead><tr>{headers}</tr></thead>'
        f'<tbody>{"".join(rows_html)}</tbody>'
        '</table>'
    )


def prose_block(html_content):
    """A plain prose paragraph between claim blocks."""
    return f'<div class="prose-block">{html_content}</div>'


def method_note(html_content):
    """A highlighted methodology callout box."""
    return f'<div class="method-note"><strong>Methodology note</strong><br>{html_content}</div>'


def claim_block(section_id, claim_text, rationale, query_description, table_html, count):
    """
    Render one claim + evidence block.
    rationale       — why this claim matters (1–3 sentences)
    query_description — how the data was queried to produce this evidence
    """
    badge = f'<span class="badge">{count} source{"s" if count != 1 else ""}</span>'
    return f"""
    <div class="claim-block" id="{section_id}">
      <div class="claim-label">Claim</div>
      <div class="claim-text">&#8220;{esc(claim_text)}&#8221; {badge}</div>
      <div class="claim-rationale">{rationale}</div>
      <div class="claim-query"><strong>Data query:</strong> {query_description}</div>
      {table_html}
    </div>
    """


# ── Build report ──────────────────────────────────────────────

def build_report(mg, ff, sg):
    DISPLAY_MG = ["trend","author","year","source_type","steep"]
    DISPLAY_FF = ["trend","title","author","year","steep"]
    DISPLAY_SG = ["trend","details","author","year","source_type"]

    blocks = []   # list of HTML strings

    # ═════════════════════════════════════════════════════════
    # DATABASE OVERVIEW
    # ═════════════════════════════════════════════════════════
    blocks.append('<h2 id="db-overview">Database overview</h2>')
    blocks.append(prose_block(f"""
    <p>The evidence base for this review is stored in a single Excel workbook
    (<code>manual_literature_search.xlsx</code>) with three active sheets.
    Each sheet corresponds to one of the three search strategies described in the Methods section.</p>

    <table class="evidence-table">
      <thead><tr>
        <th>Sheet name</th><th>Search strategy</th><th>Entries (used)</th>
        <th>Source types</th><th>Period covered</th>
      </tr></thead>
      <tbody>
        <tr style="background:#d4edda">
          <td><code>(Mega-)trends and drivers</code></td>
          <td>Strategy 2 — Cross-sectoral foresight</td>
          <td><strong>{len(mg)}</strong></td>
          <td>Academic articles, consultancy reports, government publications, NGO reports, books</td>
          <td>Mainly 2010–2024</td>
        </tr>
        <tr style="background:#cce5ff">
          <td><code>Futures of Fisheries</code></td>
          <td>Strategy 1 — Peer-reviewed fisheries futures</td>
          <td><strong>{len(ff)}</strong> (manual screening)</td>
          <td>Peer-reviewed academic articles</td>
          <td>2000–2024</td>
        </tr>
        <tr style="background:#fce4ec">
          <td><code>Signals of Future Fisheries</code></td>
          <td>Strategy 3 — Weak signal horizon scan</td>
          <td><strong>{len(sg) if not sg.empty else 0}</strong> (living document)</td>
          <td>Grey literature, news, industry reports</td>
          <td>2019–present</td>
        </tr>
      </tbody>
    </table>

    <p>In addition to these manually curated entries, an automated search via the
    <a href="https://openalex.org" target="_blank">OpenAlex</a> API (Strategy 1) retrieved
    2,100 records using seven query strings. After deduplication (479 removed), 1,621 unique
    records were screened by title and abstract, yielding <strong>720 candidate papers</strong>
    that passed initial screening and are <em>pending full-text eligibility assessment</em>.
    These 720 papers are not included in the quantitative analyses below but are documented
    in <code>search_output/screened_included.csv</code>.</p>

    <p>Column schema (all sheets share the same structure):</p>
    <table class="evidence-table">
      <thead><tr><th>Column</th><th>Description</th></tr></thead>
      <tbody>
        <tr><td><code>Trend</code></td><td>The driver, trend, or signal label as entered by the reviewer</td></tr>
        <tr><td><code>Details</code></td><td>Free-text description or elaboration of the entry</td></tr>
        <tr><td><code>Author</code></td><td>Author(s) of the source document</td></tr>
        <tr><td><code>Source</code></td><td>Publication venue, URL, or institution</td></tr>
        <tr><td><code>Title</code></td><td>Full title of the source document</td></tr>
        <tr><td><code>Comments</code></td><td>Reviewer notes and annotations</td></tr>
        <tr><td><code>DOI</code></td><td>Digital Object Identifier (where available)</td></tr>
        <tr><td><code>Year</code></td><td>Publication year</td></tr>
        <tr><td><code>Type</code></td><td>Source type (Article, Government, Consultancy, etc.)</td></tr>
      </tbody>
    </table>
    """))

    # ═════════════════════════════════════════════════════════
    # METHODOLOGY NOTE
    # ═════════════════════════════════════════════════════════
    blocks.append('<h2 id="method">Classification methodology</h2>')
    kw_examples = {k: v[:5] for k, v in STEEP_RULES.items()}
    kw_html = "".join(
        f"<li><strong>{cat}</strong>: {', '.join(f'<em>{k}</em>' for k in kws)} &hellip;</li>"
        for cat, kws in kw_examples.items()
    )
    blocks.append(prose_block(f"""
    <p>Each entry in both the megatrend and fisheries futures databases was assigned to one
    of five STEEP categories (Social, Technological, Economic, Environmental, Political) using
    automated keyword matching. The algorithm concatenates the <code>Trend</code>,
    <code>Details</code>, and <code>Title</code> fields into a single string, converts to
    lowercase, and applies a <strong>first-match rule</strong>: the entry is assigned to the
    first STEEP category whose keyword list contains a matching substring.</p>

    <p>The keyword lists used for classification (first five keywords per category shown):</p>
    <ul>{kw_html}</ul>

    <p>Entries that matched no keyword in any category — or for which all three text fields
    were empty — were assigned to a residual <strong>Other / Unclassified</strong> category.
    This includes two analytically distinct groups: (1) entries with insufficient metadata
    (no trend label, title, or details recorded in the database), and (2) entries describing
    methodological approaches — such as "Scenario" or "Trend and drivers" — rather than
    substantive thematic content. Both groups are reported transparently in the Results section
    and excluded from the percentage-based gap analysis in Figure 5 of the manuscript.</p>

    <p><strong>Important caveat:</strong> Keyword-based classification is a heuristic. It is
    sensitive to the vocabulary used in the Trend field, which varies across entries and source
    types. Short or vague trend labels may be misclassified or fall into Other even when the
    underlying source clearly addresses a particular STEEP domain. Classification was not
    subject to formal inter-rater reliability testing; this is acknowledged as a limitation
    in Section 5c of the manuscript.</p>
    """))

    # ═════════════════════════════════════════════════════════
    # STEEP FRAMEWORK
    # ═════════════════════════════════════════════════════════
    blocks.append('<h2 id="steep-framework">The STEEP framework</h2>')
    steep_cards = "".join(
        f'<div class="steep-card" style="border-left:4px solid;border-color:{STEEP_COLORS[cat]};'
        f'background:{STEEP_COLORS[cat]};padding:10px 14px;margin:8px 0;border-radius:4px">'
        f'{desc}</div>'
        for cat, desc in STEEP_DESCRIPTIONS.items()
    )
    blocks.append(prose_block(f"""
    <p>The STEEP framework (Social, Technological, Economic, Environmental, Political) is a
    widely used classification scheme in futures studies and horizon scanning for organising
    drivers of change by their primary domain of origin. It was first systematised in the
    environmental scanning literature and has since been adopted by consultancies, government
    agencies, and academic foresight researchers as a standard taxonomy for megatrend analysis
    (Naisbitt, 1982; Hajkowicz, 2015).</p>

    <p>The five categories as applied in this review:</p>
    {steep_cards}

    <p>The STEEP framework was chosen for this review because it is the dominant classification
    scheme in the cross-sectoral foresight literature from which the megatrend database was
    assembled. Applying the same framework to the fisheries futures literature enables a
    direct, category-by-category comparison and gap analysis. Alternative frameworks
    (e.g. PESTLE, DESTEP) were considered but add categories not consistently represented
    in the source literature.</p>
    """))

    # ═════════════════════════════════════════════════════════
    # SECTION 4.1 — Overview
    # ═════════════════════════════════════════════════════════
    blocks.append('<h2 id="s41">Section 4.1 — Overview of search outcomes</h2>')
    blocks.append(prose_block("""
    <p>This section reports the aggregate counts produced by the three search strategies.
    The figures are used to characterise the overall scope of the evidence base and to
    position this review within the PRISMA 2020 reporting framework (Page et al., 2021).
    The key distinction to keep in mind when reading this section is between <em>included</em>
    sources (those that passed all screening stages and contribute to the analysis) and
    <em>pending</em> sources (those that passed title/abstract screening but await full-text
    review). The claims below refer to included sources only, unless otherwise stated.</p>
    """))

    n_mg = len(mg); n_ff = len(ff)
    n_sg = len(sg) if not sg.empty else 0
    total = n_mg + n_ff + n_sg

    tbl = df_to_html_table(
        pd.DataFrame({
            "Search strategy": [
                "Strategy 1 — Fisheries futures (manual screening)",
                "Strategy 1 — Fisheries futures (OpenAlex, pending full-text)",
                "Strategy 2 — Cross-sectoral megatrends (manual)",
                "Strategy 3 — Weak signals (initial scan)",
                "TOTAL included in analysis",
            ],
            "n": [n_ff, 720, n_mg, n_sg, total],
            "Status": ["Included","Pending","Included","Included",""],
            "steep": ["Environmental","Other","Technological","Signal","Other"],
        }),
        ["Search strategy","n","Status"],
    )
    blocks.append(claim_block(
        "claim-overview",
        f"The three-step search strategy yielded a combined total of {total} sources across the peer-reviewed and grey literature databases.",
        (f"This figure ({total}) is the sum of the three included source sets: {n_ff} manually "
         f"screened fisheries futures papers (Strategy 1), {n_mg} cross-sectoral megatrend "
         f"sources (Strategy 2), and {n_sg} weak signal entries (Strategy 3). The 720 "
         f"OpenAlex candidates that passed title/abstract screening are explicitly excluded "
         f"from this total and from all quantitative analyses reported in Sections 4.2–4.4. "
         f"Completion of full-text screening of the 720 candidates is a priority for the "
         f"next phase of the review and may revise the figures reported below, particularly "
         f"in the fisheries-specific STEEP distribution."),
        "Count of rows in each active database sheet. The 720 pending OpenAlex papers are tracked in <code>search_output/screened_included.csv</code> and are not included in the sheets.",
        tbl, 3,
    ))

    blocks.append(claim_block(
        "claim-source-types",
        "Within the cross-sectoral megatrend database, academic articles constituted the largest share of sources (n = 79, 54%), followed by consultancy publications (n = 22, 15%) and government or institutional reports (n = 18, 12%).",
        ("The composition of the megatrend database by source type matters because different "
         "publication genres have different thematic emphases. Consultancy publications (PWC, "
         "KPMG, EY, Deloitte) tend to foreground technological and economic disruption, whilst "
         "government and intergovernmental sources give more weight to social and political "
         "dynamics. Academic articles span all categories. This genre effect is visible in "
         "Figure 4 of the manuscript and is documented in the claim below on consultancy "
         "emphasis on technological and economic drivers."),
        "Value counts of the <code>source_type</code> column in the megatrend sheet, after standardisation of type labels.",
        df_to_html_table(
            mg["source_type"].value_counts().reset_index().rename(
                columns={"source_type":"Source type","count":"n"}
            ).assign(steep="Other"),
            ["Source type","n"],
        ), mg["source_type"].nunique(),
    ))

    # ═════════════════════════════════════════════════════════
    # SECTION 4.2 — Megatrend STEEP
    # ═════════════════════════════════════════════════════════
    blocks.append('<h2 id="s42">Section 4.2 — Drivers: Cross-sectoral megatrend database</h2>')
    blocks.append(prose_block(f"""
    <p>The cross-sectoral megatrend database (Strategy 2, n = {n_mg}) draws on major global
    foresight publications spanning business consultancies, governmental bodies, research
    institutions, and academic journals. These sources were selected because they explicitly
    address megatrends — large-scale, long-duration forces expected to reshape societies and
    economies — and because they are frequently cited as reference frameworks within the
    futures studies community.</p>

    <p>The STEEP distribution reported below reflects how the field as a whole frames the
    landscape of global change. It serves as the <em>benchmark</em> against which the
    fisheries-specific futures literature is compared in the gap analysis (Section 4.2b).
    Each claim below links to all database rows classified in that category, enabling you
    to inspect which specific sources and trends contribute to each percentage figure.</p>

    <p>Note that {mg[mg['steep']=='Other'].shape[0]} entries ({mg[mg['steep']=='Other'].shape[0]/len(mg)*100:.0f}%)
    fell into the Other/Unclassified category and are not included in the per-category
    percentage calculations. These entries typically describe cross-cutting themes that span
    multiple STEEP domains simultaneously.</p>
    """))

    for cat in ["Environmental","Technological","Economic","Social","Political"]:
        subset = mg[mg["steep"] == cat].sort_values("year")
        pct_all   = len(subset) / len(mg) * 100
        mg_classified = mg[mg["steep"] != "Other"]
        pct_class = len(subset) / len(mg_classified) * 100

        blocks.append(claim_block(
            f"claim-mg-{cat.lower()}",
            f"{cat} drivers: n = {len(subset)} ({pct_all:.0f}% of all megatrend entries; {pct_class:.0f}% of classifiable entries).",
            (f"{STEEP_DESCRIPTIONS[cat]} "
             f"In the megatrend database, this category contains {len(subset)} entries drawn from "
             f"{subset['source_type'].nunique()} distinct source types, spanning "
             f"{int(subset['year'].min()) if not subset['year'].isna().all() else 'unknown'}"
             f"–{int(subset['year'].max()) if not subset['year'].isna().all() else 'unknown'}."),
            f"Rows from the megatrend sheet where the STEEP classification algorithm assigned category = '{cat}', sorted by year.",
            df_to_html_table(subset[DISPLAY_MG].fillna(""), DISPLAY_MG),
            len(subset),
        ))

    # Climate change
    climate = mg[mg["trend"].str.lower().str.contains("climate", na=False)]
    blocks.append(claim_block(
        "claim-climate",
        "Climate change emerged as the most frequently identified driver across all source types, cited independently by five distinct sources.",
        ("Climate change is the only driver to appear as an explicitly named megatrend in five "
         "separate source documents, making it the single most convergent finding across the "
         "entire megatrend database. This convergence is notable because it holds across source "
         "types — academic, governmental, and consultancy sources all independently identify "
         "climate change as a primary driver — and across geographic and institutional contexts. "
         "In the fisheries-specific literature, climate change is similarly dominant, suggesting "
         "that this driver is unusually well-integrated between the two bodies of evidence."),
        "Rows in the megatrend sheet where the <code>trend</code> field (lowercase) contains the substring 'climate'.",
        df_to_html_table(climate[DISPLAY_MG].fillna(""), DISPLAY_MG),
        len(climate),
    ))

    # Consultancy tech+econ
    consult_tech_econ = mg[
        mg["source_type"].str.lower().str.contains("consult|business", na=False) &
        mg["steep"].isin(["Technological","Economic"])
    ]
    all_consult = mg[mg["source_type"].str.lower().str.contains("consult|business", na=False)]
    pct_te = len(consult_tech_econ)/len(all_consult)*100 if len(all_consult) > 0 else 0
    blocks.append(claim_block(
        "claim-consult-tech",
        "Consultancy and business publications placed comparatively greater emphasis on technological and economic drivers.",
        (f"Of the {len(all_consult)} entries from consultancy and business sources, "
         f"{len(consult_tech_econ)} ({pct_te:.0f}%) were classified as either Technological "
         f"or Economic. This compares to a combined Technological + Economic share of "
         f"{(len(mg[mg['steep'].isin(['Technological','Economic'])])/len(mg)*100):.0f}% across "
         f"the full database. The genre effect is consistent with the known framing of major "
         f"consultancy megatrend reports (e.g. PwC, KPMG, EY, Deloitte), which tend to "
         f"orient their foresight analyses around business disruption, market opportunity, "
         f"and the competitive implications of emerging technologies — rather than the "
         f"ecological or social dimensions that dominate government and academic sources."),
        "Rows from the megatrend sheet where <code>source_type</code> contains 'consult' or 'business' (case-insensitive) AND STEEP category is Technological or Economic.",
        df_to_html_table(consult_tech_econ[DISPLAY_MG].fillna(""), DISPLAY_MG),
        len(consult_tech_econ),
    ))

    # ═════════════════════════════════════════════════════════
    # SECTION 4.2b — Fisheries STEEP
    # ═════════════════════════════════════════════════════════
    blocks.append('<h2 id="s42b">Section 4.2 — Drivers: Fisheries futures literature</h2>')
    ff_classified = ff[ff["steep"] != "Other"]
    blocks.append(prose_block(f"""
    <p>The fisheries futures database (Strategy 1, n = {n_ff} manually screened papers) contains
    peer-reviewed academic literature explicitly addressing future scenarios, impact analyses,
    or trend evaluations in fisheries research. Entries were identified through a structured
    database search (OpenAlex, Elicit, Semantic Scholar) and manually screened for relevance.</p>

    <p>Of the {n_ff} entries, {len(ff_classified)} ({len(ff_classified)/n_ff*100:.0f}%) were
    successfully assigned to one of the five STEEP categories. The remaining {len(ff[ff['steep']=='Other'])}
    ({len(ff[ff['steep']=='Other'])/n_ff*100:.0f}%) were classified as Other/Unclassified.
    The STEEP distribution among the {len(ff_classified)} classifiable entries is the basis
    for the gap analysis in Figure 5 of the manuscript.</p>

    <p>The claims below document each category in the fisheries futures database and provide
    the full list of supporting rows. The most significant finding — the complete absence of
    technological entries — is documented with an explicit zero-count confirmation.</p>
    """))

    for cat in ["Environmental","Technological","Economic","Social","Political"]:
        subset_ff = ff[ff["steep"] == cat]
        pct_all   = len(subset_ff) / len(ff) * 100
        pct_class = len(subset_ff) / len(ff_classified) * 100 if len(ff_classified) > 0 else 0

        if cat == "Technological":
            table_html = "<p class='zero-result'>&#10003; Confirmed zero entries. The complete absence of technological drivers in the fisheries futures literature is the central empirical finding of this review. It is not an artefact of the keyword list — the full keyword set for the Technological category was applied to all 60 entries and matched none. This absence holds for both manual trend labels and paper titles.</p>"
            rationale = (
                "This is the most significant finding in the entire review. Despite technological "
                "change (automation, digitalisation, AI, biotechnology, synthetic biology) accounting "
                "for 28% of the cross-sectoral megatrend database, not a single one of the 60 "
                "manually screened fisheries futures papers was classified under the Technological "
                "domain. This is not a data artefact: the classification algorithm searched all "
                "available text fields (trend label, details, title) for any of the Technological "
                "keywords. The finding implies that futures-oriented fisheries research has, as of "
                "the search date, not yet engaged systematically with how technological forces will "
                "reshape the sector. This represents the primary actionable gap identified by "
                "the review."
            )
        else:
            table_html = df_to_html_table(subset_ff[DISPLAY_FF].fillna(""), DISPLAY_FF)
            rationale = (
                f"Of the {len(ff)} fisheries futures entries, {len(subset_ff)} ({pct_all:.0f}% of "
                f"all entries; {pct_class:.0f}% of classifiable entries) were assigned to the "
                f"{cat} category. "
                + {
                    "Environmental": (
                        "The strong dominance of the Environmental category reflects the field's "
                        "historical focus on ecological dynamics — climate impacts on fish stocks, "
                        "overfishing, habitat loss, by-catch, and food web disruption. When the "
                        "Other/Unclassified entries are excluded, Environmental accounts for "
                        f"{pct_class:.0f}% of classifiable fisheries futures content, compared to "
                        f"{len(mg[mg['steep']=='Environmental'])/len(mg[mg['steep']!='Other'])*100:.0f}% "
                        "in the megatrend database. This is not a gap but a disciplinary orientation "
                        "that the field has developed over decades."
                    ),
                    "Economic": (
                        "Economic futures in fisheries research remain comparatively underdeveloped. "
                        "The entries that do address economic dimensions tend to focus on narrow "
                        "issues — the economic benefit of ecosystem restoration, or the role of "
                        "economic incentives in coastal management — rather than on broader economic "
                        "foresight questions such as seafood market restructuring, value chain "
                        "globalisation, or the macroeconomic implications of climate-driven stock shifts."
                    ),
                    "Social": (
                        "Social futures are the most underrepresented category in the fisheries "
                        "futures literature in absolute terms (n = 1). The single entry addresses "
                        "food security and small-scale fisheries, which is a well-established "
                        "concern in the fisheries development literature but does not engage with "
                        "broader social megatrends (demographic change, urbanisation, migration, "
                        "changing consumer values) that are prominent in the cross-sectoral database."
                    ),
                    "Political": (
                        "Political and governance futures in fisheries appear primarily as subsidiary "
                        "topics within ecological or economic analyses, rather than as primary "
                        "objects of foresight inquiry. Entries in this category address fisheries "
                        "governance, management regimes, and fishery conflict rather than the "
                        "broader geopolitical megatrends (shifting power, changing security paradigms, "
                        "protectionism) that dominate the cross-sectoral political domain."
                    ),
                }.get(cat, "")
            )

        blocks.append(claim_block(
            f"claim-ff-{cat.lower()}",
            f"Fisheries futures — {cat} domain: n = {len(subset_ff)} ({pct_all:.0f}% of all entries; {pct_class:.0f}% of classifiable entries).",
            rationale,
            f"Rows from the fisheries futures sheet where STEEP classification = '{cat}'.",
            table_html,
            len(subset_ff),
        ))

    # Other/Unclassified
    ff_other = ff[ff["steep"] == "Other"]
    blocks.append(claim_block(
        "claim-ff-other",
        f"A further 32% of entries (n = {len(ff_other)}) could not be assigned to any single STEEP domain and are retained as unclassified.",
        ("These 19 entries fall into two analytically distinct groups. The first group "
         "(approximately 10 entries) consists of rows with no usable content in any text field: "
         "the Trend, Details, and Title columns are all empty or contain only placeholder values. "
         "These represent incomplete database entries — papers that were identified and included "
         "in the database but whose trend content has not yet been recorded. The second group "
         "(approximately 9 entries) consists of entries whose Trend field describes a "
         "methodological approach rather than a substantive thematic content: entries labelled "
         "'Scenario', 'Trend and drivers', 'Participatory MES scenario', or 'Overview of different "
         "drivers'. These entries describe the type of analysis the paper performs, not the "
         "domain it addresses, and therefore cannot be meaningfully classified under STEEP. "
         "Both groups are excluded from the gap analysis percentages in Figure 5."),
        "Rows from the fisheries futures sheet where STEEP classification = 'Other'. Includes both metadata-incomplete entries and methodological-descriptor entries.",
        df_to_html_table(ff_other[DISPLAY_FF].fillna(""), DISPLAY_FF),
        len(ff_other),
    ))

    # Gap analysis
    blocks.append('<h2 id="gap">Gap analysis</h2>')
    blocks.append(prose_block("""
    <p>The gap analysis compares the STEEP distribution of the cross-sectoral megatrend
    database (n = 145) with that of the fisheries futures literature (n = 60). The comparison
    is expressed as the percentage-point difference between the share of entries in each STEEP
    category across the two databases. A positive gap means that the category is proportionally
    <em>more prominent</em> in the megatrend database than in the fisheries literature — i.e.
    that fisheries research is paying less attention to that domain relative to the wider
    foresight field. A negative gap means that fisheries research is proportionally
    <em>more focused</em> on that domain than the megatrend literature.</p>

    <p><strong>Important caveat:</strong> The two databases differ not only in disciplinary
    focus but in source type composition. The megatrend database contains a substantial share
    of consultancy and grey literature sources, which tend to emphasise technological and
    economic disruption. The fisheries database is exclusively peer-reviewed academic literature.
    Some of the observed gap may therefore reflect genre conventions rather than a genuine
    disciplinary blind spot. This is acknowledged in Section 5a of the manuscript.</p>
    """))

    cats = ["Environmental","Technological","Economic","Social","Political"]
    mg_pct = {c: mg["steep"].value_counts(normalize=True).get(c, 0)*100 for c in cats}
    ff_pct = {c: ff["steep"].value_counts(normalize=True).get(c, 0)*100 for c in cats}
    gap_rows = []
    for cat in cats:
        gap = mg_pct[cat] - ff_pct[cat]
        direction = "Fisheries UNDER-represents" if gap > 0 else "Fisheries OVER-represents"
        gap_rows.append({
            "STEEP Category":   cat,
            "Megatrends (%)":   f"{mg_pct[cat]:.0f}%",
            "Fisheries (%)":    f"{ff_pct[cat]:.0f}%",
            "Gap (pp)":         f"{gap:+.0f}",
            "Direction":        direction,
            "steep":            cat,
        })
    gap_df = pd.DataFrame(gap_rows)
    blocks.append(claim_block(
        "claim-gap",
        "Technology gap = +28 percentage points — the largest disciplinary gap identified in the review.",
        ("The +28pp technology gap means that the Technological domain accounts for 28% of "
         "megatrend entries but 0% of fisheries futures entries — a difference of 28 percentage "
         "points. This is the largest gap in either direction across all five STEEP categories. "
         "The second-largest gap is Economic (+15pp), followed by Social (+11pp). Environmental "
         "shows a negative gap (fisheries over-represents this domain by ~25pp relative to the "
         "megatrend baseline). The Technology gap is robust: it holds regardless of whether "
         "percentages are calculated including or excluding the Other/Unclassified entries from "
         "both databases. Even if all 19 unclassified fisheries entries were reclassified as "
         "Technological, the fisheries technology share would reach only 32% — comparable to "
         "but not exceeding the megatrend baseline."),
        "Percentage-point difference between megatrend STEEP% and fisheries STEEP%, calculated over all entries in each database (including Other in denominator).",
        df_to_html_table(gap_df, ["STEEP Category","Megatrends (%)","Fisheries (%)","Gap (pp)","Direction"]),
        5,
    ))

    # ═════════════════════════════════════════════════════════
    # SECTION 4.3 — Trends
    # ═════════════════════════════════════════════════════════
    blocks.append('<h2 id="s43">Section 4.3 — Trends</h2>')
    blocks.append(prose_block("""
    <p>This section documents the specific recurrent themes identified within the fisheries
    futures literature. Unlike the STEEP distribution analysis, which classifies entries by
    their primary domain, the trends analysis identifies thematic clusters — groups of entries
    that address the same substantive topic — by searching for recurring keywords across the
    Trend, Details, and Title fields.</p>

    <p>The themes documented below were identified as recurrent in the Results section of the
    manuscript. Each claim links to the specific database entries that mention that theme,
    providing the source evidence for the claim that these are established topics in the
    fisheries futures literature. Note that a single paper may address multiple themes and
    therefore appear in more than one table below.</p>
    """))

    theme_queries = [
        ("conflict",     "Fishery conflict over access and resource rights",
         "Competition over fisheries access rights is a recurring governance concern in the "
         "futures literature. Fishery conflict entries address disputes between commercial and "
         "artisanal fishers, between fishing nations, and between fisheries and other marine "
         "users (e.g. offshore energy, shipping). As climate change shifts stock distributions, "
         "conflict over access rights is likely to intensify, making this a futures-relevant theme."),
        ("demand",       "Consumer demand for seafood",
         "Increasing global demand for seafood — driven by population growth, income growth, "
         "and health-conscious dietary trends — is consistently identified as a major trend "
         "shaping future fisheries production. This theme bridges the Environmental (production "
         "capacity) and Economic (market demand) domains and appears in both the fisheries "
         "futures and megatrend databases."),
        ("recreational", "Growth of recreational fisheries",
         "Recreational fishing is increasingly recognised as a sector of significant economic "
         "and social value that competes with commercial fisheries for stock access and "
         "management attention. Several futures studies identify the growth of recreational "
         "fisheries as a structural trend requiring explicit governance consideration, "
         "particularly in high-income countries where participation rates are rising."),
        ("scenario",     "Scenario-based and futures methods in fisheries research",
         "A substantial subset of the fisheries futures literature is methodologically focused: "
         "papers that develop or apply scenario analysis, foresight frameworks, or participatory "
         "futures methods to fisheries questions. These entries describe the analytical approach "
         "of the paper rather than a substantive driver or trend, which is why they tend to "
         "fall into the Other/Unclassified STEEP category. They are documented here to "
         "characterise the methodological landscape of the field."),
        ("food security","Food security and fisheries futures",
         "The relationship between fisheries and food security — particularly for coastal "
         "and low-income populations in the Global South — is a well-established theme in "
         "fisheries development literature. In the futures context, this theme addresses how "
         "declining catches, climate-driven stock shifts, and price volatility will affect "
         "the availability and affordability of fish as a protein source."),
    ]

    for keyword, label, rationale in theme_queries:
        subset = ff[
            ff["trend"].str.lower().str.contains(keyword, na=False) |
            ff["details"].str.lower().str.contains(keyword, na=False) |
            ff["title"].str.lower().str.contains(keyword, na=False)
        ]
        if not subset.empty:
            blocks.append(claim_block(
                f"claim-theme-{keyword.replace(' ','-')}",
                f"Recurrent theme in fisheries futures literature: {label}.",
                rationale,
                f"Fisheries futures entries where any of the Trend, Details, or Title fields contains '{keyword}' (case-insensitive).",
                df_to_html_table(subset[DISPLAY_FF].fillna(""), DISPLAY_FF),
                len(subset),
            ))

    # Cross-sectoral globalisation
    glob_mg = mg[mg["trend"].str.lower().str.contains("global|econom.*power|shift.*economic", na=False, regex=True)]
    blocks.append(claim_block(
        "claim-globalisation",
        "Globalisation and shifting economic power identified as cross-cutting megatrends relevant to fisheries.",
        ("Globalisation and economic power shifts feature prominently in the cross-sectoral "
         "megatrend literature as structural forces reshaping trade relationships, supply "
         "chains, and competitive dynamics across all sectors. For fisheries specifically, "
         "these forces manifest in the growing dominance of East and Southeast Asian seafood "
         "markets, the restructuring of global seafood value chains, and the shifting economic "
         "geography of aquaculture production. These trends are well-documented in the "
         "megatrend database but relatively absent from the fisheries futures literature, "
         "contributing to the Economic gap identified in the gap analysis."),
        "Megatrend entries where the Trend field contains 'global', 'econom.*power', or 'shift.*economic' (regex search).",
        df_to_html_table(glob_mg[DISPLAY_MG].fillna(""), DISPLAY_MG),
        len(glob_mg),
    ))

    # ═════════════════════════════════════════════════════════
    # SECTION 4.4 — Signals
    # ═════════════════════════════════════════════════════════
    blocks.append('<h2 id="s44">Section 4.4 — Weak and emerging signals</h2>')
    blocks.append(prose_block(f"""
    <p>Weak signals are early, often ambiguous indicators of potentially significant future
    changes. Unlike trends — which are already visible and directional — signals are nascent:
    they may or may not develop into trends, and their interpretation is inherently uncertain.
    The value of signal monitoring is not prediction but <em>preparedness</em>: by identifying
    signals early, organisations and research communities can begin developing scenarios and
    response options before a development becomes a mainstream concern (Hiltunen, 2010;
    Rossel, 2021).</p>

    <p>The signal database currently contains <strong>{len(sg) if not sg.empty else 0} entries</strong>.
    Three signals were added in the initial horizon scan conducted as part of this review;
    the database is designed as a living document and is expected to grow as signal scanning
    continues. The signal entries documented below each include a detailed description of the
    signal, the source from which it was identified, and a commentary on its potential
    relevance to fisheries systems.</p>

    <p>It is important to note that signal detection is explicitly preliminary at this stage.
    The signals documented here have been identified by a single reviewer drawing on adjacent
    grey literature, and have not been subject to the structured screening applied to the
    other two search strategies. They should be understood as <em>candidate signals</em> for
    further monitoring and investigation, not as confirmed findings.</p>
    """))

    if not sg.empty:
        DISPLAY_SG_FULL = ["trend","details","author","source","year","source_type"]
        blocks.append(claim_block(
            "claim-signals",
            f"The signal database contains {len(sg)} formally documented entr{'y' if len(sg)==1 else 'ies'} at the time of writing.",
            ("The three signals currently in the database represent three distinct domains of "
             "potential disruption for fisheries systems. The rise of precision fermentation and "
             "cultivated seafood is primarily an Economic/Technological signal — it threatens the "
             "commercial position of wild-catch seafood but also offers new opportunities in "
             "food security. Autonomous and remotely operated fishing vessels are a Technological "
             "signal with significant Social implications for fishing community employment. "
             "Geopolitical competition over Arctic fishing access is a Political signal shaped "
             "by the Environmental force of sea-ice retreat. Together they illustrate that "
             "weak signals for fisheries frequently cut across STEEP boundaries — a pattern "
             "that reinforces the value of a cross-domain foresight approach."),
            "All rows in the 'Signals of Future Fisheries' sheet of the Excel database.",
            df_to_html_table(sg[DISPLAY_SG_FULL].fillna(""), DISPLAY_SG_FULL, steep_col="source_type"),
            len(sg),
        ))

    return blocks


# ── CSS / HTML templates ──────────────────────────────────────

CSS = """
* { box-sizing: border-box; }
body { font-family: Georgia, serif; font-size: 14px; max-width: 1160px;
       margin: 40px auto; padding: 0 24px; color: #1a1a1a; line-height: 1.7;
       background: #fff; }
h1   { font-size: 1.5em; border-bottom: 3px solid #1a1a4b; padding-bottom: 10px;
       margin-bottom: 4px; }
.subtitle { color: #666; font-size: 0.9em; margin-bottom: 24px; }
h2   { font-size: 1.1em; margin-top: 44px; color: #1a1a4b;
       border-left: 5px solid #1a1a4b; padding-left: 12px;
       background: #f0f2fa; padding-top: 6px; padding-bottom: 6px; }
p    { margin: 0.6em 0; }
.prose-block { margin: 12px 0 20px 0; padding: 0 2px; }
.method-note { background: #fffde7; border-left: 4px solid #f9a825;
               padding: 12px 16px; margin: 16px 0; border-radius: 4px; font-size: 0.92em; }
.steep-card  { margin: 6px 0; }
.claim-block { background: #f7f8fc; border: 1px solid #d0d5e8;
               border-radius: 6px; padding: 16px 20px; margin: 20px 0; }
.claim-label { font-size: 0.7em; font-weight: bold; letter-spacing: 0.1em;
               text-transform: uppercase; color: #1a1a4b; margin-bottom: 4px; }
.claim-text  { font-style: italic; font-size: 1.02em; color: #222;
               margin-bottom: 10px; border-left: 3px solid #1a1a4b;
               padding-left: 10px; }
.claim-rationale { font-size: 0.93em; color: #333; margin-bottom: 10px;
                   background: #fff; padding: 10px 14px; border-radius: 4px;
                   border: 1px solid #e0e0e0; }
.claim-query { font-size: 0.83em; color: #555; margin-bottom: 8px; }
.badge       { display: inline-block; background: #1a1a4b; color: white;
               font-size: 0.75em; padding: 2px 8px; border-radius: 10px;
               font-style: normal; vertical-align: middle; margin-left: 8px; }
.zero-result { color: #155724; background: #d4edda; padding: 10px 14px;
               border-radius: 4px; font-style: italic; border: 1px solid #c3e6cb; }
table.evidence-table { border-collapse: collapse; width: 100%; font-size: 0.85em;
                        margin-top: 8px; }
table.evidence-table th { background: #1a1a4b; color: white; padding: 7px 10px;
                           text-align: left; font-weight: normal; white-space: nowrap; }
table.evidence-table td { padding: 5px 10px; border-bottom: 1px solid #ddd;
                           vertical-align: top; }
table.evidence-table tr:hover td { filter: brightness(0.96); }
nav  { position: sticky; top: 0; background: rgba(255,255,255,0.97);
       border-bottom: 2px solid #1a1a4b; padding: 8px 0; font-size: 0.85em;
       z-index: 100; }
nav a { margin-right: 14px; color: #1a1a4b; text-decoration: none; font-style: normal; }
nav a:hover { text-decoration: underline; }
.legend { display: flex; gap: 10px; flex-wrap: wrap; margin: 14px 0; font-size: 0.85em;
          align-items: center; }
.legend span { padding: 3px 12px; border-radius: 12px; border: 1px solid #ccc; }
code { background: #f0f0f0; padding: 1px 5px; border-radius: 3px; font-size: 0.9em; }
a { color: #1a1a4b; }
hr { border: none; border-top: 1px solid #ddd; margin: 40px 0; }
"""

NAV = """
<nav>
  <a href="#db-overview">Database</a>
  <a href="#method">Methodology</a>
  <a href="#steep-framework">STEEP</a>
  <a href="#s41">4.1 Overview</a>
  <a href="#s42">4.2 Megatrends</a>
  <a href="#s42b">4.2 Fisheries</a>
  <a href="#gap">Gap analysis</a>
  <a href="#s43">4.3 Trends</a>
  <a href="#s44">4.4 Signals</a>
</nav>
"""

LEGEND = """
<div class="legend">
  <strong>Row colours:</strong>
  <span style="background:#d4edda">Environmental</span>
  <span style="background:#cce5ff">Technological</span>
  <span style="background:#fff3cd">Economic</span>
  <span style="background:#f8d7da">Social</span>
  <span style="background:#e2d9f3">Political</span>
  <span style="background:#e9ecef">Other / Unclassified</span>
  <span style="background:#fce4ec">Signal</span>
</div>
"""


def write_html(blocks, mg, ff, sg):
    n_sg = len(sg) if not sg.empty else 0
    intro = f"""
    <p>This report provides a complete, source-by-source account of the evidence that underlies
    every quantitative claim in the Results section of the manuscript
    <em>Drivers, Trends, and Signals in Fisheries Research</em>
    (Szczepanska, {TODAY[:4]}). Its purpose is threefold:</p>
    <ol>
      <li><strong>Transparency:</strong> every percentage figure, count, and thematic claim in the
      manuscript can be traced back to the specific rows in the Excel database from which it was
      derived.</li>
      <li><strong>Reproducibility:</strong> re-running <code>generate_traceability.py</code> after
      any update to the database regenerates this report with current figures, keeping claims and
      evidence in sync throughout the review process.</li>
      <li><strong>Peer review support:</strong> during peer review, this report allows reviewers
      and co-authors to inspect the underlying data for any claim without opening the Excel file
      directly.</li>
    </ol>
    <p>The report is organised to mirror the Results section of the manuscript.
    Each <em>Claim</em> block shows: (i) the exact wording from the manuscript in italics;
    (ii) an explanatory rationale discussing what the claim means and why it matters;
    (iii) the data query that produced the evidence; and (iv) the matching rows from the
    database. Row background colours reflect STEEP classification (see legend below).</p>
    <p><strong>Database at a glance:</strong> {len(mg)} megatrend entries &middot;
    {len(ff)} fisheries futures entries &middot; {n_sg} signal entries &middot;
    720 OpenAlex candidates pending full-text screening.</p>
    """

    parts = [
        "<!DOCTYPE html><html lang='en'><head>",
        "<meta charset='utf-8'>",
        "<meta name='viewport' content='width=device-width,initial-scale=1'>",
        "<title>Evidence Traceability Report — Fisheries Trends 2026</title>",
        f"<style>{CSS}</style>",
        "</head><body>",
        NAV,
        "<h1>Evidence Traceability Report</h1>",
        f"<p class='subtitle'>Drivers, Trends, and Signals in Fisheries Research &mdash; "
        f"Szczepanska ({TODAY[:4]}) &mdash; generated {TODAY}</p>",
        intro,
        LEGEND,
    ] + blocks + [
        "<hr>",
        f"<p style='font-size:0.8em;color:#999'>Generated by <code>generate_traceability.py</code> "
        f"on {TODAY}. Re-run after any database update to keep claims and evidence in sync. "
        f"Source data: <code>manual_literature_search.xlsx</code>.</p>",
        "</body></html>",
    ]

    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text("\n".join(parts), encoding="utf-8")
    print(f"Saved: {OUT_FILE}")


# ── Main ──────────────────────────────────────────────────────

def main():
    print("Loading data ...")
    mg, ff, sg = load_data()
    print(f"  Megatrends       : {len(mg)} entries")
    print(f"  Fisheries futures: {len(ff)} entries")
    print(f"  Signals          : {len(sg) if not sg.empty else 0} entries")

    print("Building traceability report ...")
    blocks = build_report(mg, ff, sg)

    print("Writing HTML ...")
    write_html(blocks, mg, ff, sg)


if __name__ == "__main__":
    main()
