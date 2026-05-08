"""
============================================================
Systematic Literature Review — PRISMA 2020 Compliant
============================================================
Topic   : Drivers, Trends, and Signals in Fisheries Research
Author  : Timo Szczepanska
          UiT The Arctic University of Norway
          tsz000@uit.no
          https://orcid.org/0000-0003-2442-8223

Databases searched : OpenAlex, Semantic Scholar
Method             : PRISMA 2020 (Page et al., 2021)
                     https://doi.org/10.1136/bmj.n71

Search date        : auto-recorded at runtime
Output directory   : ./search_output/

References
----------
Page MJ, McKenzie JE, Bossuyt PM, et al. (2021). The PRISMA 2020
statement: an updated guideline for reporting systematic reviews.
BMJ, 372, n71. https://doi.org/10.1136/bmj.n71
============================================================
"""

import io
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path

# Force UTF-8 output on Windows to avoid cp1252 encoding errors
if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import pandas as pd
import requests
from rapidfuzz import fuzz

# ============================================================
# SECTION 1 — CONFIGURATION
# All search parameters are defined here for full transparency
# and reproducibility (PRISMA requirement).
# ============================================================

# Search queries exactly as listed in the Methods section of the paper.
# Each query is run independently on every database.
SEARCH_QUERIES = [
    "Future Trend Impact Analysis Fisheries",
    "Future studies Fisheries Research",
    "Future Fisheries Scenarios",
    "Future of Fisheries",
    "Future Trends and Drivers Fisheries",
    "Megatrend Impacts Fisheries",
    "Exploring Megatrends Fisheries",
]

# Inclusion criteria (PRISMA: eligibility criteria must be pre-specified)
INCLUSION = {
    # At least one of these must appear in title or abstract
    "fisheries_terms": [
        "fisheries", "fishery", "fishing", "fish stock",
        "aquaculture", "seafood", "marine resource",
    ],
    # At least one of these must appear in title or abstract
    "futures_terms": [
        "future", "trend", "driver", "scenario", "megatrend",
        "foresight", "horizon scanning", "weak signal", "forecast",
        "projection", "outlook", "emerging", "foresight",
    ],
    # Minimum publication year
    "min_year": 2000,
}

# Exclusion criteria (reasons recorded in PRISMA log)
EXCLUSION_REASONS = {
    "not_fisheries":        "Not related to fisheries or marine resources",
    "no_future_orientation": "No future-oriented perspective (retrospective only)",
    "population_only":      "Biological population dynamics only, no broader drivers/trends",
    "too_old":              "Published before year 2000",
    "no_abstract":          "No abstract available for screening",
}

# API settings
OPENALEX_EMAIL  = "tsz000@uit.no"   # used for OpenAlex polite pool
MAX_RESULTS_PER_QUERY = 300         # cap per query per database
SEMANTIC_SCHOLAR_LIMIT = 100        # S2 returns max 100 per request

# Semantic Scholar API key (optional).
# Leave as None to skip S2 and rely on OpenAlex only.
# Once you receive your key, paste it here and re-run the script.
# Register at: https://www.semanticscholar.org/product/api
SEMANTIC_SCHOLAR_API_KEY: str | None = None

# Output directory (one level up from analysis/ so it sits at project root)
OUTPUT_DIR = Path(__file__).parent.parent / "search_output"

# ============================================================
# SECTION 2 — SEARCH FUNCTIONS
# ============================================================

def search_openalex(query: str, max_results: int = MAX_RESULTS_PER_QUERY) -> list[dict]:
    """
    Query the OpenAlex Works endpoint.

    OpenAlex is a fully open, free index of ~250 million scholarly works
    (Priem et al., 2022; https://openalex.org). No authentication required;
    providing an email places requests in the 'polite pool' with higher limits.

    Returns
    -------
    list of raw OpenAlex work objects
    """
    base_url = "https://api.openalex.org/works"
    results = []
    page = 1
    per_page = min(100, max_results)

    while len(results) < max_results:
        params = {
            "search":    query,
            "per-page":  per_page,
            "page":      page,
            "mailto":    OPENALEX_EMAIL,
            "select": (
                "id,title,abstract_inverted_index,authorships,"
                "publication_year,primary_location,doi,keywords,concepts"
            ),
        }
        try:
            r = requests.get(base_url, params=params, timeout=30)
            r.raise_for_status()
        except requests.RequestException as e:
            print(f"    [OpenAlex] Request error on page {page}: {e}")
            break

        data = r.json()
        works = data.get("results", [])
        if not works:
            break

        results.extend(works)
        total_available = data.get("meta", {}).get("count", 0)

        if len(results) >= total_available or len(results) >= max_results:
            break

        page += 1
        time.sleep(0.15)  # polite delay

    return results[:max_results]


def _reconstruct_abstract(inverted_index: dict | None) -> str:
    """
    Reconstruct plain-text abstract from OpenAlex inverted-index format.

    OpenAlex stores abstracts as {word: [position, ...]} to reduce storage.
    This function reverses that to a readable string.
    """
    if not inverted_index:
        return ""
    word_positions = [
        (pos, word)
        for word, positions in inverted_index.items()
        for pos in positions
    ]
    word_positions.sort()
    return " ".join(word for _, word in word_positions)


def search_semantic_scholar(query: str, limit: int = SEMANTIC_SCHOLAR_LIMIT) -> list[dict]:
    """
    Query the Semantic Scholar Academic Graph API.

    Semantic Scholar (Kinney et al., 2023; https://www.semanticscholar.org)
    is a free scholarly search engine with a public API.
    Requires SEMANTIC_SCHOLAR_API_KEY to be set; skipped otherwise.
    Rate limit with key: 1 request/second.

    Returns
    -------
    list of raw Semantic Scholar paper objects, or [] if key not set
    """
    if SEMANTIC_SCHOLAR_API_KEY is None:
        return []

    base_url = "https://api.semanticscholar.org/graph/v1/paper/search"
    headers = {"x-api-key": SEMANTIC_SCHOLAR_API_KEY}
    params = {
        "query":  query,
        "limit":  min(limit, 100),
        "fields": (
            "title,abstract,authors,year,journal,"
            "externalIds,openAccessPdf,publicationTypes,fieldsOfStudy"
        ),
    }
    max_retries = 3
    for attempt in range(max_retries):
        try:
            r = requests.get(base_url, params=params, headers=headers, timeout=30)
            if r.status_code == 429:
                wait = 10 * (attempt + 1)
                print(f"\n    [Semantic Scholar] Rate limit hit. Waiting {wait}s ...", end=" ", flush=True)
                time.sleep(wait)
                continue
            r.raise_for_status()
            return r.json().get("data", [])
        except requests.RequestException as e:
            print(f"\n    [Semantic Scholar] Request error (attempt {attempt+1}): {e}")
            time.sleep(5)
    print("\n    [Semantic Scholar] Failed after retries - skipping query.")
    return []


# ============================================================
# SECTION 3 — NORMALIZATION
# Convert database-specific formats to a single common schema.
# ============================================================

COMMON_SCHEMA = [
    "source_db", "search_query", "title", "abstract",
    "authors", "year", "journal", "doi",
    "openalex_id", "ss_id",
    "screen_decision", "exclusion_reason",
]


def normalize_openalex(works: list[dict], query: str) -> list[dict]:
    """Normalize OpenAlex work objects to common schema."""
    rows = []
    for w in works:
        abstract = _reconstruct_abstract(w.get("abstract_inverted_index"))
        authors = [
            a.get("author", {}).get("display_name", "")
            for a in w.get("authorships", [])
        ]
        location = w.get("primary_location") or {}
        source   = location.get("source") or {}

        rows.append({
            "source_db":      "OpenAlex",
            "search_query":   query,
            "title":          w.get("title", ""),
            "abstract":       abstract,
            "authors":        "; ".join(filter(None, authors)),
            "year":           w.get("publication_year"),
            "journal":        source.get("display_name", ""),
            "doi":            (w.get("doi") or "").replace("https://doi.org/", ""),
            "openalex_id":    w.get("id", ""),
            "ss_id":          "",
            "screen_decision": "",
            "exclusion_reason": "",
        })
    return rows


def normalize_semantic_scholar(papers: list[dict], query: str) -> list[dict]:
    """Normalize Semantic Scholar paper objects to common schema."""
    rows = []
    for p in papers:
        authors = [a.get("name", "") for a in p.get("authors", [])]
        journal = p.get("journal") or {}

        rows.append({
            "source_db":      "Semantic Scholar",
            "search_query":   query,
            "title":          p.get("title", ""),
            "abstract":       p.get("abstract") or "",
            "authors":        "; ".join(filter(None, authors)),
            "year":           p.get("year"),
            "journal":        journal.get("name", "") if isinstance(journal, dict) else "",
            "doi":            (p.get("externalIds") or {}).get("DOI", ""),
            "openalex_id":    "",
            "ss_id":          p.get("paperId", ""),
            "screen_decision": "",
            "exclusion_reason": "",
        })
    return rows


# ============================================================
# SECTION 4 — DEDUPLICATION
# PRISMA requires reporting how many duplicates were removed.
# Two-pass strategy: exact DOI match, then fuzzy title match.
# ============================================================

def _normalise_title(title: str) -> str:
    """Strip punctuation/case for comparison."""
    if not title:
        return ""
    t = title.lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def deduplicate(df: pd.DataFrame, fuzzy_threshold: int = 92) -> tuple[pd.DataFrame, int]:
    """
    Remove duplicate records.

    Pass 1 : exact DOI match (case-insensitive)
    Pass 2 : fuzzy title similarity >= fuzzy_threshold (RapidFuzz partial ratio)

    Parameters
    ----------
    df               : combined results DataFrame
    fuzzy_threshold  : minimum similarity score (0–100) to call two titles duplicates

    Returns
    -------
    (deduplicated DataFrame, number of records removed)
    """
    n_before = len(df)

    # Pass 1 — DOI deduplication
    df["_doi_key"] = df["doi"].str.lower().str.strip()
    with_doi    = df[df["_doi_key"].notna() & (df["_doi_key"] != "")]
    without_doi = df[df["_doi_key"].isna()  | (df["_doi_key"] == "")]
    with_doi    = with_doi.drop_duplicates(subset=["_doi_key"], keep="first")
    df = pd.concat([with_doi, without_doi], ignore_index=True)

    # Pass 2 — fuzzy title deduplication
    # Use first-word bucketing to avoid O(n^2) comparisons across all pairs.
    df["_title_key"] = df["title"].apply(_normalise_title)
    keep = [True] * len(df)
    titles = df["_title_key"].tolist()

    # Group by first word for efficiency, then compare within each bucket
    from collections import defaultdict
    buckets: dict[str, list[int]] = defaultdict(list)
    for idx, t in enumerate(titles):
        first = t.split()[0] if t.split() else "__empty__"
        buckets[first].append(idx)

    for bucket_indices in buckets.values():
        for ii, i in enumerate(bucket_indices):
            if not keep[i]:
                continue
            for j in bucket_indices[ii + 1:]:
                if not keep[j]:
                    continue
                if titles[i] and titles[j]:
                    score = fuzz.ratio(titles[i], titles[j])
                    if score >= fuzzy_threshold:
                        keep[j] = False

    df = df[keep].reset_index(drop=True)
    df = df.drop(columns=["_doi_key", "_title_key"], errors="ignore")

    return df, n_before - len(df)


# ============================================================
# SECTION 5 — TITLE / ABSTRACT SCREENING
# Applies pre-specified inclusion and exclusion criteria.
# Each excluded record receives a documented reason.
# ============================================================

def screen_record(row: pd.Series) -> tuple[str, str]:
    """
    Apply inclusion/exclusion criteria to a single record.

    Returns
    -------
    (decision, reason)  where decision is 'included' or 'excluded'
    """
    title    = str(row.get("title", ""))
    abstract = str(row.get("abstract", ""))
    text     = (title + " " + abstract).lower()

    # Missing abstract — flag for manual review (not auto-excluded)
    if not abstract.strip() or abstract.strip() == "nan":
        # Still check title alone
        text = title.lower()

    # Year check
    year = row.get("year")
    if year:
        try:
            if int(year) < INCLUSION["min_year"]:
                return "excluded", EXCLUSION_REASONS["too_old"]
        except (ValueError, TypeError):
            pass

    # Mandatory: fisheries-related
    if not any(t in text for t in INCLUSION["fisheries_terms"]):
        return "excluded", EXCLUSION_REASONS["not_fisheries"]

    # Mandatory: future-oriented
    if not any(t in text for t in INCLUSION["futures_terms"]):
        return "excluded", EXCLUSION_REASONS["no_future_orientation"]

    return "included", ""


def apply_screening(df: pd.DataFrame) -> pd.DataFrame:
    """Apply screen_record to every row and record the decision."""
    decisions = df.apply(screen_record, axis=1, result_type="expand")
    decisions.columns = ["screen_decision", "exclusion_reason"]
    df[["screen_decision", "exclusion_reason"]] = decisions
    return df


# ============================================================
# SECTION 6 — PRISMA 2020 FLOW DIAGRAM
# Follows the layout from Page et al. (2021), Figure 1.
# ============================================================

def draw_prisma_diagram(counts: dict, output_path: Path) -> None:
    """
    Generate a PRISMA 2020 flow diagram and save to PNG.

    Parameters
    ----------
    counts : dict with keys:
        n_openalex          — records from OpenAlex (before dedup)
        n_semanticscholar   — records from Semantic Scholar (before dedup)
        n_existing_excel    — records from existing manual search
        n_total_identified  — sum of above
        n_duplicates        — duplicates removed
        n_after_dedup       — records after deduplication
        n_screened          — records screened (title/abstract)
        n_excluded_screen   — records excluded at screening
        exclusion_breakdown — dict {reason: count}
        n_included          — records included after screening
    output_path : Path to save the PNG
    """
    fig, ax = plt.subplots(figsize=(13, 18))
    ax.set_xlim(0, 13)
    ax.set_ylim(0, 22)
    ax.axis("off")
    ax.set_facecolor("white")
    fig.patch.set_facecolor("white")

    BLUE  = "#2166ac"
    LBLUE = "#d1e5f0"
    GREY  = "#f7f7f7"
    SIDE  = "#fddbc7"

    def box(x, y, w, h, text, bg=LBLUE, fontsize=8.5, bold=False):
        rect = mpatches.FancyBboxPatch(
            (x - w / 2, y - h / 2), w, h,
            boxstyle="round,pad=0.12",
            facecolor=bg, edgecolor=BLUE, linewidth=1.4,
        )
        ax.add_patch(rect)
        ax.text(
            x, y, text,
            ha="center", va="center", fontsize=fontsize,
            wrap=True, multialignment="center",
            fontweight="bold" if bold else "normal",
        )

    def side_box(x, y, w, h, text, bg=SIDE, fontsize=8):
        rect = mpatches.FancyBboxPatch(
            (x - w / 2, y - h / 2), w, h,
            boxstyle="round,pad=0.12",
            facecolor=bg, edgecolor="#b2182b", linewidth=1.2,
        )
        ax.add_patch(rect)
        ax.text(x, y, text, ha="center", va="center", fontsize=fontsize,
                multialignment="center")

    def arrow(x1, y1, x2, y2):
        ax.annotate(
            "", xy=(x2, y2), xytext=(x1, y1),
            arrowprops=dict(arrowstyle="->", color=BLUE, lw=1.5),
        )

    def horiz_arrow(x1, y, x2):
        ax.annotate(
            "", xy=(x2, y), xytext=(x1, y),
            arrowprops=dict(arrowstyle="->", color="#b2182b", lw=1.2),
        )

    def section_label(y, text):
        ax.text(
            0.15, y, text,
            ha="left", va="center", fontsize=9,
            fontweight="bold", color=BLUE,
        )
        ax.plot([0.1, 0.1], [y - 0.8, y + 0.8], color=BLUE, lw=3)

    # ── Section labels ──────────────────────────────────────────
    section_label(20.5, "IDENTIFICATION")
    section_label(14.5, "SCREENING")
    section_label(8.5,  "ELIGIBILITY\n(full-text; manual)")
    section_label(3.2,  "INCLUDED")

    # ── Identification boxes ────────────────────────────────────
    id_text = (
        f"Records identified from databases\n"
        f"  OpenAlex            (n = {counts['n_openalex']:,})\n"
        f"  Semantic Scholar    (n = {counts['n_semanticscholar']:,})\n"
        f"Total                 (n = {counts['n_openalex'] + counts['n_semanticscholar']:,})"
    )
    box(4.5, 20.5, 6.5, 1.6, id_text, fontsize=8.5)

    other_text = (
        f"Records from existing\nmanual search (Excel)\n"
        f"(n = {counts['n_existing_excel']:,})"
    )
    box(10.5, 20.5, 3.5, 1.6, other_text, fontsize=8.5)

    arrow(4.5, 19.7, 4.5, 18.7)
    arrow(10.5, 19.7, 10.5, 18.7)
    # merge arrows down
    ax.plot([4.5, 4.5, 7.0], [18.7, 18.3, 18.3], color=BLUE, lw=1.5)
    ax.plot([10.5, 10.5, 7.0], [18.7, 18.3, 18.3], color=BLUE, lw=1.5)
    arrow(7.0, 18.3, 7.0, 17.6)

    # Duplicates removed (side box)
    dup_text = (
        f"Duplicate records removed\n(n = {counts['n_duplicates']:,})\n"
        f"(DOI match + fuzzy title match)"
    )
    side_box(11.2, 17.1, 3.3, 1.2, dup_text, fontsize=7.8)
    horiz_arrow(8.5, 17.1, 9.55)

    # After dedup
    dedup_text = f"Records after duplicates removed\n(n = {counts['n_after_dedup']:,})"
    box(7.0, 17.1, 4.5, 1.0, dedup_text)

    arrow(7.0, 16.6, 7.0, 15.5)

    # ── Screening ────────────────────────────────────────────────
    screen_text = f"Records screened\n(title and abstract)\n(n = {counts['n_screened']:,})"
    box(7.0, 15.0, 4.5, 1.1, screen_text)

    # Excluded (side box)
    breakdown = counts.get("exclusion_breakdown", {})
    excl_lines = [f"Records excluded  (n = {counts['n_excluded_screen']:,})"]
    for reason, n in sorted(breakdown.items(), key=lambda x: -x[1]):
        short = reason[:45] + "…" if len(reason) > 45 else reason
        excl_lines.append(f"  • {short}: {n}")
    excl_text = "\n".join(excl_lines)
    side_box(11.2, 14.1, 3.3, 1.8, excl_text, fontsize=7.2)
    horiz_arrow(9.25, 14.1, 9.55)

    arrow(7.0, 14.45, 7.0, 13.1)

    # ── Eligibility (manual full-text review) ───────────────────
    eligib_text = (
        f"Reports assessed for eligibility\n"
        f"(full-text; manual review)\n"
        f"(n = {counts['n_included']:,})"
    )
    box(7.0, 12.6, 4.5, 1.1, eligib_text)

    manual_excl = (
        "Reports excluded after\nfull-text review\n(n = TBD — manual step)\n"
        "Reasons:\n  • Out of scope\n  • Duplicate concept\n  • Quality"
    )
    side_box(11.2, 11.5, 3.3, 2.0, manual_excl, fontsize=7.2)
    horiz_arrow(9.25, 11.5, 9.55)

    arrow(7.0, 12.05, 7.0, 10.5)

    # ── Included ─────────────────────────────────────────────────
    incl_text = (
        f"Studies included in review\n"
        f"(after automated screening)\n"
        f"(n = {counts['n_included']:,})\n"
        f"⚠ Requires manual full-text review"
    )
    box(7.0, 9.9, 5.0, 1.3, incl_text, bg="#c7e9b4", fontsize=8.5)

    # ── Caption ──────────────────────────────────────────────────
    caption = (
        "Figure: PRISMA 2020 flow diagram (Page et al., 2021).\n"
        "Automated screening applies pre-specified inclusion/exclusion criteria to\n"
        "title and abstract. Full-text eligibility requires manual review."
    )
    ax.text(6.5, 0.8, caption, ha="center", va="bottom", fontsize=7.5,
            style="italic", color="#555555")

    plt.tight_layout(pad=0.5)
    plt.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"  PRISMA diagram saved → {output_path}")


# ============================================================
# SECTION 7 — EXPORT & DOCUMENTATION
# ============================================================

def export_results(
    df_raw: pd.DataFrame,
    df_dedup: pd.DataFrame,
    df_screened: pd.DataFrame,
    search_log: list[dict],
    counts: dict,
    output_dir: Path,
) -> None:
    """
    Save all outputs to the output directory.

    Files produced
    --------------
    raw_results.csv          — all records before deduplication
    deduplicated.csv         — records after deduplication
    screened_included.csv    — records passing title/abstract screening
    screened_excluded.csv    — records failing screening (with reasons)
    search_log.json          — full API call log (query, db, date, n_results)
    prisma_counts.json       — counts for PRISMA diagram
    prisma_diagram.png       — PRISMA 2020 flow diagram (300 dpi)
    search_documentation.xlsx — human-readable summary for the paper's Methods section
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    df_raw.to_csv(output_dir / "raw_results.csv", index=False)
    df_dedup.to_csv(output_dir / "deduplicated.csv", index=False)

    included = df_screened[df_screened["screen_decision"] == "included"]
    excluded = df_screened[df_screened["screen_decision"] == "excluded"]
    included.to_csv(output_dir / "screened_included.csv", index=False)
    excluded.to_csv(output_dir / "screened_excluded.csv", index=False)

    with open(output_dir / "search_log.json", "w", encoding="utf-8") as f:
        json.dump(search_log, f, indent=2, default=str)

    with open(output_dir / "prisma_counts.json", "w", encoding="utf-8") as f:
        json.dump(counts, f, indent=2)

    draw_prisma_diagram(counts, output_dir / "prisma_diagram.png")

    # Search documentation table — suitable for Methods section
    doc_rows = []
    for entry in search_log:
        doc_rows.append({
            "Database":         entry["database"],
            "Search Query":     entry["query"],
            "Search Date":      entry["date"],
            "Records Retrieved":entry["n_results"],
            "API Endpoint":     entry["endpoint"],
        })
    df_doc = pd.DataFrame(doc_rows)

    criteria_rows = [
        {"Category": "Inclusion", "Criterion": "Fisheries-related terms in title/abstract",
         "Terms": ", ".join(INCLUSION["fisheries_terms"])},
        {"Category": "Inclusion", "Criterion": "Future-oriented terms in title/abstract",
         "Terms": ", ".join(INCLUSION["futures_terms"])},
        {"Category": "Inclusion", "Criterion": "Publication year",
         "Terms": f">= {INCLUSION['min_year']}"},
        {"Category": "Exclusion", "Criterion": "Not fisheries-related", "Terms": "—"},
        {"Category": "Exclusion", "Criterion": "No future orientation", "Terms": "—"},
        {"Category": "Exclusion", "Criterion": "Population dynamics only", "Terms": "—"},
        {"Category": "Exclusion", "Criterion": "Published before 2000", "Terms": "—"},
    ]
    df_criteria = pd.DataFrame(criteria_rows)

    prisma_rows = [
        {"Stage": "Identification — OpenAlex",        "n": counts["n_openalex"]},
        {"Stage": "Identification — Semantic Scholar", "n": counts["n_semanticscholar"]},
        {"Stage": "Identification — Existing Excel",   "n": counts["n_existing_excel"]},
        {"Stage": "Total identified",                  "n": counts["n_total_identified"]},
        {"Stage": "Duplicates removed",                "n": counts["n_duplicates"]},
        {"Stage": "After deduplication",               "n": counts["n_after_dedup"]},
        {"Stage": "Screened (title/abstract)",         "n": counts["n_screened"]},
        {"Stage": "Excluded at screening",             "n": counts["n_excluded_screen"]},
        {"Stage": "Included (automated screening)",    "n": counts["n_included"]},
    ]
    df_prisma = pd.DataFrame(prisma_rows)

    with pd.ExcelWriter(output_dir / "search_documentation.xlsx", engine="openpyxl") as writer:
        df_doc.to_excel(writer,      sheet_name="Search Log",         index=False)
        df_criteria.to_excel(writer, sheet_name="Eligibility Criteria", index=False)
        df_prisma.to_excel(writer,   sheet_name="PRISMA Counts",       index=False)
        included.to_excel(writer,    sheet_name="Included Papers",     index=False)
        excluded.to_excel(writer,    sheet_name="Excluded Papers",     index=False)

    print(f"\n  All outputs saved to: {output_dir.resolve()}")
    print(f"  Included papers : {len(included):,}")
    print(f"  Excluded papers : {len(excluded):,}")


# ============================================================
# SECTION 8 — MAIN EXECUTION
# ============================================================

def main() -> None:
    search_date = datetime.now().strftime("%Y-%m-%d %H:%M")
    print("=" * 62)
    print("Systematic Literature Search — PRISMA 2020")
    print(f"Search date : {search_date}")
    print(f"Databases   : OpenAlex, Semantic Scholar")
    print(f"Queries     : {len(SEARCH_QUERIES)}")
    print("=" * 62)

    all_records: list[dict] = []
    search_log:  list[dict] = []

    n_openalex = 0
    n_ss       = 0

    for query in SEARCH_QUERIES:
        print(f"\nQuery: \"{query}\"")

        # — OpenAlex —
        print("  Searching OpenAlex …", end=" ", flush=True)
        oa_raw = search_openalex(query)
        oa_norm = normalize_openalex(oa_raw, query)
        n_openalex += len(oa_norm)
        all_records.extend(oa_norm)
        print(f"{len(oa_norm)} records")

        search_log.append({
            "database":  "OpenAlex",
            "query":     query,
            "date":      search_date,
            "n_results": len(oa_norm),
            "endpoint":  "https://api.openalex.org/works",
        })

        time.sleep(0.3)

        # — Semantic Scholar —
        if SEMANTIC_SCHOLAR_API_KEY is None:
            print("  Semantic Scholar: skipped (no API key set)")
        else:
            print("  Searching Semantic Scholar …", end=" ", flush=True)
            ss_raw = search_semantic_scholar(query)
            ss_norm = normalize_semantic_scholar(ss_raw, query)
            n_ss += len(ss_norm)
            all_records.extend(ss_norm)
            print(f"{len(ss_norm)} records")
            search_log.append({
                "database":  "Semantic Scholar",
                "query":     query,
                "date":      search_date,
                "n_results": len(ss_norm),
                "endpoint":  "https://api.semanticscholar.org/graph/v1/paper/search",
            })

        time.sleep(4.0)  # respect Semantic Scholar rate limit (~100 req/5min)

    # ── Build combined DataFrame ─────────────────────────────────
    df_raw = pd.DataFrame(all_records, columns=COMMON_SCHEMA)
    print(f"\nTotal records before deduplication : {len(df_raw):,}")

    # ── Deduplication ────────────────────────────────────────────
    print("Deduplicating …", end=" ", flush=True)
    df_dedup, n_dupes = deduplicate(df_raw.copy())
    print(f"{n_dupes:,} duplicates removed ->{len(df_dedup):,} unique records")

    # ── Screening ────────────────────────────────────────────────
    print("Applying inclusion/exclusion criteria …", end=" ", flush=True)
    df_screened = apply_screening(df_dedup.copy())
    n_included = (df_screened["screen_decision"] == "included").sum()
    n_excluded = (df_screened["screen_decision"] == "excluded").sum()
    print(f"included={n_included:,}, excluded={n_excluded:,}")

    exclusion_breakdown = (
        df_screened[df_screened["screen_decision"] == "excluded"]
        ["exclusion_reason"]
        .value_counts()
        .to_dict()
    )

    # ── Existing Excel count (records from the manual search) ────
    try:
        xl = pd.read_excel(
            Path(__file__).parent / "manual_literature_search.xlsx",
            sheet_name=None,
        )
        n_existing = sum(
            max(0, len(sheet) - 1)   # subtract header row
            for name, sheet in xl.items()
            if name not in ("Glossary", "Info", "Signals of Future Fisheries",
                            "Eu project (Fish and drivers)")
        )
    except Exception:
        n_existing = 0

    # ── PRISMA counts ────────────────────────────────────────────
    counts = {
        "n_openalex":          n_openalex,
        "n_semanticscholar":   n_ss,
        "n_existing_excel":    n_existing,
        "n_total_identified":  n_openalex + n_ss + n_existing,
        "n_duplicates":        n_dupes,
        "n_after_dedup":       len(df_dedup),
        "n_screened":          len(df_screened),
        "n_excluded_screen":   int(n_excluded),
        "exclusion_breakdown": exclusion_breakdown,
        "n_included":          int(n_included),
        "search_date":         search_date,
    }

    # ── Summary ──────────────────────────────────────────────────
    print("\n" + "=" * 62)
    print("PRISMA 2020 Summary")
    print("=" * 62)
    print(f"  Records identified (OpenAlex)          : {n_openalex:>6,}")
    print(f"  Records identified (Semantic Scholar)  : {n_ss:>6,}")
    print(f"  Records from existing manual search    : {n_existing:>6,}")
    print(f"  Total identified                       : {n_openalex + n_ss + n_existing:>6,}")
    print(f"  Duplicates removed                     : {n_dupes:>6,}")
    print(f"  After deduplication                    : {len(df_dedup):>6,}")
    print(f"  Excluded at title/abstract screening   : {n_excluded:>6,}")
    print(f"  Included (automated screening)         : {n_included:>6,}")
    print("-" * 62)
    print("  Exclusion breakdown:")
    for reason, n in sorted(exclusion_breakdown.items(), key=lambda x: -x[1]):
        print(f"    {reason:<45} : {n:>5,}")
    print("=" * 62)
    print("\n⚠  Automated screening is a first filter only.")
    print("   Full-text eligibility assessment requires manual review.")

    # ── Export ───────────────────────────────────────────────────
    print("\nExporting results …")
    export_results(df_raw, df_dedup, df_screened, search_log, counts, OUTPUT_DIR)

    print("\nDone. Next step: open search_output/screened_included.csv")
    print("and manually review full texts against eligibility criteria.")


if __name__ == "__main__":
    main()
