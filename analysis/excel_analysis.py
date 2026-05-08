"""
============================================================
Excel Database Analysis — Drivers, Trends, and Signals
============================================================
Topic   : Drivers, Trends, and Signals in Fisheries Research
Author  : Timo Szczepanska
          UiT The Arctic University of Norway
          tsz000@uit.no

Input   : ../data/manual_literature_search.xlsx
Output  : ../search_output/figures/   (publication-ready figures)
          ../search_output/excel_analysis.xlsx  (cleaned datasets)

This script analyses the three search strategy outputs stored in the
Excel database:
  Sheet 1 — (Mega-)trends and drivers  : cross-sectoral foresight (n=149)
  Sheet 2 — Futures of Fisheries       : fisheries futures literature (n~47)
  Sheet 3 — Signals of Future Fisheries: weak signals (n=1, nascent)

Thematic classification follows the STEEP framework
(Social, Technological, Economic, Environmental, Political/Governance),
which is standard in horizon scanning and foresight literature.
============================================================
"""

from pathlib import Path
import re

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd
import seaborn as sns

# ── Paths ────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent.parent
DATA_FILE  = BASE_DIR / "data" / "manual_literature_search.xlsx"
OUTPUT_DIR = BASE_DIR / "search_output"
FIG_DIR    = OUTPUT_DIR / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

# ── Plot style ───────────────────────────────────────────────
plt.rcParams.update({
    "font.family":     "serif",
    "font.size":       10,
    "axes.titlesize":  11,
    "axes.labelsize":  10,
    "figure.dpi":      150,
    "savefig.dpi":     300,
    "savefig.bbox":    "tight",
})
PALETTE = {
    "Environmental": "#2166ac",
    "Social":        "#d73027",
    "Technological": "#1a9641",
    "Economic":      "#f46d43",
    "Political":     "#762a83",
    "Other":         "#999999",
}

# ============================================================
# SECTION 1 — STEEP CLASSIFICATION
# Keywords used to assign each trend to a STEEP category.
# A trend is assigned to the FIRST matching category.
# ============================================================

STEEP_RULES = {
    "Environmental": [
        "climate", "environment", "biodiversity", "ecosystem", "ocean",
        "marine", "sea", "water", "pollution", "habitat", "carbon",
        "emission", "green", "nature conservation", "nature-based",
        "natural resource", "natural environment", "resource scarcity", "erosion",
        "deforestation", "renewable energy", "sustainability", "ecology",
        "fish stock", "overfishing", "fishing", "fisheries", "fishery",
        "fisher", "fishers", "aquaculture", "aquatic", "seafood", "harvest",
        "catch", "bycatch", "by-catch", "food web", "carrying capacity",
        "stock assessment", "spawn", "recruitment", "trophic",
    ],
    "Technological": [
        "technolog", "digital", "automation", "artificial intelligence",
        "ai ", "data", "innovation", "internet", "cyber", "robot",
        "autonomous", "biotech", "synthetic biology", "microbiome",
        "hyperconnectiv", "electrif", "knowledge",
    ],
    "Economic": [
        "economic", "trade", "market", "finance", "fiscal", "gdp",
        "consumption", "globali", "growth", "debt", "income",
        "employment", "labour", "labor", "industry",
        "investment", "supply chain", "demand", "profit", "revenue",
        "subsid", "cost", "price", "value chain",
    ],
    "Social": [
        "social", "demograph", "population", "urban", "aging", "health",
        "inequalit", "migration", "culture", "education", "wellbeing",
        "well-being", "individual", "community", "consumer",
        "lifestyle", "food security", "gender", "governance of people",
        "livelihood", "artisanal", "small-scale", "indigenous",
    ],
    "Political": [
        "politic", "governance", "geopolit", "security", "power",
        "regulation", "law", "policy", "democracy", "conflict",
        "supranational", "institution", "global order", "protectionism",
        "management", "quota", "treaty", "iuu", "compliance",
    ],
}


def classify_steep(text: str) -> str:
    """Assign a STEEP category to a text string."""
    if not text or str(text).strip().lower() in ("none", "nan", "?", "n=a", ""):
        return "Other"
    t = str(text).lower()
    for category, keywords in STEEP_RULES.items():
        if any(kw in t for kw in keywords):
            return category
    return "Other"


# ============================================================
# SECTION 2 — LOAD AND CLEAN DATA
# ============================================================

def load_megatrends() -> pd.DataFrame:
    """
    Load and clean the '(Mega-)trends and drivers' sheet.

    Standardises source types, removes empty rows, assigns STEEP.
    """
    df = pd.read_excel(DATA_FILE, sheet_name="(Mega-)trends and drivers")
    df.columns = ["trend", "details", "author", "source", "title",
                  "comments", "doi", "year", "source_type"]
    df = df.dropna(subset=["trend"])
    df = df[~df["trend"].astype(str).str.strip().isin(["", "nan"])]

    # Standardise source types
    type_map = {
        "Government? (finnish innovation fund)": "Government",
        "Oxfam discussion paper":                "NGO / Think Tank",
        "Book Chapter":                          "Book / Chapter",
        "Book":                                  "Book / Chapter",
        "?":                                     "Unclassified",
    }
    df["source_type"] = df["source_type"].replace(type_map)

    # STEEP classification using trend name + details
    df["steep"] = df.apply(
        lambda r: classify_steep(str(r["trend"]) + " " + str(r["details"])), axis=1
    )

    # Clean year (some may be strings or missing)
    df["year"] = pd.to_numeric(df["year"], errors="coerce")

    return df.reset_index(drop=True)


def load_fisheries_futures() -> pd.DataFrame:
    """
    Load and clean the 'Futures of Fisheries' sheet.

    Strategy: keep any row that has a real title or author (= a real
    literature entry), regardless of whether the trend field is filled.
    Remove only rows that are clearly note-taking placeholders with no
    associated reference (no title, no author, no source URL).
    """
    df = pd.read_excel(DATA_FILE, sheet_name="Futures of Fisheries")
    df.columns = ["trend", "details", "author", "source", "title",
                  "comments", "doi", "year", "source_type", "_a", "_b"]
    df = df.drop(columns=["_a", "_b"])

    # Patterns that mark a TREND FIELD as a note, not a real trend label
    # (we still keep the row if it has a title/author)
    NOTE_PATTERNS = [
        r"^n=a$", r"^none$", r"^topic$", r"^keyword",
        r"^future\s*$", r"^political variables$", r"^economic variables$",
        r"^social variables$", r"^environmental variables$",
        r"^find something", r"^interesting article", r"^book on",
        r"^manifesto$", r"^brief vision", r"^bell prediction",
    ]

    def is_note_trend(val) -> bool:
        if pd.isna(val):
            return True
        v = str(val).strip().lower()
        return any(re.match(p, v) for p in NOTE_PATTERNS)

    def has_reference(row) -> bool:
        """True if the row has at least a title, author, or real source URL."""
        has_title  = pd.notna(row["title"])  and str(row["title"]).strip().lower() not in ("", "nan")
        has_author = pd.notna(row["author"]) and str(row["author"]).strip().lower() not in ("", "nan")
        has_source = (pd.notna(row["source"]) and
                      str(row["source"]).strip().lower() not in ("", "nan") and
                      str(row["source"]).strip().lower() != row["author"])
        return has_title or has_author or has_source

    # Keep a row if it has a reference, regardless of trend field
    # OR if the trend field has meaningful content even without a full reference
    df["_has_ref"]      = df.apply(has_reference, axis=1)
    df["_note_trend"]   = df["trend"].apply(is_note_trend)

    df = df[df["_has_ref"] | (~df["_note_trend"])]
    df = df.drop(columns=["_has_ref", "_note_trend"])

    # For rows where trend is '?' or a placeholder, replace with NaN
    # so it doesn't mislead STEEP classification
    def clean_trend(val):
        if pd.isna(val):
            return None
        v = str(val).strip()
        if v in ("?", "N=A", "None") or is_note_trend(val):
            return None
        return v

    df["trend"] = df["trend"].apply(clean_trend)

    # STEEP: classify on trend + details + title (all available text)
    df["steep"] = df.apply(
        lambda r: classify_steep(
            " ".join(str(r[c]) for c in ["trend", "details", "title"]
                     if pd.notna(r[c]))
        ), axis=1
    )

    df["year"] = pd.to_numeric(df["year"], errors="coerce")

    return df.reset_index(drop=True)


# ============================================================
# SECTION 3 — ANALYSIS FUNCTIONS
# ============================================================

def steep_frequency_table(df: pd.DataFrame, source_col: str = "source_type") -> pd.DataFrame:
    """Cross-tabulate STEEP categories by source type."""
    return pd.crosstab(df["steep"], df[source_col])


def deduplicate_by_theme(df: pd.DataFrame) -> pd.DataFrame:
    """
    Identify thematically duplicate entries (same STEEP + very similar trend name).
    Returns df with a 'canonical_theme' column grouping near-duplicates.
    """
    from rapidfuzz import fuzz

    df = df.copy()
    df["canonical_theme"] = df["trend"].astype(str)

    groups: dict[str, list[int]] = {}
    for i, row in df.iterrows():
        placed = False
        t = str(row["trend"]).lower().strip()
        for canon, indices in groups.items():
            if fuzz.ratio(t, canon) >= 80:
                groups[canon].append(i)
                df.at[i, "canonical_theme"] = canon
                placed = True
                break
        if not placed:
            groups[t] = [i]
            df.at[i, "canonical_theme"] = t
    return df


# ============================================================
# SECTION 4 — FIGURES
# ============================================================

def fig_steep_distribution(df_mega: pd.DataFrame, df_fish: pd.DataFrame) -> None:
    """
    Figure 1: STEEP category distribution for both datasets side by side.
    Suitable for the Results section overview.
    """
    cats = list(PALETTE.keys())
    mega_counts = df_mega["steep"].value_counts().reindex(cats, fill_value=0)
    fish_counts = df_fish["steep"].value_counts().reindex(cats, fill_value=0)

    x = np.arange(len(cats))
    width = 0.35

    fig, ax = plt.subplots(figsize=(9, 5))
    bars1 = ax.bar(x - width/2, mega_counts, width,
                   label="Cross-sectoral megatrends\n(Search strategy 2)",
                   color=[PALETTE[c] for c in cats], alpha=0.85)
    bars2 = ax.bar(x + width/2, fish_counts, width,
                   label="Fisheries futures literature\n(Search strategy 1)",
                   color=[PALETTE[c] for c in cats], alpha=0.5,
                   edgecolor=[PALETTE[c] for c in cats], linewidth=1.5)

    ax.set_xticks(x)
    ax.set_xticklabels(cats, rotation=20, ha="right")
    ax.set_ylabel("Number of entries")
    ax.set_title("Distribution of entries by STEEP category")
    ax.legend(frameon=True, loc="upper right")
    ax.spines[["top", "right"]].set_visible(False)

    # Value labels
    for bar in bars1:
        if bar.get_height() > 0:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                    str(int(bar.get_height())), ha="center", va="bottom", fontsize=8)
    for bar in bars2:
        if bar.get_height() > 0:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                    str(int(bar.get_height())), ha="center", va="bottom", fontsize=8)

    plt.tight_layout()
    plt.savefig(FIG_DIR / "fig1_steep_distribution.png")
    plt.close()
    print("  Saved: fig1_steep_distribution.png")


def fig_source_type_breakdown(df_mega: pd.DataFrame) -> None:
    """
    Figure 2: Source type composition of the cross-sectoral megatrend database.
    """
    counts = df_mega["source_type"].value_counts()
    colors = ["#2166ac", "#4dac26", "#d01c8b", "#f1b6da", "#b8e186", "#999999"]

    fig, ax = plt.subplots(figsize=(7, 5))
    wedges, texts, autotexts = ax.pie(
        counts,
        labels=counts.index,
        autopct="%1.0f%%",
        colors=colors[:len(counts)],
        startangle=140,
        pctdistance=0.82,
    )
    for t in autotexts:
        t.set_fontsize(9)
    ax.set_title("Cross-sectoral megatrend sources by publication type\n"
                 f"(n = {len(df_mega)} entries)")
    plt.tight_layout()
    plt.savefig(FIG_DIR / "fig2_source_type_breakdown.png")
    plt.close()
    print("  Saved: fig2_source_type_breakdown.png")


def fig_temporal_coverage(df_mega: pd.DataFrame) -> None:
    """
    Figure 3: Publication year distribution of the megatrend database.
    """
    years = df_mega["year"].dropna().astype(int)
    if years.empty:
        print("  Skipped fig3: no year data available.")
        return

    fig, ax = plt.subplots(figsize=(8, 4))
    years.hist(ax=ax, bins=range(int(years.min()), int(years.max()) + 2),
               color="#2166ac", edgecolor="white", rwidth=0.85)
    ax.set_xlabel("Publication year")
    ax.set_ylabel("Number of sources")
    ax.set_title("Temporal coverage of the cross-sectoral megatrend database")
    ax.spines[["top", "right"]].set_visible(False)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "fig3_temporal_coverage.png")
    plt.close()
    print("  Saved: fig3_temporal_coverage.png")


def fig_steep_by_source_type(df_mega: pd.DataFrame) -> None:
    """
    Figure 4: Stacked bar — STEEP categories per source type.
    Shows which source types cover which domains.
    """
    ct = pd.crosstab(df_mega["source_type"], df_mega["steep"])
    # Order columns by PALETTE
    cols_ordered = [c for c in PALETTE if c in ct.columns]
    ct = ct[cols_ordered]

    fig, ax = plt.subplots(figsize=(9, 5))
    ct.plot(kind="bar", stacked=True, ax=ax,
            color=[PALETTE[c] for c in cols_ordered],
            edgecolor="white", linewidth=0.5)
    ax.set_xlabel("Source type")
    ax.set_ylabel("Number of entries")
    ax.set_title("STEEP coverage by source type (cross-sectoral megatrend database)")
    ax.tick_params(axis="x", rotation=30)
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(title="STEEP category", bbox_to_anchor=(1.01, 1), loc="upper left")
    plt.tight_layout()
    plt.savefig(FIG_DIR / "fig4_steep_by_source_type.png")
    plt.close()
    print("  Saved: fig4_steep_by_source_type.png")


def fig_megatrend_fisheries_overlap(df_mega: pd.DataFrame, df_fish: pd.DataFrame) -> None:
    """
    Figure 5: Overlap heatmap — which STEEP domains are covered by both
    cross-sectoral megatrend reports AND fisheries futures literature?
    Highlights research gaps.
    """
    cats = [c for c in PALETTE if c != "Other"]
    mega_pct = df_mega[df_mega["steep"] != "Other"]["steep"].value_counts(normalize=True) * 100
    fish_pct = df_fish[df_fish["steep"] != "Other"]["steep"].value_counts(normalize=True) * 100

    mega_vals = [mega_pct.get(c, 0) for c in cats]
    fish_vals = [fish_pct.get(c, 0) for c in cats]

    x = np.arange(len(cats))
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(x, mega_vals, "o-", color="#2166ac", linewidth=2, markersize=8,
            label="Cross-sectoral megatrends (% of entries)")
    ax.plot(x, fish_vals, "s--", color="#d73027", linewidth=2, markersize=8,
            label="Fisheries futures literature (% of entries)")

    ax.set_xticks(x)
    ax.set_xticklabels(cats)
    ax.set_ylabel("Share of entries (%)")
    ax.set_title("STEEP domain coverage: cross-sectoral megatrends vs. fisheries literature\n"
                 "(divergence indicates potential research gaps)")
    ax.legend(frameon=True)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_ylim(0, max(max(mega_vals), max(fish_vals)) * 1.2)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "fig5_megatrend_fisheries_overlap.png")
    plt.close()
    print("  Saved: fig5_megatrend_fisheries_overlap.png")


# ============================================================
# SECTION 5 — THEMATIC SYNTHESIS
# Generates a structured text summary for Results drafting.
# ============================================================

def generate_synthesis(df_mega: pd.DataFrame, df_fish: pd.DataFrame) -> str:
    """
    Produce a structured plain-text synthesis of both datasets.
    Intended as raw material for drafting the Results section.
    """
    lines = []

    lines.append("=" * 65)
    lines.append("THEMATIC SYNTHESIS — FOR RESULTS SECTION DRAFTING")
    lines.append("=" * 65)

    # --- Cross-sectoral megatrends ---
    lines.append("\n## 3.1 Cross-sectoral megatrends and drivers")
    lines.append(f"\nTotal entries: {len(df_mega)}")
    lines.append(f"Sources: {df_mega['source_type'].nunique()} types "
                 f"({', '.join(df_mega['source_type'].value_counts().index[:4])})")

    lines.append("\nSTEEP distribution:")
    for cat, n in df_mega["steep"].value_counts().items():
        pct = n / len(df_mega) * 100
        lines.append(f"  {cat:<18}: {n:>3}  ({pct:.0f}%)")

    lines.append("\nTop recurring themes (deduplicated by fuzzy title matching):")
    df_mega_dup = deduplicate_by_theme(df_mega)
    theme_counts = df_mega_dup.groupby("canonical_theme").size().sort_values(ascending=False)
    for theme, count in theme_counts.head(15).items():
        lines.append(f"  [{count:>2}x] {str(theme).title()[:60]}")

    lines.append("\nKey megatrends by STEEP category:")
    for cat in [c for c in PALETTE if c != "Other"]:
        sub = df_mega[df_mega["steep"] == cat]["trend"].dropna()
        unique_examples = sub.str.strip().drop_duplicates().head(4).tolist()
        lines.append(f"\n  {cat}:")
        for ex in unique_examples:
            lines.append(f"    — {str(ex)[:70]}")

    # --- Fisheries futures literature ---
    lines.append("\n\n## 3.2 Futures of Fisheries literature")
    lines.append(f"\nTotal usable entries: {len(df_fish)}")

    lines.append("\nSTEEP distribution:")
    for cat, n in df_fish["steep"].value_counts().items():
        pct = n / len(df_fish) * 100
        lines.append(f"  {cat:<18}: {n:>3}  ({pct:.0f}%)")

    lines.append("\nKey topics identified:")
    for _, row in df_fish.iterrows():
        trend = str(row.get("trend", "")).strip()
        title = str(row.get("title", "")).strip()
        if trend and trend.lower() not in ("none", "nan"):
            lines.append(f"  — {trend[:80]}")
            if title and title.lower() not in ("none", "nan"):
                lines.append(f"    ({title[:70]})")

    # --- Gap analysis ---
    lines.append("\n\n## 3.3 Gap analysis: megatrends not reflected in fisheries literature")
    cats = [c for c in PALETTE if c != "Other"]
    mega_pct = df_mega["steep"].value_counts(normalize=True) * 100
    fish_pct = df_fish["steep"].value_counts(normalize=True) * 100

    lines.append("\nDifference in STEEP coverage (megatrend % minus fisheries %):")
    gaps = []
    for cat in cats:
        m = mega_pct.get(cat, 0)
        f = fish_pct.get(cat, 0)
        diff = m - f
        gaps.append((cat, m, f, diff))
        lines.append(f"  {cat:<18}: megatrend={m:.0f}%  fisheries={f:.0f}%  gap={diff:+.0f}pp")

    gaps.sort(key=lambda x: -abs(x[3]))
    lines.append(f"\nLargest gap: {gaps[0][0]} (megatrends={gaps[0][1]:.0f}%, "
                 f"fisheries={gaps[0][2]:.0f}%, gap={gaps[0][3]:+.0f}pp)")

    lines.append("\n" + "=" * 65)

    return "\n".join(lines)


# ============================================================
# SECTION 6 — EXPORT
# ============================================================

def export_cleaned(df_mega: pd.DataFrame, df_fish: pd.DataFrame,
                   synthesis: str) -> None:
    """Save cleaned datasets and synthesis to Excel and text files."""
    with pd.ExcelWriter(OUTPUT_DIR / "excel_analysis.xlsx", engine="openpyxl") as writer:
        df_mega.to_excel(writer, sheet_name="Megatrends_cleaned", index=False)
        df_fish.to_excel(writer, sheet_name="FisheriesFutures_cleaned", index=False)

        # STEEP summary table
        steep_summary = pd.DataFrame({
            "STEEP Category":       list(PALETTE.keys()),
            "Megatrends (n)":       [df_mega["steep"].value_counts().get(c, 0) for c in PALETTE],
            "Fisheries futures (n)":[df_fish["steep"].value_counts().get(c, 0) for c in PALETTE],
        })
        steep_summary.to_excel(writer, sheet_name="STEEP_summary", index=False)

    with open(OUTPUT_DIR / "results_synthesis.txt", "w", encoding="utf-8") as f:
        f.write(synthesis)

    print(f"  Saved: excel_analysis.xlsx")
    print(f"  Saved: results_synthesis.txt")


# ============================================================
# SECTION 7 — MAIN
# ============================================================

def main() -> None:
    print("=" * 65)
    print("Excel Database Analysis — Drivers, Trends and Signals")
    print("=" * 65)

    print("\nLoading data ...")
    df_mega = load_megatrends()
    df_fish = load_fisheries_futures()

    print(f"  (Mega-)trends and drivers : {len(df_mega)} entries loaded")
    print(f"  Futures of Fisheries      : {len(df_fish)} entries loaded")

    print("\nSTEEP classification summary:")
    print("  Megatrends:")
    for cat, n in df_mega["steep"].value_counts().items():
        print(f"    {cat:<18}: {n}")
    print("  Fisheries futures:")
    for cat, n in df_fish["steep"].value_counts().items():
        print(f"    {cat:<18}: {n}")

    print("\nGenerating figures ...")
    fig_steep_distribution(df_mega, df_fish)
    fig_source_type_breakdown(df_mega)
    fig_temporal_coverage(df_mega)
    fig_steep_by_source_type(df_mega)
    fig_megatrend_fisheries_overlap(df_mega, df_fish)

    print("\nGenerating thematic synthesis ...")
    synthesis = generate_synthesis(df_mega, df_fish)
    print(synthesis)

    print("\nExporting cleaned datasets ...")
    export_cleaned(df_mega, df_fish, synthesis)

    print(f"\nDone. All outputs in: {OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
