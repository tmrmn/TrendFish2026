"""
Generate a publication-quality PRISMA 2020 flow diagram using Graphviz.
Writes: ../docs/PRISMA_flow.svg, .pdf, .png

PRISMA 2020 numbers (must match Methods section):
  OpenAlex database search:        2,100 identified
  Duplicates removed:                479
  After deduplication / screened:  1,621
  Excluded at title/abstract:        901
  Pending full-text:                 720
  Strategy 1 (manual):                60 included
  Strategy 2 (cross-sectoral):       145 included
  Strategy 3 (signals):                3 included (initial)
  Total included in synthesis:       208
"""

from graphviz import Digraph
from pathlib import Path

OUT_DIR = Path(__file__).parent.parent / "docs"
OUT_DIR.mkdir(exist_ok=True)

# Numbers — keep in sync with write_methods.py
N_OPENALEX  = 2_100
N_DUP       = 479
N_SCREENED  = N_OPENALEX - N_DUP   # 1,621
N_EXCL_TA   = 901
N_PENDING   = N_SCREENED - N_EXCL_TA  # 720
N_S1_MAN    = 60
N_S2        = 145
N_S3        = 3
N_TOTAL     = N_S1_MAN + N_S2 + N_S3  # 208

# Palette
C_BLUE  = "#3949AB"
C_RED   = "#B71C1C"
C_GREEN = "#1B5E20"
C_AMBER = "#E65100"
C_WHITE = "#FFFFFF"

# Backgrounds
BG_BLUE  = "#E8F0FE"
BG_RED   = "#FFEBEE"
BG_GREEN = "#E8F5E9"
BG_AMBER = "#FFF8E1"


def htbl(*rows):
    """Return an HTML-table label for a Graphviz node.

    Each item in `rows` is a string of HTML; pass empty string for a
    vertical spacer row.
    """
    cells = "".join(
        f'<TR><TD ALIGN="LEFT" BALIGN="LEFT">{r if r else " "}</TD></TR>'
        for r in rows
    )
    return (
        f'<<TABLE BORDER="0" CELLBORDER="0" CELLPADDING="4"'
        f' CELLSPACING="0">{cells}</TABLE>>'
    )


def n_main(**kw):
    return dict(
        shape="box", style="rounded,filled",
        fillcolor=BG_BLUE, color=C_BLUE,
        fontname="Helvetica", fontsize="10", penwidth="1.8", **kw,
    )


def n_excl(**kw):
    return dict(
        shape="box", style="rounded,filled",
        fillcolor=BG_RED, color=C_RED,
        fontname="Helvetica", fontsize="10", penwidth="1.4", **kw,
    )


def n_other(**kw):
    return dict(
        shape="box", style="rounded,filled",
        fillcolor=BG_GREEN, color=C_GREEN,
        fontname="Helvetica", fontsize="10", penwidth="1.8", **kw,
    )


def n_result(**kw):
    return dict(
        shape="box", style="rounded,filled",
        fillcolor=BG_AMBER, color=C_AMBER,
        fontname="Helvetica", fontsize="10", penwidth="2.2", **kw,
    )


# ── Build diagram ─────────────────────────────────────────────

dot = Digraph("PRISMA_2020", engine="dot")
dot.attr(
    rankdir="TB",
    splines="ortho",
    nodesep="0.9",
    ranksep="0.7",
    fontname="Arial",
    bgcolor="white",
    pad="0.55",
)

# ─── Phase clusters ──────────────────────────────────────────

CLUSTER_ATTR = dict(
    style="rounded",
    penwidth="1.5",
    fontname="Helvetica-Bold",
    fontsize="10",
    labeljust="l",
    margin="16",
)

with dot.subgraph(name="cluster_id") as c:
    c.attr(
        label="IDENTIFICATION",
        color=C_BLUE, fontcolor=C_BLUE,
        bgcolor="#F5F7FF",
        **CLUSTER_ATTR,
    )

    c.node("db", htbl(
        "<B>Records identified from databases</B>",
        "OpenAlex API -- 7 query strings",
        "Up to 300 results per query (polite pool)",
        f"<B>n = {N_OPENALEX:,}</B>",
    ), **n_main())

    c.node("other_id", htbl(
        "<B>Records identified from other sources</B>",
        f"Strategy 1 -- manual database search:    n = {N_S1_MAN}",
        f"Strategy 2 -- cross-sectoral foresight:  n = {N_S2}",
        f"Strategy 3 -- weak signal horizon scan:  n = {N_S3} (initial)",
        "<I>Strategies 2 and 3 screened separately (see Methods 3.3-3.4)</I>",
    ), **n_other())

    c.node("dedup", htbl(
        "<B>Records removed before screening</B>",
        f"Duplicate records removed: n = {N_DUP}",
        "Method: exact DOI match + fuzzy title (RapidFuzz, 90% threshold)",
    ), **n_excl())

with dot.subgraph(name="cluster_scr") as c:
    c.attr(
        label="SCREENING",
        color=C_BLUE, fontcolor=C_BLUE,
        bgcolor="#F5F7FF",
        **CLUSTER_ATTR,
    )

    c.node("screened", htbl(
        "<B>Records screened</B>",
        "Title and abstract review by primary reviewer",
        f"<B>n = {N_SCREENED:,}</B>",
    ), **n_main())

    c.node("excl_ta", htbl(
        "<B>Records excluded at title/abstract</B>",
        f"n = {N_EXCL_TA}",
        "Reason 1: not fisheries or marine capture fisheries",
        "Reason 2: no future-oriented analytical perspective",
        "   (retrospective or purely descriptive)",
    ), **n_excl())

with dot.subgraph(name="cluster_elig") as c:
    c.attr(
        label="ELIGIBILITY",
        color=C_BLUE, fontcolor=C_BLUE,
        bgcolor="#F5F7FF",
        **CLUSTER_ATTR,
    )

    c.node("fulltext", htbl(
        "<B>Reports assessed for eligibility (full-text)</B>",
        f"<B>n = {N_PENDING} -- PENDING</B>",
        "<I>Not included in quantitative analyses reported in this article</I>",
    ), **n_main())

with dot.subgraph(name="cluster_inc") as c:
    c.attr(
        label="INCLUDED",
        color=C_AMBER, fontcolor=C_AMBER,
        bgcolor="#FFFDF5",
        **CLUSTER_ATTR,
    )

    c.node("included", htbl(
        "<B>Studies included in synthesis</B>",
        f"Strategy 1  (peer-reviewed, manual screening):   n = {N_S1_MAN}",
        f"Strategy 2  (cross-sectoral foresight, manual):  n = {N_S2}",
        f"Strategy 3  (weak signals, initial scan):        n = {N_S3}",
        "",
        f"<B>Total included in analysis: n = {N_TOTAL}</B>",
        f"<I>Note: {N_PENDING} papers from database search pending full-text review</I>",
    ), **n_result())

# ─── Edges ───────────────────────────────────────────────────

E_MAIN = dict(color=C_BLUE,  penwidth="1.8", arrowsize="0.9")
E_EXCL = dict(color=C_RED,   penwidth="1.4", style="dashed", arrowsize="0.8")
E_OTH  = dict(color=C_GREEN, penwidth="1.4", style="dashed", arrowsize="0.8")

# Database pipeline
dot.edge("db",       "dedup",    **E_MAIN)
dot.edge("dedup",    "screened", **E_MAIN)
dot.edge("screened", "excl_ta",  **E_EXCL)
dot.edge("screened", "fulltext", **E_MAIN)
dot.edge("fulltext", "included", **E_MAIN)

# Other sources feed directly to included
dot.edge("other_id", "included", **E_OTH)

# ─── Render ──────────────────────────────────────────────────

stem = str(OUT_DIR / "PRISMA_flow")
for fmt in ("svg", "png"):
    dot.render(stem, format=fmt, cleanup=True)
    print(f"  Saved: {stem}.{fmt}")

print(f"\nPRISMA 2020 flow diagram complete.")
print(f"  Identified (database):  n = {N_OPENALEX:,}")
print(f"  After deduplication:    n = {N_SCREENED:,}")
print(f"  Pending full-text:      n = {N_PENDING}")
print(f"  Included in synthesis:  n = {N_TOTAL}")
