"""
PRISMA 2020 Flow Diagram — Template 2 (Databases/Registers + Other Sources)
Faithfully matches the official Word template from prisma-statement.org.

Official box labels extracted directly from the Word template:
  Source: Page MJ et al. BMJ 2021;372:n71. doi: 10.1136/bmj.n71
  Template: PRISMA_2020_flow_diagram_new_SRs_v2.docx (CC BY 4.0)

Writes: ../docs/PRISMA_flow.svg, ../docs/PRISMA_flow.png
"""

import matplotlib.pyplot as plt
import matplotlib.patches as mpa
from pathlib import Path

OUT_DIR = Path(__file__).parent.parent / "docs"
OUT_DIR.mkdir(exist_ok=True)

# ── Review-specific numbers ────────────────────────────────────
N_DB         = 2_100   # OpenAlex hits
N_DUP        = 479     # duplicates removed
N_SCREENED   = 1_621   # after dedup
N_EXCL_TA    = 901     # excluded at title/abstract
N_SOUGHT_L   = 720     # reports sought (left column, PENDING)
N_NOT_RET_L  = 0
N_ASSESSED_L = 720     # PENDING full-text
N_EXCL_FT_L  = 720     # all pending → counted as excluded for now

N_S1_MAN = 60
N_S2     = 145
N_S3     = 3
N_OTHER  = N_S1_MAN + N_S2 + N_S3   # 208
N_SOUGHT_R   = 208
N_NOT_RET_R  = 0
N_ASSESSED_R = 208
N_EXCL_FT_R  = 0
N_TOTAL  = N_S1_MAN + N_S2 + N_S3   # 208

# ── Figure setup ───────────────────────────────────────────────
FW, FH = 17.0, 15.5
fig, ax = plt.subplots(figsize=(FW, FH))
ax.set_xlim(0, FW)
ax.set_ylim(0, FH)
ax.axis("off")
fig.patch.set_facecolor("white")
ax.set_facecolor("white")

# ── Colour palette (matches official template) ─────────────────
C_PHASE = "#BDD7EE"   # official PRISMA light blue
C_WHITE = "white"
C_BLACK = "black"

# ── Column x-positions ─────────────────────────────────────────
PL_X, PL_W = 0.05, 1.15          # phase-label sidebar
LM_X, LM_W = 1.35, 4.85          # left main (database pipeline)
LE_X, LE_W = 6.35, 3.65          # left exclusion
RM_X, RM_W = 10.15, 4.15         # right main (other sources)
RE_X, RE_W = 14.45, 2.40         # right exclusion

# ── Phase y-boundaries ─────────────────────────────────────────
PH_ID   = (11.60, 15.50)   # Identification
PH_SCR  = ( 6.70, 11.60)   # Screening
PH_ELIG = ( 2.30,  6.70)   # Eligibility
PH_INC  = ( 0.15,  2.30)   # Included

# ── Box positions: (x, y_bottom, w, h) ────────────────────────
# Left main column
B_L_ID   = (LM_X, 12.95, LM_W, 1.95)  # Records identified from databases
B_L_SCR  = (LM_X,  9.75, LM_W, 1.00)  # Records screened
B_L_SFT  = (LM_X,  7.65, LM_W, 1.00)  # Reports sought for retrieval
B_L_ELIG = (LM_X,  3.00, LM_W, 2.80)  # Reports assessed for eligibility

# Left exclusion column (same y as their corresponding left-main box)
B_L_RM   = (LE_X, B_L_ID[1],   LE_W, B_L_ID[3])    # Records removed
B_L_ETA  = (LE_X, B_L_SCR[1],  LE_W, B_L_SCR[3])   # Records excluded (T/A)
B_L_NR   = (LE_X, B_L_SFT[1],  LE_W, B_L_SFT[3])   # Reports not retrieved
B_L_EFT  = (LE_X, B_L_ELIG[1], LE_W, B_L_ELIG[3])  # Reports excluded (FT)

# Right main column
B_R_ID   = (RM_X, 12.45, RM_W, 2.45)  # Records identified from other sources
B_R_SFT  = (RM_X,  7.65, RM_W, 1.00)  # Reports sought for retrieval
B_R_ELIG = (RM_X,  3.00, RM_W, 2.80)  # Reports assessed for eligibility

# Right exclusion column
B_R_NR   = (RE_X, B_R_SFT[1],  RE_W, B_R_SFT[3])
B_R_EFT  = (RE_X, B_R_ELIG[1], RE_W, B_R_ELIG[3])

# Included box (spans both columns)
INC_X = LM_X
INC_W = RE_X + RE_W - LM_X
B_INC = (INC_X, 0.35, INC_W, 1.75)


# ── Drawing helpers ────────────────────────────────────────────

def rect(x, y, w, h, fc="white", ec="black", lw=0.75, zorder=2):
    ax.add_patch(mpa.Rectangle(
        (x, y), w, h,
        facecolor=fc, edgecolor=ec, linewidth=lw, zorder=zorder,
    ))


FS = 8.4          # body font size (pt)
LH = FS / 72 * 1.55  # line height in inches (~1.55× font size)
PAD = 0.14        # inner top/left padding in inches


def put_text(box, lines, fs=FS, bold_idx=None):
    """Render `lines` inside `box` = (x, y_bot, w, h), top-aligned."""
    x, y_bot, _, h = box
    bold_idx = bold_idx or []
    top = y_bot + h - PAD
    for i, line in enumerate(lines):
        ly = top - i * (fs / 72 * 1.55)
        if ly < y_bot + PAD * 0.4:
            break
        fw = "bold" if i in bold_idx else "normal"
        ax.text(
            x + PAD, ly, line,
            ha="left", va="top",
            fontsize=fs, fontweight=fw,
            fontfamily="DejaVu Sans",
            zorder=4,
        )


def draw_box(box, lines, fs=FS, bold_idx=None, fc="white"):
    x, y, w, h = box
    rect(x, y, w, h, fc=fc)
    put_text(box, lines, fs=fs, bold_idx=bold_idx)


def phase_label(y_bot, h, text):
    rect(PL_X, y_bot, PL_W, h, fc=C_PHASE, lw=0.75)
    ax.text(
        PL_X + PL_W / 2, y_bot + h / 2, text,
        ha="center", va="center",
        fontsize=10.5, fontweight="bold",
        fontfamily="DejaVu Sans",
        rotation=90, zorder=4,
    )


def arr_down(x_c, y_from, y_to):
    """Downward arrow from (x_c, y_from) to (x_c, y_to)."""
    ax.annotate(
        "", xy=(x_c, y_to), xytext=(x_c, y_from),
        arrowprops=dict(
            arrowstyle="->,head_width=0.18,head_length=0.13",
            color="black", lw=0.9, shrinkA=0, shrinkB=0,
        ),
        zorder=5,
    )


def arr_right(x_from, x_to, y_c):
    """Rightward arrow from (x_from, y_c) to (x_to, y_c)."""
    ax.annotate(
        "", xy=(x_to, y_c), xytext=(x_from, y_c),
        arrowprops=dict(
            arrowstyle="->,head_width=0.18,head_length=0.13",
            color="black", lw=0.9, shrinkA=0, shrinkB=0,
        ),
        zorder=5,
    )


# ── Phase labels ───────────────────────────────────────────────
phase_label(PH_ID[0],   PH_ID[1]   - PH_ID[0],   "Identification")
phase_label(PH_SCR[0],  PH_SCR[1]  - PH_SCR[0],  "Screening")
phase_label(PH_ELIG[0], PH_ELIG[1] - PH_ELIG[0], "Eligibility")
phase_label(PH_INC[0],  PH_INC[1]  - PH_INC[0],  "Included")

# ── Column header italics (above each stream) ──────────────────
ax.text(
    LM_X + LM_W / 2, PH_ID[1] - 0.22,
    "Identification of studies via databases and registers",
    ha="center", va="top", fontsize=8.5, fontstyle="italic",
    fontfamily="DejaVu Sans", zorder=4,
)
ax.text(
    RM_X + RM_W / 2, PH_ID[1] - 0.22,
    "Identification of studies via other methods",
    ha="center", va="top", fontsize=8.5, fontstyle="italic",
    fontfamily="DejaVu Sans", zorder=4,
)

# ══════════════════════════════════════════════════════════════
# IDENTIFICATION PHASE
# ══════════════════════════════════════════════════════════════

# Left: Records identified from databases/registers
draw_box(B_L_ID, [
    "Records identified from*:",
    f"   Databases (n = {N_DB:,})",
    f"   [OpenAlex API, 7 query strings, 300 results/query]",
    "   Registers (n = 0)",
], bold_idx=[0])

# Right: Records identified from other sources
draw_box(B_R_ID, [
    "Records identified from:",
    f"   Websites (n = 0)",
    f"   Organisations (n = 0)",
    f"   Citation searching (n = {N_S1_MAN})",
    f"     [Strategy 1 — manual peer-reviewed search]",
    f"   Other methods (n = {N_S2 + N_S3})",
    f"     [Strategy 2 foresight: {N_S2}; Strategy 3 signals: {N_S3}]",
], bold_idx=[0])

# Left exclusion: Records removed before screening
draw_box(B_L_RM, [
    "Records removed before screening:",
    f"   Duplicate records removed (n = {N_DUP})",
    "   [DOI match + fuzzy title, 90% threshold]",
    "   Records marked as ineligible by",
    "   automation tools (n = 0)",
    "   Records removed for other reasons",
    "   (n = 0)",
], bold_idx=[0])

# ══════════════════════════════════════════════════════════════
# SCREENING PHASE
# ══════════════════════════════════════════════════════════════

# Left: Records screened
draw_box(B_L_SCR, [
    "Records screened",
    f"   (n = {N_SCREENED:,})",
], bold_idx=[0])

# Left exclusion: Records excluded at title/abstract
draw_box(B_L_ETA, [
    "Records excluded**",
    f"   (n = {N_EXCL_TA})",
], bold_idx=[0])

# Left: Reports sought for retrieval
draw_box(B_L_SFT, [
    "Reports sought for retrieval",
    f"   (n = {N_SOUGHT_L})",
], bold_idx=[0])

# Left exclusion: Reports not retrieved
draw_box(B_L_NR, [
    "Reports not retrieved",
    f"   (n = {N_NOT_RET_L})",
], bold_idx=[0])

# Right: Reports sought for retrieval
draw_box(B_R_SFT, [
    "Reports sought for retrieval",
    f"   (n = {N_SOUGHT_R})",
], bold_idx=[0])

# Right exclusion: Not retrieved
draw_box(B_R_NR, [
    "Reports not retrieved",
    f"   (n = {N_NOT_RET_R})",
], bold_idx=[0], fs=7.8)

# ══════════════════════════════════════════════════════════════
# ELIGIBILITY PHASE
# ══════════════════════════════════════════════════════════════

# Left: Reports assessed for eligibility
draw_box(B_L_ELIG, [
    "Reports assessed for eligibility",
    f"   (n = {N_ASSESSED_L}) — PENDING",
    "",
    "Full-text review pending; these papers",
    "are NOT included in the quantitative",
    "analyses reported in this article.",
    "Tracked in search_output/",
    "screened_included.csv",
], bold_idx=[0])

# Left exclusion: Reports excluded (full-text)
draw_box(B_L_EFT, [
    "Reports excluded:",
    "   Full-text review pending",
    f"   (n = {N_EXCL_FT_L})",
    "",
    "   Reason: not yet assessed",
], bold_idx=[0])

# Right: Reports assessed for eligibility
draw_box(B_R_ELIG, [
    "Reports assessed for eligibility",
    f"   (n = {N_ASSESSED_R})",
    "",
    f"   Strategy 1 (peer-reviewed, manual): {N_S1_MAN}",
    f"   Strategy 2 (cross-sectoral foresight): {N_S2}",
    f"   Strategy 3 (weak signal scan): {N_S3}",
    "",
    "   All screened for relevance by",
    "   primary reviewer (see Methods 3.2–3.4)",
], bold_idx=[0])

# Right exclusion: Reports excluded (full-text)
draw_box(B_R_EFT, [
    "Reports excluded:",
    f"   Reason 1 (n = {N_EXCL_FT_R})",
], bold_idx=[0], fs=7.8)

# ══════════════════════════════════════════════════════════════
# INCLUDED PHASE
# ══════════════════════════════════════════════════════════════

draw_box(B_INC, [
    f"Studies included in review  (n = {N_TOTAL})     "
    f"Reports of included studies  (n = {N_TOTAL})",
    f"   Strategy 1 — peer-reviewed fisheries futures (manual screening):   "
    f"n = {N_S1_MAN}",
    f"   Strategy 2 — cross-sectoral foresight publications (purposive):    "
    f"n = {N_S2}",
    f"   Strategy 3 — weak signal horizon scan (initial):                   "
    f"n = {N_S3}",
    f"   Note: {N_SOUGHT_L} papers from database search pending full-text review",
], bold_idx=[0])

# ══════════════════════════════════════════════════════════════
# ARROWS
# ══════════════════════════════════════════════════════════════

# Centre-x for each main column
LM_CX = LM_X + LM_W / 2
RM_CX = RM_X + RM_W / 2

# ── Left pipeline (downward) ──────────────────────────────────
arr_down(LM_CX, B_L_ID[1],   B_L_SCR[1]  + B_L_SCR[3])   # L_ID  → L_SCR
arr_down(LM_CX, B_L_SCR[1],  B_L_SFT[1]  + B_L_SFT[3])   # L_SCR → L_SFT
arr_down(LM_CX, B_L_SFT[1],  B_L_ELIG[1] + B_L_ELIG[3])  # L_SFT → L_ELIG
arr_down(LM_CX, B_L_ELIG[1], B_INC[1]    + B_INC[3])      # L_ELIG→ INC

# ── Left exclusion arrows (rightward from right edge of main) ─
LM_RX = LM_X + LM_W
arr_right(LM_RX, LE_X, B_L_ID[1]   + B_L_ID[3]   / 2)   # → removed
arr_right(LM_RX, LE_X, B_L_SCR[1]  + B_L_SCR[3]  / 2)   # → excl T/A
arr_right(LM_RX, LE_X, B_L_SFT[1]  + B_L_SFT[3]  / 2)   # → not retrieved
arr_right(LM_RX, LE_X, B_L_ELIG[1] + B_L_ELIG[3] / 2)   # → excl FT

# ── Right pipeline (downward) ─────────────────────────────────
arr_down(RM_CX, B_R_ID[1],   B_R_SFT[1]  + B_R_SFT[3])   # R_ID  → R_SFT
arr_down(RM_CX, B_R_SFT[1],  B_R_ELIG[1] + B_R_ELIG[3])  # R_SFT → R_ELIG
arr_down(RM_CX, B_R_ELIG[1], B_INC[1]    + B_INC[3])      # R_ELIG→ INC

# ── Right exclusion arrows ────────────────────────────────────
RM_RX = RM_X + RM_W
arr_right(RM_RX, RE_X, B_R_SFT[1]  + B_R_SFT[3]  / 2)
arr_right(RM_RX, RE_X, B_R_ELIG[1] + B_R_ELIG[3] / 2)

# ── Footnotes ─────────────────────────────────────────────────
ax.text(
    LM_X, 0.02,
    "*Consider, if feasible, reporting the number of records identified from each "
    "database or register separately (rather than the total across all databases/registers).\n"
    "**If automation tools were used, indicate how many records were excluded by a human "
    "and how many were excluded by automation tools.",
    ha="left", va="bottom", fontsize=6.8, fontstyle="italic",
    fontfamily="DejaVu Sans", zorder=4, color="#444444",
)
ax.text(
    INC_X + INC_W, 0.02,
    "Source: Page MJ et al. BMJ 2021;372:n71. doi: 10.1136/bmj.n71  |  "
    "This work is licensed under CC BY 4.0",
    ha="right", va="bottom", fontsize=6.5,
    fontfamily="DejaVu Sans", zorder=4, color="#666666",
)

# ══════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════

plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
for fmt, out in [("svg", OUT_DIR / "PRISMA_flow.svg"),
                 ("png", OUT_DIR / "PRISMA_flow.png")]:
    plt.savefig(out, format=fmt, dpi=300, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    print(f"  Saved: {out}")

plt.close()
print("\nPRISMA 2020 flow diagram complete.")
print(f"  Database search:        n = {N_DB:,} → {N_SCREENED:,} screened → "
      f"{N_SOUGHT_L} pending full-text")
print(f"  Other sources:          n = {N_OTHER} (all included)")
print(f"  Total included:         n = {N_TOTAL}")
