"""
Prepare the full-text review dataset for the 720 screened papers.

Step 1  Fetch missing abstracts from OpenAlex API (polite pool).
Step 2  Apply keyword-based screening → green / orange / red verdict per paper.
Step 3  Write (or update) data/full_text_review.csv.

Re-running is safe: existing user_verdict / user_notes are preserved.

Usage:
    python -X utf8 analysis/prepare_review.py
"""

import json
import time
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime

POLITE_EMAIL = "tsz000@uit.no"

BASE     = Path(__file__).parent.parent
IN_CSV   = BASE / "analysis" / "search_output" / "screened_included.csv"
OUT_CSV  = BASE / "data" / "full_text_review.csv"

# ── Inclusion / exclusion criteria (from Methods §3.2) ────────
#
# INCLUDE if BOTH:
#   I1. Explicitly addresses future scenarios, drivers, or trend
#       analyses in fisheries / marine capture fisheries
#   I2. Future-oriented analytical perspective (scenarios, projections,
#       foresight, trend assessment) — NOT purely retrospective
#
# EXCLUDE if ANY:
#   E1. Not related to fisheries / marine capture fisheries
#   E2. Lacks a future-oriented perspective
#   E3. Only biological population dynamics without broader drivers

# ── Keyword lists ─────────────────────────────────────────────

FISH_STRONG = [
    "fisheries", "fishery", "fishing industry", "commercial fishing",
    "marine capture", "fish stock", "fish catch", "artisanal fish",
    "small-scale fish", "seafood market", "aquaculture production",
    "mariculture",
]
FISH_MED = [
    "fishing", "fish catch", "fisher", "aquaculture", "seafood",
    "trawl", "seine", "longline", "gillnet", "msc certified",
    "fisheries management", "fisheries governance", "fish production",
]
FISH_CONTEXT = [
    "marine ecosystem", "coastal community", "pelagic", "demersal",
    "bycatch", "by-catch", "overfish", "fish population", "fish species",
    "ocean harvest", "sea fish",
]

FUTURE_STRONG = [
    "future scenario", "fisheries scenario", "future of fisheries",
    "fisheries futures", "foresight", "horizon scan", "prospective",
    "future projection", "future outlook", "futures study",
    "fisheries futures", "futures analysis",
]
FUTURE_MED = [
    "scenario", "projection", "forecast", "by 2030", "by 2040",
    "by 2050", "by 2060", "by 2100", "under rcp", "under ssp",
    "climate scenario", "future trend", "expected impact",
    "anticipated change", "long-term outlook", "medium-term",
    "future impact", "future change", "future state",
    "future condition", "will be", "will likely", "is expected",
    "are projected",
]
FUTURE_WEAK = [
    "trend", "driver", "future", "model", "outlook", "prospect",
    "long-term", "simulate", "predict",
]

EXCLUSION_RETRO = [
    "retrospective", "over the past", "historical trend",
    "historical analysis", "past decade", "years of data",
    "long-term data", "time series analys",
]
EXCLUSION_BIO_ONLY = [
    "phylogenetic", "phylogeny", "genomic", "genome sequen",
    "dna barcod", "genetic diversity", "microsatellite",
    "population genetic", "allele", "genotype",
]
EXCLUSION_NON_FISH = [
    "agriculture", "crop production", "livestock", "poultry",
    "bovine", "swine", "wheat", "maize", "terrestrial",
    "cattle farm", "dairy",
]


# ── Abstract reconstruction ───────────────────────────────────

def reconstruct_abstract(inv: dict) -> str:
    """Rebuild abstract text from OpenAlex inverted-index format."""
    pairs = [(pos, word) for word, positions in inv.items()
             for pos in positions]
    return " ".join(w for _, w in sorted(pairs))


def fetch_abstract(openalex_url: str) -> str | None:
    """Return reconstructed abstract from OpenAlex, or None."""
    work_id = str(openalex_url).strip().rstrip("/").split("/")[-1]
    try:
        r = requests.get(
            f"https://api.openalex.org/works/{work_id}",
            params={"select": "abstract_inverted_index", "mailto": POLITE_EMAIL},
            timeout=15,
        )
        if r.status_code == 200:
            inv = r.json().get("abstract_inverted_index")
            if inv:
                return reconstruct_abstract(inv)
    except Exception:
        pass
    return None


# ── Keyword scoring ───────────────────────────────────────────

def score(title: str, abstract: str) -> tuple[str, str, dict]:
    """
    Returns (verdict, reason, score_dict).
    verdict:  'green' | 'orange' | 'red'
    """
    text = ((title or "") + " " + (abstract or "")).lower()
    has_abstract = bool(abstract and len(str(abstract).strip()) > 50)

    # Fisheries relevance score (0–8)
    f_strong = sum(1 for t in FISH_STRONG if t in text)
    f_med    = sum(1 for t in FISH_MED    if t in text)
    f_ctx    = sum(1 for t in FISH_CONTEXT if t in text)
    fish = min(f_strong * 3 + f_med * 2 + f_ctx, 8)

    # Future orientation score (0–10)
    fut_s = sum(1 for t in FUTURE_STRONG if t in text)
    fut_m = sum(1 for t in FUTURE_MED    if t in text)
    fut_w = sum(1 for t in FUTURE_WEAK   if t in text)
    future = min(fut_s * 4 + fut_m * 2 + fut_w, 10)

    # Exclusion hits
    retro  = [t for t in EXCLUSION_RETRO    if t in text]
    bio    = [t for t in EXCLUSION_BIO_ONLY if t in text]
    nonfis = [t for t in EXCLUSION_NON_FISH if t in text]

    scores = {
        "fish": fish, "future": future,
        "retro": retro, "bio": bio, "nonfis": nonfis,
    }

    # ── Decision tree ─────────────────────────────────────────
    # Hard exclusions
    if fish == 0 and len(nonfis) >= 1:
        return "red", f"Non-fisheries topic ({nonfis[0]})", scores
    if fish == 0:
        return "red", "No fisheries relevance detected", scores
    if future == 0 and len(retro) >= 1:
        return "red", f"Retrospective analysis, no future orientation ({retro[0]})", scores
    if bio and fish <= 2 and future < 2:
        return "red", f"Biological/genetic focus without fisheries futures context ({bio[0]})", scores

    # Clear inclusions
    if fish >= 4 and future >= 6 and not retro:
        return "green", "Strong fisheries + future orientation", scores
    if fish >= 3 and future >= 4 and not retro:
        return "green", "Clear fisheries context with future-oriented analysis", scores
    if fish >= 2 and future >= 6:
        return "green", "Good fisheries relevance with strong future signals", scores

    # Uncertain
    if not has_abstract:
        return "orange", "Abstract unavailable — verdict based on title only", scores
    if future == 0:
        return "orange", "Fisheries-relevant but no future orientation detected", scores
    if fish >= 2 and future >= 2:
        return "orange", "Borderline — fisheries + future signals present but weak", scores

    return "orange", f"Insufficient signals for clear verdict (fish={fish}, future={future})", scores


# ── Main ──────────────────────────────────────────────────────

def main():
    df = pd.read_csv(IN_CSV)
    print(f"Loaded {len(df)} screened papers from {IN_CSV.name}")

    # Preserve existing user verdicts on re-run
    saved_verdicts: dict[str, dict] = {}
    if OUT_CSV.exists():
        old = pd.read_csv(OUT_CSV)
        for _, r in old.iterrows():
            if pd.notna(r.get("user_verdict")) and str(r.get("user_verdict", "")).strip():
                saved_verdicts[str(r["openalex_id"])] = {
                    "user_verdict": r["user_verdict"],
                    "user_notes":   r.get("user_notes", ""),
                    "reviewed_at":  r.get("reviewed_at", ""),
                }
        print(f"  Preserving {len(saved_verdicts)} existing user verdicts.\n")

    # ── Step 1: Fetch missing abstracts ───────────────────────
    missing = df["abstract"].isna() | (df["abstract"].astype(str).str.strip() == "")
    n_miss  = int(missing.sum())
    print(f"Step 1 — Fetching abstracts for {n_miss} papers without one ...")

    df["abstract_source"] = "original"
    fetched = 0
    for idx in df[missing].index:
        oa = df.loc[idx, "openalex_id"]
        ab = fetch_abstract(str(oa)) if pd.notna(oa) else None
        if ab:
            df.loc[idx, "abstract"] = ab
            df.loc[idx, "abstract_source"] = "openalex"
            fetched += 1
            print(f"  [{fetched:>3}/{n_miss}] {str(df.loc[idx,'title'])[:65]}")
        else:
            df.loc[idx, "abstract_source"] = "unavailable"
        time.sleep(0.12)  # ~8 req/s — polite pool

    print(f"  Fetched {fetched}/{n_miss}. "
          f"{n_miss - fetched} abstracts remain unavailable.\n")

    # ── Step 2: Keyword scoring ───────────────────────────────
    print("Step 2 — Running keyword screening ...")
    verdicts, reasons, score_jsons = [], [], []
    for _, row in df.iterrows():
        v, r, s = score(str(row.get("title", "")), str(row.get("abstract", "")))
        verdicts.append(v)
        reasons.append(r)
        score_jsons.append(json.dumps(s))

    df["kw_verdict"] = verdicts
    df["kw_reason"]  = reasons
    df["kw_scores"]  = score_jsons

    # ── Step 3: Add/restore user columns ─────────────────────
    df["user_verdict"] = pd.NA
    df["user_notes"]   = ""
    df["reviewed_at"]  = ""

    for i, row in df.iterrows():
        key = str(row["openalex_id"])
        if key in saved_verdicts:
            s = saved_verdicts[key]
            df.loc[i, "user_verdict"] = s["user_verdict"]
            df.loc[i, "user_notes"]   = s.get("user_notes", "")
            df.loc[i, "reviewed_at"]  = s.get("reviewed_at", "")

    # ── Step 4: Save ─────────────────────────────────────────
    OUT_CSV.parent.mkdir(exist_ok=True)
    df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    print(f"\nSaved {len(df)} rows → {OUT_CSV}")

    # Summary
    vc   = df["kw_verdict"].value_counts()
    revd = df["user_verdict"].notna().sum()
    print("\n── Keyword screening results ─────────────────────────")
    for v in ["green", "orange", "red"]:
        n = vc.get(v, 0)
        icon = {"green": "🟢", "orange": "🟠", "red": "🔴"}[v]
        bar  = "█" * (n // 10) + "░" * ((n // 10 - n // 10) + 1)
        print(f"  {icon} {v.upper():<8}: {n:>4}  ({n / len(df) * 100:.0f}%)")
    print(f"\n  Already reviewed by user: {revd}")
    print(f"\nNext step:\n  streamlit run analysis/review_app.py")


if __name__ == "__main__":
    main()
