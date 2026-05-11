"""
TrendFish 2026 — unified literature review management app.

Tabs:
  📋 Article   — article structure, metadata, section outline
  ✍️ Sections  — write and edit section drafts with live word counts
  🔎 Strategy  — document the 3-step search strategy
  🔍 Search    — query OpenAlex, build boolean searches, manage pool
  📄 Review    — paper-by-paper screening with keyword scoring
  📊 Analysis  — STEEP gap, megatrends, signals, keywords
  📤 Outputs   — generate manuscript, export data

Usage:
    python -m streamlit run analysis/review_app.py
"""

import json
import subprocess
import sys
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path

import openpyxl
import pandas as pd
import plotly.express as px
import requests
import streamlit as st

BASE          = Path(__file__).parent.parent
DATA_DIR      = BASE / "data"
DOCS_DIR      = BASE / "docs"
IN_CSV        = DATA_DIR / "full_text_review.csv"
KW_FILE       = DATA_DIR / "review_keywords.json"
SRCH_FILE     = DATA_DIR / "search_config.json"
EXCEL_FILE    = DATA_DIR / "manual_literature_search.xlsx"
SIGNALS_FILE  = DATA_DIR / "signals.json"
ARTICLE_FILE  = DATA_DIR / "article_structure.json"
SECTIONS_FILE = DATA_DIR / "sections.json"
STRATEGY_FILE = DATA_DIR / "search_strategy.json"

DATA_DIR.mkdir(exist_ok=True)
DOCS_DIR.mkdir(exist_ok=True)

STEEP_CATS   = ["Environmental", "Technological", "Economic", "Social", "Political", "Other"]
STEEP_COLORS = {
    "Environmental": "#2166ac", "Technological": "#1a9641",
    "Economic":      "#f46d43", "Social":        "#d73027",
    "Political":     "#762a83", "Other":         "#999999",
}
STEEP_RULES = {
    "Environmental": [
        "climate","environment","biodiversity","ecosystem","ocean","marine",
        "sea","water","pollution","habitat","carbon","emission","green",
        "nature conservation","nature-based","natural resource","resource scarcity",
        "ecology","fish stock","overfishing","fishing","fisheries","fishery",
        "fisher","aquaculture","aquatic","seafood","harvest","catch",
        "bycatch","by-catch","food web","trophic",
    ],
    "Technological": [
        "technolog","digital","automation","artificial intelligence","ai ",
        "innovation","internet","cyber","robot","autonomous","biotech",
        "synthetic biology","hyperconnectiv","data-driven","machine learning",
    ],
    "Economic": [
        "economic","trade","market","finance","fiscal","gdp","consumption",
        "globali","growth","income","employment","labour","labor","industry",
        "investment","supply chain","demand","subsid","cost","price","value chain",
    ],
    "Social": [
        "social","demograph","population","urban","aging","health","inequalit",
        "migration","culture","education","wellbeing","well-being","community",
        "consumer","lifestyle","food security","gender","livelihood",
        "artisanal","small-scale","indigenous",
    ],
    "Political": [
        "politic","governance","geopolit","security","power","regulation","law",
        "policy","democracy","conflict","management","quota","treaty","iuu","compliance",
    ],
}

DEFAULT_KW = {
    "fish_strong": [
        "fisheries", "fishery", "fishing industry", "commercial fishing",
        "marine capture", "fish stock", "fish catch", "artisanal fish",
        "small-scale fish", "seafood market", "aquaculture production", "mariculture",
    ],
    "fish_med": [
        "fishing", "fisher", "aquaculture", "seafood", "trawl", "seine",
        "longline", "gillnet", "msc certified", "fisheries management",
        "fisheries governance", "fish production",
    ],
    "fish_context": [
        "marine ecosystem", "coastal community", "pelagic", "demersal",
        "bycatch", "by-catch", "overfish", "fish population", "fish species",
        "ocean harvest", "sea fish",
    ],
    "future_strong": [
        "future scenario", "fisheries scenario", "future of fisheries",
        "fisheries futures", "foresight", "horizon scan", "prospective",
        "future projection", "future outlook", "futures study", "futures analysis",
    ],
    "future_med": [
        "scenario", "projection", "forecast", "by 2030", "by 2040", "by 2050",
        "by 2060", "by 2100", "under rcp", "under ssp", "climate scenario",
        "future trend", "expected impact", "anticipated change", "long-term outlook",
        "future impact", "future change", "will be", "will likely",
        "is expected", "are projected",
    ],
    "future_weak": [
        "trend", "driver", "future", "model", "outlook", "prospect",
        "long-term", "simulate", "predict",
    ],
    "excl_retro": [
        "retrospective", "over the past", "historical trend", "historical analysis",
        "past decade", "years of data", "long-term data", "time series analys",
    ],
    "excl_bio": [
        "phylogenetic", "phylogeny", "genomic", "genome sequen", "dna barcod",
        "genetic diversity", "microsatellite", "population genetic", "allele", "genotype",
    ],
    "excl_nonfish": [
        "agriculture", "crop production", "livestock", "poultry", "bovine",
        "swine", "wheat", "maize", "terrestrial", "cattle farm", "dairy",
    ],
}

DEFAULT_QUERIES = [
    "Future Trend Impact Analysis Fisheries",
    "Future studies Fisheries Research",
    "Future Fisheries Scenarios",
    "Future of Fisheries",
    "Future Trends and Drivers Fisheries",
    "Megatrend Impacts Fisheries",
    "Exploring Megatrends Fisheries",
]

KEYWORD_META = {
    "fish_strong":  ("🐟 Fisheries — strong",      "×3 weight", "#0d6efd"),
    "fish_med":     ("🐟 Fisheries — medium",      "×2 weight", "#0d6efd"),
    "fish_context": ("🐟 Fisheries — context",     "×1 weight", "#0d6efd"),
    "future_strong":("🔭 Future — strong",         "×4 weight", "#6f42c1"),
    "future_med":   ("🔭 Future — medium",         "×2 weight", "#6f42c1"),
    "future_weak":  ("🔭 Future — weak",           "×1 weight", "#6f42c1"),
    "excl_retro":   ("🚫 Excl. retrospective",     "triggers red if future=0",      "#dc3545"),
    "excl_bio":     ("🚫 Excl. biology-only",      "triggers red if fish≤2+future<2","#dc3545"),
    "excl_nonfish": ("🚫 Excl. non-fisheries",     "triggers red if fish=0",        "#dc3545"),
}

VERDICT_COLORS = {"green": "#28a745", "orange": "#fd7e14", "red": "#dc3545"}
VERDICTS = {
    "include": ("🟢", "INCLUDE", "#d4edda", VERDICT_COLORS["green"]),
    "unsure":  ("🟠", "UNSURE",  "#fff3cd", VERDICT_COLORS["orange"]),
    "exclude": ("🔴", "EXCLUDE", "#f8d7da", VERDICT_COLORS["red"]),
}
OA_STATUS = {
    "true":  ("🔓", "Open access", "#28a745"),
    "false": ("🔒", "Subscription", "#6c757d"),
}
POOL_COLS = [
    "openalex_id", "title", "year", "authors", "journal", "doi",
    "abstract", "abstract_source", "search_query", "date_added",
    "kw_verdict", "kw_reason", "user_verdict", "user_notes", "reviewed_at",
    "pdf_url", "is_oa",
]

DEFAULT_ARTICLE = {
    "title": "Drivers, Trends, and Signals in Fisheries Research: A Structured Literature Review and Horizon Scan",
    "subtitle": "",
    "target_journal": "Reviews in Fish Biology and Fisheries",
    "max_words": 12000,
    "authors": [
        {
            "name": "Timo Szczepanska",
            "affiliation": "Norwegian College of Fishery Science, UiT The Arctic University of Norway",
            "email": "tsz000@uit.no",
            "orcid": "0000-0003-2442-8223",
        }
    ],
    "keywords": [
        "horizon scanning", "megatrends", "fisheries futures",
        "STEEP framework", "systematic literature review", "weak signals", "foresight",
    ],
    "research_questions": [
        "What are the dominant drivers and megatrends that cross-sectoral foresight literature identifies as shaping the global environment?",
        "How do fisheries-specific futures studies compare to the cross-sectoral literature in their STEEP coverage?",
        "What weak and emerging signals at the margins of current fisheries knowledge warrant systematic attention?",
    ],
    "sections": [
        {"id": "abstract",              "title": "Abstract",                "target_words": 300},
        {"id": "introduction",          "title": "1. Introduction",         "target_words": 800},
        {"id": "conceptual_background", "title": "2. Conceptual Background","target_words": 700},
        {"id": "methods",               "title": "3. Methods",              "target_words": 1200},
        {"id": "results",               "title": "4. Results",              "target_words": 1200},
        {"id": "discussion",            "title": "5. Discussion",           "target_words": 1200},
        {"id": "conclusions",           "title": "6. Conclusions",          "target_words": 400},
        {"id": "references",            "title": "References",              "target_words": 0},
    ],
}

SECTION_DOCX_MAP = {
    "abstract":              "Abstract_Trends_2026.docx",
    "introduction":          "Introduction_Trends_2026.docx",
    "conceptual_background": "ConceptualBackground_Trends_2026.docx",
    "methods":               "Methods_Trends_2026.docx",
    "results":               "Results_Trends_2026.docx",
    "discussion":            "Discussion_Trends_2026.docx",
    "conclusions":           "Conclusions_Trends_2026.docx",
}

DEFAULT_STRATEGY = {
    "strategy1": {
        "purpose": "Structured search of peer-reviewed fisheries futures literature",
        "database": "OpenAlex",
        "date": "August 2024",
        "queries": list(DEFAULT_QUERIES),
        "n_retrieved": 2100,
        "n_deduped": 1621,
        "n_screened": 720,
        "n_manual": 60,
        "inclusion_criteria": [
            "Published in peer-reviewed journal",
            "Addresses the future of fisheries (forward-looking)",
            "Published 2000–2024",
        ],
        "exclusion_criteria": [
            "Purely retrospective — no future orientation",
            "No fisheries relevance detected",
            "Non-fisheries topic (agriculture, livestock, etc.)",
        ],
    },
    "strategy2": {
        "purpose": "Cross-sectoral scan of megatrend and foresight publications",
        "sources": [
            "Consultancy reports (McKinsey, Deloitte, PwC, KPMG)",
            "Government foresight agencies (OECD, EU, national bodies)",
            "Academic megatrend and foresight reviews",
            "NGO and think tank publications",
        ],
        "date": "10–12 August 2024",
        "n_entries": 145,
    },
    "strategy3": {
        "purpose": "Horizon scan for weak and emerging signals",
        "method": "Continuous monitoring of news, policy, and grey literature for early indicators of change in fisheries systems",
        "status": "Living document",
    },
}


# ─────────────────────────────────────────────────────────────
# Helpers — keywords, search config, scoring
# ─────────────────────────────────────────────────────────────

def load_keywords() -> dict:
    if KW_FILE.exists():
        try:
            return json.loads(KW_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {k: list(v) for k, v in DEFAULT_KW.items()}


def save_keywords(kw: dict):
    KW_FILE.write_text(json.dumps(kw, indent=2, ensure_ascii=False), encoding="utf-8")


def load_search_config() -> dict:
    if SRCH_FILE.exists():
        try:
            data = json.loads(SRCH_FILE.read_text(encoding="utf-8"))
            if "queries" in data and "saved_queries" not in data:
                data["saved_queries"] = data.pop("queries")
            return data
        except Exception:
            pass
    return {"saved_queries": list(DEFAULT_QUERIES), "max_results": 200, "email": "tsz000@uit.no"}


def save_search_config(cfg: dict):
    SRCH_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


def score_paper(title: str, abstract: str, kw: dict) -> dict:
    text         = ((title or "") + " " + (abstract or "")).lower()
    has_abstract = bool(abstract and len(str(abstract).strip()) > 50)

    f_strong_hits = [t for t in kw.get("fish_strong",  []) if t in text]
    f_med_hits    = [t for t in kw.get("fish_med",     []) if t in text]
    f_ctx_hits    = [t for t in kw.get("fish_context", []) if t in text]
    fish = min(len(f_strong_hits)*3 + len(f_med_hits)*2 + len(f_ctx_hits), 8)

    fs_hits = [t for t in kw.get("future_strong", []) if t in text]
    fm_hits = [t for t in kw.get("future_med",    []) if t in text]
    fw_hits = [t for t in kw.get("future_weak",   []) if t in text]
    future = min(len(fs_hits)*4 + len(fm_hits)*2 + len(fw_hits), 10)

    retro = [t for t in kw.get("excl_retro",   []) if t in text]
    bio   = [t for t in kw.get("excl_bio",     []) if t in text]
    nonf  = [t for t in kw.get("excl_nonfish", []) if t in text]

    if fish == 0 and nonf:
        verdict, reason = "red",    f"Non-fisheries topic ({nonf[0]})"
    elif fish == 0:
        verdict, reason = "red",    "No fisheries relevance detected"
    elif future == 0 and retro:
        verdict, reason = "red",    f"Retrospective, no future orientation ({retro[0]})"
    elif bio and fish <= 2 and future < 2:
        verdict, reason = "red",    f"Biology/genetics, no fisheries futures ({bio[0]})"
    elif fish >= 4 and future >= 6 and not retro:
        verdict, reason = "green",  "Strong fisheries + future orientation"
    elif fish >= 3 and future >= 4 and not retro:
        verdict, reason = "green",  "Clear fisheries with future-oriented analysis"
    elif fish >= 2 and future >= 6:
        verdict, reason = "green",  "Good fisheries with strong future signals"
    elif not has_abstract:
        verdict, reason = "orange", "Abstract unavailable — title only"
    elif future == 0:
        verdict, reason = "orange", "Fisheries-relevant but no future orientation"
    elif fish >= 2 and future >= 2:
        verdict, reason = "orange", "Borderline — weak fisheries + future signals"
    else:
        verdict, reason = "orange", f"Insufficient signals (fish={fish}, future={future})"

    return {
        "verdict": verdict, "reason": reason,
        "fish": fish, "future": future,
        "fish_hits":   f_strong_hits + f_med_hits + f_ctx_hits,
        "future_hits": fs_hits + fm_hits + fw_hits,
        "retro": retro, "bio": bio, "nonfish": nonf,
    }


def build_query_from_rows(rows: list) -> str:
    parts: list = []
    for i, row in enumerate(rows):
        term = (row.get("term") or "").strip()
        if not term:
            continue
        quoted = f'"{term}"' if " " in term else term
        if i == 0 or not row.get("op"):
            parts.append(quoted)
        else:
            parts.append(f"{row['op']} {quoted}")
    return " ".join(parts)


# ─────────────────────────────────────────────────────────────
# Helpers — OpenAlex
# ─────────────────────────────────────────────────────────────

def reconstruct_abstract(inv: dict) -> str:
    pairs = [(pos, word) for word, positions in inv.items() for pos in positions]
    return " ".join(w for _, w in sorted(pairs))


def _parse_work(w: dict, query: str) -> dict:
    oa_id = w.get("id", "")
    auth_list = [
        a.get("author", {}).get("display_name", "")
        for a in (w.get("authorships") or [])
    ]
    authors = "; ".join(filter(None, auth_list[:5]))
    if len(auth_list) > 5:
        authors += " et al."
    loc     = (w.get("primary_location") or {})
    src     = (loc.get("source") or {})
    journal = src.get("display_name", "")
    doi     = (w.get("doi") or "").replace("https://doi.org/", "")
    inv     = w.get("abstract_inverted_index")
    abstract = reconstruct_abstract(inv) if inv else ""
    oa      = w.get("open_access") or {}
    pdf_url = oa.get("oa_url") or ""
    is_oa   = str(oa.get("is_oa", "")).lower()
    return {
        "openalex_id":     oa_id,
        "title":           w.get("title", ""),
        "year":            str(w.get("publication_year", "")),
        "authors":         authors,
        "journal":         journal,
        "doi":             doi,
        "abstract":        abstract,
        "abstract_source": "openalex" if abstract else "unavailable",
        "search_query":    query,
        "date_added":      datetime.now().strftime("%Y-%m-%d"),
        "kw_verdict": "", "kw_reason": "",
        "user_verdict": "", "user_notes": "", "reviewed_at": "",
        "pdf_url": pdf_url, "is_oa": is_oa,
    }


def fetch_openalex(query: str, max_results: int, email: str, log_fn=None) -> list:
    collected = []
    per_page  = min(200, max_results)
    page      = 1
    while len(collected) < max_results:
        try:
            r = requests.get(
                "https://api.openalex.org/works",
                params={
                    "search":    query,
                    "per-page":  per_page,
                    "page":      page,
                    "mailto":    email,
                    "select":    ("id,title,publication_year,authorships,"
                                  "primary_location,doi,abstract_inverted_index,open_access"),
                },
                timeout=20,
            )
            if r.status_code != 200:
                if log_fn: log_fn(f"  HTTP {r.status_code} — stopping pagination")
                break
            batch = r.json().get("results", [])
            if not batch:
                break
            collected.extend(batch)
            if log_fn: log_fn(f"  Page {page}: {len(batch)} records (total {len(collected)})")
            if len(batch) < per_page:
                break
            page += 1
            time.sleep(0.12)
        except Exception as exc:
            if log_fn: log_fn(f"  Request error: {exc}")
            break
    return [_parse_work(w, query) for w in collected[:max_results]]


# ─────────────────────────────────────────────────────────────
# Helpers — pool (CSV)
# ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=30)
def load_pool() -> pd.DataFrame:
    if IN_CSV.exists():
        df = pd.read_csv(IN_CSV, dtype=str).fillna("")
        for c in POOL_COLS:
            if c not in df.columns:
                df[c] = ""
        return df
    return pd.DataFrame(columns=POOL_COLS)


def save_pool(df: pd.DataFrame):
    df.to_csv(IN_CSV, index=False, encoding="utf-8-sig")
    st.cache_data.clear()


def save_verdict_to_pool(df: pd.DataFrame, idx: int, verdict: str, notes: str):
    df.loc[idx, "user_verdict"] = verdict
    df.loc[idx, "user_notes"]   = notes
    df.loc[idx, "reviewed_at"]  = datetime.now().strftime("%Y-%m-%d %H:%M")
    save_pool(df)


def merge_records(new_records: list, kw: dict) -> tuple:
    pool = load_pool()
    existing = set(pool["openalex_id"].tolist())
    to_add, n_dup = [], 0
    seen_now: set = set()
    for rec in new_records:
        oid = rec["openalex_id"]
        if oid in existing or oid in seen_now:
            n_dup += 1
            continue
        seen_now.add(oid)
        sc = score_paper(rec["title"], rec["abstract"], kw)
        rec["kw_verdict"] = sc["verdict"]
        rec["kw_reason"]  = sc["reason"]
        to_add.append(rec)
    if to_add:
        new_df = pd.DataFrame(to_add)[POOL_COLS]
        combined = pd.concat([pool, new_df], ignore_index=True)
        save_pool(combined)
    return len(to_add), n_dup


def _bibtex_escape(s: str) -> str:
    return s.replace("\\", "").replace("{", "").replace("}", "").replace('"', "'")


def _bibtex_key(row: pd.Series, seen=None) -> str:
    authors = str(row.get("authors", "")).strip()
    year    = str(row.get("year", "")).strip()
    title   = str(row.get("title", "")).strip()
    first_author = authors.split(";")[0].strip()
    last_name = first_author.split()[-1] if first_author else "unknown"
    last_name = "".join(c for c in last_name.lower() if c.isalpha())
    skip = {"a", "an", "the", "of", "in", "on", "for", "and", "or"}
    words = [w for w in title.lower().split() if w.isalpha() and w not in skip]
    first_word = words[0] if words else "notitle"
    base = f"{last_name}_{year}_{first_word}"
    if seen is None:
        return base
    key, suffix = base, 2
    while key in seen:
        key = f"{base}{suffix}"
        suffix += 1
    return key


def _make_bibtex_entry(row: pd.Series, key: str) -> str:
    title   = _bibtex_escape(str(row.get("title",   "")).strip())
    authors_raw = str(row.get("authors", "")).strip()
    year    = str(row.get("year",    "")).strip()
    journal = _bibtex_escape(str(row.get("journal", "")).strip())
    doi     = str(row.get("doi",     "")).strip()
    pdf_url = str(row.get("pdf_url", "")).strip()
    oa_id   = str(row.get("openalex_id", "")).strip()
    parts = [p.strip().rstrip(".") for p in authors_raw.replace(" et al.", "").split(";") if p.strip()]
    bib_authors = []
    for p in parts:
        tokens = p.split()
        if len(tokens) >= 2:
            bib_authors.append(f"{tokens[-1]}, {' '.join(tokens[:-1])}")
        else:
            bib_authors.append(p)
    author_str = " and ".join(bib_authors)
    url = f"https://doi.org/{doi}" if doi else pdf_url or oa_id
    lines = [f"@article{{{key},"]
    if title:      lines.append(f"  title   = {{{title}}},")
    if author_str: lines.append(f"  author  = {{{author_str}}},")
    if year:       lines.append(f"  year    = {{{year}}},")
    if journal:    lines.append(f"  journal = {{{journal}}},")
    if doi:        lines.append(f"  doi     = {{{doi}}},")
    if url:        lines.append(f"  url     = {{{url}}},")
    lines.append("}")
    return "\n".join(lines)


def generate_bibtex(df: pd.DataFrame) -> str:
    seen: set = set()
    entries = []
    for _, row in df.iterrows():
        key = _bibtex_key(row, seen)
        seen.add(key)
        entries.append(_make_bibtex_entry(row, key))
    return "\n\n".join(entries)


@st.cache_data
def compute_live_verdicts(pool_mtime: float, kw_json: str) -> list:
    _pool = load_pool()
    _kw   = json.loads(kw_json)
    return [score_paper(r["title"], r["abstract"], _kw)["verdict"] for _, r in _pool.iterrows()]


@st.cache_data
def _cached_bibtex(bib_filter: str, pool_mtime: float) -> bytes:
    _pool = load_pool()
    masks = {
        "Included papers":   _pool["user_verdict"] == "include",
        "Included + Unsure": _pool["user_verdict"].isin(["include", "unsure"]),
        "All reviewed":      _pool["user_verdict"].str.strip().astype(bool),
        "Entire pool":       pd.Series([True] * len(_pool), index=_pool.index),
    }
    return generate_bibtex(_pool[masks[bib_filter]]).encode("utf-8")


def enrich_unpaywall(email: str, log_fn=None) -> tuple:
    pool = load_pool()
    needs = pool[
        (pool["doi"].str.strip() != "") &
        (pool["pdf_url"].str.strip() == "")
    ]
    n_found = 0
    for i, (idx, row) in enumerate(needs.iterrows()):
        doi = row["doi"].strip()
        if log_fn: log_fn(f"[{i+1}/{len(needs)}] {doi}")
        try:
            r = requests.get(
                f"https://api.unpaywall.org/v2/{doi}",
                params={"email": email},
                timeout=10,
            )
            if r.status_code == 200:
                data = r.json()
                loc  = data.get("best_oa_location") or {}
                url  = loc.get("url_for_pdf") or loc.get("url") or ""
                pool.loc[idx, "pdf_url"] = url
                pool.loc[idx, "is_oa"]   = str(data.get("is_oa", "")).lower()
                if url: n_found += 1
        except Exception:
            pass
        time.sleep(0.12)
    save_pool(pool)
    return n_found, len(needs)


# ─────────────────────────────────────────────────────────────
# Helpers — STEEP + Excel
# ─────────────────────────────────────────────────────────────

def classify_steep(text: str) -> str:
    if not text or str(text).strip().lower() in ("none", "nan", "?", ""):
        return "Other"
    t = str(text).lower()
    for cat, kws in STEEP_RULES.items():
        if any(kw in t for kw in kws):
            return cat
    return "Other"


@st.cache_data
def load_megatrends(_mtime: float = 0) -> pd.DataFrame:
    if not EXCEL_FILE.exists():
        return pd.DataFrame(columns=["trend","details","author","source","title",
                                     "comments","doi","year","source_type","steep"])
    df = pd.read_excel(EXCEL_FILE, sheet_name="(Mega-)trends and drivers")
    df.columns = ["trend","details","author","source","title","comments","doi","year","source_type"]
    df = df.dropna(subset=["trend"])
    df = df[~df["trend"].astype(str).str.strip().isin(["","nan"])]
    type_map = {
        "Government? (finnish innovation fund)": "Government",
        "Oxfam discussion paper": "NGO / Think Tank",
        "Book Chapter": "Book / Chapter",
        "Book":         "Book / Chapter",
        "?":            "Unclassified",
    }
    df["source_type"] = df["source_type"].replace(type_map)
    df["steep"] = df.apply(
        lambda r: classify_steep(str(r["trend"]) + " " + str(r.get("details",""))), axis=1
    )
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    return df.reset_index(drop=True)


@st.cache_data
def load_fisheries_futures(_mtime: float = 0) -> pd.DataFrame:
    if not EXCEL_FILE.exists():
        return pd.DataFrame(columns=["trend","author","title","year","steep"])
    df = pd.read_excel(EXCEL_FILE, sheet_name="Futures of Fisheries")
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    cols = ["trend","details","author","source","title","comments","doi","year","source_type"]
    df.columns = cols[:len(df.columns)]
    keep = [c for c in ["trend","details","author","title","year","doi"] if c in df.columns]
    df = df[keep].dropna(how="all")
    df["steep"] = df.apply(
        lambda r: classify_steep(
            " ".join(str(v) for v in [r.get("trend",""), r.get("details",""), r.get("title","")]
                     if pd.notna(v))
        ), axis=1,
    )
    df["year"] = pd.to_numeric(df.get("year", pd.Series(dtype=float)), errors="coerce")
    return df.reset_index(drop=True)


def load_signals() -> list:
    if SIGNALS_FILE.exists():
        try:
            return json.loads(SIGNALS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []


def save_signals(signals: list):
    SIGNALS_FILE.write_text(json.dumps(signals, indent=2, ensure_ascii=False), encoding="utf-8")


def add_megatrend_to_excel(row: dict):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["(Mega-)trends and drivers"]
    ws.append([
        row.get("trend",""), row.get("details",""), row.get("author",""),
        row.get("source",""), row.get("title",""), row.get("comments",""),
        row.get("doi",""), row.get("year") or None, row.get("source_type",""),
    ])
    wb.save(EXCEL_FILE)
    load_megatrends.clear()
    load_fisheries_futures.clear()


# ─────────────────────────────────────────────────────────────
# Helpers — article, sections, strategy
# ─────────────────────────────────────────────────────────────

def load_article_structure() -> dict:
    if ARTICLE_FILE.exists():
        try:
            data = json.loads(ARTICLE_FILE.read_text(encoding="utf-8"))
            # Ensure a references section exists (migration for older article_structure.json files)
            existing_ids = {s["id"] for s in data.get("sections", [])}
            if "references" not in existing_ids:
                data.setdefault("sections", []).append(
                    {"id": "references", "title": "References", "target_words": 0, "done": False}
                )
                save_article_structure(data)
            return data
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_ARTICLE))


def save_article_structure(d: dict):
    ARTICLE_FILE.write_text(json.dumps(d, indent=2, ensure_ascii=False), encoding="utf-8")


def seed_sections_from_docx() -> dict:
    try:
        from docx import Document as _DocxDoc
    except ImportError:
        return {sid: "" for sid in SECTION_DOCX_MAP}
    result = {}
    heading_prefixes = ("Abstract", "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.")
    for sec_id, fname in SECTION_DOCX_MAP.items():
        fpath = DOCS_DIR / fname
        if fpath.exists():
            try:
                d = _DocxDoc(str(fpath))
                paras = [p.text for p in d.paragraphs if p.text.strip()]
                # Drop the first paragraph if it's a section heading
                if paras and any(paras[0].startswith(px) for px in heading_prefixes):
                    paras = paras[1:]
                result[sec_id] = "\n\n".join(paras)
            except Exception:
                result[sec_id] = ""
        else:
            result[sec_id] = ""
    return result


def load_sections() -> dict:
    if SECTIONS_FILE.exists():
        try:
            return json.loads(SECTIONS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    seeded = seed_sections_from_docx()
    save_sections(seeded)
    return seeded


def save_sections(d: dict):
    SECTIONS_FILE.write_text(json.dumps(d, indent=2, ensure_ascii=False), encoding="utf-8")


def load_search_strategy() -> dict:
    if STRATEGY_FILE.exists():
        try:
            return json.loads(STRATEGY_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_STRATEGY))


def save_search_strategy(d: dict):
    STRATEGY_FILE.write_text(json.dumps(d, indent=2, ensure_ascii=False), encoding="utf-8")


def generate_manuscript_bytes(article: dict, sections: dict) -> bytes:
    try:
        from docx import Document as _DocxDoc
        from docx.shared import Pt, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    except ImportError:
        return b""

    doc = _DocxDoc()
    for sec in doc.sections:
        sec.top_margin    = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin   = Cm(3)
        sec.right_margin  = Cm(3)

    # Title page
    p = doc.add_paragraph()
    r = p.add_run(article.get("title", ""))
    r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(14)
    p.paragraph_format.space_after = Pt(18)

    for author in article.get("authors", []):
        p = doc.add_paragraph()
        parts = [author.get("name",""), author.get("affiliation",""), author.get("email","")]
        if author.get("orcid"):
            parts.append(f"ORCID: {author['orcid']}")
        r = p.add_run("\n".join(x for x in parts if x))
        r.font.name = "Times New Roman"; r.font.size = Pt(11)
        p.paragraph_format.space_after = Pt(6)

    journal = article.get("target_journal", "")
    if journal:
        p = doc.add_paragraph()
        r = p.add_run(f"Target journal: {journal}")
        r.italic = True; r.font.name = "Times New Roman"; r.font.size = Pt(10)
        p.paragraph_format.space_after = Pt(12)

    kws = article.get("keywords", [])
    if kws:
        p = doc.add_paragraph()
        r = p.add_run(f"Keywords: {'; '.join(kws)}")
        r.font.name = "Times New Roman"; r.font.size = Pt(10)

    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)

    for sec_def in article.get("sections", []):
        sec_id    = sec_def["id"]
        sec_title = sec_def["title"]
        text      = sections.get(sec_id, "").strip()
        if not text:
            continue
        p = doc.add_paragraph()
        r = p.add_run(sec_title)
        r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(12)
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after  = Pt(6)
        for block in text.split("\n\n"):
            block = block.strip()
            if not block:
                continue
            p = doc.add_paragraph(block)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for run in p.runs:
                run.font.name = "Times New Roman"; run.font.size = Pt(11)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.first_line_indent = Cm(0.75)

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# App layout
# ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="TrendFish 2026",
    page_icon="🐟",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
/* Legacy styles */
.abstract-box {
    background:#f8f9fa; border-left:4px solid #dee2e6;
    padding:12px 16px; border-radius:4px;
    font-size:0.92rem; line-height:1.6;
    max-height:260px; overflow-y:auto;
}
.score-bar-wrap { background:#e9ecef; border-radius:6px; height:14px; margin:2px 0 6px 0; }
.score-bar      { border-radius:6px; height:14px; }
.kw-match       { font-size:0.78rem; color:#555; margin:2px 0; }

/* Design system */
.lr-card {
    background:#fff; border:1px solid #dee2e6; border-radius:8px;
    padding:16px 20px; margin-bottom:12px;
}
.lr-section-header {
    font-size:1.05rem; font-weight:700; color:#212529;
    border-bottom:2px solid #0d6efd; padding-bottom:4px; margin-bottom:12px;
}
.lr-badge {
    display:inline-block; padding:2px 10px; border-radius:12px;
    font-size:0.78rem; font-weight:600; color:#fff;
}
.lr-metric { text-align:center; padding:10px; }
.lr-metric-val { font-size:1.8rem; font-weight:700; }
.lr-metric-lbl { font-size:0.78rem; color:#6c757d; }
.lr-progress-wrap { background:#e9ecef; border-radius:4px; height:10px; }
.lr-progress-bar  { background:#0d6efd; border-radius:4px; height:10px; }
</style>
""", unsafe_allow_html=True)


# ── Session state ─────────────────────────────────────────────

if "paper_idx"        not in st.session_state: st.session_state.paper_idx        = 0
if "filter_verdict"   not in st.session_state: st.session_state.filter_verdict   = "All"
if "filter_reviewed"  not in st.session_state: st.session_state.filter_reviewed  = "All"
if "kw"               not in st.session_state: st.session_state.kw               = load_keywords()
if "search_cfg"       not in st.session_state: st.session_state.search_cfg       = load_search_config()
if "search_preview"   not in st.session_state: st.session_state.search_preview   = []
if "search_rows"      not in st.session_state:
    st.session_state.search_rows    = [{"id": 0, "op": None, "term": ""}]
    st.session_state.search_next_id = 1
if "signals"          not in st.session_state: st.session_state.signals          = load_signals()
if "editing_signal"   not in st.session_state: st.session_state.editing_signal   = None
if "active_section"   not in st.session_state: st.session_state.active_section   = "abstract"

kw  = st.session_state.kw
cfg = st.session_state.search_cfg

# ── Load pool ─────────────────────────────────────────────────

pool          = load_pool()
total         = len(pool)
reviewed_mask = pool["user_verdict"].str.strip().astype(bool)
n_reviewed    = int(reviewed_mask.sum())

pool_mtime = IN_CSV.stat().st_mtime if IN_CSV.exists() else 0.0
kw_json    = json.dumps(st.session_state.kw, sort_keys=True)
live_verdicts = pd.Series(
    compute_live_verdicts(pool_mtime, kw_json) if total > 0 else [],
    index=pool.index,
)

# ── Sidebar ───────────────────────────────────────────────────

with st.sidebar:
    st.title("🐟 TrendFish 2026")

    with st.expander("❓ Quick start", expanded=False):
        st.markdown("""
**Seven tabs — use in order:**

1. **📋 Article** — define article structure and metadata
2. **✍️ Sections** — write and edit section drafts
3. **🔎 Strategy** — document your search strategy
4. **🔍 Search** — fetch papers from OpenAlex
5. **📄 Review** — screen papers one by one
6. **📊 Analysis** — STEEP gap, megatrends, signals
7. **📤 Outputs** — generate manuscript and exports

Progress saved to `data/` automatically.
""")

    st.metric("Papers in pool", total)
    st.metric("Reviewed", f"{n_reviewed} / {total}")
    if total:
        st.progress(n_reviewed / total)

    st.markdown("**Your verdicts**")
    for key, (icon, label, bg, col) in VERDICTS.items():
        n = int((pool["user_verdict"] == key).sum()) if total else 0
        st.markdown(
            f"<span style='color:{col};font-weight:700'>{icon} {label}</span>: {n}",
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.subheader("Review filters")

    if total:
        vc = live_verdicts.value_counts()
        st.markdown(
            f"<span style='color:#28a745'>● {vc.get('green',0)}</span> &nbsp;"
            f"<span style='color:#fd7e14'>● {vc.get('orange',0)}</span> &nbsp;"
            f"<span style='color:#dc3545'>● {vc.get('red',0)}</span>",
            unsafe_allow_html=True,
        )

    ai_filter = st.selectbox(
        "AI keyword verdict",
        ["All", "green", "orange", "red"],
        index=["All", "green", "orange", "red"].index(st.session_state.filter_verdict),
    )
    reviewed_filter = st.selectbox(
        "Review status",
        ["All", "Unreviewed", "Reviewed"],
        index=["All", "Unreviewed", "Reviewed"].index(st.session_state.filter_reviewed),
    )
    st.session_state.filter_verdict  = ai_filter
    st.session_state.filter_reviewed = reviewed_filter

# ── Tabs ──────────────────────────────────────────────────────

(tab_article, tab_sections, tab_strategy,
 tab_search, tab_review, tab_analysis, tab_draft, tab_outputs) = st.tabs([
    "📋 Article", "📝 Sections", "🔎 Search Strategy",
    "🔍 Search", "📄 Screening", "📊 Analysis", "✍️ Draft", "📤 Outputs",
])


# ══════════════════════════════════════════════════════════════
# TAB 1 — ARTICLE
# ══════════════════════════════════════════════════════════════

with tab_article:
    st.markdown("<div class='lr-section-header'>📋 Article Planner</div>", unsafe_allow_html=True)

    article = load_article_structure()

    sub = st.radio(
        "view", ["📋 Metadata", "📑 Section Outline"],
        horizontal=True, label_visibility="collapsed",
    )

    if sub == "📋 Metadata":
        with st.form("article_meta_form"):
            a1, a2 = st.columns(2)
            with a1:
                new_title   = st.text_input("Title", value=article.get("title",""))
                new_journal = st.text_input("Target journal", value=article.get("target_journal",""))
            with a2:
                new_subtitle = st.text_input("Subtitle", value=article.get("subtitle",""))

            st.markdown("**Authors**")
            authors_df = pd.DataFrame(article.get("authors", []))
            if authors_df.empty:
                authors_df = pd.DataFrame(columns=["name","affiliation","email","orcid"])
            edited_authors = st.data_editor(
                authors_df, num_rows="dynamic", use_container_width=True,
                key="authors_de",
                column_config={
                    "name":        st.column_config.TextColumn("Name"),
                    "affiliation": st.column_config.TextColumn("Affiliation"),
                    "email":       st.column_config.TextColumn("Email"),
                    "orcid":       st.column_config.TextColumn("ORCID"),
                },
            )
            st.caption("To remove a row: select its checkbox on the left, then press Delete.")

            st.markdown("**Keywords**")
            kw_df = pd.DataFrame({"keyword": article.get("keywords", [])})
            edited_kws = st.data_editor(
                kw_df, num_rows="dynamic", use_container_width=True, key="keywords_de",
                column_config={"keyword": st.column_config.TextColumn("Keyword")},
            )
            st.caption("To remove a row: select its checkbox on the left, then press Delete.")

            st.markdown("**Research questions**")
            rq_df = pd.DataFrame({"question": article.get("research_questions", [])})
            edited_rqs = st.data_editor(
                rq_df, num_rows="dynamic", use_container_width=True, key="rqs_de",
                column_config={"question": st.column_config.TextColumn("Research question", width="large")},
            )
            st.caption("To remove a row: select its checkbox on the left, then press Delete.")

            if st.form_submit_button("💾 Save metadata", type="primary"):
                article["title"]              = new_title
                article["subtitle"]           = new_subtitle
                article["target_journal"]     = new_journal
                article["authors"]            = edited_authors.dropna(how="all").to_dict("records")
                article["keywords"]           = [k for k in edited_kws["keyword"].dropna().tolist() if k]
                article["research_questions"] = [q for q in edited_rqs["question"].dropna().tolist() if q]
                save_article_structure(article)
                st.success("Metadata saved.")

    else:  # Section Outline
        sections_data = load_sections()
        sec_rows = []
        for s in article.get("sections", []):
            text = sections_data.get(s["id"], "")
            actual_wc = len(text.split()) if text.strip() else 0
            sec_rows.append({
                "done":         bool(s.get("done", False)),
                "id":           s["id"],
                "title":        s["title"],
                "target_words": s["target_words"],
                "actual_words": actual_wc,
            })
        sec_df = pd.DataFrame(sec_rows)
        total_actual   = int(sec_df["actual_words"].sum()) if not sec_df.empty else 0
        total_planned  = int(sec_df["target_words"].sum()) if not sec_df.empty else 0
        max_words      = int(article.get("max_words", 12000))
        pct_used       = min(total_actual / max_words, 1.0) if max_words else 0

        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            st.metric("Word limit", max_words)
        with mc2:
            st.metric("Planned words", total_planned)
        with mc3:
            st.metric("Total words written", total_actual)
        st.progress(pct_used, text=f"{total_actual:,} / {max_words:,} words ({pct_used*100:.0f}%)")

        st.markdown("**Section outline** — edit target word counts below and save.")
        with st.form("section_outline_form"):
            new_max_wc = st.number_input("Word limit", min_value=1000, max_value=30000,
                                         value=max_words, step=500)
            edited_sec = st.data_editor(
                sec_df[["done","id","title","target_words","actual_words"]],
                num_rows="dynamic", use_container_width=True, key="secoutline_de",
                column_config={
                    "done":         st.column_config.CheckboxColumn("✓ Done", width="small"),
                    "id":           st.column_config.TextColumn("Section ID"),
                    "title":        st.column_config.TextColumn("Title", width="large"),
                    "target_words": st.column_config.NumberColumn("Target words", format="%d"),
                    "actual_words": st.column_config.NumberColumn("Written", format="%d"),
                },
                disabled=["id","actual_words"],
            )
            st.caption("To remove a row: select its checkbox on the left, then press Delete.")
            if st.form_submit_button("💾 Save outline", type="primary"):
                new_sections = []
                for _, row in edited_sec.iterrows():
                    if pd.notna(row["id"]) and str(row["id"]).strip():
                        new_sections.append({
                            "id":           str(row["id"]).strip(),
                            "title":        str(row.get("title","")).strip(),
                            "target_words": int(row.get("target_words", 0) or 0),
                            "done":         bool(row.get("done", False)),
                        })
                article["sections"]  = new_sections
                article["max_words"] = int(new_max_wc)
                save_article_structure(article)
                st.success("Section outline saved.")


# ══════════════════════════════════════════════════════════════
# TAB 2 — SECTIONS
# ══════════════════════════════════════════════════════════════

with tab_sections:
    st.markdown("<div class='lr-section-header'>📝 Section Planner</div>", unsafe_allow_html=True)
    st.caption("Use this tab to plan each section: add notes, outlines, and instructions that guide the writing. The full text editor is in the ✍️ Draft tab.")

    article_s  = load_article_structure()
    sec_list   = article_s.get("sections", [])

    # ── Inline section manager ─────────────────────────────────
    with st.expander("⚙️ Manage sections — add, remove, reorder", expanded=(not sec_list)):
        sections_data_mgr = load_sections()
        mgr_rows = []
        for s in sec_list:
            text = sections_data_mgr.get(s["id"], "")
            mgr_rows.append({
                "done":         bool(s.get("done", False)),
                "id":           s["id"],
                "title":        s["title"],
                "target_words": s["target_words"],
                "words_written": len(text.split()) if text.strip() else 0,
            })
        mgr_df = pd.DataFrame(
            mgr_rows if mgr_rows else
            [{"done": False, "id": "", "title": "", "target_words": 0, "words_written": 0}]
        )
        with st.form("sections_manager_form"):
            edited_mgr = st.data_editor(
                mgr_df,
                num_rows="dynamic",
                use_container_width=True,
                key="sections_mgr_de",
                column_config={
                    "done":         st.column_config.CheckboxColumn("✓ Done", width="small"),
                    "id":           st.column_config.TextColumn(
                        "Section ID",
                        help="Unique key used as filename, e.g. 'introduction'. Lowercase, no spaces.",
                        width="medium",
                    ),
                    "title":        st.column_config.TextColumn("Display title", width="large"),
                    "target_words": st.column_config.NumberColumn("Target words", format="%d", width="small"),
                    "words_written":st.column_config.NumberColumn("Written", format="%d", width="small"),
                },
                disabled=["words_written"],
            )
            st.caption("To remove a row: select its checkbox on the left, then press Delete.")
            if st.form_submit_button("💾 Save sections", type="primary"):
                new_sections = []
                for _, row in edited_mgr.iterrows():
                    sid_val = str(row.get("id") or "").strip()
                    if not sid_val:
                        continue
                    sid_val = sid_val.lower().replace(" ", "_")
                    new_sections.append({
                        "id":           sid_val,
                        "title":        str(row.get("title") or "").strip(),
                        "target_words": int(row.get("target_words") or 0),
                        "done":         bool(row.get("done", False)),
                    })
                article_s["sections"] = new_sections
                save_article_structure(article_s)
                st.success(f"Saved {len(new_sections)} sections.")
                st.rerun()

    # Reload after any save
    article_s = load_article_structure()
    sec_list  = article_s.get("sections", [])

    if not sec_list:
        st.info("No sections yet — use the manager above to add them.")
    else:
        sec_options = {s["id"]: s["title"] for s in sec_list}
        sec_ids     = list(sec_options.keys())
        sec_titles  = list(sec_options.values())

        if st.session_state.active_section not in sec_ids:
            st.session_state.active_section = sec_ids[0]

        sel_title = st.selectbox(
            "Section",
            sec_titles,
            index=sec_ids.index(st.session_state.active_section),
            key="section_selector",
        )
        sec_id = sec_ids[sec_titles.index(sel_title)]
        st.session_state.active_section = sec_id

        notes_key = f"secnotes_{sec_id}"
        if notes_key not in st.session_state:
            secs_all = load_sections()
            st.session_state[notes_key] = secs_all.get(notes_key, "")

        def _autosave_notes():
            notes = st.session_state.get(notes_key, "")
            secs = load_sections()
            secs[notes_key] = notes
            save_sections(secs)

        st.text_area(
            "Draft / Notes",
            key=notes_key,
            height=380,
            on_change=_autosave_notes,
            placeholder=(
                "Write your notes, outline, or instructions for this section here.\n\n"
                "Example:\n"
                "- Cover the STEEP framework and its application in foresight\n"
                "- ~700 words, academic tone\n"
                "- Cite Sutherland et al. 2011 and Cuhls 2020"
            ),
        )

        st.markdown("---")
        st.markdown("**All section notes**")
        _sp_secs = load_sections()
        _sp_any  = False
        for sec_def in sec_list:
            sid    = sec_def["id"]
            stitle = sec_def["title"]
            notes  = _sp_secs.get(f"secnotes_{sid}", "")
            if not notes.strip():
                continue
            _sp_any = True
            st.markdown(f"**{stitle}**")
            st.markdown(notes)
            st.markdown("---")
        if not _sp_any:
            st.caption("No notes yet — add them in the notes field above.")


# ══════════════════════════════════════════════════════════════
# TAB 3 — STRATEGY
# ══════════════════════════════════════════════════════════════

with tab_strategy:
    st.markdown("<div class='lr-section-header'>🔎 Search Strategy Planner</div>", unsafe_allow_html=True)
    st.caption("Document the three-step search strategy. Changes are saved to `data/search_strategy.json`.")

    strategy = load_search_strategy()

    strat_sub = st.radio(
        "view", ["📚 Strategy 1", "🌍 Strategy 2", "📡 Strategy 3", "📊 PRISMA Summary"],
        horizontal=True, label_visibility="collapsed",
    )

    # ── Strategy 1 ────────────────────────────────────────────
    if strat_sub == "📚 Strategy 1":
        st.markdown("**Strategy 1 — Peer-reviewed literature (OpenAlex)**")
        s1 = strategy.get("strategy1", {})
        with st.form("strat1_form"):
            st1_purpose  = st.text_area("Purpose", value=s1.get("purpose",""), height=70)
            c1, c2 = st.columns(2)
            with c1:
                st1_db   = st.text_input("Database", value=s1.get("database","OpenAlex"))
                st1_date = st.text_input("Date conducted", value=s1.get("date",""))
                st1_retr = st.number_input("Records retrieved",    min_value=0, value=int(s1.get("n_retrieved",0)))
                st1_dedup= st.number_input("After deduplication",  min_value=0, value=int(s1.get("n_deduped",0)))
                st1_scr  = st.number_input("Title/abstract screen",min_value=0, value=int(s1.get("n_screened",0)))
                st1_man  = st.number_input("Manual inclusions",    min_value=0, value=int(s1.get("n_manual",0)))
            with c2:
                st.markdown("**Search queries**")
                q_df = pd.DataFrame({"query": s1.get("queries", [])})
                new_q = st.data_editor(q_df, num_rows="dynamic", use_container_width=True,
                                       key="strat1_queries",
                                       column_config={"query": st.column_config.TextColumn("Query string", width="large")})
                st.caption("Select checkbox + Delete to remove a row.")
                st.markdown("**Inclusion criteria**")
                inc_df = pd.DataFrame({"criterion": s1.get("inclusion_criteria", [])})
                new_inc = st.data_editor(inc_df, num_rows="dynamic", use_container_width=True, key="strat1_inc",
                                         column_config={"criterion": st.column_config.TextColumn("Criterion", width="large")})
                st.caption("Select checkbox + Delete to remove a row.")
                st.markdown("**Exclusion criteria**")
                exc_df = pd.DataFrame({"criterion": s1.get("exclusion_criteria", [])})
                new_exc = st.data_editor(exc_df, num_rows="dynamic", use_container_width=True, key="strat1_exc",
                                         column_config={"criterion": st.column_config.TextColumn("Criterion", width="large")})
                st.caption("Select checkbox + Delete to remove a row.")

            if st.form_submit_button("💾 Save Strategy 1", type="primary"):
                strategy["strategy1"] = {
                    "purpose":            st1_purpose,
                    "database":           st1_db,
                    "date":               st1_date,
                    "queries":            [q for q in new_q["query"].dropna().tolist() if q],
                    "n_retrieved":        int(st1_retr),
                    "n_deduped":          int(st1_dedup),
                    "n_screened":         int(st1_scr),
                    "n_manual":           int(st1_man),
                    "inclusion_criteria": [c for c in new_inc["criterion"].dropna().tolist() if c],
                    "exclusion_criteria": [c for c in new_exc["criterion"].dropna().tolist() if c],
                }
                save_search_strategy(strategy)
                st.success("Strategy 1 saved.")

    # ── Strategy 2 ────────────────────────────────────────────
    elif strat_sub == "🌍 Strategy 2":
        st.markdown("**Strategy 2 — Cross-sectoral megatrends (purposive scan)**")
        s2 = strategy.get("strategy2", {})
        with st.form("strat2_form"):
            st2_purpose = st.text_area("Purpose", value=s2.get("purpose",""), height=70)
            s2c1, s2c2 = st.columns(2)
            with s2c1:
                st2_date = st.text_input("Date conducted", value=s2.get("date",""))
                st2_n    = st.number_input("Number of entries", min_value=0, value=int(s2.get("n_entries",0)))
            with s2c2:
                st.markdown("**Sources scanned**")
                src_df = pd.DataFrame({"source": s2.get("sources", [])})
                new_src = st.data_editor(src_df, num_rows="dynamic", use_container_width=True, key="strat2_src",
                                          column_config={"source": st.column_config.TextColumn("Source", width="large")})
                st.caption("Select checkbox + Delete to remove a row.")
            if st.form_submit_button("💾 Save Strategy 2", type="primary"):
                strategy["strategy2"] = {
                    "purpose":  st2_purpose,
                    "date":     st2_date,
                    "n_entries":int(st2_n),
                    "sources":  [s for s in new_src["source"].dropna().tolist() if s],
                }
                save_search_strategy(strategy)
                st.success("Strategy 2 saved.")

    # ── Strategy 3 ────────────────────────────────────────────
    elif strat_sub == "📡 Strategy 3":
        st.markdown("**Strategy 3 — Horizon scan for weak signals (living document)**")
        s3 = strategy.get("strategy3", {})
        with st.form("strat3_form"):
            st3_purpose = st.text_area("Purpose", value=s3.get("purpose",""), height=70)
            st3_method  = st.text_area("Method", value=s3.get("method",""), height=80)
            st3_status  = st.selectbox("Status", ["Living document","Completed","On hold"],
                                       index=["Living document","Completed","On hold"].index(
                                           s3.get("status","Living document"))
                                       if s3.get("status","Living document") in ["Living document","Completed","On hold"] else 0)
            if st.form_submit_button("💾 Save Strategy 3", type="primary"):
                strategy["strategy3"] = {
                    "purpose": st3_purpose,
                    "method":  st3_method,
                    "status":  st3_status,
                }
                save_search_strategy(strategy)
                st.success("Strategy 3 saved.")

    # ── PRISMA summary ─────────────────────────────────────────
    else:
        st.markdown("**PRISMA flow summary (live)**")
        s1_data  = strategy.get("strategy1", {})
        excel_mt = EXCEL_FILE.stat().st_mtime if EXCEL_FILE.exists() else 0.0
        mega_cnt = len(load_megatrends(excel_mt))
        sig_cnt  = len(st.session_state.signals)
        n_inc    = int((pool["user_verdict"] == "include").sum()) if total else 0

        pc1, pc2, pc3 = st.columns(3)
        with pc1:
            st.markdown("**Strategy 1 — automated**")
            for label, val in [
                ("Retrieved", s1_data.get("n_retrieved", "?")),
                ("After dedup", s1_data.get("n_deduped", "?")),
                ("Screened", s1_data.get("n_screened", "?")),
                ("In pool", total),
                ("Reviewed", n_reviewed),
                ("Included", n_inc),
            ]:
                st.metric(label, val)
        with pc2:
            st.markdown("**Strategy 1 — manual**")
            st.metric("Manual inclusions", s1_data.get("n_manual", "?"))
            st.markdown("**Strategy 2**")
            st.metric("Megatrend entries", mega_cnt)
        with pc3:
            st.markdown("**Strategy 3**")
            st.metric("Signals documented", sig_cnt)


# ══════════════════════════════════════════════════════════════
# TAB 4 — SEARCH
# ══════════════════════════════════════════════════════════════

with tab_search:
    st.markdown("<div class='lr-section-header'>🔍 Search OpenAlex</div>", unsafe_allow_html=True)

    with st.expander("ℹ️ How the Search tab works", expanded=False):
        st.markdown("""
**Workflow: Build query → Run → Preview → Add to pool → Review**

1. **Build your query** — add one row per search term. Choose AND / OR / NOT.
   Multi-word phrases are automatically quoted.

2. **Run Search** — fetches up to the configured maximum from OpenAlex.

3. **Saved queries** — run any saved query again with ▶, or delete with 🗑.

4. **Add to pool** — after previewing results, click to merge new papers.

> All progress is saved to `data/full_text_review.csv`.
""")

    sc1, sc2, sc3 = st.columns([2, 2, 1])
    with sc1:
        max_results = st.number_input(
            "Max results per query", min_value=10, max_value=1000, step=10,
            value=int(cfg.get("max_results", 200)),
        )
    with sc2:
        email = st.text_input("Email (polite pool)", value=cfg.get("email", "tsz000@uit.no"))
    with sc3:
        st.markdown("<div style='margin-top:26px'>", unsafe_allow_html=True)
        if st.button("💾 Save settings", use_container_width=True):
            cfg["max_results"] = max_results
            cfg["email"]       = email
            st.session_state.search_cfg = cfg
            save_search_config(cfg)
            st.success("Saved.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### Build a search query")
    st.caption("Add one search term per row. Use AND / OR / NOT to define relationships.")

    rows_to_delete = None
    for i, row in enumerate(st.session_state.search_rows):
        rid = row["id"]
        if i == 0:
            c_label, c_term, c_del = st.columns([0.6, 5, 0.5])
            with c_label:
                st.markdown(
                    "<div style='text-align:right;padding-top:34px;color:#888;font-size:0.85rem'>Find</div>",
                    unsafe_allow_html=True,
                )
        else:
            c_op, c_term, c_del = st.columns([0.9, 4.7, 0.5])
            with c_op:
                new_op = st.selectbox(
                    "op", ["AND", "OR", "NOT"],
                    index=["AND", "OR", "NOT"].index(row.get("op") or "AND"),
                    key=f"s_op_{rid}", label_visibility="collapsed",
                )
                st.session_state.search_rows[i]["op"] = new_op

        with c_term:
            new_term = st.text_input(
                "term", value=row["term"], key=f"s_term_{rid}",
                placeholder='Term or phrase — multi-word phrases are auto-quoted',
                label_visibility="collapsed",
            )
            st.session_state.search_rows[i]["term"] = new_term

        with c_del:
            if i > 0:
                st.markdown("<div style='padding-top:28px'>", unsafe_allow_html=True)
                if st.button("✕", key=f"s_del_{rid}", help="Remove this row"):
                    rows_to_delete = i
                st.markdown("</div>", unsafe_allow_html=True)

    if rows_to_delete is not None:
        st.session_state.search_rows.pop(rows_to_delete)
        st.rerun()

    add_c, _, clear_c = st.columns([1, 4, 1])
    with add_c:
        if st.button("＋ Add row", use_container_width=True):
            nid = st.session_state.search_next_id
            st.session_state.search_next_id += 1
            st.session_state.search_rows.append({"id": nid, "op": "AND", "term": ""})
            st.rerun()
    with clear_c:
        if st.button("✕ Clear all", use_container_width=True):
            st.session_state.search_rows    = [{"id": 0, "op": None, "term": ""}]
            st.session_state.search_next_id = 1
            st.rerun()

    current_query = build_query_from_rows(st.session_state.search_rows)
    if current_query:
        st.markdown("**Query preview:**")
        st.code(current_query, language=None)
    else:
        st.caption("Fill in at least one row above to build a query.")

    st.markdown("")
    run_c, save_c, _ = st.columns([1.2, 1.2, 3])
    with run_c:
        run_btn = st.button(
            "🔍 Run Search", type="primary",
            disabled=(not current_query), use_container_width=True,
        )
    with save_c:
        if st.button("💾 Save query", disabled=(not current_query), use_container_width=True):
            saved = cfg.get("saved_queries", [])
            if current_query not in saved:
                saved.append(current_query)
                cfg["saved_queries"] = saved
                st.session_state.search_cfg = cfg
                save_search_config(cfg)
            st.success("Query saved.")

    with st.expander("💡 Query tips & examples", expanded=False):
        st.markdown("""
| Operator | Effect | Example |
|----------|--------|---------|
| **AND** | Both must appear — narrows | `fisheries` AND `climate change` |
| **OR** | Either may appear — broadens | `aquaculture` OR `"fish farming"` |
| **NOT** | Excludes this term | `fisheries` AND `management` NOT `aquaculture` |

Multi-word phrases are auto-quoted. Parentheses are supported by OpenAlex.
""")

    saved_queries = cfg.get("saved_queries", [])
    if saved_queries:
        st.markdown("---")
        st.markdown("#### Saved queries")
        for qi, sq in enumerate(saved_queries):
            sq_c1, sq_c2, sq_c3 = st.columns([5, 0.5, 0.5])
            with sq_c1:
                st.code(sq, language=None)
            with sq_c2:
                if st.button("▶", key=f"run_sq_{qi}", help="Run this query"):
                    st.session_state.pending_run_query = sq
                    st.rerun()
            with sq_c3:
                if st.button("🗑", key=f"del_sq_{qi}", help="Delete"):
                    cfg["saved_queries"].pop(qi)
                    st.session_state.search_cfg = cfg
                    save_search_config(cfg)
                    st.rerun()

    query_to_run = None
    if run_btn and current_query:
        query_to_run = current_query
    elif "pending_run_query" in st.session_state:
        query_to_run = st.session_state.pop("pending_run_query")

    if query_to_run:
        all_records: list = []
        with st.status(f'Searching: {query_to_run[:60]}{"…" if len(query_to_run)>60 else ""}',
                       expanded=True) as status:
            def log(msg): st.write(msg)
            recs = fetch_openalex(query_to_run, max_results, email, log_fn=log)
            all_records.extend(recs)
            seen_ids: set = set()
            deduped = [r for r in all_records
                       if r["openalex_id"] not in seen_ids
                       and not seen_ids.add(r["openalex_id"])]
            st.write(f"**Total:** {len(all_records)} records → {len(deduped)} unique")
            status.update(label=f"Done — {len(deduped)} unique records found", state="complete")
        st.session_state.search_preview = deduped

    preview = st.session_state.search_preview
    if preview:
        existing_ids = set(pool["openalex_id"].tolist())
        n_new = sum(1 for r in preview if r["openalex_id"] not in existing_ids)
        n_dup = len(preview) - n_new

        st.markdown(f"#### Search results: {len(preview)} records")
        info_col, btn_col = st.columns([3, 1])
        with info_col:
            st.info(
                f"**{n_new}** new  ·  **{n_dup}** already in pool  ·  "
                f"Pool grows from **{total}** → **{total + n_new}**"
            )
        with btn_col:
            if st.button(f"➕ Add {n_new} papers", type="primary",
                         use_container_width=True, disabled=(n_new == 0)):
                with st.spinner("Scoring and saving…"):
                    added, skipped = merge_records(preview, kw)
                st.success(f"Added {added}. {skipped} duplicates skipped.")
                st.session_state.search_preview = []
                st.rerun()

        prev_df = pd.DataFrame(preview)[["title","year","journal","authors","openalex_id"]]
        prev_df.insert(0, "In pool", prev_df["openalex_id"].isin(existing_ids).map({True:"✓",False:"new"}))
        prev_df = prev_df.drop(columns=["openalex_id"])
        st.dataframe(prev_df, use_container_width=True, height=350)

    st.markdown("---")
    st.markdown("#### Current review pool")
    if total == 0:
        st.info("Pool is empty. Run a search to add papers.")
    else:
        overview_cols = st.columns(4)
        vc = live_verdicts.value_counts() if total else {}
        with overview_cols[0]: st.metric("Total papers", total)
        with overview_cols[1]: st.metric("🟢 AI green",  int(vc.get("green",  0)))
        with overview_cols[2]: st.metric("🟠 AI orange", int(vc.get("orange", 0)))
        with overview_cols[3]: st.metric("🔴 AI red",    int(vc.get("red",    0)))

        if st.checkbox("Show full pool table"):
            show_df = pool[["title","year","journal","search_query",
                            "date_added","kw_verdict","user_verdict"]].copy()
            st.dataframe(show_df, use_container_width=True, height=400)

        n_with_doi    = int((pool["doi"].str.strip() != "").sum())
        n_with_pdf    = int((pool["pdf_url"].str.strip() != "").sum())
        n_needs_enrich = n_with_doi - n_with_pdf

        st.markdown("---")
        enrich_col, enrich_info = st.columns([1, 2])
        with enrich_col:
            enrich_btn = st.button("🔗 Fetch open-access links",
                                   disabled=(n_needs_enrich == 0), use_container_width=True)
        with enrich_info:
            st.caption(f"PDF links: **{n_with_pdf}** of {n_with_doi} with DOIs  ·  {n_needs_enrich} remaining")

        if enrich_btn:
            with st.status("Querying Unpaywall…", expanded=True) as status:
                def _log(msg): st.write(msg)
                found, checked = enrich_unpaywall(email=cfg.get("email","tsz000@uit.no"), log_fn=_log)
                status.update(label=f"Done — {found} / {checked} PDF links found", state="complete")
            st.rerun()

        st.markdown("")
        st.markdown("**Export**")
        exp_col1, exp_col2, exp_col3 = st.columns([1, 1, 2])
        with exp_col1:
            st.download_button("⬇️ Full pool (CSV)",
                data=pool.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                file_name="full_text_review.csv", mime="text/csv", use_container_width=True)
        with exp_col2:
            bib_filter = st.selectbox("BibTeX export",
                ["Included papers","Included + Unsure","All reviewed","Entire pool"],
                label_visibility="collapsed")
            bib_count_map = {
                "Included papers":   int((pool["user_verdict"]=="include").sum()),
                "Included + Unsure": int(pool["user_verdict"].isin(["include","unsure"]).sum()),
                "All reviewed":      int(pool["user_verdict"].str.strip().astype(bool).sum()),
                "Entire pool":       len(pool),
            }
            n_bib = bib_count_map[bib_filter]
            bib_bytes = _cached_bibtex(bib_filter, pool_mtime) if n_bib > 0 else b""
            st.download_button(f"⬇️ Export BibTeX ({n_bib})", data=bib_bytes,
                file_name="trendfish_references.bib", mime="text/plain",
                disabled=(n_bib == 0), use_container_width=True)
        with exp_col3:
            st.caption(
                f"Included: **{int((pool['user_verdict']=='include').sum())}**  ·  "
                f"Unsure: **{int((pool['user_verdict']=='unsure').sum())}**  ·  "
                f"Excluded: **{int((pool['user_verdict']=='exclude').sum())}**  ·  "
                f"Unreviewed: **{int((pool['user_verdict'].str.strip()=='').sum())}**"
            )


# ══════════════════════════════════════════════════════════════
# TAB 5 — REVIEW
# ══════════════════════════════════════════════════════════════

with tab_review:
    st.markdown("<div class='lr-section-header'>📄 Screening</div>", unsafe_allow_html=True)

    with st.expander("ℹ️ How the Review tab works", expanded=False):
        st.markdown("""
**Goal:** Go through each paper and record your inclusion/exclusion verdict.

- 🟢 **GREEN** — strong fisheries + future-orientation signals → likely include
- 🟠 **ORANGE** — borderline or missing abstract → needs your judgement
- 🔴 **RED** — no fisheries relevance, or purely retrospective → likely exclude

Use **sidebar filters** to focus on a subset. Verdicts save automatically.

> **Tip:** Start with the 🔴 RED filter to clear obvious exclusions, then work through 🟠 ORANGE.
""")

    if total == 0:
        st.info("No papers in pool yet. Use the 🔍 Search tab to add papers.")
        st.stop()

    mask = pd.Series([True] * total, index=pool.index)
    if ai_filter != "All":
        mask &= (live_verdicts == ai_filter)
    if reviewed_filter == "Unreviewed":
        mask &= ~reviewed_mask
    elif reviewed_filter == "Reviewed":
        mask &= reviewed_mask

    filtered_idx = pool[mask].index.tolist()
    n_filtered   = len(filtered_idx)

    if n_filtered == 0:
        st.warning("No papers match the current filters.")
        st.stop()

    pos = min(st.session_state.paper_idx, n_filtered - 1)
    st.session_state.paper_idx = pos
    row_idx = filtered_idx[pos]
    row     = pool.loc[row_idx]

    nav_l, nav_m, nav_r = st.columns([1, 3, 1])
    with nav_l:
        if st.button("◀ Previous", disabled=(pos == 0), use_container_width=True):
            st.session_state.paper_idx = pos - 1
            st.rerun()
    with nav_m:
        jump = st.number_input(
            f"Paper (1–{n_filtered})  of {n_filtered} shown  |  {total} total",
            min_value=1, max_value=n_filtered, value=pos + 1, step=1,
        )
        if int(jump) - 1 != pos:
            st.session_state.paper_idx = int(jump) - 1
            st.rerun()
    with nav_r:
        if st.button("Next ▶", disabled=(pos == n_filtered - 1), use_container_width=True):
            st.session_state.paper_idx = pos + 1
            st.rerun()

    st.markdown("---")

    sc      = score_paper(str(row.get("title","")), str(row.get("abstract","")), kw)
    vc_col  = VERDICT_COLORS[sc["verdict"]]

    sc_left, sc_right = st.columns([3, 1])
    with sc_left:
        st.markdown(
            f"<div style='background:{vc_col}22;border:1px solid {vc_col};"
            f"border-radius:8px;padding:10px 14px'>"
            f"<span style='font-size:1.1rem;font-weight:700;color:{vc_col}'>"
            f"● {sc['verdict'].upper()}</span>"
            f"<span style='color:#555;font-size:0.9rem;margin-left:10px'>{sc['reason']}</span>"
            f"</div>", unsafe_allow_html=True,
        )
    with sc_right:
        st.markdown(
            f"<div style='text-align:center;font-size:0.85rem;padding-top:10px'>"
            f"🐟 <b style='color:#0d6efd'>{sc['fish']}/8</b> &nbsp;&nbsp;"
            f"🔭 <b style='color:#6f42c1'>{sc['future']}/10</b></div>",
            unsafe_allow_html=True,
        )

    bar_l, bar_r = st.columns(2)
    for col, (score, max_s, label, color) in zip(
        [bar_l, bar_r],
        [(sc["fish"], 8, "🐟 Fisheries", "#0d6efd"), (sc["future"], 10, "🔭 Future", "#6f42c1")],
    ):
        with col:
            pct = score / max_s * 100
            st.markdown(
                f"<div style='font-size:0.8rem;color:{color}'>{label} score: {score}/{max_s}</div>"
                f"<div class='score-bar-wrap'><div class='score-bar' "
                f"style='width:{pct:.0f}%;background:{color}'></div></div>",
                unsafe_allow_html=True,
            )

    with st.expander("Matched keywords", expanded=False):
        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            st.markdown("**Fisheries hits**")
            for t in sc["fish_hits"]: st.markdown(f"<div class='kw-match'>🐟 {t}</div>", unsafe_allow_html=True)
            if not sc["fish_hits"]: st.caption("none")
        with mc2:
            st.markdown("**Future hits**")
            for t in sc["future_hits"]: st.markdown(f"<div class='kw-match'>🔭 {t}</div>", unsafe_allow_html=True)
            if not sc["future_hits"]: st.caption("none")
        with mc3:
            st.markdown("**Exclusion hits**")
            for t in sc["retro"]: st.markdown(f"<div class='kw-match' style='color:#dc3545'>⏪ {t}</div>", unsafe_allow_html=True)
            for t in sc["bio"]:   st.markdown(f"<div class='kw-match' style='color:#dc3545'>🧬 {t}</div>", unsafe_allow_html=True)
            for t in sc["nonfish"]: st.markdown(f"<div class='kw-match' style='color:#dc3545'>🚫 {t}</div>", unsafe_allow_html=True)
            if not (sc["retro"] + sc["bio"] + sc["nonfish"]): st.caption("none")

    title   = str(row.get("title",   "")).strip() or "(no title)"
    authors = str(row.get("authors", "")).strip()
    year    = str(row.get("year",    "")).strip()
    journal = str(row.get("journal", "")).strip()
    doi     = str(row.get("doi",     "")).strip()
    oa_id   = str(row.get("openalex_id","")).strip()
    sq      = str(row.get("search_query","")).strip()
    pdf_url = str(row.get("pdf_url","")).strip()
    is_oa   = str(row.get("is_oa","")).strip().lower()

    st.markdown(f"### {title}")
    meta = []
    if authors: meta.append(authors[:120] + ("…" if len(authors) > 120 else ""))
    if year:    meta.append(year)
    if journal: meta.append(f"*{journal}*")
    if meta:
        st.markdown("<div style='color:#555;font-size:0.9rem'>" + "  ·  ".join(meta) + "</div>",
                    unsafe_allow_html=True)

    link_row, badge_row = st.columns([3, 1])
    with link_row:
        if pdf_url: st.link_button("📄 Open PDF ↗", pdf_url, type="primary")
        sec_links = []
        if doi:   sec_links.append(f"[DOI ↗]({'https://doi.org/'+doi if not doi.startswith('http') else doi})")
        if oa_id: sec_links.append(f"[OpenAlex ↗]({oa_id})")
        if sec_links: st.markdown("  ·  ".join(sec_links))
    with badge_row:
        badge = OA_STATUS.get(is_oa)
        if badge:
            icon, label, color = badge
            st.markdown(f"<div style='text-align:right;color:{color};font-size:0.85rem;padding-top:6px'>{icon} {label}</div>",
                        unsafe_allow_html=True)

    if sq: st.caption(f"Found via: *{sq}*")

    abstract = str(row.get("abstract","")).strip()
    ab_src   = str(row.get("abstract_source","")).strip()
    if abstract and len(abstract) > 20:
        src_note = " *(fetched from OpenAlex)*" if ab_src == "openalex" else ""
        st.markdown(f"**Abstract**{src_note}")
        st.markdown(f"<div class='abstract-box'>{abstract}</div>", unsafe_allow_html=True)
    else:
        st.warning("No abstract available — check DOI before deciding.")

    st.markdown("---")
    current_verdict = str(row.get("user_verdict","")).strip()
    current_notes   = str(row.get("user_notes",  "")).strip()

    st.subheader("Your verdict")
    btn_cols = st.columns(3)
    clicked  = None
    for col, (key, (icon, label, bg, col_hex)) in zip(btn_cols, VERDICTS.items()):
        border = "3px solid #000" if current_verdict == key else f"2px solid {col_hex}"
        with col:
            st.markdown(
                f"<div style='background:{bg};border:{border};border-radius:8px;"
                f"padding:6px;text-align:center;font-weight:700;font-size:1.05rem'>"
                f"{icon} {label}</div>", unsafe_allow_html=True,
            )
            if st.button(f"Set {label}", key=f"btn_{key}", use_container_width=True):
                clicked = key

    notes_input = st.text_area("Notes (optional)", value=current_notes, height=80,
                               placeholder="Comments, flags, or reasons…")

    if clicked is not None:
        save_verdict_to_pool(pool, row_idx, clicked, notes_input)
        st.success(f"Saved: {VERDICTS[clicked][0]} {VERDICTS[clicked][1]}")
        if reviewed_filter == "Unreviewed" and pos + 1 < n_filtered:
            st.session_state.paper_idx = pos + 1
        st.rerun()

    save_c, _ = st.columns([1, 4])
    with save_c:
        if st.button("💾 Save notes", disabled=(not current_verdict)):
            save_verdict_to_pool(pool, row_idx, current_verdict, notes_input)
            st.success("Notes saved.")
            st.rerun()

    if current_verdict:
        icon, label, _, col_hex = VERDICTS[current_verdict]
        reviewed_at = str(row.get("reviewed_at","")).strip()
        st.markdown(
            f"<div style='font-size:0.9rem;color:{col_hex};margin-top:6px'>"
            f"Your verdict: {icon} {label}"
            + (f" — *{current_notes}*" if current_notes else "")
            + "</div>", unsafe_allow_html=True,
        )
        if reviewed_at: st.caption(f"Reviewed at: {reviewed_at}")


# ══════════════════════════════════════════════════════════════
# TAB 6 — ANALYSIS (hub)
# ══════════════════════════════════════════════════════════════

with tab_analysis:
    st.markdown("<div class='lr-section-header'>📊 Analysis</div>", unsafe_allow_html=True)

    sub_nav = st.radio(
        "analysis_sub",
        ["Overview", "STEEP Gap", "Megatrends", "Signals", "Keywords"],
        horizontal=True,
        label_visibility="collapsed",
    )

    excel_mtime_a = EXCEL_FILE.stat().st_mtime if EXCEL_FILE.exists() else 0.0
    mega_a = load_megatrends(excel_mtime_a)
    fish_a = load_fisheries_futures(excel_mtime_a)
    sigs_a = st.session_state.signals

    # ── Overview ───────────────────────────────────────────────
    if sub_nav == "Overview":
        st.markdown("#### Review progress")
        n_include = int((pool["user_verdict"] == "include").sum())
        n_unsure  = int((pool["user_verdict"] == "unsure").sum())
        n_exclude = int((pool["user_verdict"] == "exclude").sum())
        n_pending = total - n_reviewed
        pct = n_reviewed / total * 100 if total > 0 else 0

        prog_cols = st.columns(6)
        for col, label, val, color in zip(
            prog_cols,
            ["Total pool","Reviewed","Pending","✅ Included","⚠️ Unsure","❌ Excluded"],
            [total, n_reviewed, n_pending, n_include, n_unsure, n_exclude],
            ["#0d6efd","#6c757d","#fd7e14","#28a745","#fd7e14","#dc3545"],
        ):
            with col:
                st.markdown(
                    f"<div style='text-align:center'>"
                    f"<div style='font-size:1.6rem;font-weight:700;color:{color}'>{val}</div>"
                    f"<div style='font-size:0.8rem;color:#555'>{label}</div></div>",
                    unsafe_allow_html=True,
                )
        st.progress(pct / 100, text=f"{pct:.1f}% screened ({n_reviewed} / {total})")

        st.markdown("---")
        st.markdown("#### PRISMA counts")
        strategy_o = load_search_strategy()
        s1o = strategy_o.get("strategy1", {})
        pr1, pr2, pr3 = st.columns(3)
        with pr1:
            st.markdown("**Strategy 1 — automated**")
            for label, val in [
                ("Retrieved", s1o.get("n_retrieved","?")),
                ("After dedup", s1o.get("n_deduped","?")),
                ("Screened", s1o.get("n_screened","?")),
                ("In pool", total),
                ("Reviewed", n_reviewed),
                ("Included", n_include),
            ]:
                st.metric(label, val)
        with pr2:
            st.markdown("**Strategy 1 — manual**")
            st.metric("Manual inclusions", s1o.get("n_manual","?"))
            st.markdown("**Strategy 2**")
            st.metric("Megatrend entries", len(mega_a))
        with pr3:
            st.markdown("**Strategy 3**")
            st.metric("Signals documented", len(sigs_a))
            st.markdown("**Combined**")
            st.metric("Total sources", s1o.get("n_manual",0) + len(mega_a) + len(sigs_a))

        st.markdown("---")
        st.markdown("#### Source type breakdown (cross-sectoral megatrends)")
        type_counts = mega_a["source_type"].value_counts().reset_index()
        type_counts.columns = ["Source type", "Count"]
        fig_pie = px.pie(type_counts, names="Source type", values="Count",
                         color_discrete_sequence=px.colors.qualitative.Set2, height=350)
        fig_pie.update_layout(margin=dict(t=20, b=20))
        st.plotly_chart(fig_pie, use_container_width=True)

    # ── STEEP Gap ──────────────────────────────────────────────
    elif sub_nav == "STEEP Gap":
        st.markdown("#### STEEP gap analysis")
        st.caption("Percentage-point difference between cross-sectoral megatrends and fisheries futures literature.")

        cats_5 = [c for c in STEEP_CATS if c != "Other"]

        def steep_pct(df: pd.DataFrame) -> dict:
            if df.empty:
                return {c: 0.0 for c in cats_5}
            counts = df["steep"].value_counts()
            t = len(df)
            return {c: round(counts.get(c, 0) / t * 100, 1) for c in cats_5}

        mega_pct = steep_pct(mega_a)
        fish_pct = steep_pct(fish_a)
        gap_pct  = {c: round(mega_pct[c] - fish_pct[c], 1) for c in cats_5}

        chart_cols = st.columns(3)

        def make_bar(title, data_dict, colors):
            df_c = pd.DataFrame({"STEEP": list(data_dict.keys()), "Pct": list(data_dict.values())})
            fig  = px.bar(df_c, x="STEEP", y="Pct", color="STEEP",
                          color_discrete_map=colors,
                          text=df_c["Pct"].apply(lambda v: f"{v:.0f}%"),
                          title=title, height=320)
            fig.update_layout(showlegend=False, margin=dict(t=40,b=20),
                              yaxis=dict(range=[0,65], title="%"))
            fig.update_traces(textposition="outside")
            return fig

        with chart_cols[0]:
            st.plotly_chart(make_bar("Cross-sectoral megatrends", mega_pct, STEEP_COLORS),
                            use_container_width=True)
        with chart_cols[1]:
            st.plotly_chart(make_bar("Fisheries futures literature", fish_pct, STEEP_COLORS),
                            use_container_width=True)
        with chart_cols[2]:
            gap_colors = {c: ("#dc3545" if v >= 5 else "#28a745" if v <= -5 else "#adb5bd")
                          for c, v in gap_pct.items()}
            df_gap = pd.DataFrame({"STEEP": list(gap_pct.keys()), "Gap (pp)": list(gap_pct.values())})
            fig_gap = px.bar(df_gap, x="STEEP", y="Gap (pp)", color="STEEP",
                             color_discrete_map=gap_colors,
                             text=df_gap["Gap (pp)"].apply(lambda v: f"{v:+.0f}pp"),
                             title="Gap (megatrends − fisheries)", height=320)
            fig_gap.update_layout(showlegend=False, margin=dict(t=40,b=20),
                                  yaxis=dict(title="Percentage points"))
            fig_gap.update_traces(textposition="outside")
            st.plotly_chart(fig_gap, use_container_width=True)

        st.caption("🔴 Red = underrepresented in fisheries research  ·  🟢 Green = overrepresented")
        gap_table = pd.DataFrame({
            "STEEP":        cats_5,
            "Megatrends %": [mega_pct[c] for c in cats_5],
            "Fisheries %":  [fish_pct[c] for c in cats_5],
            "Gap (pp)":     [gap_pct[c]  for c in cats_5],
        })
        st.dataframe(gap_table, use_container_width=True, hide_index=True,
                     column_config={"Gap (pp)": st.column_config.NumberColumn(format="%.1f")})

    # ── Megatrends ─────────────────────────────────────────────
    elif sub_nav == "Megatrends":
        st.markdown("#### Cross-sectoral Megatrends & Drivers")
        st.caption("Strategy 2 — `data/manual_literature_search.xlsx`")

        f1, f2, f3 = st.columns([2, 2, 3])
        with f1:
            steep_filter = st.multiselect("STEEP category", STEEP_CATS, default=STEEP_CATS, key="mega_steep")
        with f2:
            all_types = sorted(mega_a["source_type"].dropna().unique().tolist())
            type_filter = st.multiselect("Source type", all_types, default=all_types, key="mega_type")
        with f3:
            text_search = st.text_input("Search trends", key="mega_search", placeholder="Type to filter…")

        filtered = mega_a[mega_a["steep"].isin(steep_filter) & mega_a["source_type"].isin(type_filter)]
        if text_search:
            mask_m = (filtered["trend"].str.contains(text_search, case=False, na=False) |
                      filtered["details"].str.contains(text_search, case=False, na=False))
            filtered = filtered[mask_m]

        st.caption(f"Showing **{len(filtered)}** of {len(mega_a)} entries")

        steep_counts = (filtered["steep"].value_counts()
                        .reindex(STEEP_CATS, fill_value=0).reset_index())
        steep_counts.columns = ["STEEP","Count"]
        steep_counts["Pct"] = (steep_counts["Count"] / max(len(filtered),1) * 100).round(1)
        fig_mega = px.bar(steep_counts, x="STEEP", y="Count", color="STEEP",
                          color_discrete_map=STEEP_COLORS,
                          text=steep_counts["Pct"].apply(lambda v: f"{v:.0f}%"),
                          title="STEEP distribution — cross-sectoral megatrends", height=300)
        fig_mega.update_layout(showlegend=False, margin=dict(t=40,b=20))
        fig_mega.update_traces(textposition="outside")
        st.plotly_chart(fig_mega, use_container_width=True)

        display_cols = ["trend","steep","source_type","year","author","source"]
        st.dataframe(
            filtered[[c for c in display_cols if c in filtered.columns]],
            use_container_width=True, height=380,
            column_config={
                "trend":       st.column_config.TextColumn("Trend / Driver", width="large"),
                "steep":       st.column_config.TextColumn("STEEP", width="small"),
                "source_type": st.column_config.TextColumn("Type", width="small"),
                "year":        st.column_config.NumberColumn("Year", format="%d", width="small"),
                "author":      st.column_config.TextColumn("Author", width="medium"),
                "source":      st.column_config.LinkColumn("Source", width="medium"),
            },
        )

        st.markdown("---")
        with st.expander("➕ Add new megatrend entry", expanded=False):
            with st.form("add_mega_form", clear_on_submit=True):
                ac1, ac2 = st.columns(2)
                with ac1:
                    new_trend   = st.text_input("Trend / driver label *")
                    new_author  = st.text_input("Author(s)")
                    new_year    = st.number_input("Year", min_value=1990, max_value=2030, value=2024, step=1)
                with ac2:
                    new_details = st.text_area("Details / description", height=100)
                    new_source  = st.text_input("Source URL or name")
                    new_stype   = st.selectbox("Source type",
                        ["Article","Consultancy","Government","NGO / Think Tank",
                         "Book / Chapter","Business","Unclassified"])
                new_title    = st.text_input("Publication title")
                new_comments = st.text_area("Comments", height=60)
                if st.form_submit_button("💾 Save entry", type="primary"):
                    if not new_trend.strip():
                        st.error("Trend label is required.")
                    else:
                        add_megatrend_to_excel({
                            "trend": new_trend, "details": new_details,
                            "author": new_author, "source": new_source,
                            "title": new_title, "comments": new_comments,
                            "doi": "", "year": int(new_year), "source_type": new_stype,
                        })
                        st.success(f"Entry '{new_trend}' added.")
                        st.rerun()

    # ── Signals ────────────────────────────────────────────────
    elif sub_nav == "Signals":
        st.markdown("#### Weak & Emerging Signals")
        st.caption("Strategy 3 — `data/signals.json`")

        signals = st.session_state.signals
        STATUS_COLORS = {"Monitoring":"#fd7e14","Confirmed":"#28a745","Dismissed":"#6c757d"}

        with st.expander("➕ Add new signal", expanded=(len(signals)==0)):
            with st.form("add_signal_form", clear_on_submit=True):
                s1c, s2c = st.columns(2)
                with s1c:
                    sig_title  = st.text_input("Signal title *")
                    sig_source = st.text_input("Source name")
                    sig_date   = st.text_input("Date observed (YYYY-MM-DD)",
                                               value=datetime.today().strftime("%Y-%m-%d"))
                    sig_steep  = st.selectbox("STEEP category", STEEP_CATS)
                    sig_status = st.selectbox("Status", ["Monitoring","Confirmed","Dismissed"])
                with s2c:
                    sig_desc  = st.text_area("Description *", height=110,
                        placeholder="What is happening? Why might it matter?")
                    sig_url   = st.text_input("Source URL")
                    sig_rel   = st.text_area("Fisheries relevance", height=80)
                if st.form_submit_button("💾 Save signal", type="primary"):
                    if not sig_title.strip() or not sig_desc.strip():
                        st.error("Title and description are required.")
                    else:
                        new_id = max((s["id"] for s in signals), default=0) + 1
                        signals.append({
                            "id": new_id, "title": sig_title, "description": sig_desc,
                            "source": sig_source, "url": sig_url, "date": sig_date,
                            "steep": sig_steep, "fisheries_relevance": sig_rel,
                            "status": sig_status,
                        })
                        st.session_state.signals = signals
                        save_signals(signals)
                        st.success(f"Signal '{sig_title}' saved.")
                        st.rerun()

        st.markdown("---")
        if not signals:
            st.info("No signals documented yet.")
        else:
            st.markdown(f"**{len(signals)} signal{'s' if len(signals)!=1 else ''} documented**")
            for sig in signals:
                sid      = sig["id"]
                s_color  = STEEP_COLORS.get(sig["steep"], "#999")
                st_color = STATUS_COLORS.get(sig["status"], "#999")
                is_editing = (st.session_state.editing_signal == sid)
                with st.container(border=True):
                    h1, h2 = st.columns([5,1])
                    with h1:
                        st.markdown(
                            f"<span style='background:{s_color};color:white;padding:2px 8px;"
                            f"border-radius:4px;font-size:0.8rem'>{sig['steep']}</span>&nbsp;&nbsp;"
                            f"**{sig['title']}**&nbsp;&nbsp;"
                            f"<span style='color:#888;font-size:0.85rem'>{sig.get('date','')}</span>",
                            unsafe_allow_html=True,
                        )
                    with h2:
                        b1, b2 = st.columns(2)
                        with b1:
                            if st.button("✏️", key=f"edit_sig_{sid}"):
                                st.session_state.editing_signal = None if is_editing else sid
                                st.rerun()
                        with b2:
                            if st.button("🗑", key=f"del_sig_{sid}"):
                                st.session_state.signals = [s for s in signals if s["id"]!=sid]
                                save_signals(st.session_state.signals)
                                st.rerun()

                    if not is_editing:
                        st.markdown(sig["description"])
                        if sig.get("fisheries_relevance"):
                            st.caption(f"*Fisheries relevance:* {sig['fisheries_relevance']}")
                        foot1, foot2 = st.columns([3,1])
                        with foot1:
                            if sig.get("url"):
                                st.markdown(f"[{sig.get('source') or sig['url']}]({sig['url']})")
                            elif sig.get("source"):
                                st.caption(sig["source"])
                        with foot2:
                            st.markdown(
                                f"<div style='text-align:right;color:{st_color};font-size:0.85rem'>● {sig['status']}</div>",
                                unsafe_allow_html=True,
                            )
                    else:
                        with st.form(f"edit_sig_{sid}_form", clear_on_submit=False):
                            e1, e2 = st.columns(2)
                            with e1:
                                e_title  = st.text_input("Title",  value=sig["title"])
                                e_source = st.text_input("Source", value=sig.get("source",""))
                                e_date   = st.text_input("Date",   value=sig.get("date",""))
                                e_steep  = st.selectbox("STEEP",   STEEP_CATS,
                                    index=STEEP_CATS.index(sig["steep"]) if sig["steep"] in STEEP_CATS else 0)
                                e_status = st.selectbox("Status",  ["Monitoring","Confirmed","Dismissed"],
                                    index=["Monitoring","Confirmed","Dismissed"].index(sig["status"]))
                            with e2:
                                e_desc = st.text_area("Description", value=sig["description"], height=110)
                                e_url  = st.text_input("URL",  value=sig.get("url",""))
                                e_rel  = st.text_area("Fisheries relevance",
                                                      value=sig.get("fisheries_relevance",""), height=80)
                            if st.form_submit_button("💾 Save changes"):
                                for s in st.session_state.signals:
                                    if s["id"] == sid:
                                        s.update({"title":e_title,"description":e_desc,
                                                  "source":e_source,"url":e_url,"date":e_date,
                                                  "steep":e_steep,"fisheries_relevance":e_rel,
                                                  "status":e_status})
                                save_signals(st.session_state.signals)
                                st.session_state.editing_signal = None
                                st.rerun()

    # ── Keywords ───────────────────────────────────────────────
    elif sub_nav == "Keywords":
        st.markdown("#### Keyword Configuration")
        st.caption("Edit keyword lists used for AI pre-screening. Changes apply instantly to Review tab scoring.")

        kw_changed = False
        new_kw     = {}
        groups = [
            ("🐟 Fisheries relevance", ["fish_strong","fish_med","fish_context"]),
            ("🔭 Future orientation",  ["future_strong","future_med","future_weak"]),
            ("🚫 Exclusion criteria",  ["excl_retro","excl_bio","excl_nonfish"]),
        ]
        for group_label, keys in groups:
            st.markdown(f"#### {group_label}")
            cols = st.columns(len(keys))
            for col, key in zip(cols, keys):
                label, weight, color = KEYWORD_META[key]
                with col:
                    st.markdown(
                        f"<div style='color:{color};font-weight:600;margin-bottom:2px'>{label}</div>"
                        f"<div style='font-size:0.75rem;color:#888;margin-bottom:6px'>{weight}</div>",
                        unsafe_allow_html=True,
                    )
                    edited = st.data_editor(
                        pd.DataFrame({"keyword": kw.get(key, [])}),
                        num_rows="dynamic", use_container_width=True,
                        hide_index=True, key=f"de_{key}",
                        column_config={"keyword": st.column_config.TextColumn("Keyword / phrase", max_chars=120)},
                    )
                    st.caption("Checkbox + Delete to remove.")
                    parsed = [v.strip().lower() for v in edited["keyword"].dropna() if str(v).strip()]
                    new_kw[key] = parsed
                    if parsed != kw.get(key):
                        kw_changed = True
            st.markdown("")

        if kw_changed:
            st.session_state.kw = new_kw

        st.markdown("---")
        act_l, act_r, _ = st.columns([1, 1, 3])
        with act_l:
            if st.button("💾 Save to file", use_container_width=True):
                save_keywords(st.session_state.kw)
                st.success(f"Saved → {KW_FILE.name}")
        with act_r:
            if st.button("↺ Reset to defaults", use_container_width=True):
                for key in DEFAULT_KW:
                    st.session_state.pop(f"de_{key}", None)
                st.session_state.kw = {k: list(v) for k, v in DEFAULT_KW.items()}
                st.rerun()


# ══════════════════════════════════════════════════════════════
# TAB 7 — DRAFT
# ══════════════════════════════════════════════════════════════

with tab_draft:
    st.markdown("<div class='lr-section-header'>✍️ Draft</div>", unsafe_allow_html=True)

    _d_article  = load_article_structure()
    _d_sec_list = _d_article.get("sections", [])

    if not _d_sec_list:
        st.info("No sections defined yet — go to the 📝 Sections tab to add them.")
    else:
        _d_options = {s["id"]: s["title"] for s in _d_sec_list}
        _d_ids     = list(_d_options.keys())
        _d_titles  = list(_d_options.values())

        if st.session_state.active_section not in _d_ids:
            st.session_state.active_section = _d_ids[0]

        _d_sel_title = st.selectbox(
            "Section",
            _d_titles,
            index=_d_ids.index(st.session_state.active_section),
            key="draft_section_selector",
        )
        sec_id = _d_ids[_d_titles.index(_d_sel_title)]
        st.session_state.active_section = sec_id

        target_words = next((s["target_words"] for s in _d_sec_list if s["id"] == sec_id), 0)
        sec_key      = f"sectext_{sec_id}"

        if sec_key not in st.session_state:
            secs = load_sections()
            st.session_state[sec_key] = secs.get(sec_id, "")

        current_text = st.session_state[sec_key]
        wc = len(current_text.split()) if current_text.strip() else 0
        if target_words > 0:
            ratio = wc / target_words
            wc_color = "#28a745" if 0.9 <= ratio <= 1.15 else ("#dc3545" if ratio > 1.15 else "#fd7e14")
        else:
            wc_color = "#6c757d"

        st.markdown(
            f"<span class='lr-badge' style='background:{wc_color}'>"
            f"{wc:,} / {target_words:,} words</span>",
            unsafe_allow_html=True,
        )

        def _autosave_section():
            text = st.session_state.get(sec_key, "")
            secs = load_sections()
            secs[sec_id] = text
            save_sections(secs)

        st.text_area(
            _d_sel_title,
            key=sec_key,
            height=540,
            on_change=_autosave_section,
            label_visibility="collapsed",
            placeholder="Start writing or paste your draft here…",
        )

        _dsave_c, _dreset_c, _ = st.columns([1, 1, 3])
        with _dsave_c:
            if st.button("💾 Save section", use_container_width=True, type="primary",
                         key="draft_save_btn"):
                text = st.session_state.get(sec_key, "")
                secs = load_sections()
                secs[sec_id] = text
                save_sections(secs)
                st.success("Saved.")
        with _dreset_c:
            if st.button("↩ Reset from .docx", use_container_width=True,
                         help="Re-seed from the corresponding Word document",
                         key="draft_reset_btn"):
                fresh = seed_sections_from_docx()
                st.session_state[sec_key] = fresh.get(sec_id, "")
                secs = load_sections()
                secs[sec_id] = st.session_state[sec_key]
                save_sections(secs)
                st.rerun()

        # Auto-generate references (references section only)
        _REF_IDS = {"references", "literature", "bibliography", "reference_list"}
        if sec_id in _REF_IDS or any(x in sec_id for x in ("ref", "lit", "bib")):
            with st.expander("📚 Auto-generate reference list", expanded=False):
                st.caption(
                    "Builds a formatted reference list from your review pool and/or megatrend "
                    "sources. The result is placed in the text area above where you can edit it."
                )
                _rg1, _rg2 = st.columns(2)
                with _rg1:
                    rg_verdict = st.selectbox(
                        "Include pool papers",
                        ["Included only", "Included + Unsure", "None"],
                        key="rg_verdict_sel",
                    )
                with _rg2:
                    rg_mega = st.checkbox("Add megatrend sources (Strategy 2)", value=False,
                                          key="rg_mega_cb")

                if st.button("⚙️ Generate reference list", type="primary", key="rg_gen_btn"):
                    refs = []

                    if rg_verdict != "None":
                        _pool = load_pool()
                        if rg_verdict == "Included only":
                            _pool = _pool[_pool["user_verdict"] == "include"]
                        else:
                            _pool = _pool[_pool["user_verdict"].isin(["include", "unsure"])]

                        for _, row in _pool.iterrows():
                            authors_raw = str(row.get("authors", "")).strip()
                            year   = str(row.get("year",    "")).strip()
                            title  = str(row.get("title",   "")).strip()
                            journal= str(row.get("journal", "")).strip()
                            doi    = str(row.get("doi",     "")).strip()

                            raw_parts = [p.strip().rstrip(".") for p in
                                         authors_raw.replace(" et al.", "").split(";") if p.strip()]
                            apa = []
                            for p in raw_parts:
                                tokens = p.split()
                                if len(tokens) >= 2:
                                    initials = " ".join(t[0].upper() + "." for t in tokens[:-1])
                                    apa.append(f"{tokens[-1]}, {initials}")
                                else:
                                    apa.append(p)
                            if "et al." in authors_raw and apa:
                                author_str = apa[0] + " et al."
                            elif len(apa) > 1:
                                author_str = ", ".join(apa[:-1]) + ", & " + apa[-1]
                            else:
                                author_str = apa[0] if apa else "Unknown"

                            ref = f"{author_str} ({year}). {title}."
                            if journal:
                                ref += f" *{journal}*."
                            if doi:
                                ref += f" https://doi.org/{doi}"
                            refs.append(ref)

                    if rg_mega:
                        excel_mt = EXCEL_FILE.stat().st_mtime if EXCEL_FILE.exists() else 0.0
                        mega_refs = load_megatrends(excel_mt)
                        seen_titles: set = set()
                        for _, row in mega_refs.iterrows():
                            author = str(row.get("author", "")).strip()
                            title  = str(row.get("title",  "")).strip()
                            source = str(row.get("source", "")).strip()
                            doi    = str(row.get("doi",    "")).strip()
                            year_v = row.get("year")
                            year   = str(int(year_v)) if pd.notna(year_v) else ""
                            _key   = (author, title)
                            if _key in seen_titles or not (author or title):
                                continue
                            seen_titles.add(_key)
                            ref = f"{author} ({year}). {title}."
                            if source:
                                ref += f" {source}."
                            if doi:
                                ref += f" https://doi.org/{doi}"
                            refs.append(ref)

                    refs.sort(key=str.lower)
                    ref_text = "\n\n".join(refs)
                    st.session_state[sec_key] = ref_text
                    _secs = load_sections()
                    _secs[sec_id] = ref_text
                    save_sections(_secs)
                    st.success(f"Generated {len(refs)} references — edit in the text area above.")
                    st.rerun()

        st.markdown("---")
        st.markdown("**Full article draft**")
        _all_secs = load_sections()
        _any_content = False
        for sec_def in _d_sec_list:
            sid    = sec_def["id"]
            stitle = sec_def["title"]
            body   = st.session_state.get(f"sectext_{sid}", "") if sid == sec_id else _all_secs.get(sid, "")
            if not body.strip():
                continue
            _any_content = True
            st.markdown(f"### {stitle}")
            st.markdown(body)
            st.markdown("---")
        if not _any_content:
            st.caption("No sections have content yet.")


# ══════════════════════════════════════════════════════════════
# TAB 8 — OUTPUTS
# ══════════════════════════════════════════════════════════════

with tab_outputs:
    st.markdown("<div class='lr-section-header'>📤 Outputs</div>", unsafe_allow_html=True)

    article_o  = load_article_structure()
    sections_o = load_sections()

    # ── Manuscript ────────────────────────────────────────────
    st.markdown("#### Manuscript (Word)")
    sec_status = []
    for s in article_o.get("sections", []):
        text = sections_o.get(s["id"], "")
        wc   = len(text.split()) if text.strip() else 0
        sec_status.append({
            "Done":    bool(s.get("done", False)),
            "Section": s["title"],
            "Words":   wc,
            "Target":  s["target_words"],
        })

    st.dataframe(pd.DataFrame(sec_status), use_container_width=True, hide_index=True,
                 column_config={
                     "Done":   st.column_config.CheckboxColumn("✓ Done", width="small"),
                     "Words":  st.column_config.NumberColumn(format="%d"),
                     "Target": st.column_config.NumberColumn(format="%d"),
                 })

    out_c1, out_c2 = st.columns(2)
    with out_c1:
        if st.button("⚙️ Generate manuscript", type="primary", use_container_width=True):
            with st.spinner("Building Word document…"):
                docx_bytes = generate_manuscript_bytes(article_o, sections_o)
            if docx_bytes:
                st.session_state["manuscript_bytes"] = docx_bytes
                st.success("Manuscript generated. Click download below.")
            else:
                st.error("Could not generate document — ensure python-docx is installed.")

    with out_c2:
        mb = st.session_state.get("manuscript_bytes", b"")
        st.download_button(
            "⬇️ Download Manuscript (DOCX)",
            data=mb,
            file_name="Manuscript_Trends_2026.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            disabled=(not mb),
            use_container_width=True,
        )

    # ── Also save to docs/ ─────────────────────────────────────
    if st.session_state.get("manuscript_bytes"):
        if st.button("💾 Save to docs/Manuscript_Trends_2026.docx"):
            out_path = DOCS_DIR / "Manuscript_Trends_2026.docx"
            out_path.write_bytes(st.session_state["manuscript_bytes"])
            st.success(f"Saved to {out_path}")

    st.markdown("---")

    # ── PRISMA diagram ─────────────────────────────────────────
    st.markdown("#### PRISMA Diagram")
    prisma_script = BASE / "analysis" / "generate_prisma.py"
    prisma_svg    = DOCS_DIR / "PRISMA_flow.svg"
    prisma_png    = DOCS_DIR / "PRISMA_flow.png"

    if prisma_script.exists():
        if st.button("⚙️ Generate PRISMA diagram", use_container_width=True):
            with st.spinner("Running generate_prisma.py…"):
                try:
                    result = subprocess.run(
                        [sys.executable, str(prisma_script)],
                        capture_output=True, text=True, cwd=str(BASE), timeout=60,
                    )
                    if result.returncode == 0:
                        st.success("PRISMA diagram generated.")
                    else:
                        st.error(f"Script error:\n{result.stderr[:500]}")
                except Exception as e:
                    st.error(str(e))

    dl1, dl2 = st.columns(2)
    with dl1:
        if prisma_svg.exists():
            st.download_button("⬇️ PRISMA SVG", data=prisma_svg.read_bytes(),
                               file_name="PRISMA_flow.svg", mime="image/svg+xml",
                               use_container_width=True)
        else:
            st.caption("PRISMA SVG not yet generated.")
    with dl2:
        if prisma_png.exists():
            st.download_button("⬇️ PRISMA PNG", data=prisma_png.read_bytes(),
                               file_name="PRISMA_flow.png", mime="image/png",
                               use_container_width=True)
        else:
            st.caption("PRISMA PNG not yet generated.")

    st.markdown("---")

    # ── Data exports ───────────────────────────────────────────
    st.markdown("#### Data exports")
    ex1, ex2, ex3, ex4 = st.columns(4)
    with ex1:
        st.download_button(
            "⬇️ Review pool (CSV)",
            data=pool.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
            file_name="full_text_review.csv", mime="text/csv",
            use_container_width=True,
        )
    with ex2:
        n_inc_dl = int((pool["user_verdict"]=="include").sum())
        bib_dl   = _cached_bibtex("Included papers", pool_mtime) if n_inc_dl > 0 else b""
        st.download_button(
            f"⬇️ BibTeX — included ({n_inc_dl})",
            data=bib_dl, file_name="included_references.bib", mime="text/plain",
            disabled=(n_inc_dl == 0), use_container_width=True,
        )
    with ex3:
        signals_json = json.dumps(st.session_state.signals, indent=2, ensure_ascii=False).encode("utf-8")
        st.download_button(
            f"⬇️ Signals JSON ({len(st.session_state.signals)})",
            data=signals_json, file_name="signals.json", mime="application/json",
            use_container_width=True,
        )
    with ex4:
        if EXCEL_FILE.exists():
            st.download_button(
                "⬇️ Megatrends (Excel)",
                data=EXCEL_FILE.read_bytes(),
                file_name="manual_literature_search.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.caption("Excel file not found.")
