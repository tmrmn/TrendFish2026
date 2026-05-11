"""
TrendFish 2026 — unified search + review Streamlit app.

Tabs:
  🔍 Search   — query OpenAlex, fetch abstracts, add to review pool
  📄 Review   — paper-by-paper review with live keyword scoring
  ⚙️  Keywords — configure keyword lists

All progress is saved automatically to data/full_text_review.csv.
Search settings persist in data/search_config.json.
Keywords persist in data/review_keywords.json.

Usage:
    python -m streamlit run analysis/review_app.py
"""

import json
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

BASE      = Path(__file__).parent.parent
DATA_DIR  = BASE / "data"
IN_CSV    = DATA_DIR / "full_text_review.csv"
KW_FILE   = DATA_DIR / "review_keywords.json"
SRCH_FILE = DATA_DIR / "search_config.json"

DATA_DIR.mkdir(exist_ok=True)

# ── Default keyword lists ─────────────────────────────────────

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

# ─────────────────────────────────────────────────────────────
# Helpers
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
            # Migrate old "queries" list to "saved_queries"
            if "queries" in data and "saved_queries" not in data:
                data["saved_queries"] = data.pop("queries")
            return data
        except Exception:
            pass
    return {"saved_queries": DEFAULT_QUERIES, "max_results": 200, "email": "tsz000@uit.no"}


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


# ── Query builder ─────────────────────────────────────────────

def build_query_from_rows(rows: list[dict]) -> str:
    """Build an OpenAlex boolean query from builder rows.
    Multi-word terms are auto-quoted; operators are AND / OR / NOT (uppercase)."""
    parts: list[str] = []
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


# ── OpenAlex ──────────────────────────────────────────────────

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


def fetch_openalex(query: str, max_results: int, email: str,
                   log_fn=None) -> list[dict]:
    """Return up to max_results parsed records for a single query string."""
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
                if log_fn:
                    log_fn(f"  HTTP {r.status_code} — stopping pagination")
                break
            batch = r.json().get("results", [])
            if not batch:
                break
            collected.extend(batch)
            if log_fn:
                log_fn(f"  Page {page}: {len(batch)} records (total {len(collected)})")
            if len(batch) < per_page:
                break
            page += 1
            time.sleep(0.12)
        except Exception as exc:
            if log_fn:
                log_fn(f"  Request error: {exc}")
            break
    return [_parse_work(w, query) for w in collected[:max_results]]


# ── Pool (CSV) helpers ────────────────────────────────────────

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


def merge_records(new_records: list[dict], kw: dict) -> tuple[int, int]:
    """Merge new records into CSV. Returns (n_added, n_duplicate)."""
    pool = load_pool()
    existing = set(pool["openalex_id"].tolist())
    to_add, n_dup = [], 0
    seen_now: set[str] = set()
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


def _bibtex_key(row: pd.Series, seen: set[str] | None = None) -> str:
    """Generate a unique citekey: firstauthorlastname_year_firsttitleword.
    Pass seen to guarantee uniqueness across a batch by appending a numeric suffix."""
    authors = str(row.get("authors", "")).strip()
    year    = str(row.get("year", "")).strip()
    title   = str(row.get("title", "")).strip()

    # BibTeX convention: use only the primary author for the citekey
    first_author = authors.split(";")[0].strip()
    last_name = first_author.split()[-1] if first_author else "unknown"
    last_name = "".join(c for c in last_name.lower() if c.isalpha())

    # Skip articles/prepositions so citekey is meaningful (e.g. "fisheries_2023_future" not "a_2020_a")
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

    # Convert "First Last; First Last et al." → "Last, First and Last, First"
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
    if title:   lines.append(f"  title   = {{{title}}},")
    if author_str: lines.append(f"  author  = {{{author_str}}},")
    if year:    lines.append(f"  year    = {{{year}}},")
    if journal: lines.append(f"  journal = {{{journal}}},")
    if doi:     lines.append(f"  doi     = {{{doi}}},")
    if url:     lines.append(f"  url     = {{{url}}},")
    lines.append("}")
    return "\n".join(lines)


def generate_bibtex(df: pd.DataFrame) -> str:
    """Convert a filtered DataFrame to a BibTeX string."""
    seen: set[str] = set()
    entries = []
    for _, row in df.iterrows():
        key = _bibtex_key(row, seen)
        seen.add(key)
        entries.append(_make_bibtex_entry(row, key))
    return "\n\n".join(entries)


@st.cache_data
def compute_live_verdicts(pool_mtime: float, kw_json: str) -> list[str]:
    """Score all papers in the pool. Cached by CSV mtime + keyword state."""
    _pool = load_pool()
    _kw   = json.loads(kw_json)
    return [score_paper(r["title"], r["abstract"], _kw)["verdict"] for _, r in _pool.iterrows()]


@st.cache_data
def _cached_bibtex(bib_filter: str, pool_mtime: float) -> bytes:
    """Generate BibTeX for a verdict filter. Cached by filter + CSV mtime."""
    _pool = load_pool()
    masks = {
        "Included papers":   _pool["user_verdict"] == "include",
        "Included + Unsure": _pool["user_verdict"].isin(["include", "unsure"]),
        "All reviewed":      _pool["user_verdict"].str.strip().astype(bool),
        "Entire pool":       pd.Series([True] * len(_pool), index=_pool.index),
    }
    return generate_bibtex(_pool[masks[bib_filter]]).encode("utf-8")


def enrich_unpaywall(email: str, log_fn=None) -> tuple[int, int]:
    """Query Unpaywall for every paper with a DOI but no pdf_url yet.
    Returns (n_found_pdf, n_checked)."""
    pool = load_pool()
    needs = pool[
        (pool["doi"].str.strip() != "") &
        (pool["pdf_url"].str.strip() == "")
    ]
    n_found = 0
    for i, (idx, row) in enumerate(needs.iterrows()):
        doi = row["doi"].strip()
        if log_fn:
            log_fn(f"[{i+1}/{len(needs)}] {doi}")
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
                if url:
                    n_found += 1
        except Exception:
            pass
        time.sleep(0.12)
    save_pool(pool)
    return n_found, len(needs)


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
.abstract-box {
    background:#f8f9fa; border-left:4px solid #dee2e6;
    padding:12px 16px; border-radius:4px;
    font-size:0.92rem; line-height:1.6;
    max-height:260px; overflow-y:auto;
}
.score-bar-wrap { background:#e9ecef; border-radius:6px; height:14px; margin:2px 0 6px 0; }
.score-bar      { border-radius:6px; height:14px; }
.kw-match       { font-size:0.78rem; color:#555; margin:2px 0; }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────

if "paper_idx"        not in st.session_state:
    st.session_state.paper_idx        = 0
if "filter_verdict"   not in st.session_state:
    st.session_state.filter_verdict   = "All"
if "filter_reviewed"  not in st.session_state:
    st.session_state.filter_reviewed  = "All"
if "kw"               not in st.session_state:
    st.session_state.kw               = load_keywords()
if "search_cfg"       not in st.session_state:
    st.session_state.search_cfg       = load_search_config()
if "search_preview"   not in st.session_state:
    st.session_state.search_preview   = []
if "search_rows"      not in st.session_state:
    st.session_state.search_rows      = [{"id": 0, "op": None, "term": ""}]
    st.session_state.search_next_id   = 1

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
**Three tabs — use in order:**

1. **🔍 Search** — enter search terms,
   fetch papers from OpenAlex, add
   them to your review pool.

2. **📄 Review** — read each paper's
   abstract and keyword score, then
   click Include / Unsure / Exclude.
   Auto-saves after every click.

3. **⚙️ Keywords** — tune which words
   count as fisheries or future signals.
   Changes re-score all papers live.

Each tab has a **ℹ️ How it works**
section with detailed instructions.

**Progress is always saved** to
`data/full_text_review.csv` — close
the app and continue later anytime.
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

tab_search, tab_review, tab_keywords = st.tabs(["🔍 Search", "📄 Review", "⚙️ Keywords"])


# ══════════════════════════════════════════════════════════════
# TAB 1 — SEARCH
# ══════════════════════════════════════════════════════════════

with tab_search:
    st.markdown("### Search OpenAlex")

    with st.expander("ℹ️ How the Search tab works", expanded=False):
        st.markdown("""
**Workflow: Build query → Run → Preview → Add to pool → Review**

1. **Build your query** — add one row per search term. Choose how each term relates
   to the previous one using the **AND / OR / NOT** dropdown:
   - **AND** — both terms must appear (narrows results)
   - **OR** — either term may appear (broadens results)
   - **NOT** — excludes papers containing this term
   Multi-word phrases are automatically quoted. The live preview shows exactly what
   OpenAlex receives.

2. **Run Search** — fetches up to the configured maximum from OpenAlex and shows a
   preview of results. **Save query** adds it to your saved list without running.

3. **Saved queries** — your query history. Run any saved query again with ▶,
   or delete it with 🗑. Papers already in the pool are skipped automatically.

4. **Add to pool** — after previewing results, click to merge new papers into your
   review pool. Your existing verdicts are never overwritten.

> All progress is saved to `data/full_text_review.csv`. Close and reopen anytime.
""")

    # ── Settings (compact) ─────────────────────────────────────
    sc1, sc2, sc3 = st.columns([2, 2, 1])
    with sc1:
        max_results = st.number_input(
            "Max results per query",
            min_value=10, max_value=1000, step=10,
            value=int(cfg.get("max_results", 200)),
            help="OpenAlex returns up to 200 per page; larger values trigger multiple pages.",
        )
    with sc2:
        email = st.text_input(
            "Email (polite pool)",
            value=cfg.get("email", "tsz000@uit.no"),
            help="Included in API requests for OpenAlex polite pool priority.",
        )
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

    # ── Row-based boolean query builder ───────────────────────
    builder_col, help_col = st.columns([3, 2])

    with builder_col:
        st.markdown("#### Build a search query")
        st.caption("Add one search term per row. Use the AND / OR / NOT dropdown to define the relationship between terms.")

    with help_col:
        st.markdown("#### Tips & examples")
        st.markdown("""
<div style="background:#f8f9fa;border-radius:8px;padding:12px 16px;font-size:0.85rem;line-height:1.6">

**AND** — narrow results (both terms must appear)<br>
`fisheries` **AND** `climate change`

**OR** — broaden results (either term may appear)<br>
`aquaculture` **OR** `fish farming`

**NOT** — exclude papers with this term<br>
`fisheries` **AND** `management` **NOT** `aquaculture`

**Multi-word phrases** are auto-quoted:<br>
typing *small-scale fisheries* becomes `"small-scale fisheries"`

---

**Example queries**

*Futures research on fisheries:*<br>
`fisheries` AND `future` AND `scenario`

*Broad capture fishery coverage:*<br>
`fisheries` OR `"fish stocks"` OR `"marine capture"`

*Climate impacts, excluding aquaculture:*<br>
`fisheries` AND `"climate change"` NOT `aquaculture`

</div>
""", unsafe_allow_html=True)

    rows_to_delete = None
    for i, row in enumerate(st.session_state.search_rows):
        rid = row["id"]
        if i == 0:
            # First row: no operator dropdown
            c_label, c_term, c_del = st.columns([0.6, 5, 0.5])
            with c_label:
                st.markdown(
                    "<div style='text-align:right;padding-top:34px;color:#888;"
                    "font-size:0.85rem'>Find</div>",
                    unsafe_allow_html=True,
                )
        else:
            c_op, c_term, c_del = st.columns([0.9, 4.7, 0.5])
            with c_op:
                new_op = st.selectbox(
                    "op", ["AND", "OR", "NOT"],
                    index=["AND", "OR", "NOT"].index(row.get("op") or "AND"),
                    key=f"s_op_{rid}",
                    label_visibility="collapsed",
                )
                st.session_state.search_rows[i]["op"] = new_op

        with c_term:
            new_term = st.text_input(
                "term", value=row["term"], key=f"s_term_{rid}",
                placeholder='Term or phrase — multi-word phrases are auto-quoted, e.g. marine capture',
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

    # Live query preview
    current_query = build_query_from_rows(st.session_state.search_rows)
    if current_query:
        st.markdown("**Query preview:**")
        st.code(current_query, language=None)
    else:
        st.caption("Fill in at least one row above to build a query.")

    # Action buttons
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

    # ── Saved queries ──────────────────────────────────────────
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
                if st.button("🗑", key=f"del_sq_{qi}", help="Delete this query"):
                    cfg["saved_queries"].pop(qi)
                    st.session_state.search_cfg = cfg
                    save_search_config(cfg)
                    st.rerun()

    # ── Execute search ─────────────────────────────────────────
    query_to_run = None
    if run_btn and current_query:
        query_to_run = current_query
    elif "pending_run_query" in st.session_state:
        query_to_run = st.session_state.pop("pending_run_query")

    if query_to_run:
        all_records: list[dict] = []
        with st.status(f'Searching: {query_to_run[:60]}{"…" if len(query_to_run)>60 else ""}',
                       expanded=True) as status:
            def log(msg): st.write(msg)
            recs = fetch_openalex(query_to_run, max_results, email, log_fn=log)
            all_records.extend(recs)

            seen_ids: set[str] = set()
            deduped = [r for r in all_records
                       if r["openalex_id"] not in seen_ids
                       and not seen_ids.add(r["openalex_id"])]  # type: ignore[func-returns-value]
            st.write(f"**Total:** {len(all_records)} records → {len(deduped)} unique")
            status.update(label=f"Done — {len(deduped)} unique records found", state="complete")

        st.session_state.search_preview = deduped

    # ── Preview + merge ────────────────────────────────────────

    preview = st.session_state.search_preview
    if preview:
        existing_ids = set(pool["openalex_id"].tolist())
        n_new = sum(1 for r in preview if r["openalex_id"] not in existing_ids)
        n_dup = len(preview) - n_new

        st.markdown(f"#### Search results: {len(preview)} records")
        info_col, btn_col = st.columns([3, 1])
        with info_col:
            st.info(
                f"**{n_new}** new papers not yet in your pool  ·  "
                f"**{n_dup}** already in pool (will be skipped)  ·  "
                f"Pool will grow from **{total}** → **{total + n_new}**"
            )
        with btn_col:
            if st.button(
                f"➕ Add {n_new} papers to pool",
                type="primary",
                use_container_width=True,
                disabled=(n_new == 0),
            ):
                with st.spinner("Scoring and saving…"):
                    added, skipped = merge_records(preview, kw)
                st.success(f"Added {added} papers. {skipped} duplicates skipped.")
                st.session_state.search_preview = []
                st.rerun()

        # Preview table
        prev_df = pd.DataFrame(preview)[["title", "year", "journal", "authors", "openalex_id"]]
        prev_df.insert(0, "In pool", prev_df["openalex_id"].isin(existing_ids).map({True: "✓", False: "new"}))
        prev_df = prev_df.drop(columns=["openalex_id"])
        st.dataframe(prev_df, use_container_width=True, height=350)

    # ── Pool overview ──────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### Current review pool")
    if total == 0:
        st.info("Pool is empty. Run a search to add papers.")
    else:
        overview_cols = st.columns(4)
        vc = live_verdicts.value_counts() if total else {}
        with overview_cols[0]:
            st.metric("Total papers", total)
        with overview_cols[1]:
            st.metric("🟢 AI green", int(vc.get("green",  0)))
        with overview_cols[2]:
            st.metric("🟠 AI orange", int(vc.get("orange", 0)))
        with overview_cols[3]:
            st.metric("🔴 AI red", int(vc.get("red", 0)))

        if st.checkbox("Show full pool table"):
            show_df = pool[["title", "year", "journal", "search_query",
                            "date_added", "kw_verdict", "user_verdict"]].copy()
            st.dataframe(show_df, use_container_width=True, height=400)

        # OA enrichment
        n_with_doi    = int((pool["doi"].str.strip() != "").sum())
        n_with_pdf    = int((pool["pdf_url"].str.strip() != "").sum())
        n_needs_enrich = n_with_doi - n_with_pdf

        st.markdown("---")
        enrich_col, enrich_info = st.columns([1, 2])
        with enrich_col:
            enrich_btn = st.button(
                "🔗 Fetch open-access links",
                disabled=(n_needs_enrich == 0),
                use_container_width=True,
                help="Queries Unpaywall for a direct PDF link for every paper that has a DOI but no PDF link yet.",
            )
        with enrich_info:
            st.caption(
                f"📄 PDF links found: **{n_with_pdf}** of {n_with_doi} papers with DOIs  ·  "
                f"{n_needs_enrich} remaining to check"
            )

        if enrich_btn:
            with st.status("Querying Unpaywall…", expanded=True) as status:
                def _log(msg): st.write(msg)
                found, checked = enrich_unpaywall(email=cfg.get("email", "tsz000@uit.no"), log_fn=_log)
                status.update(
                    label=f"Done — found PDF links for {found} / {checked} papers",
                    state="complete",
                )
            st.rerun()

        st.markdown("")
        st.markdown("**Export**")
        exp_col1, exp_col2, exp_col3 = st.columns([1, 1, 2])

        with exp_col1:
            csv_bytes = pool.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                "⬇️ Full pool (CSV)",
                data=csv_bytes,
                file_name="full_text_review.csv",
                mime="text/csv",
                use_container_width=True,
            )

        with exp_col2:
            bib_filter = st.selectbox(
                "BibTeX export",
                ["Included papers", "Included + Unsure", "All reviewed", "Entire pool"],
                label_visibility="collapsed",
            )
            bib_count_map = {
                "Included papers":   int((pool["user_verdict"] == "include").sum()),
                "Included + Unsure": int(pool["user_verdict"].isin(["include", "unsure"]).sum()),
                "All reviewed":      int(pool["user_verdict"].str.strip().astype(bool).sum()),
                "Entire pool":       len(pool),
            }
            n_bib = bib_count_map[bib_filter]
            bib_bytes = _cached_bibtex(bib_filter, pool_mtime) if n_bib > 0 else b""
            st.download_button(
                f"⬇️ Export BibTeX ({n_bib})",
                data=bib_bytes,
                file_name="trendfish_references.bib",
                mime="text/plain",
                disabled=(n_bib == 0),
                use_container_width=True,
            )

        with exp_col3:
            st.caption(
                f"Included: **{int((pool['user_verdict']=='include').sum())}**  ·  "
                f"Unsure: **{int((pool['user_verdict']=='unsure').sum())}**  ·  "
                f"Excluded: **{int((pool['user_verdict']=='exclude').sum())}**  ·  "
                f"Unreviewed: **{int((pool['user_verdict'].str.strip()=='').sum())}**"
            )


# ══════════════════════════════════════════════════════════════
# TAB 2 — REVIEW
# ══════════════════════════════════════════════════════════════

with tab_review:

    with st.expander("ℹ️ How the Review tab works", expanded=False):
        st.markdown("""
**Goal:** Go through each paper and record your own inclusion/exclusion verdict.

**Keyword score (top of each paper)**
The coloured banner shows the *AI pre-screening verdict* based on the keyword lists you configured
in the ⚙️ Keywords tab:
- 🟢 **GREEN** — strong fisheries + future-orientation signals → likely include
- 🟠 **ORANGE** — borderline or missing abstract → needs your judgement
- 🔴 **RED** — no fisheries relevance, or purely retrospective → likely exclude

The score bars show the raw fisheries (0–8) and future (0–10) scores.
Click **Matched keywords** to see exactly which terms triggered the score.

**Navigating papers**
- Use **◀ Previous / Next ▶** or type a number to jump directly.
- Use the **sidebar filters** to focus on a subset:
  - *AI keyword verdict* — show only green/orange/red papers
  - *Review status* — show only papers you haven't reviewed yet (most efficient workflow)

**Giving your verdict**
1. Read the abstract (and the DOI if needed).
2. Click one of the three verdict buttons:
   - 🟢 **INCLUDE** — paper meets inclusion criteria for full-text review
   - 🟠 **UNSURE** — you're not certain; flag for later or a second opinion
   - 🔴 **EXCLUDE** — paper clearly does not meet criteria
3. Add optional notes (methodology, reason for exclusion, etc.).
4. The verdict saves immediately to `data/full_text_review.csv`.
   When filtering on *Unreviewed*, the app auto-advances to the next paper after each verdict.

> **Tip:** Start with the 🔴 RED filter to quickly clear obvious exclusions, then work through 🟠 ORANGE.
""")

    if total == 0:
        st.info("No papers in pool yet. Use the 🔍 Search tab to add papers.")
        st.stop()

    # ── Apply filters ──────────────────────────────────────────
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

    # ── Navigation ─────────────────────────────────────────────
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

    # ── Live keyword score ─────────────────────────────────────
    sc      = score_paper(str(row.get("title", "")), str(row.get("abstract", "")), kw)
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
            for t in sc["fish_hits"]:
                st.markdown(f"<div class='kw-match'>🐟 {t}</div>", unsafe_allow_html=True)
            if not sc["fish_hits"]: st.caption("none")
        with mc2:
            st.markdown("**Future hits**")
            for t in sc["future_hits"]:
                st.markdown(f"<div class='kw-match'>🔭 {t}</div>", unsafe_allow_html=True)
            if not sc["future_hits"]: st.caption("none")
        with mc3:
            st.markdown("**Exclusion hits**")
            for t in sc["retro"]:
                st.markdown(f"<div class='kw-match' style='color:#dc3545'>⏪ {t}</div>", unsafe_allow_html=True)
            for t in sc["bio"]:
                st.markdown(f"<div class='kw-match' style='color:#dc3545'>🧬 {t}</div>", unsafe_allow_html=True)
            for t in sc["nonfish"]:
                st.markdown(f"<div class='kw-match' style='color:#dc3545'>🚫 {t}</div>", unsafe_allow_html=True)
            if not (sc["retro"] + sc["bio"] + sc["nonfish"]): st.caption("none")

    # ── Paper metadata ─────────────────────────────────────────
    title   = str(row.get("title",   "")).strip() or "(no title)"
    authors = str(row.get("authors", "")).strip()
    year    = str(row.get("year",    "")).strip()
    journal = str(row.get("journal", "")).strip()
    doi     = str(row.get("doi",     "")).strip()
    oa_id   = str(row.get("openalex_id", "")).strip()
    sq      = str(row.get("search_query", "")).strip()
    pdf_url = str(row.get("pdf_url", "")).strip()
    is_oa   = str(row.get("is_oa",   "")).strip().lower()

    st.markdown(f"### {title}")
    meta = []
    if authors: meta.append(authors[:120] + ("…" if len(authors) > 120 else ""))
    if year:    meta.append(year)
    if journal: meta.append(f"*{journal}*")
    if meta:
        st.markdown(
            "<div style='color:#555;font-size:0.9rem'>" + "  ·  ".join(meta) + "</div>",
            unsafe_allow_html=True,
        )

    # PDF / access links
    link_row, badge_row = st.columns([3, 1])
    with link_row:
        if pdf_url:
            st.link_button("📄 Open PDF ↗", pdf_url, type="primary")
        sec_links = []
        if doi:
            doi_url = doi if doi.startswith("http") else f"https://doi.org/{doi}"
            sec_links.append(f"[DOI ↗]({doi_url})")
        if oa_id:
            sec_links.append(f"[OpenAlex ↗]({oa_id})")
        if sec_links:
            st.markdown("  ·  ".join(sec_links))
    with badge_row:
        badge = OA_STATUS.get(is_oa)
        if badge:
            icon, label, color = badge
            st.markdown(
                f"<div style='text-align:right;color:{color};font-size:0.85rem;"
                f"padding-top:6px'>{icon} {label}</div>",
                unsafe_allow_html=True,
            )

    if sq: st.caption(f"Found via: *{sq}*")

    abstract = str(row.get("abstract", "")).strip()
    ab_src   = str(row.get("abstract_source", "")).strip()
    if abstract and len(abstract) > 20:
        src_note = " *(fetched from OpenAlex)*" if ab_src == "openalex" else ""
        st.markdown(f"**Abstract**{src_note}")
        st.markdown(f"<div class='abstract-box'>{abstract}</div>", unsafe_allow_html=True)
    else:
        st.warning("No abstract available — check DOI before deciding.")

    st.markdown("---")

    # ── Verdict panel ──────────────────────────────────────────
    current_verdict = str(row.get("user_verdict", "")).strip()
    current_notes   = str(row.get("user_notes",   "")).strip()

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

    notes_input = st.text_area(
        "Notes (optional)", value=current_notes, height=80,
        placeholder="Comments, flags, or reasons…",
    )

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
        reviewed_at = str(row.get("reviewed_at", "")).strip()
        st.markdown(
            f"<div style='font-size:0.9rem;color:{col_hex};margin-top:6px'>"
            f"Your verdict: {icon} {label}"
            + (f" — *{current_notes}*" if current_notes else "")
            + "</div>", unsafe_allow_html=True,
        )
        if reviewed_at:
            st.caption(f"Reviewed at: {reviewed_at}")


# ══════════════════════════════════════════════════════════════
# TAB 3 — KEYWORDS
# ══════════════════════════════════════════════════════════════

with tab_keywords:
    st.markdown("### Keyword Configuration")

    with st.expander("ℹ️ How keyword scoring works", expanded=False):
        st.markdown("""
**The scoring system classifies each paper as green / orange / red based on two scores:**

| Score | How it's calculated | Max |
|-------|---------------------|-----|
| 🐟 **Fisheries score** | strong hits ×3 + medium hits ×2 + context hits ×1 | 8 |
| 🔭 **Future score** | strong hits ×4 + medium hits ×2 + weak hits ×1 | 10 |

A "hit" means the keyword (or phrase) appears anywhere in the title or abstract (case-insensitive).

**Green / Orange / Red decision rules:**

| Verdict | Condition |
|---------|-----------|
| 🟢 GREEN | (fish ≥ 4 and future ≥ 6, no retro hits) OR (fish ≥ 3 and future ≥ 4, no retro) OR (fish ≥ 2 and future ≥ 6) |
| 🔴 RED | fish = 0, OR (future = 0 and a retrospective hit), OR (biology-only hit with fish ≤ 2 and future < 2) |
| 🟠 ORANGE | Everything else — borderline or abstract missing |

**Exclusion lists** act as hard overrides:
- **Retrospective** — if the paper has no future orientation AND a retrospective term → red
- **Biology-only** — genomics/genetics focus without fisheries futures context → red
- **Non-fisheries** — agriculture, livestock etc. with no fisheries signal → red

**Editing keywords**
- Click any cell to edit a keyword in place.
- Click the **+** row at the bottom of a table to add a new keyword.
- Select a row (checkbox on the left) and press `Delete` / `Backspace` to remove it.
- Changes apply instantly to the Review tab scoring.
- Click **Save to file** to persist changes across sessions.
- Click **Reset to defaults** to restore the original keyword lists.

> **Tip:** If good papers are being scored red, check which fisheries/future keywords they lack and add them to the medium or context lists. If too many irrelevant papers are green, strengthen the exclusion lists.
""")

    st.markdown(
        "Edit, add, or delete keywords. Changes apply instantly to scoring in the Review tab. "
        "Click **Save to file** to persist across sessions."
    )

    kw_changed = False
    new_kw     = {}

    groups = [
        ("🐟 Fisheries relevance", ["fish_strong", "fish_med", "fish_context"]),
        ("🔭 Future orientation",  ["future_strong", "future_med", "future_weak"]),
        ("🚫 Exclusion criteria",  ["excl_retro", "excl_bio", "excl_nonfish"]),
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
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    key=f"de_{key}",
                    column_config={
                        "keyword": st.column_config.TextColumn(
                            "Keyword / phrase",
                            help="Lowercase. Matched anywhere in title + abstract.",
                            max_chars=120,
                        )
                    },
                )
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

    st.markdown("---")
    st.markdown("**Scoring thresholds**")
    st.markdown("""
| Verdict | Condition |
|---------|-----------|
| 🟢 GREEN | fish ≥ 4 + future ≥ 6 (no retro), OR fish ≥ 3 + future ≥ 4 (no retro), OR fish ≥ 2 + future ≥ 6 |
| 🟠 ORANGE | Borderline or missing abstract |
| 🔴 RED | fish = 0, OR (future = 0 + retro hit), OR (bio hit + fish ≤ 2 + future < 2) |

Fisheries score = strong×3 + medium×2 + context×1, capped at 8.
Future score = strong×4 + medium×2 + weak×1, capped at 10.
""")
