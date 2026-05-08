"""Inspect which fisheries entries fall into the Other STEEP category."""
import re
from pathlib import Path
import pandas as pd

BASE_DIR  = Path(__file__).parent.parent
DATA_FILE = BASE_DIR / "data" / "manual_literature_search.xlsx"

STEEP_RULES = {
    "Environmental": ["climate","environment","biodiversity","ecosystem","ocean","marine","sea",
        "water","pollution","habitat","carbon","emission","green","nature","resource scarcity",
        "erosion","deforestation","renewable energy","sustainability","ecology","fish stock",
        "nature conservation","nature-based","natural resource","natural environment",
        "overfishing","fishing","fisheries","fishery","fisher","fishers","aquaculture","aquatic",
        "seafood","harvest","catch","bycatch","by-catch","food web","carrying capacity",
        "stock assessment","spawn","recruitment","trophic",
        "nature conservation","nature-based","natural resource","natural environment"],
    "Technological": ["technolog","digital","automation","artificial intelligence","ai ","data",
        "innovation","internet","cyber","robot","autonomous","biotech","synthetic biology",
        "microbiome","hyperconnectiv","electrif","knowledge"],
    "Economic": ["economic","trade","market","finance","fiscal","gdp","consumption","globali",
        "growth","debt","income","employment","labour","labor","industry",
        "investment","supply chain","demand","profit","revenue","subsid","cost","price","value chain"],
    "Social": ["social","demograph","population","urban","aging","health","inequalit","migration",
        "culture","education","wellbeing","well-being","individual","community","consumer",
        "lifestyle","food security","gender","governance of people","livelihood","artisanal",
        "small-scale","indigenous"],
    "Political": ["politic","governance","geopolit","security","power","regulation","law",
        "policy","democracy","conflict","supranational","institution","global order","protectionism",
        "management","quota","treaty","iuu","compliance"],
}

NOTE_PATTERNS = [r"^n=a$",r"^none$",r"^topic$",r"^keyword",r"^future\s*$",r"^political variables$",
    r"^economic variables$",r"^social variables$",r"^environmental variables$",
    r"^find something",r"^interesting article",r"^book on",r"^manifesto$",
    r"^brief vision",r"^bell prediction"]


def classify_steep(text: str) -> str:
    if not text or str(text).strip().lower() in ("none", "nan", "?", "n=a", ""):
        return "Other"
    t = str(text).lower()
    for cat, kws in STEEP_RULES.items():
        if any(kw in t for kw in kws):
            return cat
    return "Other"


def is_note_trend(val) -> bool:
    if pd.isna(val):
        return True
    v = str(val).strip().lower()
    return any(re.match(p, v) for p in NOTE_PATTERNS)


def has_reference(row) -> bool:
    has_title  = pd.notna(row["title"])  and str(row["title"]).strip().lower() not in ("", "nan")
    has_author = pd.notna(row["author"]) and str(row["author"]).strip().lower() not in ("", "nan")
    has_source = (pd.notna(row["source"]) and str(row["source"]).strip().lower() not in ("", "nan")
                  and str(row["source"]).strip().lower() != str(row["author"]).strip().lower())
    return has_title or has_author or has_source


df = pd.read_excel(DATA_FILE, sheet_name="Futures of Fisheries")
df.columns = ["trend","details","author","source","title","comments","doi","year","source_type","_a","_b"]
df = df.drop(columns=["_a","_b"])

df["_has_ref"]    = df.apply(has_reference, axis=1)
df["_note_trend"] = df["trend"].apply(is_note_trend)
df = df[df["_has_ref"] | (~df["_note_trend"])]

def clean_trend(val):
    if pd.isna(val): return None
    v = str(val).strip()
    if v in ("?", "N=A", "None") or is_note_trend(val): return None
    return v

df["trend"] = df["trend"].apply(clean_trend)
df["steep"] = df.apply(
    lambda r: classify_steep(
        " ".join(str(r[c]) for c in ["trend","details","title"] if pd.notna(r[c]))
    ), axis=1
)

print(f"Total fisheries entries: {len(df)}")
print(f"Other count: {len(df[df['steep'] == 'Other'])}")
print()
print("=== Entries classified as 'Other' ===")
for _, row in df[df["steep"] == "Other"].iterrows():
    trend = str(row.get("trend", "")).strip()
    title = str(row.get("title", "")).strip()
    details = str(row.get("details", "")).strip()
    print(f"  TREND  : {trend[:70]}")
    print(f"  TITLE  : {title[:80]}")
    print(f"  DETAILS: {details[:70]}")
    print()
