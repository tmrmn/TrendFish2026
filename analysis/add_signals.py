"""
Add documented weak signal entries to the 'Signals of Future Fisheries' sheet.
Appends rows to the existing Excel database — does NOT overwrite existing entries.

Run once to seed the database with the three candidate signals identified
in the horizon scan. Each entry follows the same column schema as the
other sheets: Trend, Details, Author, Source, Title, Comments, DOI, Year, Type
"""

from pathlib import Path
import openpyxl

DATA_FILE = Path(__file__).parent.parent / "data" / "manual_literature_search.xlsx"
SHEET_NAME = "Signals of Future Fisheries"

NEW_SIGNALS = [
    {
        "trend":    "Rise of alternative protein competing with wild-catch seafood",
        "details":  (
            "Precision fermentation and cultivated (cell-based) seafood are moving "
            "from laboratory to commercial scale. Companies including BlueNalu (US) "
            "and Wildtype (US) have demonstrated cell-cultivated fish fillets; "
            "Singapore approved the first commercial sale of cultivated meat in 2020. "
            "If cost curves follow those of plant-based meat, price parity with "
            "premium wild-caught species could be reached within a decade, "
            "with major implications for demand and market structure."
        ),
        "author":   "Horizon scan — compiled by reviewer",
        "source":   "https://gfi.org/resource/cultivated-seafood/",
        "title":    "Cultivated Seafood: State of the Industry Report",
        "comments": (
            "Good Food Institute annual report tracks investment and regulatory "
            "milestones. Singapore Food Agency approval (Dec 2020) is a key landmark."
        ),
        "doi":      "",
        "year":     2024,
        "type":     "Grey literature / Industry report",
    },
    {
        "trend":    "Autonomous and remotely operated vessels entering commercial fishing",
        "details":  (
            "Several technology companies and fishing nations are trialling "
            "autonomous surface vessels (ASVs) and remotely piloted fishing craft. "
            "Norway's autonomous cargo vessel Yara Birkeland (launched 2022) "
            "demonstrates the technical feasibility of uncrewed vessels in cold-water "
            "operations. Japan's Fisheries Agency has funded trials of AI-guided "
            "longliners. Automation threatens to decouple fishing capacity from crew "
            "availability, with potential implications for labour, safety regulation, "
            "and coastal community livelihoods."
        ),
        "author":   "Horizon scan — compiled by reviewer",
        "source":   "https://www.yara.com/news-and-media/news/archive/2022/",
        "title":    "Yara Birkeland autonomous ship — first commercial voyage",
        "comments": (
            "Regulatory frameworks for autonomous fishing vessels are essentially "
            "absent. IMO is developing guidelines for Maritime Autonomous Surface "
            "Ships (MASS) but these do not yet cover fishing vessels specifically."
        ),
        "doi":      "",
        "year":     2022,
        "type":     "Grey literature / News",
    },
    {
        "trend":    "Geopolitical contest over Arctic and sub-Arctic fishing access",
        "details":  (
            "Retreating Arctic sea ice is opening new fishing grounds in the central "
            "Arctic Ocean and shifting the ranges of commercially important species "
            "northward. The 2018 Agreement to Prevent Unregulated High Seas Fisheries "
            "in the Central Arctic Ocean (signed by ten parties) represents the first "
            "precautionary governance instrument for this space, but access tensions "
            "between Russia, Norway, Iceland, Canada, and the US are intensifying. "
            "Projections suggest commercially viable mackerel, herring, and cod "
            "stocks may be accessible in central Arctic waters within 10–20 years "
            "under high-emissions scenarios."
        ),
        "author":   "Horizon scan — compiled by reviewer",
        "source":   "https://arcticportal.org/arctic-governance/fisheries",
        "title":    "Arctic Fisheries Governance and the 2018 Central Arctic Ocean Agreement",
        "comments": (
            "Link to the agreement text: "
            "https://www.state.gov/central-arctic-ocean-fisheries-agreement/. "
            "ICES working group on Arctic fisheries publishes annual updates."
        ),
        "doi":      "10.1093/icesjms/fsz052",
        "year":     2019,
        "type":     "Academic article / Policy document",
    },
]


def get_last_row(sheet):
    """Return index of the last row with any data."""
    for row in range(sheet.max_row, 0, -1):
        if any(sheet.cell(row, col).value for col in range(1, sheet.max_column + 1)):
            return row
    return 1  # header only


def add_signals():
    wb = openpyxl.load_workbook(DATA_FILE)

    if SHEET_NAME not in wb.sheetnames:
        print(f"  Sheet '{SHEET_NAME}' not found — creating it.")
        ws = wb.create_sheet(SHEET_NAME)
        ws.append(["Trend", "Details", "Author", "Source", "Title",
                   "Comments", "DOI", "Year", "Type"])
    else:
        ws = wb[SHEET_NAME]

    last = get_last_row(ws)
    print(f"  Sheet has {last} row(s) before additions (incl. header).")

    added = 0
    for sig in NEW_SIGNALS:
        # Skip if this trend already exists (simple duplicate check on trend text)
        trend_lower = sig["trend"].lower()
        already = any(
            str(ws.cell(r, 1).value or "").lower() == trend_lower
            for r in range(2, last + 1)
        )
        if already:
            print(f"  Skipped (already exists): {sig['trend'][:60]}")
            continue

        ws.append([
            sig["trend"],
            sig["details"],
            sig["author"],
            sig["source"],
            sig["title"],
            sig["comments"],
            sig["doi"],
            sig["year"],
            sig["type"],
        ])
        added += 1
        print(f"  Added: {sig['trend'][:70]}")

    wb.save(DATA_FILE)
    print(f"\n  Saved {added} new signal(s) to '{DATA_FILE.name}'.")


if __name__ == "__main__":
    print(f"Adding signals to: {DATA_FILE}")
    add_signals()
