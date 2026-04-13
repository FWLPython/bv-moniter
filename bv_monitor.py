"""
Bona Vacantia List Monitor — GitHub Notifications version
==========================================================
No email sending. When changes are found, this script saves a
CHANGES_FOUND.md file to the repo. GitHub's own notification
system then emails you automatically when the file is updated.

Zero authentication issues. Zero SMTP. Just works.
"""

import re
import datetime
import requests
import pandas as pd
from pathlib import Path

SCRIPT_DIR      = Path(__file__).parent
SAVED_LIST_PATH = SCRIPT_DIR / "bv_saved_list.csv"
CHANGES_FILE    = SCRIPT_DIR / "CHANGES_FOUND.md"
LOG_FILE        = SCRIPT_DIR / "bv_monitor.log"

BV_PAGE_URL = "https://www.gov.uk/government/statistical-data-sets/unclaimed-estates-list"

# ─────────────────────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────────────────────

def log(msg):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

# ─────────────────────────────────────────────────────────────
#  STEP 1 — Find and download the latest BV spreadsheet
# ─────────────────────────────────────────────────────────────

def find_download_url() -> str:
    headers = {"User-Agent": "Mozilla/5.0 (BV-Monitor/1.0)"}
    resp = requests.get(BV_PAGE_URL, headers=headers, timeout=30)
    resp.raise_for_status()

    urls = re.findall(
        r'https://assets\.publishing\.service\.gov\.uk[^"\']+\.(?:xlsx|csv)',
        resp.text
    )

    if not urls:
        raise RuntimeError("Could not find a download link on the BV page.")

    return urls[0]


def download_bv_list(url: str) -> pd.DataFrame:
    headers = {"User-Agent": "Mozilla/5.0 (BV-Monitor/1.0)"}
    resp = requests.get(url, headers=headers, timeout=60)
    resp.raise_for_status()

    ext = ".xlsx" if url.endswith(".xlsx") else ".csv"
    tmp_path = SCRIPT_DIR / f"_tmp_bv{ext}"

    with open(tmp_path, "wb") as f:
        f.write(resp.content)

    df = pd.read_excel(tmp_path) if ext == ".xlsx" else pd.read_csv(tmp_path)
    tmp_path.unlink()

    df['BV Reference'] = df['BV Reference'].astype(str).str.strip()
    return df

# ─────────────────────────────────────────────────────────────
#  STEP 2 — Compare with saved list
# ─────────────────────────────────────────────────────────────

def load_saved_list():
    if SAVED_LIST_PATH.exists():
        df = pd.read_csv(SAVED_LIST_PATH, dtype=str)
        df['BV Reference'] = df['BV Reference'].astype(str).str.strip()
        return df
    return None


def save_list(df: pd.DataFrame):
    df.to_csv(SAVED_LIST_PATH, index=False)


def compare(old_df: pd.DataFrame, new_df: pd.DataFrame):
    old_refs = set(old_df['BV Reference'])
    new_refs = set(new_df['BV Reference'])

    sort_col = 'Surname' if 'Surname' in new_df.columns else new_df.columns[0]

    on_df  = new_df[new_df['BV Reference'].isin(new_refs - old_refs)].sort_values(sort_col)
    off_df = old_df[old_df['BV Reference'].isin(old_refs - new_refs)].sort_values(
        sort_col if sort_col in old_df.columns else old_df.columns[0]
    )

    return on_df, off_df

# ─────────────────────────────────────────────────────────────
#  STEP 3 — Write changes to CHANGES_FOUND.md
#  GitHub will email you when this file is committed
# ─────────────────────────────────────────────────────────────

def write_changes_file(on_df: pd.DataFrame, off_df: pd.DataFrame):
    now = datetime.datetime.now().strftime("%d %B %Y at %H:%M")

    lines = [
        f"# BV List Changes — {now}",
        "",
        f"**✅ NEW entries (ON): {len(on_df)}**",
        f"**❌ REMOVED entries (OFF): {len(off_df)}**",
        "",
    ]

    if not on_df.empty:
        lines.append("## New Entries (ON)")
        lines.append("| BV Reference | Forename | Surname | Date of Death | Place of Death |")
        lines.append("|---|---|---|---|---|")
        for _, r in on_df.iterrows():
            lines.append(
                f"| {r.get('BV Reference','')} "
                f"| {r.get('Forename','')} "
                f"| {r.get('Surname','')} "
                f"| {r.get('Date of Death','')} "
                f"| {r.get('Place of Death','')} |"
            )
        lines.append("")

    if not off_df.empty:
        lines.append("## Removed Entries (OFF)")
        lines.append("| BV Reference | Forename | Surname | Date of Death | Place of Death |")
        lines.append("|---|---|---|---|---|")
        for _, r in off_df.iterrows():
            lines.append(
                f"| {r.get('BV Reference','')} "
                f"| {r.get('Forename','')} "
                f"| {r.get('Surname','')} "
                f"| {r.get('Date of Death','')} "
                f"| {r.get('Place of Death','')} |"
            )
        lines.append("")

    with open(CHANGES_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    log(f"Changes written to {CHANGES_FILE.name}")

# ─────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────

def main():
    log("── BV Monitor starting ──")

    try:
        log("Finding download URL on gov.uk...")
        url = find_download_url()
        log(f"Downloading: {url}")
        new_df = download_bv_list(url)
        log(f"Downloaded {len(new_df)} entries.")

        old_df = load_saved_list()

        if old_df is None:
            log("No saved list found — saving baseline. No changes recorded.")
            save_list(new_df)
            return

        on_df, off_df = compare(old_df, new_df)
        log(f"Comparison done — ON: {len(on_df)}, OFF: {len(off_df)}")

        if on_df.empty and off_df.empty:
            log("No changes detected.")
            save_list(new_df)
            return

        log("Changes found — writing changes file...")
        write_changes_file(on_df, off_df)

        save_list(new_df)
        log("All done ✅")

    except Exception as e:
        log(f"ERROR: {e}")
        raise


if __name__ == "__main__":
    main()




