"""
Bona Vacantia List Monitor — GitHub Actions version
=====================================================
Runs every hour via GitHub Actions (free, cloud-based, no PC needed).
Downloads the latest BV Unclaimed Estates list from gov.uk, compares it
to the previously saved version, and emails a styled ON/OFF Excel report
to admin@family-wise.co.uk — but ONLY if something changed.

Credentials are stored as GitHub Secrets (not in this file).
"""

import os
import re
import smtplib
import datetime
import requests
import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ─────────────────────────────────────────────────────────────
#  CONFIGURATION
#  Credentials come from GitHub Secrets — nothing to edit here
# ─────────────────────────────────────────────────────────────

GMAIL_SENDER   = os.environ.get("GMAIL_SENDER",   "bvlistdaemon@gmail.com")
GMAIL_APP_PASS = os.environ.get("GMAIL_APP_PASS", "iinjdbkmfmkntvxl")
EMAIL_TO       = os.environ.get("EMAIL_TO",       "admin@family-wise.co.uk")

SCRIPT_DIR      = Path(__file__).parent
SAVED_LIST_PATH = SCRIPT_DIR / "bv_saved_list.csv"
OUTPUT_XLSX     = SCRIPT_DIR / "BV_Changes.xlsx"
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
        raise RuntimeError(
            "Could not find a download link on the BV page. "
            "The page layout may have changed."
        )

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
#  STEP 3 — Build styled Excel report
# ─────────────────────────────────────────────────────────────

def style_sheet(ws, df, title, header_color, tab_color):
    ws.title = title
    ws.sheet_properties.tabColor = tab_color

    preferred = ['BV Reference', 'Forename', 'Surname', 'Date of Death', 'Place of Death']
    cols = [c for c in preferred if c in df.columns] or list(df.columns)

    col_widths = {
        'BV Reference': 18, 'Forename': 15, 'Surname': 20,
        'Date of Death': 16, 'Place of Death': 35
    }

    thin        = Side(style='thin', color='CCCCCC')
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill('solid', start_color=header_color)

    ws.merge_cells(f'A1:{get_column_letter(len(cols))}1')
    ws['A1'] = title
    ws['A1'].font      = Font(name='Arial', bold=True, size=13, color='FFFFFF')
    ws['A1'].fill      = header_fill
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 22

    ws.merge_cells(f'A2:{get_column_letter(len(cols))}2')
    ws['A2'] = f'Total entries: {len(df)}'
    ws['A2'].font      = Font(name='Arial', italic=True, size=10, color='555555')
    ws['A2'].alignment = Alignment(horizontal='center')
    ws.row_dimensions[2].height = 16

    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=3, column=ci, value=col)
        cell.font      = Font(name='Arial', bold=True, size=10, color='FFFFFF')
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border    = border
        ws.column_dimensions[get_column_letter(ci)].width = col_widths.get(col, 20)
    ws.row_dimensions[3].height = 18

    for ri, (_, row) in enumerate(df[cols].iterrows(), 4):
        fill_color = 'F9F9F9' if ri % 2 == 0 else 'FFFFFF'
        for ci, col in enumerate(cols, 1):
            val = row[col]
            if hasattr(val, 'strftime'):
                val = val.strftime('%d/%m/%Y')
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font      = Font(name='Arial', size=10)
            cell.fill      = PatternFill('solid', start_color=fill_color)
            cell.border    = border
            cell.alignment = Alignment(vertical='center')

    ws.freeze_panes = 'A4'


def build_excel(on_df: pd.DataFrame, off_df: pd.DataFrame) -> Path:
    wb = Workbook()
    ws1 = wb.active
    style_sheet(ws1, on_df,  'ON List (New Entries)',      '2E7D32', '00AA00')
    ws2 = wb.create_sheet()
    style_sheet(ws2, off_df, 'OFF List (Removed Entries)', 'B71C1C', 'CC0000')
    wb.save(OUTPUT_XLSX)
    return OUTPUT_XLSX

# ─────────────────────────────────────────────────────────────
#  STEP 4 — Send email
# ─────────────────────────────────────────────────────────────

def send_email(on_df: pd.DataFrame, off_df: pd.DataFrame, xlsx_path: Path):
    today   = datetime.date.today().strftime("%d %B %Y")
    subject = f"Bona Vacantia List Update — {today}"

    lines = [
        f"BV Unclaimed Estates list changes detected — {today}.",
        "",
        f"  ✅  NEW entries (ON list):      {len(on_df)}",
        f"  ❌  REMOVED entries (OFF list): {len(off_df)}",
        "",
    ]

    if not on_df.empty and 'Surname' in on_df.columns:
        lines.append("New entries (ON):")
        for _, r in on_df.iterrows():
            lines.append(f"  + {r.get('Forename','')} {r.get('Surname','')}  [{r.get('BV Reference','')}]")
        lines.append("")

    if not off_df.empty and 'Surname' in off_df.columns:
        lines.append("Removed entries (OFF):")
        for _, r in off_df.iterrows():
            lines.append(f"  - {r.get('Forename','')} {r.get('Surname','')}  [{r.get('BV Reference','')}]")
        lines.append("")

    lines.append("Full details are in the attached Excel file.")
    body = "\n".join(lines)

    msg            = MIMEMultipart()
    msg["From"]    = GMAIL_SENDER
    msg["To"]      = EMAIL_TO
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with open(xlsx_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={xlsx_path.name}")
    msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_SENDER, GMAIL_APP_PASS)
        server.sendmail(GMAIL_SENDER, EMAIL_TO, msg.as_string())

    log(f"Email sent to {EMAIL_TO}")

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
            log("No saved list found — saving current list as baseline. No email sent.")
            save_list(new_df)
            return

        on_df, off_df = compare(old_df, new_df)
        log(f"Comparison done — ON: {len(on_df)}, OFF: {len(off_df)}")

        if on_df.empty and off_df.empty:
            log("No changes detected. No email sent.")
            save_list(new_df)
            return

        log("Changes found — building Excel report...")
        xlsx = build_excel(on_df, off_df)

        log("Sending email...")
        send_email(on_df, off_df, xlsx)

        save_list(new_df)
        log("All done ✅")

    except Exception as e:
        log(f"ERROR: {e}")
        raise


if __name__ == "__main__":
    main()

