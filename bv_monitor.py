"""
Bona Vacantia List Monitor — SMTP2GO Email version
====================================================
When changes are found, this script emails a summary via SMTP2GO
instead of writing a CHANGES_FOUND.md file for GitHub notifications.
"""

import re
import os
import smtplib
import ssl
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests
import pandas as pd
from pathlib import Path

SCRIPT_DIR      = Path(__file__).parent
SAVED_LIST_PATH = SCRIPT_DIR / "bv_saved_list.csv"
LOG_FILE        = SCRIPT_DIR / "bv_monitor.log"

BV_PAGE_URL = "https://www.gov.uk/government/statistical-data-sets/unclaimed-estates-list"

# ─────────────────────────────────────────────────────────────
#  EMAIL / SMTP2GO CONFIG
# ─────────────────────────────────────────────────────────────
# It's best practice to keep credentials out of source code.
# Set these as environment variables before running the script, e.g.:
#   export SMTP2GO_USERNAME="FamilyWiseBVLIST"
#   export SMTP2GO_PASSWORD="your-password-here"
# The os.environ.get(...) calls below will fall back to the
# hardcoded values only if the environment variables aren't set.

SMTP_HOST     = "mail.smtp2go.com"
SMTP_PORT_TLS = 2525   # STARTTLS
SMTP_PORT_SSL = 465    # SSL

SMTP_USERNAME = os.environ.get("SMTP2GO_USERNAME", "FamilyWiseBVLIST")
SMTP_PASSWORD = os.environ.get("SMTP2GO_PASSWORD", "0LGj2guAb1rcYLPJ1skgMV55")

EMAIL_FROM    = os.environ.get("BV_EMAIL_FROM", "bv@family-wise.co.uk")
EMAIL_TO      = os.environ.get("BV_EMAIL_TO", "admin@family-wise.co.uk")  # comma-separate multiple recipients
EMAIL_SUBJECT_PREFIX = "Bona Vacantia List Update"

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
#  STEP 3 — Build email content and send via SMTP2GO
# ─────────────────────────────────────────────────────────────

def build_email_body(on_df: pd.DataFrame, off_df: pd.DataFrame) -> str:
    now = datetime.datetime.now().strftime("%d %B %Y at %H:%M")

    lines = [
        f"Bona Vacantia List Changes — {now}",
        "",
        f"NEW entries (ON): {len(on_df)}",
        f"REMOVED entries (OFF): {len(off_df)}",
        "",
    ]

    def format_table(df, heading):
        table_lines = [heading, "-" * len(heading)]
        for _, r in df.iterrows():
            table_lines.append(
                f"{r.get('BV Reference','')} | "
                f"{r.get('Forename','')} | "
                f"{r.get('Surname','')} | "
                f"{r.get('Date of Death','')} | "
                f"{r.get('Place of Death','')}"
            )
        table_lines.append("")
        return table_lines

    if not on_df.empty:
        lines.extend(format_table(on_df, "New Entries (ON)"))

    if not off_df.empty:
        lines.extend(format_table(off_df, "Removed Entries (OFF)"))

    return "\n".join(lines)


def send_email(subject: str, body: str, use_ssl: bool = True):
    """
    Sends an email via SMTP2GO.
    use_ssl=True  -> connects on port 465 using SMTP_SSL
    use_ssl=False -> connects on port 2525 using STARTTLS
    """
    msg = MIMEMultipart()
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    recipients = [addr.strip() for addr in EMAIL_TO.split(",") if addr.strip()]

    if use_ssl:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT_SSL, context=context, timeout=30) as server:
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(EMAIL_FROM, recipients, msg.as_string())
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT_TLS, timeout=30) as server:
            server.starttls(context=ssl.create_default_context())
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(EMAIL_FROM, recipients, msg.as_string())

    log(f"Email sent to {EMAIL_TO} via SMTP2GO ({'SSL 465' if use_ssl else 'TLS 2525'})")

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

        log("Changes found — building and sending email...")
        body = build_email_body(on_df, off_df)
        subject = f"{EMAIL_SUBJECT_PREFIX} — {len(on_df)} new, {len(off_df)} removed"

        try:
            send_email(subject, body, use_ssl=True)
        except Exception as ssl_err:
            log(f"SSL (465) send failed: {ssl_err}. Retrying with TLS (2525)...")
            send_email(subject, body, use_ssl=False)

        save_list(new_df)
        log("All done ✅")

    except Exception as e:
        log(f"ERROR: {e}")
        raise


if __name__ == "__main__":
    main()
