import os
import csv
import smtplib
import traceback
from datetime import datetime
from email.mime.text import MIMEText

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ============================================================================
# CONFIG - edit these if paths / names change
# ============================================================================

# ---- Source Google Sheet (fabric current / open PO / WIP tabs) ----
SOURCE_SHEET_ID = "1oKefmjFRJI0N_CniKjqLNzJ0C3qqbS0XBT1BP0nA31g"
CURRENT_TAB = "Partial Possible Current"
OPEN_PO_TAB = "Open PO"
WIP_TAB = "WIP"

# Path to the service account JSON key file. In GitHub Actions this file is
# written by the workflow from the GOOGLE_SERVICE_ACCOUNT_JSON secret.
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON_PATH", "service_account.json")

# ---- Local working files (relative folder inside the repo/runner) ----
BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(BASE_DIR, exist_ok=True)
CURRENT_FILE = os.path.join(BASE_DIR, "Fabric_Current.csv")
OPEN_PO_FILE = os.path.join(BASE_DIR, "Fabric_Open_PO.csv")
WIP_FILE = os.path.join(BASE_DIR, "Fabric_WIP.csv")
OUTPUT_FILE = os.path.join(BASE_DIR, "Fabric_overall.csv")

# ---- Destination Google Sheet (where the combined SKU overall data goes) ----
DEST_SPREADSHEET_ID = "1qa25mI_xoAfSJhx6L11ZSB0xwS5nCg9xBk5AOxr9Ehw"
DEST_TAB_NAME = "Fabric Overall"

# ---- Gmail credentials, read from environment variables (set via GitHub
# Secrets in the workflow). GMAIL_APP_PASSWORD must be a Gmail "App Password",
# not your normal login password. ----
GMAIL_EMAIL_ENV = "GMAIL_EMAIL"
GMAIL_APP_PASSWORD_ENV = "GMAIL_APP_PASSWORD"

# Who receives the report email. Leave as None to send to the same Gmail
# account read from the env vars above. Or set to a string / list of strings.
EMAIL_TO = None

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587

SCOPES_READONLY = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]
SCOPES_FULL = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ============================================================================
# STEP TRACKING
# ============================================================================

class StepTracker:
    """Keeps an ordered record of every step's outcome for the email report."""

    def __init__(self):
        self.steps = []

    def success(self, name, detail="", rows=None):
        self.steps.append({
            "name": name,
            "status": "SUCCESS",
            "detail": detail,
            "rows": rows,
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        print(f"[OK] {name}" + (f" - {detail}" if detail else ""))

    def failed(self, name, error):
        self.steps.append({
            "name": name,
            "status": "FAILED",
            "detail": str(error),
            "rows": None,
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        print(f"[FAILED] {name} - {error}")

    def skipped(self, name, reason):
        self.steps.append({
            "name": name,
            "status": "SKIPPED",
            "detail": reason,
            "rows": None,
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        print(f"[SKIPPED] {name} - {reason}")

    def has_failure(self):
        return any(s["status"] == "FAILED" for s in self.steps)

    def first_failure(self):
        for s in self.steps:
            if s["status"] == "FAILED":
                return s
        return None

    def to_email_body(self):
        overall_ok = not self.has_failure()
        lines = []

        if overall_ok:
            lines.append("SKU Overall sheet has been updated with fabric data successfully.")
        else:
            fail = self.first_failure()
            lines.append(
                f"The Fabric Overall pipeline FAILED at step: \"{fail['name']}\"."
            )
            lines.append("The SKU Overall sheet was NOT updated.")

        lines.append("")
        lines.append("Run time: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        lines.append("")
        lines.append("Step-by-step detail:")
        lines.append("-" * 60)

        for i, s in enumerate(self.steps, start=1):
            mark = {"SUCCESS": "[OK]", "FAILED": "[FAILED]", "SKIPPED": "[SKIPPED]"}[s["status"]]
            row_txt = f" | Rows: {s['rows']}" if s["rows"] is not None else ""
            lines.append(f"{i}. {mark} {s['name']}{row_txt}")
            if s["detail"]:
                lines.append(f"    {s['detail']}")
            lines.append(f"    Time: {s['time']}")
            lines.append("")

        return "\n".join(lines), overall_ok


# ============================================================================
# HELPERS
# ============================================================================

def load_sheet_as_df(sheet_id, tab_name, creds_path, scopes):
    """Pull a worksheet from Google Sheets straight into memory."""
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.worksheet(tab_name)
    records = worksheet.get_all_records()
    return pd.DataFrame(records)


def _clean_cell(cell):
    """Strip normal whitespace plus stray non-breaking spaces (\\xa0) that
    Excel/Windows sometimes saves into CSVs, which otherwise pass str.strip()
    silently but break comparisons like the '@' check below."""
    return cell.replace("\xa0", " ").replace("\u200b", "").strip()


def load_gmail_credentials_from_env():
    """Read Gmail email + app password from environment variables (GitHub
    Secrets), instead of a local CSV file."""
    email_addr = os.environ.get(GMAIL_EMAIL_ENV, "").strip()
    password = os.environ.get(GMAIL_APP_PASSWORD_ENV, "").strip()

    if not email_addr or not password:
        raise ValueError(
            f"Missing Gmail credentials. Expected environment variables "
            f"'{GMAIL_EMAIL_ENV}' and '{GMAIL_APP_PASSWORD_ENV}' to be set "
            f"(via GitHub Secrets in the workflow)."
        )

    return email_addr, password


# ============================================================================
# STEP 1: Fabric_Current.csv  ("Partial Possible Current" tab)
# ============================================================================

def build_fabric_current():
    df = load_sheet_as_df(SOURCE_SHEET_ID, CURRENT_TAB, SERVICE_ACCOUNT_JSON, SCOPES_READONLY)
    df.columns = [c.strip() for c in df.columns]

    df["SKU"] = df["SKU"].astype(str).str.strip()
    df["Fabric"] = df["Fabric"].astype(str).str.strip()

    df["_row_order"] = range(len(df))
    df = df.sort_values(["SKU", "_row_order"])

    grouped = (
        df.groupby("SKU", sort=False)
          .agg(
              Fabric=("Fabric", lambda s: " / ".join(s)),
              **{"Current Inventory": ("Current Inventory", lambda s: " / ".join(str(v) for v in s))},
              **{"Min unit": ("Min unit", lambda s: int(s.iloc[0]))},
          )
          .reset_index()
    )

    grouped.to_csv(CURRENT_FILE, index=False, encoding="utf-8-sig")
    return grouped


# ============================================================================
# STEP 2: Fabric_Open_PO.csv  ("Open PO" tab)
# ============================================================================

def build_fabric_open_po():
    df = load_sheet_as_df(SOURCE_SHEET_ID, OPEN_PO_TAB, SERVICE_ACCOUNT_JSON, SCOPES_READONLY)
    df.columns = [c.strip() for c in df.columns]

    df["SKU"] = df["SKU"].astype(str).str.strip()

    df["_row_order"] = range(len(df))
    df = df.sort_values(["SKU", "_row_order"])

    result = (
        df.groupby("SKU", sort=False)["Min unit"]
          .apply(lambda s: int(s.iloc[0]))
          .reset_index()
          .rename(columns={"Min unit": "Min Unit"})
    )

    result.to_csv(OPEN_PO_FILE, index=False, encoding="utf-8-sig")
    return result


# ============================================================================
# STEP 3: Fabric_WIP.csv  ("WIP" tab)
# ============================================================================

def build_fabric_wip():
    df = load_sheet_as_df(SOURCE_SHEET_ID, WIP_TAB, SERVICE_ACCOUNT_JSON, SCOPES_READONLY)
    df.columns = [c.strip() for c in df.columns]

    df["SKU"] = df["SKU"].astype(str).str.strip()

    df["_row_order"] = range(len(df))
    df = df.sort_values(["SKU", "_row_order"])

    result = (
        df.groupby("SKU", sort=False)["Min unit"]
          .apply(lambda s: int(s.iloc[0]))
          .reset_index()
          .rename(columns={"Min unit": "Min Unit"})
    )

    result.to_csv(WIP_FILE, index=False, encoding="utf-8-sig")
    return result


# ============================================================================
# STEP 4: Join Current + Open_PO + WIP -> Fabric_overall.csv
# ============================================================================

def join_fabric():
    df_current = pd.read_csv(CURRENT_FILE, encoding="utf-8-sig", dtype=str)
    df_open_po = pd.read_csv(OPEN_PO_FILE, encoding="utf-8-sig", dtype=str)
    df_wip = pd.read_csv(WIP_FILE, encoding="utf-8-sig", dtype=str)

    df_open_po = df_open_po.rename(columns={"Min Unit": "Open_PO"})
    df_result = df_current.merge(df_open_po[["SKU", "Open_PO"]], on="SKU", how="left")
    df_result["Open_PO"] = df_result["Open_PO"].fillna("-")

    df_wip = df_wip.rename(columns={"Min Unit": "WIP"})
    df_result = df_result.merge(df_wip[["SKU", "WIP"]], on="SKU", how="left")
    df_result["WIP"] = df_result["WIP"].fillna("-")

    df_result = df_result[["SKU", "Fabric", "Current Inventory", "Min unit", "Open_PO", "WIP"]]

    df_result.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
    return df_result


# ============================================================================
# STEP 5: Upload Fabric_overall.csv to the destination Google Sheet
# ============================================================================

def upload_to_sheet():
    df = pd.read_csv(OUTPUT_FILE)

    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_JSON, scopes=SCOPES_FULL)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(DEST_SPREADSHEET_ID)
    try:
        worksheet = spreadsheet.worksheet(DEST_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=DEST_TAB_NAME, rows=max(len(df) + 10, 100), cols=max(len(df.columns) + 5, 20)
        )

    worksheet.clear()
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    worksheet.update(values, value_input_option="USER_ENTERED")

    return df


# ============================================================================
# STEP 6: Email the run report
# ============================================================================

def send_email_report(tracker):
    body, overall_ok = tracker.to_email_body()
    subject = (
        "SKU Overall sheet updated with Fabric data"
        if overall_ok
        else f"Fabric Overall pipeline FAILED - {tracker.first_failure()['name']}"
    )

    sender_email, sender_password = load_gmail_credentials_from_env()
    to_addr = EMAIL_TO or sender_email
    if isinstance(to_addr, list):
        to_header = ", ".join(to_addr)
    else:
        to_header = to_addr

    # Explicit utf-8 charset so accented characters, tracebacks, etc. never
    # trigger an ASCII encode error, regardless of OS/locale.
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = to_header

    recipients = to_addr if isinstance(to_addr, list) else [to_header]

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_bytes())


# ============================================================================
# MAIN ORCHESTRATOR
# ============================================================================

def run_step(tracker, step_name, func, *args, row_source=None):
    """Run one step, record success/failure, return (ok, result)."""
    try:
        result = func(*args)
        rows = len(result) if row_source is None and hasattr(result, "__len__") else row_source
        tracker.success(step_name, rows=rows)
        return True, result
    except Exception as e:
        tracker.failed(step_name, f"{type(e).__name__}: {e}\n{traceback.format_exc(limit=3)}")
        return False, None


def main():
    tracker = StepTracker()

    ok1, _ = run_step(tracker, "Step 1: Build Fabric_Current.csv (Partial Possible Current tab)", build_fabric_current)
    ok2, _ = run_step(tracker, "Step 2: Build Fabric_Open_PO.csv (Open PO tab)", build_fabric_open_po)
    ok3, _ = run_step(tracker, "Step 3: Build Fabric_WIP.csv (WIP tab)", build_fabric_wip)

    if ok1 and ok2 and ok3:
        ok4, _ = run_step(tracker, "Step 4: Join Current + Open_PO + WIP into Fabric_overall.csv", join_fabric)
    else:
        ok4 = False
        tracker.skipped(
            "Step 4: Join Current + Open_PO + WIP into Fabric_overall.csv",
            "Skipped because one or more source files (Step 1/2/3) failed to build.",
        )

    if ok4:
        ok5, _ = run_step(tracker, "Step 5: Upload Fabric_overall.csv to 'Fabric Overall' sheet tab", upload_to_sheet)
    else:
        tracker.skipped(
            "Step 5: Upload Fabric_overall.csv to 'Fabric Overall' sheet tab",
            "Skipped because Step 4 (join) did not complete.",
        )

    # Step 6 always runs, so you get a report even on failure.
    try:
        send_email_report(tracker)
        print("[OK] Step 6: Report email sent.")
    except Exception as e:
        print(f"[FAILED] Step 6: Report email could not be sent - {e}")
        print(traceback.format_exc(limit=3))


if __name__ == "__main__":
    main()
