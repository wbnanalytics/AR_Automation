import os
import io
import secrets
import tempfile
import requests
import threading
import time
import base64
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template, request
from dotenv import load_dotenv

# ── Load .env ─────────────────────────────────────────────────────────────────
load_dotenv()

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG  (all values come from .env — never hardcode secrets)
# ══════════════════════════════════════════════════════════════════════════════

GDRIVE_EXCEL_FILE_ID = os.getenv("GDRIVE_EXCEL_FILE_ID", "")
USE_GDRIVE           = True
DATA_FILE            = "M_data.xlsx"

# ── SendGrid ──────────────────────────────────────────────────────────────────
SENDGRID_API_KEY  = os.getenv("SENDGRID_API_KEY", "")
SENDER_EMAIL      = os.getenv("SMTP_EMAIL", "")
SENDER_NAME       = os.getenv("SENDER_NAME", "Wellbeing Nutrition AR")

_cc_raw        = os.getenv("CC_EMAILS", "")
CC_EMAILS_LIST = [e.strip() for e in _cc_raw.split(",") if e.strip()]

# ── Zoho ──────────────────────────────────────────────────────────────────────
ZOHO_CLIENT_ID       = os.getenv("ZOHO_CLIENT_ID", "")
ZOHO_CLIENT_SECRET   = os.getenv("ZOHO_CLIENT_SECRET", "")
ZOHO_REFRESH_TOKEN   = os.getenv("ZOHO_REFRESH_TOKEN", "")
ZOHO_ORGANIZATION_ID = os.getenv("ZOHO_ORGANIZATION_ID", "")
ZOHO_REGION          = os.getenv("ZOHO_REGION", "in")

ZOHO_ACCOUNTS_URL = f"https://accounts.zoho.{ZOHO_REGION}/oauth/v2/token"
ZOHO_API_BASE     = f"https://www.zohoapis.{ZOHO_REGION}/books/v3"

SENDGRID_API_URL  = "https://api.sendgrid.com/v3/mail/send"

REQUIRED_COLUMNS = {
    "Age", "Balance_Due", "Customer_Name", "Due_Date",
    "Inv_date", "Invoice_no.", "Salesperson",
    "Salesperson_mail_id", "Unused_Credits",
}

DATE_COLUMNS = ["Inv_date", "Due_Date"]
SORT_COLUMNS = ["Salesperson", "Customer_Name", "Invoice_no."]

AGING_BUCKETS = {
    "1-30 Days (Current)" : (1,   30),
    "31-60 Days"          : (31,  60),
    "61-90 Days"          : (61,  90),
    "91-120 Days"         : (91,  120),
    "Above 120 Days"      : (121, 99999),
    "Above 60 Days"       : (61,  99999),
    "Above 90 Days"       : (91,  99999),
}

# ══════════════════════════════════════════════════════════════════════════════

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

drafts       = {}
pdf_status   = {}
send_results = {}


# ══════════════════════════════════════════════════════════════════════════════
#  ZOHO ACCESS TOKEN MANAGER
# ══════════════════════════════════════════════════════════════════════════════

class ZohoTokenManager:
    def __init__(self):
        self._token      = None
        self._expires_at = 0
        self._lock       = threading.Lock()

    def get_token(self):
        with self._lock:
            if not self._token or time.time() >= (self._expires_at - 60):
                self._refresh()
            return self._token

    def _refresh(self):
        if not all([ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET, ZOHO_REFRESH_TOKEN]):
            raise RuntimeError("Zoho credentials missing in .env")
        resp = requests.post(
            ZOHO_ACCOUNTS_URL,
            data={
                "refresh_token": ZOHO_REFRESH_TOKEN,
                "client_id"    : ZOHO_CLIENT_ID,
                "client_secret": ZOHO_CLIENT_SECRET,
                "grant_type"   : "refresh_token",
            },
            timeout=20,
        )
        resp.raise_for_status()
        data = resp.json()
        if "access_token" not in data:
            raise RuntimeError(f"Zoho token refresh failed: {data}")
        self._token      = data["access_token"]
        expires_in       = int(data.get("expires_in", 3600))
        self._expires_at = time.time() + expires_in
        print(f"[Zoho] Token refreshed. Expires in {expires_in}s.")


zoho_token_manager = ZohoTokenManager()


# ══════════════════════════════════════════════════════════════════════════════
#  ZOHO BOOKS — PDF DOWNLOADER
# ══════════════════════════════════════════════════════════════════════════════

def zoho_headers():
    return {
        "Authorization": f"Zoho-oauthtoken {zoho_token_manager.get_token()}",
        "Content-Type" : "application/json",
    }


def zoho_find_invoice_id(invoice_number):
    url    = f"{ZOHO_API_BASE}/invoices"
    params = {"organization_id": ZOHO_ORGANIZATION_ID, "invoice_number": invoice_number.strip()}
    try:
        resp     = requests.get(url, headers=zoho_headers(), params=params, timeout=30)
        resp.raise_for_status()
        invoices = resp.json().get("invoices", [])
        if not invoices:
            params2  = {"organization_id": ZOHO_ORGANIZATION_ID, "reference_number": invoice_number.strip()}
            resp     = requests.get(url, headers=zoho_headers(), params=params2, timeout=30)
            resp.raise_for_status()
            invoices = resp.json().get("invoices", [])
        if invoices:
            return invoices[0]["invoice_id"]
    except Exception as e:
        print(f"[Zoho] Could not find invoice {invoice_number}: {e}")
    return None


def zoho_download_pdf(invoice_id, save_dir, filename):
    url    = f"{ZOHO_API_BASE}/invoices/{invoice_id}"
    params = {"organization_id": ZOHO_ORGANIZATION_ID, "accept": "pdf"}
    try:
        resp = requests.get(
            url,
            headers={**zoho_headers(), "Accept": "application/pdf"},
            params=params, timeout=60,
        )
        resp.raise_for_status()
        content_type = resp.headers.get("Content-Type", "")
        if "pdf" not in content_type and len(resp.content) < 200:
            return None
        safe_name = "".join(c if c.isalnum() or c in "-_." else "_" for c in filename)
        if not safe_name.lower().endswith(".pdf"):
            safe_name += ".pdf"
        file_path = os.path.join(save_dir, safe_name)
        with open(file_path, "wb") as f:
            f.write(resp.content)
        print(f"[Zoho] Downloaded: {safe_name} ({len(resp.content):,} bytes)")
        return file_path
    except Exception as e:
        print(f"[Zoho] PDF download failed for {invoice_id}: {e}")
        return None


# ══════════════════════════════════════════════════════════════════════════════
#  GOOGLE DRIVE — EXCEL LOADER
# ══════════════════════════════════════════════════════════════════════════════

def load_excel_from_gdrive():
    url  = f"https://docs.google.com/spreadsheets/d/{GDRIVE_EXCEL_FILE_ID}/export?format=xlsx"
    resp = requests.get(url, timeout=30)
    if resp.status_code != 200:
        url  = f"https://drive.google.com/uc?export=download&id={GDRIVE_EXCEL_FILE_ID}"
        resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    return pd.read_excel(io.BytesIO(resp.content))


# ══════════════════════════════════════════════════════════════════════════════
#  CORE DATA HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def format_date(value):
    return value.strftime("%d-%m-%Y") if pd.notna(value) else ""


def format_currency(value):
    """Format currency with ₹ symbol and Indian number system."""
    if pd.isna(value):
        return "₹ 0.00"
    num              = float(value)
    s                = f"{num:.2f}"
    integer, decimal = s.split(".")
    if len(integer) > 3:
        last3   = integer[-3:]
        rest    = integer[:-3]
        rest    = ",".join([rest[max(i - 2, 0):i] for i in range(len(rest), 0, -2)][::-1])
        integer = rest + "," + last3
    return f"&#8377; {integer}.{decimal}"


def format_currency_plain(value):
    """Same as format_currency but uses ₹ character (for non-HTML contexts)."""
    if pd.isna(value):
        return "₹ 0.00"
    num              = float(value)
    s                = f"{num:.2f}"
    integer, decimal = s.split(".")
    if len(integer) > 3:
        last3   = integer[-3:]
        rest    = integer[:-3]
        rest    = ",".join([rest[max(i - 2, 0):i] for i in range(len(rest), 0, -2)][::-1])
        integer = rest + "," + last3
    return f"₹ {integer}.{decimal}"


def get_unused_credit(df):
    credits = df["Unused_Credits"].dropna()
    return float(credits.iloc[0]) if not credits.empty else 0


def validate_columns(df):
    missing = REQUIRED_COLUMNS.difference(df.columns)
    if missing:
        raise KeyError(f"Missing columns: {', '.join(sorted(missing))}")


def load_data():
    if USE_GDRIVE:
        df = load_excel_from_gdrive()
    else:
        df = pd.read_excel(Path(__file__).parent / DATA_FILE)
    df.columns = [c.strip().replace(" ", "_") for c in df.columns]
    validate_columns(df)
    for col in DATE_COLUMNS:
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
    df["Age"] = pd.to_numeric(df["Age"], errors="coerce")

    def clean_numeric(series):
        return (
            series.astype(str)
                  .str.replace("₹", "", regex=False)
                  .str.replace("Rs.", "", regex=False)
                  .str.replace(",", "", regex=False)
                  .str.strip()
                  .pipe(pd.to_numeric, errors="coerce")
        )

    df["Balance_Due"]    = clean_numeric(df["Balance_Due"])
    df["Unused_Credits"] = clean_numeric(df["Unused_Credits"])
    return df.sort_values(SORT_COLUMNS).reset_index(drop=True)


def apply_aging_filter(df, aging_label):
    if not aging_label or aging_label not in AGING_BUCKETS:
        return df
    low, high = AGING_BUCKETS[aging_label]
    age_col   = pd.to_numeric(df["Age"], errors="coerce").fillna(0)
    return df[(age_col >= low) & (age_col <= high)]


def safe_age(value):
    try:
        v = pd.to_numeric(value, errors="coerce")
        return "" if pd.isna(v) else int(v)
    except Exception:
        return ""


def df_to_records(df):
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "invoice_no"  : str(r["Invoice_no."]),
            "customer"    : str(r["Customer_Name"]),
            "salesperson" : str(r["Salesperson"]),
            "email"       : str(r["Salesperson_mail_id"]),
            "inv_date"    : format_date(r["Inv_date"]),
            "due_date"    : format_date(r["Due_Date"]),
            "age"         : safe_age(r["Age"]),
            "balance"     : float(r["Balance_Due"]) if pd.notna(r["Balance_Due"]) else 0.0,
            "balance_fmt" : format_currency_plain(r["Balance_Due"]),
            "unused"      : float(r["Unused_Credits"]) if pd.notna(r["Unused_Credits"]) else 0.0,
        })
    return rows


# ══════════════════════════════════════════════════════════════════════════════
#  EMAIL HTML BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def build_customer_html(customer_df):
    customer_name   = customer_df["Customer_Name"].iloc[0]
    total           = customer_df["Balance_Due"].fillna(0).sum()
    unused_credit   = get_unused_credit(customer_df)
    net_outstanding = total - unused_credit

    html = f"""
<p style='font-family:Calibri,Arial,sans-serif;font-size:11pt;font-weight:bold;
          color:#1a3c6e;margin:16px 0 4px 0;'>{customer_name}</p>
<table border='0' cellpadding='0' cellspacing='0' width='100%'
       style='border-collapse:collapse;font-family:Calibri,Arial,sans-serif;
              font-size:9.5pt;margin-bottom:12px;max-width:620px;'>
    <thead>
        <tr style='background-color:#1a3c6e;color:#ffffff;'>
            <th style='padding:6px 10px;border:1px solid #15326a;text-align:left;font-size:9pt;width:28%;'>Invoice No</th>
            <th style='padding:6px 10px;border:1px solid #15326a;text-align:center;font-size:9pt;width:16%;'>Invoice Date</th>
            <th style='padding:6px 10px;border:1px solid #15326a;text-align:center;font-size:9pt;width:16%;'>Due Date</th>
            <th style='padding:6px 10px;border:1px solid #15326a;text-align:center;font-size:9pt;width:14%;'>Overdue Days</th>
            <th style='padding:6px 10px;border:1px solid #15326a;text-align:right;font-size:9pt;width:26%;'>Balance (INR)</th>
        </tr>
    </thead>
    <tbody>"""

    for i, (_, row) in enumerate(customer_df.iterrows()):
        age_val   = safe_age(row["Age"])
        age_color = "#a32d2d" if isinstance(age_val, int) and age_val > 30 else "#1a3c6e"
        row_bg    = "#f7f9fc" if i % 2 == 0 else "#ffffff"
        html += f"""
        <tr style='background-color:{row_bg};'>
            <td style='padding:5px 10px;border:1px solid #d9e1ec;color:#1a3c6e;font-weight:bold;'>{row['Invoice_no.']}</td>
            <td style='padding:5px 10px;border:1px solid #d9e1ec;text-align:center;'>{format_date(row['Inv_date'])}</td>
            <td style='padding:5px 10px;border:1px solid #d9e1ec;text-align:center;'>{format_date(row['Due_Date'])}</td>
            <td style='padding:5px 10px;border:1px solid #d9e1ec;text-align:center;font-weight:bold;color:{age_color};'>{age_val}</td>
            <td style='padding:5px 10px;border:1px solid #d9e1ec;text-align:right;'>{format_currency(row['Balance_Due'])}</td>
        </tr>"""

    html += f"""
    </tbody>
    <tfoot>
        <tr style='background-color:#eef2f8;'>
            <td colspan='4' style='padding:6px 10px;border:1px solid #d9e1ec;font-weight:bold;'>Total Outstanding</td>
            <td style='padding:6px 10px;border:1px solid #d9e1ec;text-align:right;font-weight:bold;'>{format_currency(total)}</td>
        </tr>"""

    if unused_credit > 0:
        html += f"""
        <tr style='background-color:#e8f5f0;'>
            <td colspan='4' style='padding:6px 10px;border:1px solid #d9e1ec;font-weight:bold;color:#0f6e56;'>Unused Credit Available</td>
            <td style='padding:6px 10px;border:1px solid #d9e1ec;text-align:right;font-weight:bold;color:#0f6e56;'>({format_currency(unused_credit)})</td>
        </tr>
        <tr style='background-color:#1a3c6e;'>
            <td colspan='4' style='padding:6px 10px;border:1px solid #15326a;font-weight:bold;color:#fff;'>Net Outstanding</td>
            <td style='padding:6px 10px;border:1px solid #15326a;text-align:right;font-weight:bold;color:#fff;'>{format_currency(net_outstanding)}</td>
        </tr>"""

    html += "</tfoot></table>"
    return html


def build_email_html(salesperson, salesperson_df):
    first_name        = str(salesperson).split()[0]
    total_outstanding = salesperson_df["Balance_Due"].fillna(0).sum()
    total_unused      = get_unused_credit(salesperson_df)
    net_outstanding   = total_outstanding - total_unused
    # ── NEW: count unique customers for this salesperson ──
    total_customers   = salesperson_df["Customer_Name"].nunique()

    html = f"""<div style='font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222;max-width:680px;'>
<p style='margin:0 0 12px 0;'>Hi {first_name},</p>
<p style='margin:0 0 20px 0;'>Please find the outstanding invoice details for Retail Customers below.
The respective invoices are attached for your reference.</p>

<table cellpadding='0' cellspacing='0' border='0' width='100%'
       style='border-collapse:collapse;margin-bottom:20px;max-width:620px;'>
    <tr style='background-color:#1a3c6e;'>
        <td style='padding:10px 14px;'>
            <span style='font-family:Calibri,Arial,sans-serif;font-size:12pt;font-weight:bold;color:#fff;'>
                Accounts Receivable Summary
            </span>
        </td>
    </tr>
    <tr><td style='padding:0;'>
        <table width='100%' cellpadding='0' cellspacing='0'
               style='border-collapse:collapse;'>
            <tr>
                <td style='padding:10px 14px;background-color:#eef2f8;border:1px solid #d9e1ec;width:25%;'>
                    <div style='font-size:8pt;color:#555;text-transform:uppercase;letter-spacing:0.5px;'>Total Outstanding</div>
                    <div style='font-size:11pt;font-weight:bold;color:#a32d2d;margin-top:4px;'>{format_currency(total_outstanding)}</div>
                </td>
                <td style='padding:10px 14px;background-color:#e8f5f0;border:1px solid #d9e1ec;width:25%;'>
                    <div style='font-size:8pt;color:#555;text-transform:uppercase;letter-spacing:0.5px;'>Unused Credit</div>
                    <div style='font-size:11pt;font-weight:bold;color:#0f6e56;margin-top:4px;'>{format_currency(total_unused)}</div>
                </td>
                <td style='padding:10px 14px;background-color:#f0f4ff;border:1px solid #d9e1ec;width:25%;'>
                    <div style='font-size:8pt;color:#555;text-transform:uppercase;letter-spacing:0.5px;'>Net Outstanding</div>
                    <div style='font-size:11pt;font-weight:bold;color:#1a3c6e;margin-top:4px;'>{format_currency(net_outstanding)}</div>
                </td>
                <td style='padding:10px 14px;background-color:#fff8f0;border:1px solid #d9e1ec;width:25%;'>
                    <div style='font-size:8pt;color:#555;text-transform:uppercase;letter-spacing:0.5px;'>Total Customers</div>
                    <div style='font-size:11pt;font-weight:bold;color:#b45309;margin-top:4px;'>{total_customers}</div>
                </td>
            </tr>
        </table>
    </td></tr>
</table>
"""
    for _, customer_df in salesperson_df.groupby("Customer_Name", dropna=False):
        html += build_customer_html(customer_df)

    html += """
<p style='font-family:Calibri,Arial,sans-serif;font-size:11pt;margin-top:20px;'>
    Kindly review the details and follow up with the respective customers to expedite the payment.
</p>
</div>"""
    return html


# ══════════════════════════════════════════════════════════════════════════════
#  SENDGRID SENDER
# ══════════════════════════════════════════════════════════════════════════════

def send_via_sendgrid(to_email, cc_list, subject, body_html, attachment_paths):
    if not SENDGRID_API_KEY:
        raise RuntimeError("SENDGRID_API_KEY is missing in .env")
    if not SENDER_EMAIL:
        raise RuntimeError("SMTP_EMAIL (sender address) is missing in .env")

    seen     = {to_email.lower()}
    clean_cc = []
    for e in (cc_list or []):
        e = e.strip()
        if e and e.lower() not in seen:
            seen.add(e.lower())
            clean_cc.append(e)

    personalizations = [{"to": [{"email": to_email}]}]
    if clean_cc:
        personalizations[0]["cc"] = [{"email": e} for e in clean_cc]

    attachments    = []
    attached_count = 0
    for path in (attachment_paths or []):
        if path and os.path.exists(path):
            with open(path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")
            attachments.append({
                "content"    : encoded,
                "type"       : "application/pdf",
                "filename"   : os.path.basename(path),
                "disposition": "attachment",
            })
            attached_count += 1

    payload = {
        "personalizations": personalizations,
        "from"            : {"email": SENDER_EMAIL, "name": SENDER_NAME},
        "subject"         : subject,
        "content"         : [{"type": "text/html", "value": body_html}],
    }
    if attachments:
        payload["attachments"] = attachments

    headers = {
        "Authorization": f"Bearer {SENDGRID_API_KEY}",
        "Content-Type" : "application/json",
    }

    print(f"[SendGrid] To={to_email} | CC={clean_cc} | PDFs={attached_count}")
    resp = requests.post(SENDGRID_API_URL, json=payload, headers=headers, timeout=60)

    if resp.status_code not in (200, 202):
        raise RuntimeError(f"SendGrid returned {resp.status_code}: {resp.text[:400]}")
    print(f"[SendGrid] Accepted — status {resp.status_code}")


# ══════════════════════════════════════════════════════════════════════════════
#  FLASK ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/options")
def options():
    try:
        df = load_data()
        return jsonify({
            "customers"     : sorted(df["Customer_Name"].dropna().unique().tolist()),
            "salespersons"  : sorted(df["Salesperson"].dropna().unique().tolist()),
            "invoices"      : sorted(df["Invoice_no."].dropna().unique().tolist()),
            "aging_buckets" : list(AGING_BUCKETS.keys()),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/search")
def search():
    try:
        df          = load_data()
        customer    = request.args.get("customer", "").strip()
        salesperson = request.args.get("salesperson", "").strip()
        invoice     = request.args.get("invoice", "").strip()
        aging       = request.args.get("aging", "").strip()

        mask = pd.Series([True] * len(df))
        if customer:    mask &= df["Customer_Name"].str.upper() == customer.upper()
        if salesperson: mask &= df["Salesperson"].str.upper() == salesperson.upper()
        if invoice:     mask &= df["Invoice_no."].str.upper() == invoice.upper()

        filtered = df[mask]
        if aging:
            filtered = apply_aging_filter(filtered, aging)

        if filtered.empty:
            return jsonify({"records": [], "summary": {}})

        total  = filtered["Balance_Due"].fillna(0).sum()
        unused = get_unused_credit(filtered)
        return jsonify({
            "records": df_to_records(filtered),
            "summary": {
                "total"  : format_currency_plain(total),
                "unused" : format_currency_plain(unused),
                "net"    : format_currency_plain(total - unused),
                "count"  : len(filtered),
            },
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/draft_email", methods=["POST"])
def draft_email():
    try:
        body        = request.get_json()
        salesperson = body.get("salesperson", "").strip()
        customer    = body.get("customer", "").strip()
        invoice     = body.get("invoice", "").strip()
        aging       = body.get("aging", "").strip()

        df   = load_data()
        mask = pd.Series([True] * len(df))
        if salesperson: mask &= df["Salesperson"].str.upper() == salesperson.upper()
        if customer:    mask &= df["Customer_Name"].str.upper() == customer.upper()
        if invoice:     mask &= df["Invoice_no."].str.upper() == invoice.upper()

        filtered = df[mask]
        if aging:
            filtered = apply_aging_filter(filtered, aging)

        if filtered.empty:
            return jsonify({"error": "No records found for the given filter."}), 404

        all_invoice_nos = filtered["Invoice_no."].dropna().unique().tolist()
        tmp_dir         = tempfile.mkdtemp()
        draft_id        = secrets.token_hex(8)
        draft_entries   = []

        for sp, sp_df in filtered.groupby("Salesperson", dropna=False):
            to_email  = str(sp_df["Salesperson_mail_id"].iloc[0]).strip()
            body_html = build_email_html(sp, sp_df)
            inv_list  = sp_df["Invoice_no."].dropna().unique().tolist()
            draft_entries.append({
                "salesperson" : str(sp),
                "to_email"    : to_email,
                "cc_list"     : list(CC_EMAILS_LIST),
                "subject"     : "Outstanding Invoices for Retail Customers",
                "body_html"   : body_html,
                "attachments" : [],
                "pdf_map"     : {},
                "pdf_count"   : 0,
                "missing_pdfs": [],
                "invoice_list": inv_list,
            })

        drafts[draft_id]     = draft_entries
        zoho_configured      = all([ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET,
                                    ZOHO_REFRESH_TOKEN, ZOHO_ORGANIZATION_ID])
        pdf_status[draft_id] = {
            "status" : "fetching" if zoho_configured else "skipped",
            "found"  : 0,
            "done"   : 0,
            "total"  : len(all_invoice_nos),
            "draft_id": draft_id,
        }

        def bg_fetch(draft_id, all_invoice_nos, tmp_dir):
            try:
                if not zoho_configured:
                    return
                total   = len(all_invoice_nos)
                pdf_map = {}
                found   = 0
                done    = 0
                print(f"[BG] Fetching {total} PDFs from Zoho...")

                for inv_no in all_invoice_nos:
                    inv_no_str = str(inv_no).strip()
                    if not inv_no_str:
                        continue
                    print(f"[BG] Fetching: {inv_no_str}")
                    try:
                        invoice_id = zoho_find_invoice_id(inv_no_str)
                        if invoice_id:
                            fp = zoho_download_pdf(invoice_id, tmp_dir, f"{inv_no_str}.pdf")
                            pdf_map[inv_no_str] = fp
                            if fp:
                                found += 1
                        else:
                            pdf_map[inv_no_str] = None
                    except Exception as ex:
                        print(f"[BG] Error on {inv_no_str}: {ex}")
                        pdf_map[inv_no_str] = None
                    finally:
                        done += 1
                    if draft_id in pdf_status:
                        pdf_status[draft_id]["done"]  = done
                        pdf_status[draft_id]["found"] = found

                if draft_id in drafts:
                    for e in drafts[draft_id]:
                        atts, missing = [], []
                        for inv_no in e["invoice_list"]:
                            fp = pdf_map.get(str(inv_no).strip())
                            if fp: atts.append(fp)
                            else:  missing.append(inv_no)
                        seen = set()
                        e["attachments"]  = [p for p in atts if not (p in seen or seen.add(p))]
                        e["missing_pdfs"] = missing
                        e["pdf_count"]    = len(e["attachments"])
                        e["pdf_map"]      = pdf_map

                pdf_status[draft_id] = {
                    "status": "done", "found": found,
                    "done": done, "total": total
                }
                print(f"[BG] Done — {found}/{total} PDFs fetched.")
            except Exception as ex:
                print(f"[BG] Error: {ex}")
                if draft_id in pdf_status:
                    pdf_status[draft_id] = {"status": "error", "message": str(ex)}

        threading.Thread(target=bg_fetch, args=(draft_id, all_invoice_nos, tmp_dir), daemon=True).start()

        previews = [{
            "salesperson" : e["salesperson"],
            "to_email"    : e["to_email"],
            "cc_list"     : e["cc_list"],
            "subject"     : e["subject"],
            "body_html"   : e["body_html"],
            "pdf_count"   : 0,
            "missing_note": "",
            "invoice_list": e["invoice_list"],
            "pdf_loading" : zoho_configured,
        } for e in draft_entries]

        return jsonify({"draft_id": draft_id, "previews": previews})

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/pdf_status/<draft_id>")
def pdf_status_route(draft_id):
    status = dict(pdf_status.get(draft_id, {"status": "unknown"}))

    if draft_id in drafts:
        tabs = []
        for e in drafts[draft_id]:
            missing_note = ""
            if e.get("missing_pdfs"):
                missing_note = (
                    f"{len(e['missing_pdfs'])} PDF(s) not found in Zoho: "
                    f"{', '.join(str(x) for x in e['missing_pdfs'][:5])}"
                )
            tabs.append({
                "salesperson" : e["salesperson"],
                "pdf_count"   : e.get("pdf_count", 0),
                "missing_note": missing_note,
                "missing_pdfs": [str(x) for x in e.get("missing_pdfs", [])],
                "invoice_list": e.get("invoice_list", []),
            })
        status["tabs"] = tabs

    return jsonify(status)


@app.route("/api/update_draft", methods=["POST"])
def update_draft():
    try:
        body      = request.get_json()
        draft_id  = body.get("draft_id", "").strip()
        tab_index = int(body.get("tab_index", 0))

        if draft_id not in drafts:
            return jsonify({"error": "Draft not found."}), 404

        entry = drafts[draft_id][tab_index]
        if "to_email"  in body: entry["to_email"]  = body["to_email"].strip()
        if "cc_list"   in body: entry["cc_list"]   = [e.strip() for e in body["cc_list"] if e.strip()]
        if "subject"   in body: entry["subject"]   = body["subject"].strip()
        if "body_html" in body: entry["body_html"] = body["body_html"]

        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/confirm_send", methods=["POST"])
def confirm_send():
    import traceback
    try:
        body     = request.get_json(force=True, silent=True) or {}
        draft_id = str(body.get("draft_id", "")).strip()
        print(f"[Send] confirm_send called. draft_id={draft_id!r}")

        if not draft_id:
            return jsonify({"error": "No draft_id provided."}), 400
        if draft_id not in drafts:
            return jsonify({"error": "Draft not found or already sent. Please preview again."}), 404

        waited = 0
        while pdf_status.get(draft_id, {}).get("status") == "fetching" and waited < 15:
            time.sleep(1); waited += 1

        draft_entries = drafts.pop(draft_id)
        pdf_status.pop(draft_id, None)

        send_id = secrets.token_hex(6)
        send_results[send_id] = {"status": "sending"}

        def do_send(send_id, draft_entries):
            try:
                sent_to = []
                for e in draft_entries:
                    valid_attachments = [
                        p for p in e.get("attachments", [])
                        if p and os.path.exists(p)
                    ]
                    print(f"[Send] Sending to {e['to_email']} with {len(valid_attachments)} PDF(s) via SendGrid")
                    send_via_sendgrid(
                        to_email         = e["to_email"],
                        cc_list          = e["cc_list"],
                        subject          = e["subject"],
                        body_html        = e["body_html"],
                        attachment_paths = valid_attachments,
                    )
                    sent_to.append(f"{e['salesperson']} <{e['to_email']}>")
                msg = f"Email sent to: {', '.join(sent_to)}"
                print(f"[Send] Success: {msg}")
                send_results[send_id] = {"status": "done", "message": msg}
            except Exception as ex:
                traceback.print_exc()
                send_results[send_id] = {"status": "error", "message": str(ex)}

        threading.Thread(target=do_send, args=(send_id, draft_entries), daemon=True).start()

        return jsonify({"success": True, "send_id": send_id, "message": "Sending in progress…"})

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/send_status/<send_id>")
def send_status(send_id):
    result = send_results.get(send_id, {"status": "unknown"})
    if result.get("status") in ("done", "error"):
        send_results.pop(send_id, None)
    return jsonify(result)


@app.route("/zoho/callback")
def zoho_callback():
    code = request.args.get("code", "")
    if not code:
        return "No code received.", 400
    return f"""<html><body style="font-family:monospace;padding:40px;max-width:900px;">
    <h2>Auth Code Received</h2>
    <pre style="background:#f0f4fa;padding:16px;border-radius:8px;">{code}</pre>
    <p>Run this curl command:</p>
    <pre style="background:#f0f4fa;padding:16px;border-radius:8px;white-space:pre-wrap;">curl -X POST "https://accounts.zoho.{ZOHO_REGION}/oauth/v2/token" ^
  -d "code={code}" ^
  -d "client_id=YOUR_CLIENT_ID" ^
  -d "client_secret=YOUR_CLIENT_SECRET" ^
  -d "redirect_uri=https://ar-mails-automate.onrender.com/zoho/callback" ^
  -d "grant_type=authorization_code"</pre>
    </body></html>"""


if __name__ == "__main__":
    app.run(debug=True, port=5000)