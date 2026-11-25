# -*- coding: utf-8 -*-
import os, re, io, json, hashlib, secrets
from pathlib import Path
from datetime import datetime, date
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from pypdf import PdfReader, PdfWriter

st.set_page_config(page_title="Paragon Reconciliation Office | 2026", page_icon="📊", layout="wide")

APP_TITLE = "PARAGON RECONCILIATION OFFICE | 2026"
ADDRESS = "Paragon House, 5, Mohakhali C/A, Dhaka 1212, Bangladesh"
TOLERANCE = 0.005
BASE_DIR = Path("paragon_data")
PDF_DIR = BASE_DIR / "PDF_Software"
INPUT_PDF_DIR = PDF_DIR / "Input_PDFs"
SALE_DIR = BASE_DIR / "Sale_Folder"
OUT_DIR = BASE_DIR / "Reconciliation"
USERS_FILE = BASE_DIR / "users.json"

for d in [PDF_DIR, INPUT_PDF_DIR, SALE_DIR, OUT_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ---------------- AUTHENTICATION & ADMIN CONTROL ----------------
def _password_hash(password, salt):
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 240000).hex()

def load_users():
    if USERS_FILE.exists():
        try:
            return json.loads(USERS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    salt = secrets.token_hex(16)
    users = {"Subroto Das": {"salt": salt, "hash": _password_hash("12345", salt), "role": "Host"}}
    USERS_FILE.write_text(json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8")
    return users

def save_users(users):
    USERS_FILE.write_text(json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8")

# ---------------- CORE HELPER FUNCTIONS ----------------
def norm_id(v):
    if pd.isna(v): return None
    s = str(v).strip()
    if not s or s.casefold() in {"nan", "none", "n/a", "na", "null"}: return None
    try:
        n = float(s.replace(",", ""))
        return str(int(n)) if n.is_integer() else str(n)
    except Exception:
        return s

def parse_date(v):
    if pd.isna(v) or str(v).strip() == "": return pd.NaT
    if isinstance(v, (pd.Timestamp, datetime, date)): return pd.Timestamp(v).normalize()
    s = str(v).strip()
    try:
        if len(s) >= 10 and s[:4].isdigit() and s[4] == "-":
            return pd.to_datetime(s, errors="coerce", yearfirst=True).normalize()
        return pd.to_datetime(s, errors="coerce", dayfirst=True).normalize()
    except Exception:
        return pd.NaT

def norm_dates(s):
    return pd.to_datetime(s.map(parse_date), errors="coerce")

def find_col(df, candidates, label):
    lookup = {str(c).strip().casefold(): c for c in df.columns}
    for x in candidates:
        if x.casefold() in lookup: return lookup[x.casefold()]
    raise ValueError(f"{label} কলাম পাওয়া যায়নি। পাওয়া কলাম: {list(df.columns)}")

def extract_outlet_code(name):
    if pd.isna(name): return None
    value = str(name).strip()
    if not value: return None
    matches = re.findall(r"\(\s*(\d+(?:\.0)?)\s*\)\s*$", value)
    if matches: return norm_id(matches[-1])
    match = re.match(r"^\s*(\d+(?:\.0)?)\s+(?=\D)", value)
    if match: return norm_id(match.group(1))
    return None

def read_mapping(path, sheet):
    xl = pd.ExcelFile(path)
    names = {n.strip().casefold(): n for n in xl.sheet_names}
    if sheet.casefold() not in names:
        raise ValueError(f"Mapping Excel.xlsx-এ '{sheet}' Sheet নেই।")
    df = pd.read_excel(path, sheet_name=names[sheet.casefold()])
    c = find_col(df, ["Outlet", "Outlet Code", "OutletCode"], "Mapping Outlet Code")
    n = find_col(df, ["Outlet Name", "OutletName"], "Mapping Outlet Name")
    s = find_col(df, ["Store ID", "StoreID"], "Mapping Store ID")
    out = df[[c, n, s]].copy()
    out.columns = ["Outlet Code", "Outlet Name", "Store ID"]
    out["Outlet Code"] = out["Outlet Code"].map(norm_id)
    out["Store ID"] = out["Store ID"].map(norm_id)
    out["Outlet Name"] = out["Outlet Name"].fillna("").astype(str).str.strip()
    return out.dropna(subset=["Outlet Code", "Store ID"])

def _exception(source, row, dt, store, outlet, amount, issue, details):
    return {"Source": source, "Excel Row": row, "Date": dt, "Store ID": store, "Outlet Code": outlet, "Amount": amount, "Issue": issue, "Details": details}

def find_incomplete_mapping_rows(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet)
    c = find_col(df, ["Outlet", "Outlet Code", "OutletCode"], "Mapping Outlet Code")
    n = find_col(df, ["Outlet Name", "OutletName"], "Mapping Outlet Name")
    s = find_col(df, ["Store ID", "StoreID"], "Mapping Store ID")
    issues = []
    for idx, row in df.iterrows():
        code, store = norm_id(row[c]), norm_id(row[s])
        name = "" if pd.isna(row[n]) else str(row[n]).strip()
        if (code is not None or store is not None or name) and (code is None or store is None):
            missing = "Store ID" if store is None else "Outlet Code"
            issues.append(_exception(f"Mapping Excel.xlsx | {sheet}", idx + 2, None, store, code, None, "Incomplete Mapping Row", f"{sheet} Sheet-এ {missing} নেই"))
    return issues

def detect_and_import_pos(uploaded_file):
    preview = pd.read_excel(uploaded_file, header=None, nrows=30, dtype=str).fillna("")
    text = " ".join(str(v).strip().casefold() for v in preview.to_numpy().ravel() if str(v).strip())
    channel = None
    if "paragon momo" in text or ("pubali pos" not in text and "pubali" in text):
        channel = "Momo"
    elif "paragon mart" in text or "pubali pos" in text:
        channel = "Mart"
    if not channel:
        raise ValueError("Excel ফাইলটি Mart নাকি Momo শনাক্ত করা যায়নি।")

    header_row = None
    for i in range(len(preview)):
        vals = {str(v).strip().casefold() for v in preview.iloc[i].dropna().tolist()}
        if "outlet name" in vals and ("outlet code" in vals or "total" in vals):
            header_row = i
            break
    if header_row is None:
        raise ValueError("POS report header পাওয়া যায়নি।")

    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=header_row)
    cmap = {str(c).strip().casefold(): c for c in df.columns}
    name_col = cmap["outlet name"]
    code_col = cmap["outlet code"]
    date_col = cmap["date"]
    amt_col = cmap.get("pubali pos") if channel == "Mart" else cmap.get("pubali", cmap.get("pubali pos"))

    fresh = pd.DataFrame({
        "Outlet Name": df[name_col].fillna("").astype(str).str.strip(),
        "Outlet Code": df[code_col],
        "Amount": pd.to_numeric(df[amt_col], errors="coerce"),
        "Date": pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.normalize()
    })
    fresh["Outlet Code"] = fresh.apply(lambda r: str(r["Outlet Code"]).replace(".0", "").strip() if pd.notna(r["Outlet Code"]) and str(r["Outlet Code"]).strip() else extract_outlet_code(r["Outlet Name"]), axis=1)
    fresh = fresh.dropna(subset=["Date", "Amount"])
    fresh = fresh[fresh["Outlet Name"].str.len() > 0]

    if channel == "Mart":
        fresh = fresh.rename(columns={"Amount": "Pubali POS"})[["Outlet Name", "Outlet Code", "Pubali POS", "Date"]]
        sale_file = SALE_DIR / "Sale Mart.xlsx"
        dcol = "Date"
    else:
        fresh = fresh.rename(columns={"Amount": "Total Amount", "Date": "Txn. Date"})[["Outlet Name", "Outlet Code", "Total Amount", "Txn. Date"]]
        sale_file = SALE_DIR / "Sale Momo.xlsx"
        dcol = "Txn. Date"

    if sale_file.exists():
        old = pd.read_excel(sale_file)
        old_dcol = "Date" if "Date" in old.columns else "Txn. Date"
        dates = set(pd.to_datetime(fresh[dcol], errors="coerce").dropna().dt.normalize())
        old_dates = pd.to_datetime(old[old_dcol], errors="coerce", dayfirst=True).dt.normalize()
        kept = old.loc[~old_dates.isin(dates)].copy()
        for c in old.columns:
            if c not in fresh.columns: fresh[c] = None
        fresh = fresh[old.columns]
        combined = pd.concat([kept, fresh], ignore_index=True)
    else:
        combined = fresh

    with pd.ExcelWriter(sale_file, engine="openpyxl", datetime_format="DD-MM-YYYY") as w:
        combined.to_excel(w, index=False)
    return channel, len(fresh)

def build_bank_from_pdfs():
    pdf_files = sorted(INPUT_PDF_DIR.glob("*.pdf"))
    if not pdf_files:
        raise FileNotFoundError("কোনো PDF আপলোড করা হয়নি।")
    writer = PdfWriter()
    records = []
    for pdf_path in pdf_files:
        reader = PdfReader(str(pdf_path))
        writer.append(str(pdf_path))
        full_text = "\n".join((page.extract_text() or "") for page in reader.pages)
        date_match = re.search(r"Settlement Date\s*:\s*(\d{2}-[A-Z]{3}-\d{4})", full_text, re.I)
        raw_date = date_match.group(1).upper() if date_match else ""
        if not raw_date:
            name_match = re.search(r"(\d{2}-[A-Z]{3}-\d{4})", pdf_path.name.upper())
            raw_date = name_match.group(1) if name_match else ""
        try:
            txn_date = datetime.strptime(raw_date, "%d-%b-%Y").strftime("%Y-%m-%d")
        except Exception:
            txn_date = raw_date
        store_ids = re.findall(r"Store ID:\s*(\d+)", full_text)
        totals = re.findall(r"Total:\s*([\d,]+\.\d{2})", full_text)
        for i, amount in enumerate(totals):
            store_id = store_ids[i].zfill(3) if i < len(store_ids) else None
            records.append({"Txn. Date": txn_date, "Store ID": store_id, "Total Amount": float(amount.replace(",", ""))})
    combined_out = PDF_DIR / "Combined_Master.pdf"
    with open(combined_out, "wb") as f:
        writer.write(f)
    writer.close()

    df = pd.DataFrame(records, columns=["Txn. Date", "Store ID", "Total Amount"])
    summary_out = PDF_DIR / "Final_Summary.xlsx"
    with pd.ExcelWriter(summary_out, engine="openpyxl", datetime_format="DD-MM-YYYY") as xw:
        df.to_excel(xw, index=False, sheet_name="Summary")
    return len(pdf_files), len(df)

