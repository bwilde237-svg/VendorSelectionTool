import os
import streamlit as st
import pandas as pd
import io

# -------------------------------
# CONFIGURATION
# -------------------------------
VENDOR_FILE = "Connected Booking System Evaluation Template (NDA).xlsx - Scoring Upload.csv"

REQUIREMENT_WEIGHTS = {
    "critical": 4,
    "important": 3,
    "useful": 2,
    "nice to have": 1,
    "not required": 0
}

RESPONSE_SCORES = {
    "yes": 1,
    "not provided": 0.5,
    "no": 0
}

MEETS_THRESHOLD = 0.75  # require 75% or higher to "Meet"
EXPECTED_CRITERIA_COLS = {"function", "requirement", "business area"}

# -------------------------------
# HELPER FUNCTIONS
# -------------------------------
def normalize_case(s):
    return str(s).strip().lower()

def _ensure_columns(df, expected_cols):
    normalized = {normalize_case(c): c for c in df.columns}
    missing = [c for c in expected_cols if c not in normalized]
    return (len(missing) == 0, missing, normalized)

def _detect_header_and_load_csv(file_like):
    """
    Try to robustly load a CSV that may have leading rows before the real header.
    Strategy:
      - First try pd.read_csv(header=0)
      - If column names are all 'Unnamed: *' or missing expected columns, read with header=None
        and search the first N rows for a row that contains the expected column names (case-insensitive).
      - If found, re-read CSV using that row as header (skiprows=header_row_index).
    """
    # try straightforward csv read first
    try:
        df = pd.read_csv(file_like)
        ok, missing, _ = _ensure_columns(df, EXPECTED_CRITERIA_COLS)
        if ok:
            return df
    except Exception:
        # fall through to more robust attempt
        file_like.seek(0)

    # read without header to inspect top rows
    try:
        raw = pd.read_csv(file_like, header=None, dtype=str, keep_default_na=False)
    except Exception as e:
        file_like.seek(0)
        # try more tolerant engine or raise
        raise ValueError(f"Unable to parse CSV: {e}")

    # Search top rows for a header row that contains all expected column names (case-insensitive)
    max_rows_to_scan = min(20, len(raw))
    header_row_index = None
    for i in range(max_rows_to_scan):
        row_vals = [normalize_case(str(x)) for x in raw.iloc[i].tolist()]
        matches = sum(1 for c in EXPECTED_CRITERIA_COLS if any(c in v for v in row_vals))
        if matches == len(EXPECTED_CRITERIA_COLS):
            header_row_index = i
            break

    if header_row_index is None:
        # As a fallback, see if any single row contains at least one expected column name,
        # and warn user of which columns were found.
        found = {}
        for i in range(max_rows_to_scan):
            row_vals = [normalize_case(str(x)) for x in raw.iloc[i].tolist()]
            for c in EXPECTED_CRITERIA_COLS:
                if any(c in v for v in row_vals):
                    found.setdefault(i, []).append(c)
        if found:
            raise ValueError(
                "Could not automatically identify a single row containing all required header names "
                f"({', '.join(EXPECTED_CRITERIA_COLS)}). Partial header matches found in rows: {found}. "
                "Please ensure your CSV has a header row containing Function, Requirement, Business Area."
            )
        else:
            raise ValueError(
                "No header row containing Function/Requirement/Business Area was found in the first "
                f"{max_rows_to_scan} rows. Please ensure the uploaded file is the correct Criteria CSV "
                "and that the header row appears within the top 20 rows."
            )

    # re-read using located header row
    file_like.seek(0)
    df = pd.read_csv(file_like, header=header_row_index)
    return df

def load_criteria_file(uploaded_file):
    # uploaded_file is an UploadedFile (has .name and is file-like)
    name = getattr(uploaded_file, "name", "")
    uploaded_file.seek(0)
    if name.lower().endswith((".xls", ".xlsx")):
        # excel: try to read and expect the header row to be correct; if not, user must correct sheet
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            raise ValueError(f"Failed to read Excel criteria file: {e}")
    else:
        # CSV: attempt robust header detection
        df = _detect_header_and_load_csv(uploaded_file)
    return df

def calculate_scores(vendor_df, criteria_df):
    vendor_df = vendor_df.copy()
    criteria_df = criteria_df.copy()
    vendor_df.columns = [normalize_case(c) for c in vendor_df.columns]
    criteria_df.columns = [normalize_case(c) for c in criteria_df.columns]

    ok, missing, _ = _ensure_columns(criteria_df, EXPECTED_CRITERIA_COLS)
    if not ok:
        raise ValueError(f"Criteria file is missing required columns (case-insensitive): {missing}. Found columns: {list(criteria_df.columns)}")

    if "vendor" not in vendor_df.columns:
        raise ValueError(f"Vendor file must include a 'Vendor' column (case-insensitive). Found columns: {list(vendor_df.columns)}")

    func_to_req = dict(zip([normalize_case(x) for x in criteria_df["function"]], [normalize_case(x) for x in criteria_df["requirement"]]))
    func_to_area = dict(zip([normalize_case(x) for x in criteria_df["function"]], [normalize_case(x) for x in criteria_df["business area"]]))

    vendor_scores, detailed_records = [], []

    for _, row in vendor_df.iterrows():
        vendor_name = row["vendor"]
        total_score, total_weight = 0.0, 0.0
        area_scores = {}

        for func_col in vendor_df.columns:
            if func_col == "vendor":
                continue

            func_name = normalize_case(func_col)
            if func_name not in func_to_req:
                continue

            req = func_to_req[func_name]
            weight = REQUIREMENT_WEIGHTS.get(req, 0)
            if weight == 0:
                continue

            raw_resp = row[func_col]
            resp = normalize_case(raw_resp)
            resp_score = RESPONSE_SCORES.get(resp, None)
            if resp_score is None:
                resp_score = 0.0
                resp_note = f"Unknown response '{raw_resp}'"
            else:
                resp_note = ""

            weighted_score = resp_score * weight
            total_score += weighted_score
            total_weight += weight

            area = func_to_area.get(func_name, "unspecified")
            if area not in area_scores:
                area_scores[area] = {"score": 0.0, "weight": 0.0}
            area_scores[area]["score"] += weighted_score
            area_scores[area]["weight"] += weight

            meets = "Meets Criteria" if resp_score >= MEETS_THRESHOLD else "Does Not Meet"
            weighted_pct = round((weighted_score / weight) * 100 if weight > 0 else 0, 2)
            detailed_records.append({
                "Vendor": vendor_name,
                "Business Area": area,
                "Function": func_col,
                "Requirement": req,
                "Response": raw_resp,
                "Response Note": resp_note,
                "Weighted Score (%) of that function": weighted_pct,
                "Meets Criteria": meets
            })

        vendor_total_pct = round((total_score / total_weight) * 100 if total_weight > 0 else 0, 2)
        summary_row = {"Vendor": vendor_name, "Total Score (%)": vendor_total_pct}
        for area, vals in area_scores.items():
            area_pct = round((vals["score"] / vals["weight"]) * 100 if vals["weight"] > 0 else 0, 2)
            summary_row[f"{area} (%)"] = area_pct
        vendor_scores.append(summary_row)

    return pd.DataFrame(vendor_scores), pd.DataFrame(detailed_records)

def convert_df(df):
    return df.to_csv(index=False).encode("utf-8")

# -------------------------------
# STREAMLIT APP
# -------------------------------
st.set_page_config(page_title="Vendor Functionality Scoring", layout="wide")
st.title("📊 Vendor Functionality Scoring Tool")

st.markdown("""
Upload a **System Criteria CSV** to score vendors against the fixed vendor functionality file.

Notes:
- Expected columns (case-insensitive) in Criteria file: Function, Requirement, Business Area
- If your CSV has extra header rows (e.g., a title row), the app will try to auto-detect the real header.
- If auto-detection fails, open the file in Excel/LibreOffice, move the row containing Function/Requirement/Business Area to be the top header row and save as CSV, then re-upload.
""")

# Load fixed vendor file if present
vendor_df = None
if os.path.exists(VENDOR_FILE):
   
