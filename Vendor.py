import os
import streamlit as st
import pandas as pd

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
    """Robust CSV header detection.

    Reads the uploaded CSV and searches the first several rows for a row that
    contains the expected header names (Function, Requirement, Business Area).
    If found, re-reads the CSV using that row as the header.
    """
    # Try a simple read first
    try:
        file_like.seek(0)
        df = pd.read_csv(file_like)
        ok, missing, _ = _ensure_columns(df, EXPECTED_CRITERIA_COLS)
        if ok:
            return df
    except Exception:
        file_like.seek(0)

    # Read without header to inspect the top rows
    try:
        file_like.seek(0)
        raw = pd.read_csv(file_like, header=None, dtype=str, keep_default_na=False)
    except Exception as e:
        file_like.seek(0)
        raise ValueError(f"Unable to parse CSV: {e}")

    max_rows_to_scan = min(20, len(raw))
    header_row_index = None
    for i in range(max_rows_to_scan):
        row_vals = [normalize_case(str(x)) for x in raw.iloc[i].tolist()]
        matches = sum(1 for c in EXPECTED_CRITERIA_COLS if any(c in v for v in row_vals))
        if matches == len(EXPECTED_CRITERIA_COLS):
            header_row_index = i
            break

    if header_row_index is None:
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

    # Re-read using the located header row
    file_like.seek(0)
    df = pd.read_csv(file_like, header=header_row_index)
    return df

def load_criteria_file(uploaded_file):
    """Load criteria CSV/XLSX robustly."""
    name = getattr(uploaded_file, "name", "")
    uploaded_file.seek(0)
    if name.lower().endswith((".xls", ".xlsx")):
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            raise ValueError(f"Failed to read Excel criteria file: {e}")
    else:
        df = _detect_header_and_load_csv(uploaded_file)
    return df

def calculate_scores(vendor_df, criteria_df):
    vendor_df = vendor_df.copy()
    criteria_df = criteria_df.copy()
    vendor_df.columns = [normalize_case(c) for c in vendor_df.columns]
    criteria_df.columns = [normalize_case(c) for c in criteria_df.columns]

    ok, missing, _ = _ensure_columns(criteria_df, EXPECTED_CRITERIA_COLS)
    if not ok:
        raise ValueError(
            f"Criteria file is missing required columns (case-insensitive): {missing}. "
            f"Found columns: {list(criteria_df.columns)}"
        )

    if "vendor" not in vendor_df.columns:
        raise ValueError(
            f"Vendor file must include a 'Vendor' column (case-insensitive). "
            f"Found columns: {list(vendor_df.columns)}"
        )

    func_to_req = dict(zip(
        [normalize_case(x) for x in criteria_df["function"]],
        [normalize_case(x) for x in criteria_df["requirement"]]
    ))
    func_to_area = dict(zip(
        [normalize_case(x) for x in criteria_df["function"]],
        [normalize_case(x) for x in criteria_df["business area"]]
    ))

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

st.markdown(
    "Upload a **System Criteria CSV** to score vendors against the fixed vendor functionality file.\n\n"
    "Expected columns in Criteria file (case-insensitive): Function, Requirement, Business Area\n"
    "If your CSV has extra header rows (e.g., a title row) the app will try to auto-detect the real header."
)

# Load fixed vendor file if present
vendor_df = None
if os.path.exists(VENDOR_FILE):
    try:
        vendor_df = pd.read_csv(VENDOR_FILE)
        st.info(f"Loaded vendor file from: {VENDOR_FILE}")
    except Exception as e_csv:
        try:
            vendor_df = pd.read_excel(VENDOR_FILE)
            st.info(f"Loaded vendor file (Excel) from: {VENDOR_FILE}")
        except Exception as e:
            st.warning(f"Failed to load bundled vendor file '{VENDOR_FILE}': {e_csv}; excel fallback: {e}. You can upload a vendor CSV below.")
else:
    st.info(f"Bundled vendor file not found at: {VENDOR_FILE}. You can upload a vendor CSV below.")

# vendor upload if not present
if vendor_df is None:
    vendor_upload = st.file_uploader("Upload Vendor file (CSV or Excel)", type=["csv", "xlsx", "xls"], key="vendor_upload")
    if vendor_upload is not None:
        try:
            if vendor_upload.name.lower().endswith((".xls", ".xlsx")):
                vendor_df = pd.read_excel(vendor_upload)
            else:
                vendor_df = pd.read_csv(vendor_upload)
            st.success("Vendor file uploaded successfully.")
        except Exception as e:
            st.error(f"Error reading uploaded vendor file: {e}")

criteria_file = st.file_uploader("Upload your System Criteria CSV", type=["csv", "xlsx", "xls"], key="criteria_upload")

if criteria_file is not None and vendor_df is not None:
    try:
        criteria_df = load_criteria_file(criteria_file)
    except Exception as e:
        st.error(f"Error reading criteria file: {e}")
        st.stop()

    try:
        with st.spinner("Calculating scores..."):
            summary_df, detailed_df = calculate_scores(vendor_df, criteria_df)
    except Exception as e:
        st.error(f"Error during scoring: {e}")
        st.stop()

    st.success("✅ Scoring complete!")

    st.subheader("Overall Vendor Rankings")
    if not summary_df.empty:
        st.dataframe(summary_df.sort_values("Total Score (%)", ascending=False), use_container_width=True)
    else:
        st.info("No scored vendors to display.")

    # Business area averages
    area_cols = [c for c in summary_df.columns if c not in ["Vendor", "Total Score (%)"]]
    if area_cols:
        st.subheader("Business Area Breakdown")
        st.dataframe(summary_df[["Vendor"] + area_cols], use_container_width=True)

    st.subheader("Detailed Results")
    st.dataframe(detailed_df, use_container_width=True)

    # Download buttons
    st.download_button(
        label="⬇️ Download Summary CSV",
        data=convert_df(summary_df),
        file_name="vendor_scores_summary.csv",
        mime="text/csv"
    )

    st.download_button(
        label="⬇️ Download Detailed CSV",
        data=convert_df(detailed_df),
        file_name="vendor_scores_detailed.csv",
        mime="text/csv"
    )
elif criteria_file is not None and vendor_df is None:
    st.error("No vendor file available. Upload a vendor CSV (or place the vendor file next to the app with the configured filename).")
else:
    st.info("Please upload a System Criteria CSV file to begin scoring.")
