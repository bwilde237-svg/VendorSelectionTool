import os
import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# -------------------------------
# CONFIGURATION
# -------------------------------
VENDOR_FILE = "Connected Booking System Evaluation Template (NDA).xlsx - Scoring Upload (1).csv"

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
# HELPERS
# -------------------------------
def normalize_case(s):
    return str(s).strip().lower()

def _ensure_columns(df, expected_cols):
    normalized = {normalize_case(c): c for c in df.columns}
    missing = [c for c in expected_cols if c not in normalized]
    return (len(missing) == 0, missing, normalized)

def _detect_header_and_load_csv(file_like):
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

    file_like.seek(0)
    df = pd.read_csv(file_like, header=header_row_index)
    return df

def load_criteria_file(uploaded_file):
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

# universal display helper that ensures no left-hand index is shown and enforces readable colors.
def display_df_no_index(df: pd.DataFrame, height: int | None = None):
    if df is None:
        return

    df2 = df.copy().reset_index(drop=True)

    # drop obvious exported index column if present
    if df2.shape[1] >= 1:
        first_col_name = str(df2.columns[0]).strip().lower()
        first_col_vals = df2.iloc[:, 0].tolist()
        name_flag = first_col_name in ["", "unnamed: 0", "unnamed:0", "unnamed", "index", "0"]
        seq_flag = False
        try:
            n = len(first_col_vals)
            seq_strs = [str(i) for i in range(n)]
            seq_ints = list(range(n))
            if [str(v) for v in first_col_vals] == seq_strs or first_col_vals == seq_ints:
                seq_flag = True
        except Exception:
            seq_flag = False
        if name_flag or seq_flag:
            if df2.shape[1] >= 2:
                df2 = df2.iloc[:, 1:]

    # Try pandas Styler.hide_index() (best)
    try:
        styler = df2.style.hide_index()
        kwargs = {"use_container_width": True}
        if height is not None:
            kwargs["height"] = height
        st.dataframe(styler, **kwargs)
        return
    except Exception:
        pass

    # Render as HTML with CSS that respects light/dark (stronger colors)
    try:
        html_table = df2.to_html(index=False, border=0, classes="mytable")
        css = """
        <style>
        table.mytable { width:100%; border-collapse:collapse; font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial; }
        table.mytable th, table.mytable td { padding:8px 12px; border-bottom:1px solid rgba(0,0,0,0.08) !important; color: #111 !important; }
        table.mytable th { font-weight:700; background:transparent !important; }
        @media (prefers-color-scheme: dark) {
            table.mytable th, table.mytable td { color: #e6eef8 !important; border-bottom:1px solid rgba(255,255,255,0.06) !important; }
            table.mytable th { color: #fff !important; }
            table.mytable td { color: #ddd !important; }
            table.mytable { background: transparent !important; }
        }
        .table-wrapper { width:100%; overflow:auto; }
        </style>
        """
        html_block = f"<div class='table-wrapper'>{css}{html_table}</div>"
        if height is None:
            components.html(html_block, scrolling=True, height=300)
        else:
            components.html(html_block, scrolling=True, height=height)
        return
    except Exception:
        pass

    # Last fallback
    records = df2.to_dict(orient="records")
    df_show = pd.DataFrame.from_records(records)
    kwargs = {"use_container_width": True}
    if height is not None:
        kwargs["height"] = height
    st.dataframe(df_show, **kwargs)

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

    # --- CultureHosts filter checkbox (optional) ---
    def find_col_case_insensitive(df, target_name):
        target_norm = normalize_case(target_name)
        for c in df.columns:
            if normalize_case(c) == target_norm:
                return c
        return None

    filter_ch_yes = st.checkbox("Only show vendors with CultureHosts Connection == Yes", value=False, key="filter_ch_yes")

    # compute list of vendors that match the CultureHosts Connection == yes
    filtered_vendor_names = None
    if filter_ch_yes:
        vendor_col_actual = find_col_case_insensitive(vendor_df, "vendor")
        ch_col_actual = find_col_case_insensitive(vendor_df, "CultureHosts Connection")
        if vendor_col_actual is None:
            st.warning("Vendor file does not contain a 'Vendor' column (case-insensitive); cannot apply CultureHosts filter.")
            filtered_vendor_names = []
        elif ch_col_actual is None:
            st.warning("Vendor file does not contain a 'CultureHosts Connection' column (case-insensitive).")
            filtered_vendor_names = []
        else:
            series = vendor_df[ch_col_actual].astype(str).apply(lambda x: normalize_case(x))
            accepted_yes = {"yes", "y", "true", "1"}
            filtered_vendor_names = vendor_df.loc[series.isin(accepted_yes), vendor_col_actual].astype(str).tolist()
            if not filtered_vendor_names:
                st.info("No vendors found with CultureHosts Connection == 'Yes'")

    # --- Pricing Model multi-select filter (new) ---
    pricing_col_actual = find_col_case_insensitive(vendor_df, "Pricing Model")
    pricing_filtered_vendor_names = None
    pricing_selection = None

    if pricing_col_actual is not None:
        raw_vals = vendor_df[pricing_col_actual].fillna("").astype(str).map(lambda s: s.strip())
        unique_vals = sorted([v for v in pd.unique(raw_vals) if v != ""], key=lambda s: s.lower())
        options = ["All"] + unique_vals
        # Use multiselect so user can pick multiple pricing models; default is "All"
        pricing_selection_list = st.multiselect(
            "Filter by Pricing Model (multi-select)",
            options=options,
            default=["All"],
            help="Choose one or more Pricing Models to filter vendors (All = no filter)"
        )

        # interpret selection
        if pricing_selection_list and ("All" in pricing_selection_list):
            pricing_filtered_vendor_names = None  # All selected => no filter
        elif pricing_selection_list:
            # Build case-insensitive set of chosen models
            chosen_lower = {s.strip().lower() for s in pricing_selection_list}
            mask = raw_vals.str.lower().isin(chosen_lower)
            pricing_filtered_vendor_names = vendor_df.loc[mask, find_col_case_insensitive(vendor_df, "vendor")].astype(str).tolist()
            if not pricing_filtered_vendor_names:
                st.info("No vendors match the selected Pricing Model(s).")
        else:
            pricing_filtered_vendor_names = None  # nothing selected => treat as no filter
    else:
        st.info("Note: no 'Pricing Model' column found in the vendor file — Pricing filter unavailable.")

    # Compose filters: pricing + CultureHosts (intersection if both active)
    final_filtered_vendor_names = None
    if pricing_filtered_vendor_names is not None and filtered_vendor_names is not None:
        final_filtered_vendor_names = list(set(pricing_filtered_vendor_names).intersection(set(filtered_vendor_names)))
    elif pricing_filtered_vendor_names is not None:
        final_filtered_vendor_names = pricing_filtered_vendor_names
    elif filtered_vendor_names is not None:
        final_filtered_vendor_names = filtered_vendor_names
    else:
        final_filtered_vendor_names = None

    # --- Top-N controls (applies the composed filters) ---
    if summary_df.empty:
        st.info("No scored vendors to display.")
    else:
        max_top = len(summary_df)
        default_top = 5 if max_top >= 5 else max_top
        n_top = st.slider(
            "Number of top vendors to display",
            min_value=1,
            max_value=max_top,
            value=default_top,
            help="Select how many top vendors to show (by Total Score (%))"
        )

        # If filters are active, restrict the summary_source first so Top-N is among the filtered set
        if final_filtered_vendor_names is not None:
            summary_source = summary_df[summary_df["Vendor"].isin(final_filtered_vendor_names)].copy()
        else:
            summary_source = summary_df.copy()

        top_summary = summary_source.sort_values("Total Score (%)", ascending=False).head(n_top)

        # minimal top summary (Rank + Vendor + Total Score)
        top_summary_minimal = top_summary[["Vendor", "Total Score (%)"]].copy().reset_index(drop=True)
        top_summary_minimal.insert(0, "Rank", range(1, 1 + len(top_summary_minimal)))

        st.subheader(f"Top {n_top} Vendors — Overall Rankings")
        display_df_no_index(top_summary_minimal)

        # Business area breakdown restricted to top N
        area_cols = [c for c in summary_df.columns if c not in ["Vendor", "Total Score (%)"]]
        if area_cols:
            st.subheader("Business Area Breakdown (Top selection)")
            display_df_no_index(top_summary[["Vendor"] + area_cols])

        # Vendor multi-select (from Top-N)
        st.subheader("Inspect Vendor(s) Criteria (select from the Top selection)")
        top_vendors = top_summary["Vendor"].tolist()
        vendor_select = st.multiselect(
            "Select vendor(s) to inspect",
            options=top_vendors,
            default=[top_vendors[0]] if top_vendors else [],
            help="Pick one or more vendors to inspect. Each vendor will get its own tab."
        )

        if not vendor_select:
            st.info("Select at least one vendor to inspect their criteria.")
        else:
            tabs = st.tabs(vendor_select)
            for tab_label, tab in zip(vendor_select, tabs):
                with tab:
                    vendor_rows = detailed_df[detailed_df["Vendor"] == tab_label]
                    if vendor_rows.empty:
                        st.warning(
                            f"No detailed criteria rows found for '{tab_label}'. This can happen if the vendor "
                            "file doesn't have matching function columns for the criteria file."
                        )
                        continue

                    met_df = vendor_rows[vendor_rows["Meets Criteria"] == "Meets Criteria"].reset_index(drop=True)
                    not_met_df = vendor_rows[vendor_rows["Meets Criteria"] != "Meets Criteria"].reset_index(drop=True)

                    st.markdown(f"**{tab_label} — Summary:** {len(met_df)} functions meet criteria; {len(not_met_df)} do not meet.")
                    cols_left, cols_right = st.columns(2)
                    with cols_left:
                        with st.expander("Functions that Meet Criteria", expanded=True):
                            if not met_df.empty:
                                display_df_no_index(met_df)
                            else:
                                st.info("No functions meeting criteria for this vendor.")
                    with cols_right:
                        with st.expander("Functions that Do Not Meet Criteria", expanded=True):
                            if not not_met_df.empty:
                                display_df_no_index(not_met_df)
                            else:
                                st.info("All scored functions meet criteria for this vendor!")

                    st.download_button(
                        label=f"⬇️ Download {tab_label} Detailed CSV",
                        data=convert_df(vendor_rows),
                        file_name=f"{tab_label}_detailed.csv",
                        mime="text/csv"
                    )

        # Detailed Results (Top selection) - after inspect UI
        st.subheader("Detailed Results (Top selection)")
        filtered_detailed = detailed_df[detailed_df["Vendor"].isin(top_vendors)].reset_index(drop=True)
        display_df_no_index(filtered_detailed, height=500)

        # Download buttons
        st.download_button(
            label="⬇️ Download Full Summary CSV",
            data=convert_df(summary_df),
            file_name="vendor_scores_summary_full.csv",
            mime="text/csv"
        )

        st.download_button(
            label="⬇️ Download Full Detailed CSV",
            data=convert_df(detailed_df),
            file_name="vendor_scores_detailed_full.csv",
            mime="text/csv"
        )

        st.download_button(
            label=f"⬇️ Download Top {n_top} Summary CSV (Rank + Vendor + Total Score)",
            data=convert_df(top_summary_minimal),
            file_name=f"vendor_scores_top_{n_top}_summary_minimal.csv",
            mime="text/csv"
        )

elif criteria_file is not None and vendor_df is None:
    st.error("No vendor file available. Upload a vendor CSV (or place the vendor file next to the app with the configured filename).")
else:
    st.info("Please upload a System Criteria CSV file to begin scoring.")
