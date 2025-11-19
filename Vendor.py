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

def _canonical_requirement_key(req_text: str) -> str:
    if req_text is None:
        return ""
    req_norm = normalize_case(req_text)
    for key in REQUIREMENT_WEIGHTS.keys():
        if key in req_norm:
            return key
    if req_norm in REQUIREMENT_WEIGHTS:
        return req_norm
    return ""

def _interpret_response(resp_raw: object) -> (float, str):
    if pd.isna(resp_raw):
        return RESPONSE_SCORES.get("not provided", 0.5), "not provided"

    s = normalize_case(resp_raw)
    yes_vals = {"yes", "y", "true", "1", "available", "supported"}
    if s in yes_vals:
        return RESPONSE_SCORES.get("yes", 1), "yes"

    not_provided_vals = {"not provided", "n/a", "na", "-", "none", "unknown", "no response"}
    partial_vals = {"partial", "partially", "limited", "some", "sometimes", "conditional"}
    if s in not_provided_vals:
        return RESPONSE_SCORES.get("not provided", 0.5), "not provided"
    if s in partial_vals:
        return RESPONSE_SCORES.get("not provided", 0.5), "not provided"

    no_vals = {"no", "n", "false", "0", "not supported", "not available"}
    if s in no_vals:
        return RESPONSE_SCORES.get("no", 0), "no"

    try:
        f = float(s)
        f = max(0.0, min(1.0, f))
        if f >= 0.75:
            return 1.0, "yes"
        if f >= 0.5:
            return RESPONSE_SCORES.get("not provided", 0.5), "not provided"
        return 0.0, "no"
    except Exception:
        return RESPONSE_SCORES.get("not provided", 0.5), "not provided"

def calculate_scores(vendor_df, criteria_df):
    vendor_df = vendor_df.copy()
    criteria_df = criteria_df.copy()
    vendor_col_map = {normalize_case(c): c for c in vendor_df.columns}

    ok, missing, crit_map = _ensure_columns(criteria_df, EXPECTED_CRITERIA_COLS)
    if not ok:
        raise ValueError(
            f"Criteria file is missing required columns (case-insensitive): {missing}. "
            f"Found columns: {list(criteria_df.columns)}"
        )

    vendor_col_actual = None
    for c in vendor_df.columns:
        if normalize_case(c) == "vendor":
            vendor_col_actual = c
            break
    if vendor_col_actual is None:
        raise ValueError(
            f"Vendor file must include a 'Vendor' column (case-insensitive). Found columns: {list(vendor_df.columns)}"
        )

    vendor_scores_acc = {str(v): {"score": 0.0, "weight": 0.0, "areas": {}} for v in vendor_df[vendor_col_actual].astype(str).tolist()}
    detailed_records = []

    for _, crit in criteria_df.iterrows():
        func_orig = crit[crit_map["function"]]
        req_orig = crit[crit_map["requirement"]]
        area_orig = crit[crit_map["business area"]]

        func_norm = normalize_case(func_orig)
        vendor_col_for_func = vendor_col_map.get(func_norm)
        if not vendor_col_for_func:
            continue

        req_key = _canonical_requirement_key(req_orig)
        weight = REQUIREMENT_WEIGHTS.get(req_key, 0)
        if weight == 0:
            continue

        for idx, vendor_row in vendor_df.iterrows():
            vendor_name = str(vendor_row[vendor_col_actual])
            raw_resp = vendor_row.get(vendor_col_for_func, "")
            resp_score, resp_canon = _interpret_response(raw_resp)
            weighted_score = resp_score * weight

            vacc = vendor_scores_acc.setdefault(vendor_name, {"score": 0.0, "weight": 0.0, "areas": {}})
            vacc["score"] += weighted_score
            vacc["weight"] += weight

            area_key = normalize_case(area_orig) if pd.notna(area_orig) else "unspecified"
            if area_key not in vacc["areas"]:
                vacc["areas"][area_key] = {"score": 0.0, "weight": 0.0}
            vacc["areas"][area_key]["score"] += weighted_score
            vacc["areas"][area_key]["weight"] += weight

            meets = "Meets Criteria" if resp_score >= MEETS_THRESHOLD else "Does Not Meet"
            weighted_pct = round((weighted_score / weight) * 100 if weight > 0 else 0, 2)
            detailed_records.append({
                "Vendor": vendor_name,
                "Business Area": area_orig,
                "Function": func_orig,
                "Requirement": req_orig,
                "Response": raw_resp,
                "Weighted Score (%) of that function": weighted_pct,
                "Meets Criteria": meets
            })

    summary_rows = []
    for vendor_name, vals in vendor_scores_acc.items():
        total_pct = round((vals["score"] / vals["weight"]) * 100 if vals["weight"] > 0 else 0, 2)
        row = {"Vendor": vendor_name, "Total Score (%)": total_pct}
        for area, sub in vals["areas"].items():
            area_pct = round((sub["score"] / sub["weight"]) * 100 if sub["weight"] > 0 else 0, 2)
            row[f"{area} (%)"] = area_pct
        summary_rows.append(row)

    summary_df = pd.DataFrame(summary_rows)
    detailed_df = pd.DataFrame(detailed_records)

    if not summary_df.empty:
        cols = [c for c in summary_df.columns if c != "Vendor"]
        summary_df = summary_df[["Vendor"] + cols]

    return summary_df, detailed_df

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

    # --- Culturehost filter checkbox (optional) ---
    def find_col_case_insensitive(df, target_name):
        target_norm = normalize_case(target_name)
        for c in df.columns:
            if normalize_case(c) == target_norm:
                return c
        return None

    filter_ch_yes = st.checkbox("Only show vendors with Culturehost Connection == Yes", value=False, key="filter_ch_yes")

    # compute list of vendors that match the CultureHost Connection == yes
    filtered_vendor_names = None
    if filter_ch_yes:
        vendor_col_actual = find_col_case_insensitive(vendor_df, "vendor")
        ch_col_actual = find_col_case_insensitive(vendor_df, "CultureHost Connection")
        if vendor_col_actual is None:
            st.warning("Vendor file does not contain a 'Vendor' column (case-insensitive); cannot apply CultureHost filter.")
            filtered_vendor_names = []
        elif ch_col_actual is None:
            st.warning("Vendor file does not contain a 'CultureHost Connection' column (case-insensitive).")
            filtered_vendor_names = []
        else:
            series = vendor_df[ch_col_actual].astype(str).apply(lambda x: normalize_case(x))
            accepted_yes = {"yes", "y", "true", "1"}
            filtered_vendor_names = vendor_df.loc[series.isin(accepted_yes), vendor_col_actual].astype(str).tolist()
            if not filtered_vendor_names:
                st.info("No vendors found with CultureHost Connection == 'Yes'")

    # --- Pricing Model multi-select filter (new) ---
    pricing_col_actual = find_col_case_insensitive(vendor_df, "Pricing Model")
    pricing_filtered_vendor_names = None
    pricing_selection = None

    if pricing_col_actual is not None:
        raw_vals = vendor_df[pricing_col_actual].fillna("").astype(str).map(lambda s: s.strip())
        unique_vals = sorted([v for v in pd.unique(raw_vals) if v != ""], key=lambda s: s.lower())
        options = ["All"] + unique_vals
        pricing_selection_list = st.multiselect(
            "Filter by Pricing Model (multi-select)",
            options=options,
            default=["All"],
            help="Choose one or more Pricing Models to filter vendors (All = no filter)"
        )

        if pricing_selection_list and ("All" in pricing_selection_list):
            pricing_filtered_vendor_names = None
        elif pricing_selection_list:
            chosen_lower = {s.strip().lower() for s in pricing_selection_list}
            mask = raw_vals.str.lower().isin(chosen_lower)
            pricing_filtered_vendor_names = vendor_df.loc[mask, find_col_case_insensitive(vendor_df, "vendor")].astype(str).tolist()
            if not pricing_filtered_vendor_names:
                st.info("No vendors match the selected Pricing Model(s).")
        else:
            pricing_filtered_vendor_names = None
    else:
        st.info("Note: no 'Pricing Model' column found in the vendor file — Pricing filter unavailable.")

    # --- Functionality requirement filter (new) ---
    # Allow user to require that vendors have specific function(s) (meet criteria)
    function_filtered_vendor_names = None
    if not detailed_df.empty:
        available_functions = sorted(pd.unique(detailed_df["Function"].astype(str)))
        st.markdown("---")
        st.write("Functionality requirement filter — require vendors to have (meet) selected functions.")
        func_selection = st.multiselect(
            "Select functions that vendors must satisfy (vendors must meet all selected functions)",
            options=available_functions,
            default=[],
            help="Pick one or more functions. Only vendors that meet each selected function will be shown."
        )
        if func_selection:
            # Let user choose whether to require 'at least one criteria row' or 'all criteria rows' for the function
            mode = st.radio(
                "For each selected function, require:",
                options=[
                    "At least one matching criteria row to be 'Meets Criteria' (lenient)",
                    "All matching criteria rows must be 'Meets Criteria' (strict)"
                ],
                index=0
            )
            # Pre-compute vendor x function meet stats from detailed_df
            # group by Vendor and Function
            grp = detailed_df.groupby(["Vendor", "Function"])["Meets Criteria"].apply(list).reset_index()
            # Build maps vendor->function->(any_meet, all_meet)
            vendor_func_ok = {}
            for _, r in grp.iterrows():
                v = r["Vendor"]
                f = r["Function"]
                meets_list = [str(x) for x in r["Meets Criteria"]]
                any_meet = any(x == "Meets Criteria" for x in meets_list)
                all_meet = all(x == "Meets Criteria" for x in meets_list)
                vendor_func_ok.setdefault(v, {})[f] = {"any": any_meet, "all": all_meet}

            # Evaluate vendors
            selected_vendors = []
            for v in vendor_df[find_col_case_insensitive(vendor_df, "vendor")].astype(str).tolist():
                ok_vendor = True
                for f in func_selection:
                    vmap = vendor_func_ok.get(v, {})
                    stat = vmap.get(f)
                    if stat is None:
                        # vendor has no rows for this function -> fail
                        ok_vendor = False
                        break
                    if mode.startswith("At least one"):
                        if not stat["any"]:
                            ok_vendor = False
                            break
                    else:
                        if not stat["all"]:
                            ok_vendor = False
                            break
                if ok_vendor:
                    selected_vendors.append(v)
            function_filtered_vendor_names = selected_vendors
            if not function_filtered_vendor_names:
                st.info("No vendors meet the selected required functions.")
        st.markdown("-----")

    # Compose filters: pricing + CultureHost + function (intersection if multiple active)
    final_filtered_vendor_names = None
    active_filters = []
    filters_lists = []
    if pricing_filtered_vendor_names is not None:
        filters_lists.append(set(pricing_filtered_vendor_names))
    if filtered_vendor_names is not None:
        filters_lists.append(set(filtered_vendor_names))
    if function_filtered_vendor_names is not None:
        filters_lists.append(set(function_filtered_vendor_names))

    if filters_lists:
        # intersection of all active filters
        intersection = set.intersection(*filters_lists) if len(filters_lists) > 1 else filters_lists[0]
        final_filtered_vendor_names = list(intersection)

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
