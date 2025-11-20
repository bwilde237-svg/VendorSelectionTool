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

def find_col_case_insensitive(df, target_name):
    target_norm = normalize_case(target_name)
    for c in df.columns:
        if normalize_case(c) == target_norm:
            return c
    return None

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
    """Map a free-text Requirement cell to one of the canonical keys in REQUIREMENT_WEIGHTS."""
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
    """
    Interpret many variants of Yes/No/Partial/Not provided and return (score, canonical_string).

    - Accepts substring matches so "Yes - via API" or "Included" are recognized.
    - Parses percent strings like "75%" or numeric values.
    - Falls back to 'not provided' (0.5) for unknown free text.
    """
    if pd.isna(resp_raw):
        return RESPONSE_SCORES.get("not provided", 0.5), "not provided"

    s = normalize_case(resp_raw).strip()

    # Numeric / percentage handling
    try:
        if isinstance(resp_raw, str) and "%" in resp_raw:
            num = float(resp_raw.replace("%", "").strip())
            val = max(0.0, min(1.0, num / 100.0))
            if val >= 0.75:
                return 1.0, "yes"
            if val >= 0.5:
                return RESPONSE_SCORES.get("not provided", 0.5), "not provided"
            return 0.0, "no"
        num = float(s)
        if num > 1.0:
            val = max(0.0, min(1.0, num / 100.0))
        else:
            val = max(0.0, min(1.0, num))
        if val >= 0.75:
            return 1.0, "yes"
        if val >= 0.5:
            return RESPONSE_SCORES.get("not provided", 0.5), "not provided"
        return 0.0, "no"
    except Exception:
        pass

    yes_keywords = ["yes", "y", "true", "1", "available", "supported", "included", "included in", "included:"]
    not_provided_keywords = ["not provided", "n/a", "na", "-", "none", "unknown", "no response", "tbd", "tbc", "not stated"]
    partial_keywords = ["partial", "partially", "limited", "some", "sometimes", "conditional", "part"]
    no_keywords = ["no", "n", "false", "0", "not supported", "not available", "nope"]

    for kw in yes_keywords:
        if kw in s:
            return RESPONSE_SCORES.get("yes", 1), "yes"
    for kw in not_provided_keywords:
        if kw in s:
            return RESPONSE_SCORES.get("not provided", 0.5), "not provided"
    for kw in partial_keywords:
        if kw in s:
            return RESPONSE_SCORES.get("not provided", 0.5), "not provided"
    for kw in no_keywords:
        if kw in s:
            return RESPONSE_SCORES.get("no", 0), "no"

    return RESPONSE_SCORES.get("not provided", 0.5), "not provided"

def calculate_scores(vendor_df, criteria_df):
    """
    Scoring logic. IMPORTANT: the 'Pricing Model' column is excluded from scoring (used only for filtering).
    """
    vendor_df = vendor_df.copy()
    criteria_df = criteria_df.copy()

    # Build normalized->actual column name map for vendor file excluding Pricing Model
    vendor_col_map = {
        normalize_case(c): c
        for c in vendor_df.columns
        if normalize_case(c) != "pricing model"
    }

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

    # Iterate criteria rows (supports multiple rows with same Function)
    for _, crit in criteria_df.iterrows():
        func_orig = crit[crit_map["function"]]
        req_orig = crit[crit_map["requirement"]]
        area_orig = crit[crit_map["business area"]]

        func_norm = normalize_case(func_orig)
        if func_norm == "pricing model":
            # explicitly ignore pricing model as a function for scoring
            continue

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

    # Render as HTML with CSS that respects light/dark
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

# Initialize session state defaults for filters
if "filter_ch_yes" not in st.session_state:
    st.session_state["filter_ch_yes"] = False
if "pricing_selection_list" not in st.session_state:
    st.session_state["pricing_selection_list"] = ["All"]
if "pricing_selection_list_display" not in st.session_state:
    st.session_state["pricing_selection_list_display"] = ["All"]
if "func_selection" not in st.session_state:
    st.session_state["func_selection"] = []
if "top_n" not in st.session_state:
    st.session_state["top_n"] = 5

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

# -------------------------------
# Sidebar filters (reordered & separated)
# Order: Top-N slider -> Pricing Model -> Mandatory Functionality -> CultureHost
# -------------------------------
# Create a placeholder in the sidebar for the Mandatory Functionality control so ordering is stable.
with st.sidebar:
    st.header("Filters & Controls")

    def _reset_filters():
        st.session_state["filter_ch_yes"] = False
        st.session_state["pricing_selection_list"] = ["All"]
        st.session_state["pricing_selection_list_display"] = ["All"]
        st.session_state["func_selection"] = []
        st.session_state["top_n"] = 5
        # remove internal mapping if present
        st.session_state.pop("_pricing_display_map", None)
        # No explicit st.experimental_rerun() call — Streamlit will rerun automatically on widget interaction.

    st.button("Reset filters to defaults", on_click=_reset_filters)

    st.markdown("---")
    # Top-N slider at top of sidebar
    st.slider(
        "Number of vendors to display",
        min_value=1,
        max_value=100,
        value=st.session_state.get("top_n", 5),
        key="top_n",
        help="How many top vendors to show (applies after scoring and filtering)."
    )

    st.markdown("---")
    # Pricing Model filter next
    st.write("Pricing Model filter")
    if vendor_df is not None:
        pricing_col_actual = find_col_case_insensitive(vendor_df, "Pricing Model")
        if pricing_col_actual is not None:
            raw_vals = vendor_df[pricing_col_actual].fillna("").astype(str).map(lambda s: s.strip())
            counts = raw_vals.str.lower().value_counts().to_dict()
            unique_vals = sorted([v for v in pd.unique(raw_vals) if v != ""], key=lambda s: s.lower())
            options_display = ["All"] + [f"{v} ({counts.get(v.lower(),0)})" for v in unique_vals]
            display_to_value = {f"{v} ({counts.get(v.lower(),0)})": v for v in unique_vals}
            st.session_state["_pricing_display_map"] = display_to_value
            # multiselect writes its value into session_state['pricing_selection_list_display']
            st.multiselect(
                "Filter by Pricing Model",
                options=options_display,
                default=st.session_state.get("pricing_selection_list_display", ["All"]),
                key="pricing_selection_list_display",
                help="Choose one or more Pricing Models to filter vendors (All = no filter)"
            )
            # Map the display selection into canonical list for later filtering
            display_selected = st.session_state.get("pricing_selection_list_display", ["All"])
            if display_selected and "All" in display_selected:
                st.session_state["pricing_selection_list"] = ["All"]
            elif display_selected:
                mapped = []
                for disp in display_selected:
                    mapped_val = st.session_state["_pricing_display_map"].get(disp, disp)
                    mapped.append(mapped_val)
                st.session_state["pricing_selection_list"] = mapped
            else:
                st.session_state["pricing_selection_list"] = ["All"]
        else:
            st.info("Note: no 'Pricing Model' column found — Pricing filter unavailable.")
    else:
        st.info("Upload vendor file to enable Pricing Model filter.")

    st.markdown("---")
    # Mandatory Functionality placeholder location (fixed)
    func_placeholder = st.empty()
    func_placeholder.write("Mandatory Functionality (STRICT)")
    func_placeholder.write("After you upload criteria, select required functions (options appear post-scoring).")
    if st.session_state.get("func_selection"):
        func_placeholder.write(f"{len(st.session_state.get('func_selection'))} function(s) selected")

    st.markdown("---")
    # CultureHost last
    st.checkbox(
        "CultureHost Connection",
        value=st.session_state.get("filter_ch_yes", False),
        key="filter_ch_yes"
    )

# -------------------------------
# Main behaviour once both files present
# -------------------------------
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

    # Build filters results (apply sidebar selections)
    # CultureHost filter
    filtered_vendor_names = None
    if st.session_state.get("filter_ch_yes", False):
        vendor_col_actual = find_col_case_insensitive(vendor_df, "vendor")
        ch_col_actual = find_col_case_insensitive(vendor_df, "CultureHost Connection")
        if vendor_col_actual is None:
            st.warning("Vendor file does not contain a 'Vendor' column; cannot apply CultureHost filter.")
            filtered_vendor_names = []
        elif ch_col_actual is None:
            st.warning("Vendor file does not contain a 'CultureHost Connection' column.")
            filtered_vendor_names = []
        else:
            series = vendor_df[ch_col_actual].astype(str).apply(lambda x: normalize_case(x))
            accepted_yes = {"yes", "y", "true", "1"}
            filtered_vendor_names = vendor_df.loc[series.isin(accepted_yes), vendor_col_actual].astype(str).tolist()
            if not filtered_vendor_names:
                st.info("No vendors found with CultureHost Connection == 'Yes'")

    # Pricing filter (interpret sidebar selection stored canonical in pricing_selection_list)
    pricing_filtered_vendor_names = None
    pricing_col_actual = find_col_case_insensitive(vendor_df, "Pricing Model")
    if pricing_col_actual is not None:
        pricing_selection_list = st.session_state.get("pricing_selection_list", ["All"])
        raw_vals = vendor_df[pricing_col_actual].fillna("").astype(str).map(lambda s: s.strip())
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

    # Build available functions list from criteria_df: only include functions with weight > 0 and exclude pricing model
    available_functions = []
    try:
        crit_map_ok, _, crit_map = _ensure_columns(criteria_df, EXPECTED_CRITERIA_COLS)
        if crit_map_ok:
            func_to_weights = {}
            for _, r in criteria_df.iterrows():
                fn = normalize_case(r[crit_map["function"]])
                req = r[crit_map["requirement"]]
                key = _canonical_requirement_key(req)
                w = REQUIREMENT_WEIGHTS.get(key, 0)
                func_to_weights.setdefault(fn, 0)
                func_to_weights[fn] = max(func_to_weights[fn], w)
            if not detailed_df.empty:
                funcs_in_detail = sorted(pd.unique(detailed_df["Function"].astype(str)))
                for f in funcs_in_detail:
                    if normalize_case(f) == "pricing model":
                        continue
                    if func_to_weights.get(normalize_case(f), 0) > 0:
                        available_functions.append(f)
    except Exception:
        available_functions = sorted(pd.unique(detailed_df["Function"].astype(str))) if not detailed_df.empty else []

    # Replace the placeholder with the actual multiselect in the sidebar (preserves ordering)
    if available_functions:
        func_placeholder.multiselect(
            "Select Mandatory functions (vendors must meet ALL matching rows)",
            options=available_functions,
            default=st.session_state.get("func_selection", []),
            key="func_selection",
            help="Pick functions vendors must fully satisfy (STRICT)."
        )

    # Evaluate the function filter if selection exists
    function_filtered_vendor_names = None
    if available_functions and st.session_state.get("func_selection"):
        chosen_funcs = st.session_state.get("func_selection", [])
        grp = detailed_df.groupby(["Vendor", "Function"])["Meets Criteria"].apply(list).reset_index()
        vendor_func_ok = {}
        for _, r in grp.iterrows():
            v = r["Vendor"]
            f = r["Function"]
            meets_list = [str(x) for x in r["Meets Criteria"]]
            all_meet = all(x == "Meets Criteria" for x in meets_list)
            vendor_func_ok.setdefault(v, {})[f] = {"all": all_meet}

        selected_vendors = []
        vendor_name_col = find_col_case_insensitive(vendor_df, "vendor")
        for v in vendor_df[vendor_name_col].astype(str).tolist():
            ok_vendor = True
            for f in chosen_funcs:
                vmap = vendor_func_ok.get(v, {})
                stat = vmap.get(f)
                if stat is None or not stat["all"]:
                    ok_vendor = False
                    break
            if ok_vendor:
                selected_vendors.append(v)
        function_filtered_vendor_names = selected_vendors
        if not function_filtered_vendor_names:
            st.info("No vendors meet the selected required functions.")

    # Compose filters: intersection of active filters
    filters_lists = []
    if pricing_filtered_vendor_names is not None:
        filters_lists.append(set(pricing_filtered_vendor_names))
    if filtered_vendor_names is not None:
        filters_lists.append(set(filtered_vendor_names))
    if function_filtered_vendor_names is not None:
        filters_lists.append(set(function_filtered_vendor_names))

    if filters_lists:
        intersection = set.intersection(*filters_lists) if len(filters_lists) > 1 else filters_lists[0]
        final_filtered_vendor_names = list(intersection)
    else:
        final_filtered_vendor_names = None

    # Metrics & applied filters summary
    total_vendors = len(summary_df)
    shown_vendors = len(final_filtered_vendor_names) if final_filtered_vendor_names is not None else total_vendors

    if final_filtered_vendor_names is not None:
        filtered_summary_for_metrics = summary_df[summary_df["Vendor"].isin(final_filtered_vendor_names)].copy()
    else:
        filtered_summary_for_metrics = summary_df.copy()

    top_score = int(filtered_summary_for_metrics["Total Score (%)"].max()) if not filtered_summary_for_metrics.empty else 0

    cols = st.columns(3)
    cols[0].metric("Vendors scored", total_vendors)
    cols[1].metric("Vendors shown", shown_vendors)
    cols[2].metric("Top score (%)", f"{top_score}%")

    applied = []
    if pricing_filtered_vendor_names is not None:
        applied.append(f"Pricing models: {', '.join(st.session_state.get('pricing_selection_list', []))}")
    if st.session_state.get("filter_ch_yes", False):
        applied.append("CultureHost = Yes")
    if st.session_state.get("func_selection"):
        applied.append(f"Required functions: {', '.join(st.session_state.get('func_selection', []))}")
    if applied:
        st.info("Active filters: " + " | ".join(applied))
    else:
        st.info("No filters active — showing all scored vendors")

    # Use single Top-N slider value from sidebar (st.session_state['top_n'])
    if summary_df.empty:
        st.info("No scored vendors to display.")
    else:
        max_top = len(filtered_summary_for_metrics)
        requested_top = st.session_state.get("top_n", 5)
        n_top = min(requested_top, max_top) if max_top > 0 else 0

        if n_top <= 0:
            st.info("No vendors available for Top-N display after filtering.")
        else:
            # Show top N vendors (after filters applied)
            if final_filtered_vendor_names is not None:
                summary_source = summary_df[summary_df["Vendor"].isin(final_filtered_vendor_names)].copy()
            else:
                summary_source = summary_df.copy()

            top_summary = summary_source.sort_values("Total Score (%)", ascending=False).head(n_top)

            top_summary_minimal = top_summary[["Vendor", "Total Score (%)"]].copy().reset_index(drop=True)
            top_summary_minimal.insert(0, "Rank", range(1, 1 + len(top_summary_minimal)))

            st.subheader(f"Top {n_top} Vendors — Overall Rankings")
            display_df_no_index(top_summary_minimal)

            area_cols = [c for c in summary_df.columns if c not in ["Vendor", "Total Score (%)"]]
            if area_cols:
                st.subheader("Business Area Breakdown (Top selection)")
                display_df_no_index(top_summary[["Vendor"] + area_cols])

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

            st.subheader("Detailed Results (Top selection)")
            filtered_detailed = detailed_df[detailed_df["Vendor"].isin(top_vendors)].reset_index(drop=True)
            display_df_no_index(filtered_detailed, height=500)

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
