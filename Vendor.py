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

# -------------------------------
# HELPER FUNCTIONS
# -------------------------------
def normalize_case(s):
    return str(s).strip().lower()

def calculate_scores(vendor_df, criteria_df):
    vendor_df.columns = [normalize_case(c) for c in vendor_df.columns]
    criteria_df.columns = [normalize_case(c) for c in criteria_df.columns]

    # Mappings from criteria file
    func_to_req = dict(zip(criteria_df["function"], criteria_df["requirement"]))
    func_to_area = dict(zip(criteria_df["function"], criteria_df["business area"]))

    func_to_req = {normalize_case(k): normalize_case(v) for k, v in func_to_req.items()}
    func_to_area = {normalize_case(k): v for k, v in func_to_area.items()}

    vendor_scores, detailed_records = [], []

    for _, row in vendor_df.iterrows():
        vendor_name = row["vendor"]
        total_score, total_weight = 0, 0
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

            resp = normalize_case(row[func_col])
            resp_score = RESPONSE_SCORES.get(resp, 0)
            weighted_score = resp_score * weight

            total_score += weighted_score
            total_weight += weight

            area = func_to_area.get(func_name, "Unspecified")
            if area not in area_scores:
                area_scores[area] = {"score": 0, "weight": 0}
            area_scores[area]["score"] += weighted_score
            area_scores[area]["weight"] += weight

            meets = "Meets Criteria" if (resp_score * 100) >= 75 else "Does Not Meet"
            detailed_records.append({
                "Vendor": vendor_name,
                "Business Area": area,
                "Function": func_col,
                "Requirement": req,
                "Response": row[func_col],
                "Weighted Score": round((weighted_score / weight) * 100 if weight > 0 else 0, 2),
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
""")

# Load fixed vendor file
try:
    vendor_df = pd.read_csv(VENDOR_FILE)
except Exception as e:
    st.error(f"Error loading vendor file: {e}")
    st.stop()

criteria_file = st.file_uploader("Upload your System Criteria CSV", type=["csv"])

if criteria_file is not None:
    criteria_df = pd.read_csv(criteria_file)

    with st.spinner("Calculating scores..."):
        summary_df, detailed_df = calculate_scores(vendor_df, criteria_df)

    st.success("✅ Scoring complete!")

    st.subheader("Overall Vendor Rankings")
    st.dataframe(summary_df.sort_values("Total Score (%)", ascending=False), use_container_width=True)

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
else:
    st.info("Please upload a System Criteria CSV file to begin scoring.")
