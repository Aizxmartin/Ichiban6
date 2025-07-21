import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
from generate_report import generate_report

st.set_page_config(page_title="Market Valuation App", layout="wide")
st.title("🏡 Market Valuation App")

uploaded_file = st.file_uploader("Upload MLS CSV/XLSX File", type=["csv", "xlsx"])
uploaded_pdf = st.file_uploader("Upload Property PDF", type="pdf")
zestimate = st.number_input("Zillow Zestimate ($)", step=1000)
redfin_estimate = st.number_input("Redfin Estimate ($)", step=1000)
est_subject_value = st.number_input("Estimated Subject Value ($)", step=1000)

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    numeric_cols = ["Close Price", "Concessions", "Above Grade Finished Area"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df.get(col), errors="coerce")

    # Generate Street Address
    df["Street Address"] = df["Street Number"].astype(str).fillna("") + " " +                            df["Street Dir Prefix"].fillna("") + " " +                            df["Street Name"].fillna("") + " " +                            df["Street Suffix"].fillna("")

    st.subheader("Filtered Comparable Properties")
    st.dataframe(df[["Street Address", "Above Grade Finished Area", "Close Price"]])

    if st.button("Generate Report"):
        subject_info = {
            "Above Grade Finished Area": 1843,
            "Bedrooms": 3,
            "Bathrooms": 2,
            "Address": "2524 S Krameria St Denver, CO 80222"
        }

        # Try to pull RealAVM if available
        if uploaded_pdf:
            pdf = PdfReader(uploaded_pdf)
            text = ""
            for page in pdf.pages:
                text += page.extract_text()
            if "RealAVM™" in text:
                import re
                match = re.search(r"RealAVM™\s*\$([\d,]+)", text)
                if match:
                    real_avm = int(match.group(1).replace(",", ""))
                    est_subject_value = real_avm

        online_vals = [v for v in [zestimate, redfin_estimate, est_subject_value] if v]
        online_avg = int(sum(online_vals) / len(online_vals)) if online_vals else None

        try:
            docx_file = generate_report(df, subject_info, online_avg)
            st.success("Report generated successfully!")
            st.download_button("Download DOCX", data=docx_file, file_name="valuation_report.docx")
        except Exception as e:
            st.error(f"Error generating report: {e}")