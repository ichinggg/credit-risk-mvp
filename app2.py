import streamlit as st
import tempfile, os, subprocess, sys
import pandas as pd

st.set_page_config(page_title="Case Extractor — PDF → Excel", page_icon="📄", layout="wide")

st.title("📄 Case Extractor — PDF → Excel")
st.caption("Upload SSCT/STARs (property), SCCB/ACRA, and CBS PDFs to generate a clean Excel.")

left, right = st.columns([1,1])
with left:
    stars_file = st.file_uploader("① SSCT / STARs (Property) PDF", type=["pdf"], key="stars")
    sccb_file  = st.file_uploader("② SCCB / ACRA PDF", type=["pdf"], key="sccb")
with right:
    cbs_file   = st.file_uploader("③ CBS (Credit Bureau) PDF", type=["pdf"], key="cbs")

outfile = st.text_input("Output Excel filename", value="Case_Output.xlsx")
run = st.button("🚀 Run Extraction", type="primary", use_container_width=True)

def load_summary_sheet(xlsx_path: str):
    try:
        df = pd.read_excel(xlsx_path, sheet_name="Summary", header=None)
        df.columns = ["Field", "Value"]
        return df
    except Exception as e:
        st.warning(f"Could not preview Summary sheet: {e}")
        return None

if run:
    if not (stars_file and cbs_file and sccb_file):
        st.error("Please upload all three PDFs.")
        st.stop()

    with tempfile.TemporaryDirectory() as td:
        stars_path = os.path.join(td, "stars.pdf")
        cbs_path   = os.path.join(td, "cbs.pdf")
        sccb_path  = os.path.join(td, "sccb.pdf")
        with open(stars_path, "wb") as f: f.write(stars_file.read())
        with open(cbs_path, "wb") as f: f.write(cbs_file.read())
        with open(sccb_path, "wb") as f: f.write(sccb_file.read())

        out_path = os.path.join(td, outfile or "Case_Output.xlsx")

        with st.spinner("Extracting from PDFs…"):
            # Hint: main.py can optionally run adverse-news if env keys exist.
            # No keys handled here to keep UI clean.
            cmd = [sys.executable, "main.py", "--stars", stars_path, "--cbs", cbs_path, "--sccb", sccb_path, "--out", out_path, "--adverse"]
            res = subprocess.run(cmd, capture_output=True, text=True)
            if res.returncode != 0:
                st.error("Extraction failed.")
                with st.expander("Error log"):
                    st.code(res.stderr or res.stdout)
                st.stop()

        st.success("Done — Excel generated!")

        # Download button
        with open(out_path, "rb") as f:
            st.download_button("⬇️ Download Excel", f, file_name=os.path.basename(out_path),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        # Quick Preview — Summary sheet
        st.markdown("### Quick Preview — Summary")
        df = load_summary_sheet(out_path)
        if df is not None:
            for _, row in df.iterrows():
                if pd.isna(row["Field"]) and pd.isna(row["Value"]): continue
                st.markdown(f"**{str(row['Field']).strip()}**")
                v = "" if pd.isna(row["Value"]) else str(row["Value"]).strip()
                st.text(v)

st.markdown("---")
st.caption("Adverse-news: set GOOGLE_CSE_API_KEY and GOOGLE_CSE_ENGINE_ID as environment variables before running. If not set, it’s skipped automatically.")
