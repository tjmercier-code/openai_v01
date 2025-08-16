
# app.py — Streamlit Cloud safe upload handling (uses tempfile, no /mnt/data writes)
import os, json, tempfile
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from worldanalogs_tools import WorldAnalogs

st.set_page_config(page_title="World Analogs Macro Replica", layout="wide")
st.title("World Analogs – Macro Replica (Search • Extend • Plot • Histogram)")

st.markdown("""
Upload **WorldAnalogs.xls** (or .xlsx). This app replicates the Excel tools:
**Analog Search**, **Extend Selection**, **Analog Plot**, and **Analog Histogram**.
""")

uploaded = st.file_uploader("Upload WorldAnalogs.xls / .xlsx", type=["xls", "xlsx"])

# Require an upload on Streamlit Cloud
if uploaded is None:
    st.info("Please upload your WorldAnalogs.xls to begin.")
    st.stop()

# Save to a temporary file (works on Streamlit Cloud)
suffix = ".xlsx" if uploaded.name.lower().endswith("xlsx") else ".xls"
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
    tmp.write(uploaded.read())
    workbook_path = tmp.name

st.caption(f"Loaded workbook: {uploaded.name}")

@st.cache_data(show_spinner=False)
def load_world(workbook_path: str) -> WorldAnalogs:
    return WorldAnalogs.load(workbook_path)

try:
    wa = load_world(workbook_path)
    
except Exception as e:
    st.error(str(e))
    st.stop()

# --- Sidebar: Filters (Analog Search) ---
st.sidebar.header("Analog Search (Filters)")
class_vars = wa.list_classification_vars()

preferred = [
    "Structural Setting", "Architecture", "Depositional System",
    "Source Rock Depositional Environment", "Kerogen Type", "Trap Type",
    "Status", "Reservoir Rock Lithology"
]
ordered = [c for c in preferred if c in class_vars] + [c for c in class_vars if c not in preferred]

filters = {}
with st.sidebar.expander("Pick filters", expanded=True):
    for col in ordered:
        try:
            opts = sorted(pd.Series(wa.geology[col].astype(str).str.strip().unique()).dropna().tolist())
        except Exception:
            continue
        # Skip large cardinality and identity columns
        if len(opts) > 1 and len(opts) <= 50 and col not in ("AU_Code", "AU Name", "AU_Name"):
            sel = st.multiselect(col, opts, default=[])
            if sel:
                filters[col] = sel

# Apply filters
if filters:
    gsel = wa.filter_analogs(filters)
else:
    gsel = wa.geology

st.subheader("Selected Analogs")
st.write(f"{len(gsel)} assessment units match your filters.")
st.dataframe(gsel, use_container_width=True)

# Extend Selection
au_code_col = wa._find_col(wa.geology, ["AU_Code"])
selected_codes = gsel[au_code_col].astype(str).tolist()
extended = wa.extend_selection(selected_codes)

# Choose value + maturity
st.sidebar.header("Plot Settings")
sheet = st.sidebar.selectbox("Sheet (utility vars)", ["Oil", "Gas", "BOE"], index=2)

logical_choices = {
    "Number / 1000 km2 (≥5)": "number_density_gt5",
    "Number / 1000 km2 (≥50)": "number_density_gt50",
    "Median size (≥5)": "median_gt5",
    "Median size (≥50)": "median_gt50",
    "Maximum size (≥5)": "maximum_gt5",
    "Maximum size (≥50)": "maximum_gt50",
}
value_label = st.sidebar.selectbox("Utility variable", list(logical_choices.keys()), index=1)
value_logical = logical_choices[value_label]

maturity = st.sidebar.selectbox("Maturity", ["volume_gt50", "volume_gt5", "number_gt50", "number_gt5"], index=0)

# Plots
col1, col2 = st.columns(2)

with col1:
    st.markdown("#### Analog Plot (good for ≤20 AUs)")
    try:
        fig = wa.analog_plot(sheet, value_logical, maturity_mode=maturity, selected_au_codes=selected_codes)
        st.pyplot(fig, clear_figure=True)
    except Exception as e:
        st.warning(f"Analog Plot warning: {e}")

with col2:
    st.markdown("#### Analog Histogram (good for large sets)")
    try:
        fig2 = wa.analog_histogram(sheet, value_logical, selected_au_codes=selected_codes)
        st.pyplot(fig2, clear_figure=True)
    except Exception as e:
        st.warning(f"Analog Histogram warning: {e}")

# Exports (write to temp and surface links)
st.subheader("Exports")
if st.button("Export selected rows from all sheets to CSVs"):
    import tempfile, os
    out_dir = tempfile.mkdtemp(prefix="wa_exports_")
    files = wa.export_selection_csvs(selected_codes, out_dir)
    st.success(f"Wrote {len(files)} CSVs.")
    for p in files:
        st.markdown(f"- {os.path.basename(p)}")
    st.caption("Note: On Streamlit Cloud, files are ephemeral; download them from the app session.")
