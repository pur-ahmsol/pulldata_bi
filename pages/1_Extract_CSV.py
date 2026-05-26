import zipfile
import streamlit as st
import pandas as pd

from io import BytesIO

st.set_page_config(
    page_title="Extract CSV",
    page_icon="📤",
    layout="wide"
)

st.title("📤 Extract CSV")
st.caption("Upload an Excel file and download each sheet as a separate CSV.")

# =============================================================================
# UPLOAD
# =============================================================================

uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if not uploaded:
    st.info("⬆️ Upload an Excel file to get started.")
    st.stop()

# =============================================================================
# READ ALL SHEETS
# =============================================================================

@st.cache_data(show_spinner=False)
def load_all_sheets(file):
    xl = pd.ExcelFile(file)
    sheets = {}
    skipped = []
    for name in xl.sheet_names:
        if name.strip().lower() == "readme":
            skipped.append(name)
            continue
        sheets[name] = xl.parse(name)
    return sheets, skipped

with st.spinner("Reading sheets..."):
    sheets, skipped = load_all_sheets(uploaded)

if skipped:
    st.info(f"Skipped sheet(s): **{', '.join(skipped)}** (README)")

st.success(f"Found **{len(sheets)} sheet(s)** ready to extract.")

# =============================================================================
# SUMMARY
# =============================================================================

st.subheader("Sheet Summary")

summary_rows = [
    {"Sheet": name, "Rows": len(df), "Columns": len(df.columns)}
    for name, df in sheets.items()
]
st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

# =============================================================================
# PREVIEW
# =============================================================================

st.subheader("Preview (first 10 rows per sheet)")

if sheets:
    preview_tabs = st.tabs(list(sheets.keys()))
    for tab, (name, df) in zip(preview_tabs, sheets.items()):
        with tab:
            st.dataframe(df.head(10), use_container_width=True, hide_index=True)

# =============================================================================
# DOWNLOAD
# =============================================================================

st.subheader("Download")

# Build CSV bytes for every sheet
csv_data = {name: df.to_csv(index=False).encode("utf-8") for name, df in sheets.items()}

# ── ZIP of all ────────────────────────────────────────────────────────────────
def build_zip(csv_data):
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in csv_data.items():
            safe = name.replace("/", "_").replace("\\", "_")
            zf.writestr(f"{safe}.csv", data)
    buf.seek(0)
    return buf.read()

st.markdown("**All sheets in one ZIP**")
zip_bytes = build_zip(csv_data)
base_name = uploaded.name.rsplit(".", 1)[0]
st.download_button(
    label="⬇️ Download all as ZIP",
    data=zip_bytes,
    file_name=f"{base_name}_csv_export.zip",
    mime="application/zip",
    use_container_width=True,
)

st.markdown("---")
st.markdown("**Individual sheets**")

# Rows of 4 buttons
names = list(csv_data.keys())
for i in range(0, len(names), 4):
    row_names = names[i:i+4]
    cols = st.columns(4)
    for col, name in zip(cols, row_names):
        safe = name.replace("/", "_").replace("\\", "_")
        col.download_button(
            label=f"⬇️ {name}",
            data=csv_data[name],
            file_name=f"{safe}.csv",
            mime="text/csv",
            use_container_width=True,
            key=f"dl_{name}",
        )