
from pathlib import Path
import io
import streamlit as st
import pandas as pd

from report_generator import build_monthly_report

st.set_page_config(
    page_title="BSNL FTTH Monthly Dashboard",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).resolve().parent

st.markdown("""
<style>
.stApp {background:#f4f7fb;}
.hero {
    padding:22px 28px; border-radius:16px; margin-bottom:18px;
    background:linear-gradient(135deg,#102a43,#1f4eba,#2d7ff9);
    color:white;
}
.hero h1 {margin:0; font-size:30px;}
.hero p {margin:5px 0 0; opacity:.9;}
.kpi {
    background:white; border:1px solid #e5eaf0; border-radius:14px;
    padding:15px; text-align:center;
}
.kpi .v {font-size:26px;font-weight:800;color:#102a43;}
.kpi .l {font-size:11px;font-weight:700;color:#667085;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>📡 BSNL FTTH WARANGAL OA</h1>
<p>Monthly Total Connections Dashboard • OLT • Franchisee • BBC • DE / Manager</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("📂 Monthly Input")
    uploaded = st.file_uploader(
        "Upload Total Connections Excel",
        type=["xlsx", "xlsm", "xls", "xltx", "xltm"],
        help="Example: TOTAL CONNECTIONS AUG2026.xlsx. Filename can be anything."
    )

    st.markdown("---")
    st.header("🗂 Repository Masters")
    map_path = ROOT / "OLT_BBC_MAP.xlsx"
    master_path = ROOT / "BBC_Master.xlsx"
    st.write(("✅" if map_path.exists() else "❌") + " OLT_BBC_MAP.xlsx")
    st.write(("✅" if master_path.exists() else "❌") + " BBC_Master.xlsx")
    st.caption("Replace these two Excel files in GitHub whenever mapping changes. No Python code change is required.")

if not uploaded:
    st.info("Upload the monthly Total Connections Excel file to generate the dashboard.")
    st.markdown("""
### Expected source fields
- **OLT IP** — OLT against which the connection is working
- **FR Service Code** — franchisee/service code
- Other source columns are retained in the detailed output.

### Generated Excel workbook
- Executive Dashboard
- DE_Manager_Report
- BBC_Report
- Franchisee_Report
- OLT_Report
- OLT_BBC_Map
- BBC_Master
- Data
- Embedded chart sheets
""")
    st.stop()

if st.button("🚀 GENERATE MONTHLY EXCEL DASHBOARD", type="primary", use_container_width=True):
    with st.spinner("Reading source, applying mappings and building Excel dashboard..."):
        try:
            source_bytes = uploaded.getvalue()
            result = build_monthly_report(
                source_bytes,
                uploaded.name,
                map_path,
                master_path,
            )
            st.session_state.result = result
        except Exception as exc:
            st.error("Dashboard generation failed.")
            st.exception(exc)

if "result" not in st.session_state:
    st.stop()

r = st.session_state.result
summary = r["summary"]

cols = st.columns(6)
for c, (label, value) in zip(cols, [
    ("TOTAL CONNECTIONS", summary["total_connections"]),
    ("OLT IPs", summary["olt_count"]),
    ("FRANCHISEES", summary["franchisee_count"]),
    ("BBCs", summary["bbc_count"]),
    ("DE / MANAGERS", summary["manager_count"]),
    ("UNMAPPED OLTs", summary["unmapped_olts"]),
]):
    c.markdown(f'<div class="kpi"><div class="v">{value:,}</div><div class="l">{label}</div></div>', unsafe_allow_html=True)

st.success(f"Generated from: {uploaded.name}")

st.subheader("📊 Dashboard Preview")
st.dataframe(r["bbc_report"], use_container_width=True, hide_index=True)

c1, c2 = st.columns(2)
with c1:
    st.subheader("Connections by DE / Manager")
    st.bar_chart(r["manager_report"].set_index("DE / Manager")["Connections"])
with c2:
    st.subheader("Top BBCs by Connections")
    top = r["bbc_report"].sort_values("Connections", ascending=False).head(20)
    st.bar_chart(top.set_index("BBC")["Connections"])

st.download_button(
    "⬇️ DOWNLOAD MONTHLY EXCEL DASHBOARD",
    data=r["xlsx_bytes"],
    file_name=r["output_name"],
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

with st.expander("Mapping / Data Quality"):
    st.write(f"Mapped OLT IPs: **{summary['mapped_olts']:,}**")
    st.write(f"Unmapped OLT IPs: **{summary['unmapped_olts']:,}**")
    if r["unmapped_olt_list"]:
        st.dataframe(pd.DataFrame({"Unmapped OLT IP": r["unmapped_olt_list"]}), use_container_width=True, hide_index=True)
