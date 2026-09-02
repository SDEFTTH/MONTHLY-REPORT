from __future__ import annotations

import io
import tempfile
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="BSNL FTTH | Warangal OA",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# Theme
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .stApp { background: #f5f7fb; }
    [data-testid="stSidebar"] { background: #102a43; }
    [data-testid="stSidebar"] * { color: #ffffff !important; }
    .hero {
        background: linear-gradient(135deg,#102a43 0%,#1f4eba 58%,#2d7ff9 100%);
        padding: 24px 30px; border-radius: 18px; color: white;
        box-shadow: 0 8px 24px rgba(16,42,67,.18); margin-bottom: 18px;
    }
    .hero h1 { margin: 0; font-size: 30px; }
    .hero p { margin: 5px 0 0; opacity: .88; }
    .kpi {
        background: white; border: 1px solid #e6eaf0; border-radius: 14px;
        padding: 16px 18px; box-shadow: 0 4px 14px rgba(16,42,67,.06);
    }
    .kpi-label { font-size: 12px; color: #6b7280; font-weight: 700; letter-spacing: .05em; }
    .kpi-value { font-size: 27px; font-weight: 800; color: #102a43; margin-top: 4px; }
    .kpi-sub { font-size: 11px; color: #7b8794; margin-top: 2px; }
    .section { font-size: 19px; font-weight: 800; color: #102a43; margin: 18px 0 8px; }
    .quality { padding: 12px 14px; border-radius: 10px; background:#fff8e1; border:1px solid #ffe082; }
    </style>
    """,
    unsafe_allow_html=True,
)


def metric_card(label: str, value: str, sub: str = "") -> None:
    st.markdown(
        f'<div class="kpi"><div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div></div>',
        unsafe_allow_html=True,
    )


def validate_source(uploaded) -> tuple[bool, list[str], list[str]]:
    """Validate workbook structure before invoking the full processor."""
    required = {
        "BBC Name", "CLSR", "Ont Acquisition Type", "Disconnection reason",
        "Completion_Date", "Maintenance Franchisee", "OLT IP", "Order Id",
    }
    try:
        xl = pd.ExcelFile(io.BytesIO(uploaded.getvalue()))
        if "Sheet0" not in xl.sheet_names:
            return False, ["Required worksheet 'Sheet0' was not found."], []
        preview = pd.read_excel(io.BytesIO(uploaded.getvalue()), sheet_name="Sheet0", header=2, nrows=1, engine="calamine")
        headers = {str(x).strip() for x in preview.columns if x is not None}
        missing = sorted(required - headers)
        return not missing, missing, sorted(headers)
    except Exception as exc:
        return False, [f"Workbook validation failed: {exc}"], []


def read_dashboard(xlsx_bytes: bytes) -> pd.DataFrame:
    """Read the processor's executive BBM table from the generated workbook."""
    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name="FTTHDashboard", header=None)
    # Header rows 6-7 are merged/compound. Row 8 is Total OA and row 9 onward is detail.
    records = []
    for i in range(8, len(df)):
        row = df.iloc[i]
        if pd.isna(row.iloc[0]) or str(row.iloc[0]).startswith("*"):
            continue
        name = str(row.iloc[2]) if not pd.isna(row.iloc[2]) else ""
        if name in {"UNMAPPED / OTHER *", ""}:
            continue
        records.append({
            "Manager": row.iloc[1], "BBC": name, "Area": row.iloc[3],
            "OLTEs": pd.to_numeric(row.iloc[5], errors="coerce") or 0,
            "Target": pd.to_numeric(row.iloc[6], errors="coerce") or 0,
            "Daily": pd.to_numeric(row.iloc[7], errors="coerce") or 0,
            "Cumulative": pd.to_numeric(row.iloc[8], errors="coerce") or 0,
            "Achievement": pd.to_numeric(row.iloc[9], errors="coerce") or 0,
            "CLSVO": pd.to_numeric(row.iloc[10], errors="coerce") or 0,
            "CLSNP": pd.to_numeric(row.iloc[11], errors="coerce") or 0,
            "Disconnections": pd.to_numeric(row.iloc[12], errors="coerce") or 0,
            "NET": pd.to_numeric(row.iloc[13], errors="coerce") or 0,
            "NPC": pd.to_numeric(row.iloc[14], errors="coerce") or 0,
            "Reconnections": pd.to_numeric(row.iloc[15], errors="coerce") or 0,
        })
    return pd.DataFrame(records)


def read_data_sheet(xlsx_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name="Data")


# -----------------------------------------------------------------------------
# Header
# -----------------------------------------------------------------------------
st.markdown(
    '<div class="hero"><h1>📡 BSNL FTTH WARANGAL OA</h1>'
    '<p>Executive Performance Dashboard • Provisioning • Reconnections • Disconnections • NET</p></div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("## 📂 Input Report")
    uploaded = st.file_uploader(
        "Upload FTTH provisioning Excel",
        type=["xlsx", "xlsm", "xltx", "xltm"],
        help="The workbook can have any filename. The processor expects data in Sheet0 with headers on row 3.",
    )
    st.markdown("---")
    st.markdown("### Dashboard controls")
    st.caption("Filters become active after the report is generated.")
    st.markdown("---")
    st.caption("BSNL FTTH • Warangal OA")
    map_ok = Path(__file__).with_name("OLT_BBC_MAP.xlsx").exists()
    master_ok = Path(__file__).with_name("BBC_Master.xlsx").exists()
    st.markdown("### 🗂️ Repository masters")
    st.write(f"{'✅' if map_ok else '❌'} OLT_BBC_MAP.xlsx")
    st.write(f"{'✅' if master_ok else '❌'} BBC_Master.xlsx")
    st.caption("Replace these files in GitHub whenever OLT/BBC/DE/TIP/Target mapping changes. No Python code change is required.")

if not uploaded:
    st.info("Upload the latest FTTH Excel export from the left panel to start.")
    st.markdown("### What this version adds")
    st.markdown(
        "- Executive KPI cards and management-level drill-down\n"
        "- DE/Manager, BBC and Maintenance Franchisee analytics\n"
        "- Interactive Plotly charts and filters\n"
        "- Data-quality/unmapped OLT visibility\n"
        "- One-click Excel + HTML dashboard downloads\n"
        "- Safer workbook validation and clearer error reporting"
    )
    st.stop()

if "generated" not in st.session_state:
    st.session_state.generated = False

valid, missing, headers = validate_source(uploaded)
if not valid:
    st.error("The uploaded workbook is not compatible with this dashboard.")
    for item in missing:
        st.write(f"• {item}")
    with st.expander("Detected columns"):
        st.write(headers)
    st.stop()

st.success(f"Validated: **{uploaded.name}**")

if st.button("🚀 BUILD WARANGAL FTTH DASHBOARD", type="primary", use_container_width=True):
    with st.spinner("Analysing FTTH data and building executive reports…"):
        try:
            with tempfile.TemporaryDirectory() as td:
                work = Path(td)
                input_file = work / uploaded.name
                output_xlsx = work / "FTTH_Warangal_Dashboard.xlsx"
                output_html = work / "FTTH_Warangal_Dashboard.html"
                input_file.write_bytes(uploaded.getvalue())

                from monthly_report import run_report
                xlsx_path, html_path, log = run_report(input_file, output_xlsx, output_html)

                st.session_state.xlsx_bytes = xlsx_path.read_bytes()
                st.session_state.html_bytes = html_path.read_bytes()
                st.session_state.log = log
                st.session_state.source_filename = uploaded.name
                st.session_state.generated = True
                st.session_state.detail_df = read_dashboard(st.session_state.xlsx_bytes)
                st.session_state.data_df = read_data_sheet(st.session_state.xlsx_bytes)
            st.rerun()
        except Exception as exc:
            st.session_state.generated = False
            st.error("Dashboard generation failed.")
            st.exception(exc)

if not st.session_state.generated:
    st.stop()

bbm = st.session_state.detail_df.copy()
data = st.session_state.data_df.copy()

# -----------------------------------------------------------------------------
# Sidebar filters
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🔎 Filters")
    managers = sorted(bbm["Manager"].dropna().unique().tolist())
    areas = sorted(bbm["Area"].dropna().unique().tolist())
    bbc_names = sorted(bbm["BBC"].dropna().unique().tolist())
    mf_col = "Maintenance Franchisee"
    franchisees = sorted(data[mf_col].dropna().astype(str).unique().tolist()) if mf_col in data else []

    f_manager = st.multiselect("Manager / MT", managers, default=managers)
    f_area = st.multiselect("Area / TIP", areas, default=areas)
    f_bbc = st.multiselect("BBC", bbc_names, default=bbc_names)
    f_mf = st.multiselect("Maintenance Franchisee", franchisees)

filtered = bbm[bbm["Manager"].isin(f_manager) & bbm["Area"].isin(f_area) & bbm["BBC"].isin(f_bbc)].copy()

# Franchisee filter is applied to source data for the operational table only.
if f_mf and mf_col in data:
    op_data = data[data[mf_col].astype(str).isin(f_mf)].copy()
else:
    op_data = data

# -----------------------------------------------------------------------------
# KPI band
# -----------------------------------------------------------------------------
target = int(filtered["Target"].sum()) if not filtered.empty else 0
cum = int(filtered["Cumulative"].sum()) if not filtered.empty else 0
daily = int(filtered["Daily"].sum()) if not filtered.empty else 0
disc = int(filtered["Disconnections"].sum()) if not filtered.empty else 0
net = int(filtered["NET"].sum()) if not filtered.empty else 0
npc = int(filtered["NPC"].sum()) if not filtered.empty else 0
recon = int(filtered["Reconnections"].sum()) if not filtered.empty else 0
pct = (cum / target * 100) if target else 0

k = st.columns(8)
with k[0]: metric_card("MONTHLY TARGET", f"{target:,}", "Selected scope")
with k[1]: metric_card("CUMULATIVE", f"{cum:,}", "NPC + Reconnection")
with k[2]: metric_card("NPC", f"{npc:,}", "New provisions")
with k[3]: metric_card("RECONNECTIONS", f"{recon:,}", "Recovered connections")
with k[4]: metric_card("DAILY PROVISION", f"{daily:,}", "As-on provisioning date")
with k[5]: metric_card("DISCONNECTIONS", f"{disc:,}", "CLSVO + CLSNP")
with k[6]: metric_card("ACHIEVEMENT", f"{pct:.1f}%", "Against monthly target")
with k[7]: metric_card("NET", f"{net:+,}", "Cumulative − disconnections")

st.markdown('<div class="section">📊 Management Overview</div>', unsafe_allow_html=True)

# Use horizontal charts throughout so every employee/BBC label remains visible.
# The previous vertical charts squeezed names together and made the last few
# employees difficult to read on smaller Streamlit windows.
if filtered.empty:
    st.warning("No records match the selected filters.")
else:
    emp = filtered.sort_values("Cumulative", ascending=True).copy()
    c1, c2 = st.columns(2, gap="large")
    with c1:
        fig = px.bar(emp, x="Cumulative", y="BBC", orientation="h", text="Cumulative",
                     title="All BBC / Employee – Cumulative Achievement")
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_layout(height=max(480, 34 * len(emp) + 120), margin=dict(l=10,r=55,t=65,b=35),
                          xaxis_title="Connections", yaxis_title="", yaxis=dict(automargin=True))
        st.plotly_chart(fig, use_container_width=True, key="bbc_cumulative")
    with c2:
        disc_emp = emp.sort_values("Disconnections", ascending=True)
        fig = px.bar(disc_emp, x="Disconnections", y="BBC", orientation="h",
                     color_discrete_sequence=["#c62828"], text="Disconnections",
                     title="All BBC / Employee – Disconnections")
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_layout(height=max(480, 34 * len(disc_emp) + 120), margin=dict(l=10,r=55,t=65,b=35),
                          xaxis_title="Connections", yaxis_title="", yaxis=dict(automargin=True))
        st.plotly_chart(fig, use_container_width=True, key="bbc_disconnections")

    c3, c4 = st.columns(2, gap="large")
    with c3:
        net_emp = emp.sort_values("NET", ascending=True)
        fig = px.bar(net_emp, x="NET", y="BBC", orientation="h", text="NET",
                     title="All BBC / Employee – NET")
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_layout(height=max(480, 34 * len(net_emp) + 120), margin=dict(l=10,r=55,t=65,b=35),
                          xaxis_title="NET Connections", yaxis_title="", yaxis=dict(automargin=True),
                          xaxis_zeroline=True)
        st.plotly_chart(fig, use_container_width=True, key="bbc_net")
    with c4:
        ach_emp = emp.sort_values("Achievement", ascending=True)
        fig = px.bar(ach_emp, x="Achievement", y="BBC", orientation="h", text="Achievement",
                     title="All BBC / Employee – Target Achievement")
        fig.update_traces(texttemplate="%{x:.1%}", textposition="outside", cliponaxis=False)
        fig.update_layout(height=max(480, 34 * len(ach_emp) + 120), margin=dict(l=10,r=65,t=65,b=35),
                          xaxis_title="Achievement", yaxis_title="", yaxis_tickformat=".0%",
                          yaxis=dict(automargin=True))
        st.plotly_chart(fig, use_container_width=True, key="bbc_achievement")

    # Manager summary is kept separate from employee/BBC charts.
    st.markdown("#### Manager / MT Summary")
    mgr = filtered.groupby("Manager", as_index=False)[["Cumulative", "Target", "CLSVO", "CLSNP", "NET"]].sum()
    mgr["Achievement"] = mgr["Cumulative"] / mgr["Target"].replace(0, pd.NA)
    mgr = mgr.sort_values("NET", ascending=True)
    fig = px.bar(mgr, x="NET", y="Manager", orientation="h", text="NET", title="Manager / MT-wise NET")
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(height=max(300, 70 * len(mgr) + 100), margin=dict(l=10,r=55,t=65,b=30),
                      xaxis_title="NET Connections", yaxis_title="", yaxis=dict(automargin=True))
    st.plotly_chart(fig, use_container_width=True, key="manager_net")

# -----------------------------------------------------------------------------
# Operational analysis
# -----------------------------------------------------------------------------
st.markdown('<div class="section">🛠️ Operational Analysis</div>', unsafe_allow_html=True)
tab1, tab2, tab3 = st.tabs(["BBC Performance", "Franchisee Performance", "Raw / Classified Data"])

with tab1:
    view = filtered.copy()
    view["Achievement"] = view["Achievement"].map(lambda x: f"{x:.1%}")
    st.dataframe(view[["Manager","BBC","Area","OLTEs","Target","Daily","Cumulative","Achievement","CLSVO","CLSNP","NET","NPC","Reconnections"]], use_container_width=True, hide_index=True)

with tab2:
    if mf_col not in op_data.columns:
        st.warning("Maintenance Franchisee column is not available in the generated Data sheet.")
    else:
        mf = op_data.groupby(mf_col).size().reset_index(name="Orders").sort_values("Orders", ascending=False)
        mf = mf.sort_values("Orders", ascending=True)
        fig = px.bar(mf, x="Orders", y=mf_col, orientation="h", title="All Maintenance Franchisees – Classified Orders")
        fig.update_layout(height=max(500, 28 * len(mf) + 120), margin=dict(l=10,r=40,t=55,b=25), yaxis_title="", yaxis=dict(automargin=True))
        fig.update_traces(texttemplate="%{x:,}", textposition="outside", cliponaxis=False)
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(mf, use_container_width=True, hide_index=True)

with tab3:
    type_col = "Connection Type"
    if type_col in op_data.columns:
        type_summary = op_data[type_col].value_counts().rename_axis("Connection Type").reset_index(name="Count")
        st.dataframe(type_summary, use_container_width=True, hide_index=True)
    st.dataframe(op_data.head(1000), use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# Data quality + downloads
# -----------------------------------------------------------------------------
st.markdown('<div class="section">✅ Data Quality & Deliverables</div>', unsafe_allow_html=True)
q1, q2, q3 = st.columns(3)
with q1:
    metric_card("SOURCE ROWS", f"{len(data):,}", "Rows classified")
with q2:
    unmapped = 0
    if "BBC Name (rebuilt)" in data.columns:
        known = set(bbm["BBC"].dropna().astype(str))
        unmapped = int((~data["BBC Name (rebuilt)"].astype(str).isin(known)).sum())
    metric_card("UNMAPPED / OTHER", f"{unmapped:,}", "Review master mapping")
with q3:
    metric_card("FILTERED ORDERS", f"{len(op_data):,}", "Current operational scope")

if unmapped:
    st.markdown('<div class="quality">⚠️ Some rows could not be matched to the canonical BBC master. Review the <b>OLT_BBC_Map</b> and <b>BBC_Master</b> sheets in the Excel output before using the figures for final reporting.</div>', unsafe_allow_html=True)

st.markdown("### 📥 Download")
d1, d2 = st.columns(2)
with d1:
    st.download_button(
        "⬇️ Excel Executive Dashboard",
        data=st.session_state.xlsx_bytes,
        file_name="FTTH_Warangal_Dashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with d2:
    st.download_button(
        "⬇️ Standalone HTML Dashboard",
        data=st.session_state.html_bytes,
        file_name="FTTH_Warangal_Dashboard.html",
        mime="text/html",
        use_container_width=True,
    )

with st.expander("Processing log"):
    st.code(st.session_state.get("log", "Completed successfully."))
