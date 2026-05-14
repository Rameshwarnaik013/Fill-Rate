import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(
    page_title="Fill Rate Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Origin lookup rules ────────────────────────────────────────────────────────
ITEM_GROUP_ORIGIN = {
    "Almonds": "Indore",
    "Roasted Almonds": "Indore",
    "Apricots": "Indore",
    "Roasted & Flavoured Cashew": "Indore",
    "Dates": "Indore",
    "Healthy Desserts": "Indore",
    "Cranberries": "Indore",
    "Trail Mixes": "Indore",
    "Seed Mix": "Indore",
    "Berry Mix": "Indore",
    "Savoury Mixes": "Indore",
    "Figs": "Indore",
    "Pistachio": "Indore",
    "Raisins": "Indore",
    "Other Seeds": "Indore",
    "Chia Seeds": "Indore",
    "Walnut": "Indore",
    "Roasted & Flavoured Makhana": "Purnia",
    "Makhana Raw": "Purnia",
}

MUNCHIES_REBELA_KEYWORDS = ["makha shaka imli waves", "makha shaka cheese waves"]


def get_origin(row):
    branch = str(row.get("Branch") or "").strip()
    item_group = str(row.get("NEW MIS ITEM GROUP") or "").strip()
    item_name = str(row.get("Item Name") or "").strip().lower()
    customer_group = str(row.get("Customer Group") or "").strip()

    if branch == "Haryana":
        return "CFA(Gurgaon)"
    if item_group == "Cashew":
        return "Indore" if customer_group == "CPC KPKB" else "Udupi"
    if item_group == "Munchies":
        if any(k in item_name for k in MUNCHIES_REBELA_KEYWORDS):
            return "Rebela"
        return "UD Foods"
    return ITEM_GROUP_ORIGIN.get(item_group, "Indore")


# ── Aggregation ────────────────────────────────────────────────────────────────
NUM_COLS = ["Stock Qty in KGS", "Delivered Qty (Kgs)", "Closed Kgs", "Pending Dispatch Kgs"]


def make_summary(df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    grp = (
        df.groupby(group_col, as_index=False)
        .agg(
            Stock=("Stock Qty in KGS", "sum"),
            Delivered=("Delivered Qty (Kgs)", "sum"),
            Closed=("Closed Kgs", "sum"),
            Pending=("Pending Dispatch Kgs", "sum"),
        )
        .sort_values("Stock", ascending=False)
    )

    with np.errstate(divide="ignore", invalid="ignore"):
        grp["fill_rate"] = np.where(grp["Stock"] > 0, grp["Delivered"] / grp["Stock"], 0.0)
        grp["closed_pct"] = np.where(grp["Stock"] > 0, grp["Closed"] / grp["Stock"], 0.0)
        grp["pen_dis_pct"] = np.where(grp["Stock"] > 0, grp["Pending"] / grp["Stock"], 0.0)

    tot = grp[["Stock", "Delivered", "Closed", "Pending"]].sum()
    total_row = {
        group_col: "Grand Total",
        "Stock": tot["Stock"],
        "Delivered": tot["Delivered"],
        "Closed": tot["Closed"],
        "Pending": tot["Pending"],
        "fill_rate": tot["Delivered"] / tot["Stock"] if tot["Stock"] else 0.0,
        "closed_pct": tot["Closed"] / tot["Stock"] if tot["Stock"] else 0.0,
        "pen_dis_pct": tot["Pending"] / tot["Stock"] if tot["Stock"] else 0.0,
    }
    grp = pd.concat([grp, pd.DataFrame([total_row])], ignore_index=True)
    return grp


# ── Color helpers (thresholds mirror the Excel heat-map in screenshots) ────────
def _fill_color(v: float) -> str:
    if v >= 0.85:
        return "background-color:#70AD47;color:#000;font-weight:bold"
    if v >= 0.70:
        return "background-color:#FFD966;color:#000;font-weight:bold"
    return "background-color:#FF6B6B;color:#000;font-weight:bold"


def _closed_color(v: float) -> str:
    if v <= 0.05:
        return "background-color:#70AD47;color:#000;font-weight:bold"
    if v <= 0.10:
        return "background-color:#FFD966;color:#000;font-weight:bold"
    return "background-color:#FF6B6B;color:#000;font-weight:bold"


def _pen_color(v: float) -> str:
    if v <= 0.05:
        return "background-color:#70AD47;color:#000;font-weight:bold"
    if v <= 0.15:
        return "background-color:#FFD966;color:#000;font-weight:bold"
    return "background-color:#FF6B6B;color:#000;font-weight:bold"


# ── Render one summary table ───────────────────────────────────────────────────
def render_table(grp: pd.DataFrame, group_col: str) -> pd.DataFrame:
    disp = grp[[group_col, "Stock", "Delivered", "Closed", "Pending",
                "fill_rate", "closed_pct", "pen_dis_pct"]].copy()
    disp.columns = [
        group_col,
        "Stock Qty in KGS", "Delivered Qty (Kgs)", "Closed Kgs", "Pending Dispatch Kgs",
        "Fill Rate %", "Closed %", "Pen Dis %",
    ]

    def _style_row(row):
        n = len(disp.columns)
        styles = [""] * n
        idx = {c: i for i, c in enumerate(disp.columns)}
        styles[idx["Fill Rate %"]] = _fill_color(row["Fill Rate %"])
        styles[idx["Closed %"]] = _closed_color(row["Closed %"])
        styles[idx["Pen Dis %"]] = _pen_color(row["Pen Dis %"])
        return styles

    def _fmt_num(x):
        return f"{x:,.0f}" if x > 0 else "-"

    def _fmt_pct(x):
        return f"{int(round(x * 100))}%"

    styled = (
        disp.style
        .apply(_style_row, axis=1)
        .format(
            {
                "Stock Qty in KGS": "{:,.0f}",
                "Delivered Qty (Kgs)": "{:,.0f}",
                "Closed Kgs": _fmt_num,
                "Pending Dispatch Kgs": _fmt_num,
                "Fill Rate %": _fmt_pct,
                "Closed %": _fmt_pct,
                "Pen Dis %": _fmt_pct,
            }
        )
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)

    # Return plain formatted df for Excel export
    export = disp.copy()
    for c in ["Stock Qty in KGS", "Delivered Qty (Kgs)"]:
        export[c] = export[c].apply(lambda x: f"{x:,.0f}")
    for c in ["Closed Kgs", "Pending Dispatch Kgs"]:
        export[c] = export[c].apply(_fmt_num)
    for c in ["Fill Rate %", "Closed %", "Pen Dis %"]:
        export[c] = export[c].apply(_fmt_pct)
    return export


# ── Excel export ───────────────────────────────────────────────────────────────
def to_excel(sheets: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            # Auto-fit columns
            for col_cells in ws.columns:
                max_len = max((len(str(c.value)) for c in col_cells if c.value), default=10)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 40)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  Main App
# ══════════════════════════════════════════════════════════════════════════════
st.title("📦 Fill Rate Dashboard")
st.caption("Upload your fill-rate Excel export to explore delivery performance by item, date, customer group and customer.")

uploaded = st.file_uploader("Upload Fill Rate Excel (.xlsx / .xls)", type=["xlsx", "xls"])

if not uploaded:
    st.info("⬆️  Upload an Excel file to begin.")
    st.stop()

# ── Load & clean ───────────────────────────────────────────────────────────────
try:
    df = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"Could not read file: {e}")
    st.stop()

missing = [c for c in NUM_COLS + ["NEW MIS ITEM GROUP"] if c not in df.columns]
if missing:
    st.error(f"Missing required columns: {missing}")
    st.stop()

for c in NUM_COLS:
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

if "Sales Order Date" in df.columns:
    df["Sales Order Date"] = pd.to_datetime(df["Sales Order Date"], errors="coerce")

df["Origin"] = df.apply(get_origin, axis=1)

# ── Sidebar filters ────────────────────────────────────────────────────────────
st.sidebar.title("🔎 Filters")

origins = ["All"] + sorted(df["Origin"].dropna().unique().tolist())
sel_origin = st.sidebar.selectbox("Origin", origins)

date_range = None
if "Sales Order Date" in df.columns:
    valid_dates = df["Sales Order Date"].dropna()
    if not valid_dates.empty:
        d_min, d_max = valid_dates.min().date(), valid_dates.max().date()
        date_range = st.sidebar.date_input(
            "SO Date Range", value=(d_min, d_max), min_value=d_min, max_value=d_max
        )

customer_groups = ["All"] + sorted(df["Customer Group"].dropna().unique().tolist())
sel_cg = st.sidebar.selectbox("Customer Group", customer_groups)

# ── Apply filters ──────────────────────────────────────────────────────────────
dff = df.copy()
if sel_origin != "All":
    dff = dff[dff["Origin"] == sel_origin]
if date_range and len(date_range) == 2:
    dff = dff[
        (dff["Sales Order Date"] >= pd.Timestamp(date_range[0]))
        & (dff["Sales Order Date"] <= pd.Timestamp(date_range[1]))
    ]
if sel_cg != "All":
    dff = dff[dff["Customer Group"] == sel_cg]

if dff.empty:
    st.warning("No data matches the selected filters.")
    st.stop()

# ── KPI cards ──────────────────────────────────────────────────────────────────
ts = dff["Stock Qty in KGS"].sum()
td = dff["Delivered Qty (Kgs)"].sum()
tc = dff["Closed Kgs"].sum()
tp = dff["Pending Dispatch Kgs"].sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Stock (KGS)", f"{ts:,.0f}")
c2.metric("Total Delivered (KGS)", f"{td:,.0f}")
c3.metric("Fill Rate", f"{int(round(td / ts * 100)) if ts else 0}%")
c4.metric("Closed %", f"{int(round(tc / ts * 100)) if ts else 0}%")
c5.metric("Pen Dis %", f"{int(round(tp / ts * 100)) if ts else 0}%")

st.divider()

# ── Legend ─────────────────────────────────────────────────────────────────────
with st.expander("Color legend", expanded=False):
    leg1, leg2, leg3 = st.columns(3)
    leg1.markdown(
        "**Fill Rate %**  \n"
        "🟩 ≥ 85%  \n🟨 70–84%  \n🟥 < 70%"
    )
    leg2.markdown(
        "**Closed %**  \n"
        "🟩 0–5%  \n🟨 6–10%  \n🟥 > 10%"
    )
    leg3.markdown(
        "**Pen Dis %**  \n"
        "🟩 0–5%  \n🟨 6–15%  \n🟥 > 15%"
    )

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["🗂 Item Group", "📅 By Date", "👥 Customer Group", "🏪 Customer", "🌍 Origin"]
)
sheets: dict = {}

with tab1:
    st.subheader("Fill Rate by NEW MIS ITEM GROUP")
    grp = make_summary(dff, "NEW MIS ITEM GROUP")
    sheets["By Item Group"] = render_table(grp, "NEW MIS ITEM GROUP")

with tab2:
    st.subheader("Fill Rate by Sales Order Date")
    df2 = dff.copy()
    df2["SO Date"] = df2["Sales Order Date"].dt.strftime("%Y-%m-%d")
    grp = make_summary(df2, "SO Date")
    # Sort date rows chronologically (Grand Total stays last)
    body = grp[grp["SO Date"] != "Grand Total"].sort_values("SO Date")
    tail = grp[grp["SO Date"] == "Grand Total"]
    grp = pd.concat([body, tail], ignore_index=True)
    sheets["By Date"] = render_table(grp, "SO Date")

with tab3:
    st.subheader("Fill Rate by Customer Group")
    grp = make_summary(dff, "Customer Group")
    sheets["By Customer Group"] = render_table(grp, "Customer Group")

with tab4:
    st.subheader("Fill Rate by Customer")
    grp = make_summary(dff, "Customer")
    sheets["By Customer"] = render_table(grp, "Customer")

with tab5:
    st.subheader("Fill Rate by Origin")
    grp = make_summary(dff, "Origin")
    sheets["By Origin"] = render_table(grp, "Origin")

# ── Download ───────────────────────────────────────────────────────────────────
st.divider()
st.download_button(
    label="📥 Download Excel Report",
    data=to_excel(sheets),
    file_name="fill_rate_dashboard.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
