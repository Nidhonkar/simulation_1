
import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(page_title="Supply Chain Simulation Dashboard (v5)", layout="wide")

def coerce_num(s):
    return pd.to_numeric(s, errors="coerce")

def agg_series(df, col, rename_to):
    if col not in df.columns:
        return pd.DataFrame(columns=["Round","value","label"])
    tmp = df[["Round", col]].copy()
    tmp[col] = coerce_num(tmp[col])
    g = tmp.groupby("Round")[col].mean().reset_index().dropna()
    g = g.rename(columns={col:"value"})
    g["label"] = rename_to
    return g[["Round","value","label"]]

def line_chart(df, title, ytitle):
    if df.empty:
        st.warning(f"No data available for: {title}")
        return
    c = (alt.Chart(df, title=title)
         .mark_line(point=True)
         .encode(
            x=alt.X("Round:O", title="Round (week)"),
            y=alt.Y("value:Q", title=ytitle, scale=alt.Scale(zero=False)),
            tooltip=["Round","value"]
         ).properties(height=280))
    st.altair_chart(c, use_container_width=True)

@st.cache_data
def load_data():
    sales = pd.read_excel("TFC_0_3.xlsx", sheet_name="Salesarea - Customer - Product")
    comp = pd.read_excel("TFC_0_3.xlsx", sheet_name="Component")
    product = pd.read_excel("TFC_0_3.xlsx", sheet_name="Product")
    bottling = pd.read_excel("TFC_0_3.xlsx", sheet_name="Bottling line")
    wh = pd.read_excel("TFC_0_3.xlsx", sheet_name="Warehouse, Salesarea")
    supplier = pd.read_excel("TFC_0_3.xlsx", sheet_name="Supplier")
    supplier_comp = pd.read_excel("TFC_0_3.xlsx", sheet_name="Supplier - Component")
    fin = pd.read_excel("FinanceReport (5).xlsx", sheet_name="Output").rename(columns={"Unnamed: 0":"Metric"}).set_index("Metric")
    return sales, comp, product, bottling, wh, supplier, supplier_comp, fin

sales_df, comp_df, product_df, bottling_df, warehouse_df, supplier_df, supplier_comp_df, finance_df = load_data()

st.sidebar.header("Filters")
rounds_all = sorted(list(set(sales_df["Round"])))
rmin, rmax = st.sidebar.select_slider("Round range (week)", options=rounds_all, value=(min(rounds_all), max(rounds_all)))
show_tables = st.sidebar.checkbox("Show underlying tables", value=False)

def apply_round_filter(df):
    if "Round" not in df.columns: 
        return df
    return df[(df["Round"]>=rmin) & (df["Round"]<=rmax)]

st.title("Supply Chain Simulation: KPI Navigator (v5 — fixed)")
st.caption("This build standardizes KPI data and coercion so charts never go blank.")

tabs = st.tabs(["Finance","Sales","Supply Chain","Operations","Purchasing"])

# Finance
with tabs[0]:
    def fin_series(metric, label):
        if metric not in finance_df.index:
            return pd.DataFrame(columns=["Round","value","label"])
        s = pd.Series(finance_df.loc[metric].values, name="value")
        df = pd.DataFrame({"Round": list(range(len(s))), "value": coerce_num(s)})
        df["label"] = label
        return df
    rev = fin_series("Realized revenue - Contracted sales revenue","Revenue")
    roi = fin_series("ROI","ROI")
    gm  = fin_series("Gross margin","Gross Margin")
    oh  = fin_series("Operating profit - Indirect cost - Overhead costs","Overheads")
    m = min(len(gm), len(oh)) if (not gm.empty and not oh.empty) else 0
    np_df = pd.DataFrame(columns=["Round","value","label"])
    if m>0:
        np_df = pd.DataFrame({
            "Round": list(range(m)),
            "value": coerce_num(gm["value"].iloc[:m]).values - coerce_num(oh["value"].iloc[:m]).values,
            "label": ["Net Profit (GM−OH)"]*m
        })
    for d in (rev, roi, gm, oh, np_df):
        d[:] = apply_round_filter(d)
    c1,c2 = st.columns(2)
    with c1:
        line_chart(roi, "ROI", "ROI")
        st.caption("Comment: Negative ROI indicates costs outweigh returns; improve price/mix and cut waste.")
    with c2:
        line_chart(rev, "Revenue", "Currency")
        st.caption("Comment: Track whether top-line growth is sufficient to cover rising costs.")
    c1,c2 = st.columns(2)
    with c1:
        line_chart(gm, "Gross Margin", "Currency")
        st.caption("Comment: Squeezed margins often signal higher COGS or discounting.")
    with c2:
        line_chart(oh, "Operating Expense (Overheads)", "Currency")
        st.caption("Comment: Keep overheads from scaling 1:1 with revenue.")
    line_chart(np_df, "Net Profit (Gross Margin − Overheads)", "Currency")
    st.caption("Comment: Turn this curve positive via mix, pricing, SMED, and supplier reliability.")

# Sales
with tabs[1]:
    st.subheader("Sales KPIs")
    svc = agg_series(sales_df, "Service level (pieces)", "Service Level (pieces)")
    shl = agg_series(sales_df, "Attained shelf life", "Shelf Life")
    osa = agg_series(sales_df, "OSA", "OSA")
    dem = agg_series(sales_df, "Demand per week", "Demand per week")
    gms = agg_series(sales_df, "Gross margin", "Gross Margin (Sales)")
    for d in (svc, shl, osa, dem, gms):
        d[:] = apply_round_filter(d)
    c1,c2 = st.columns(2)
    with c1:
        line_chart(svc, "Service Level (pieces)", "Service Level")
        st.caption("Comment: % of demanded pieces delivered. Declines = stockouts or late fulfillment.")
    with c2:
        line_chart(shl, "Attained Shelf Life", "Shelf Life")
        st.caption("Comment: Freshness at delivery; improves with better rotation policies.")
    c1,c2 = st.columns(2)
    with c1:
        line_chart(osa, "Order Service Accuracy (OSA)", "OSA")
        st.caption("Comment: Accuracy/completeness of orders. Drops hint at picking/plan issues.")
    with c2:
        line_chart(dem, "Demand per Week", "Units")
        st.caption("Comment: Market pull; often follows service-level performance.")
    line_chart(gms, "Gross Margin (Sales)", "Currency")
    st.caption("Comment: Pricing power & mix; compression hints at promos or rising COGS.")
    if show_tables: st.dataframe(apply_round_filter(sales_df))

# Supply Chain
with tabs[2]:
    st.subheader("Supply Chain KPIs")
    lot = agg_series(comp_df, "Order size", "Lot Size (Raw)")
    ssrm = agg_series(comp_df, "Stock (weeks)", "Safety Stock (Raw)")
    fzn = agg_series(bottling_df, "Production plan adherence (%)", "Frozen (Proxy: Plan Adherence %)")
    pi  = agg_series(product_df, "Production batches previous round", "Production Interval (Batches)")
    ssfg = agg_series(product_df, "Stock (weeks)", "Safety Stock (FG)")
    for d in (lot, ssrm, fzn, pi, ssfg):
        d[:] = apply_round_filter(d)
    c1,c2 = st.columns(2)
    with c1:
        line_chart(lot, "Lot Size (Raw Material)", "Units per order")
        st.caption("Comment: Larger lots reduce ordering cost/frequency but increase holding risk.")
    with c2:
        line_chart(ssrm, "Safety Stock (Raw Material)", "Weeks of stock")
        st.caption("Comment: Buffer vs supplier variability; tune to lead-time volatility.")
    c1,c2 = st.columns(2)
    with c1:
        line_chart(fzn, "Frozen Period (Proxy: Plan Adherence %)", "Plan adherence %")
        st.caption("Comment: Higher adherence = more stability, less flexibility.")
    with c2:
        line_chart(pi, "Production Interval (Finished Goods)", "Batches")
        st.caption("Comment: More batches = shorter intervals; tradeoff with changeover time.")
    line_chart(ssfg, "Safety Stock (Finished Goods)", "Weeks of stock")
    st.caption("Comment: Balance service vs inventory cost & shelf-life risk.")
    if show_tables:
        st.dataframe(apply_round_filter(comp_df))
        st.dataframe(apply_round_filter(product_df))
        st.dataframe(apply_round_filter(bottling_df))

# Operations
with tabs[3]:
    st.subheader("Operations KPIs")
    inbound = warehouse_df[warehouse_df["Warehouse"]=="Raw materials warehouse"].copy()
    outb = warehouse_df[~warehouse_df["Warehouse"].isin(["Raw materials warehouse","Tank yard"])].copy()
    cap_in = agg_series(inbound, "Capacity", "Inbound Capacity")
    cap_out = agg_series(outb, "Capacity", "Outbound Capacity")
    shifts = agg_series(bottling_df, "Run time per week (hours)", "Shifts (run hrs / 8h)")
    if not shifts.empty:
        shifts["value"] = coerce_num(shifts["value"])/8.0
    smed = agg_series(bottling_df, "Changeover time per week (hours)", "SMED (Changeover hrs)")
    brk = agg_series(bottling_df, "Breakdown time per week (hours)", "Breakdown hrs")
    for d in (cap_in, cap_out, shifts, smed, brk):
        d[:] = apply_round_filter(d)
    c1,c2 = st.columns(2)
    with c1:
        line_chart(cap_in, "Inbound Warehouse Capacity", "Capacity")
        st.caption("Comment: More inbound capacity supports larger lots; reduces receiving bottlenecks.")
    with c2:
        line_chart(cap_out, "Outbound Warehouse Capacity", "Capacity")
        st.caption("Comment: More outbound space cushions FG variability and sustains service level.")
    c1,c2 = st.columns(2)
    with c1:
        line_chart(shifts, "Shifts (Run time / 8h)", "Shifts")
        st.caption("Comment: Throughput proxy. More shifts raise output and labor cost—watch utilization.")
    with c2:
        line_chart(smed, "SMED (Changeover Time)", "Hours")
        st.caption("Comment: Lower changeover hours = better SMED; enables flexible, short runs.")
    line_chart(brk, "Breakdown Trailing (Downtime)", "Hours")
    st.caption("Comment: Rising downtime = reliability gap; schedule preventive maintenance.")
    if show_tables:
        st.dataframe(apply_round_filter(warehouse_df))
        st.dataframe(apply_round_filter(bottling_df))

# Purchasing
with tabs[4]:
    st.subheader("Purchasing KPIs")
    lead = agg_series(supplier_df, "Delivery reliability (%)", "Delivery reliability (%)")
    trade = agg_series(supplier_comp_df, "Order size", "Order size")
    rej_col = "Rejection  (%)" if "Rejection  (%)" in supplier_df.columns else ("Rejection (%)" if "Rejection (%)" in supplier_df.columns else None)
    rej = agg_series(supplier_df, rej_col, rej_col) if rej_col else pd.DataFrame(columns=["Round","value","label"])
    for d in (lead, trade, rej):
        d[:] = apply_round_filter(d)
    c1,c2 = st.columns(2)
    with c1:
        line_chart(lead, "Lead Time (Proxy: Delivery Reliability %)", "Reliability %")
        st.caption("Comment: Lower reliability extends effective lead time; buffer stock/dual sourcing help.")
    with c2:
        line_chart(trade, "Trade Unit (Order Size)", "Units")
        st.caption("Comment: Keep order size aligned with warehouse & production cadence to avoid peaks.")
    line_chart(rej, "Delivery Window (Proxy: Rejection %)", "Rejection %")
    st.caption("Comment: Fewer rejections suggest on-time, in-spec deliveries; high rejections disrupt production.")
    if "Purchase  value previous round" in supplier_df.columns:
        mix = supplier_df.groupby("Supplier")["Purchase  value previous round"].mean().reset_index()
        mix["Purchase  value previous round"] = coerce_num(mix["Purchase  value previous round"])
        bar = (alt.Chart(mix, title="Supplier Mix (Avg Purchase Value)")
               .mark_bar()
               .encode(x=alt.X("Supplier:N", sort="-y", title="Supplier"),
                       y=alt.Y("Purchase  value previous round:Q", title="Avg Purchase Value"),
                       tooltip=["Supplier","Purchase  value previous round"])
               .properties(height=300))
        st.altair_chart(bar, use_container_width=True)

    if show_tables:
        st.dataframe(apply_round_filter(supplier_df))
        st.dataframe(apply_round_filter(supplier_comp_df))
