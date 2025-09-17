import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(page_title="Supply Chain Simulation Dashboard", layout="wide")

def to_num(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def group_mean(df, by="Round", cols=None):
    g = df.groupby(by).mean(numeric_only=True).reset_index()
    if cols is not None:
        g = g[[by] + [c for c in cols if c in g.columns]]
        g = to_num(g, cols)
    return g

def filter_rounds(df, rmin, rmax):
    return df[(df["Round"] >= rmin) & (df["Round"] <= rmax)]

def line(df, x, y, title, ytitle):
    df = df.copy()
    if y.split(":")[0].strip("`") in df.columns:
        col = y.split(":")[0].strip("`")
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return (alt.Chart(df, title=title)
            .mark_line(point=True)
            .encode(
                x=alt.X(x, title="Round (week)"),
                y=alt.Y(y, title=ytitle, scale=alt.Scale(zero=False)),
                tooltip=[x, y]
            )
            .properties(height=280))

@st.cache_data
def load_data():
    sales = pd.read_excel("TFC_0_3.xlsx", sheet_name="Salesarea - Customer - Product")
    comp = pd.read_excel("TFC_0_3.xlsx", sheet_name="Component")
    product = pd.read_excel("TFC_0_3.xlsx", sheet_name="Product")
    bottling = pd.read_excel("TFC_0_3.xlsx", sheet_name="Bottling line")
    wh = pd.read_excel("TFC_0_3.xlsx", sheet_name="Warehouse, Salesarea")
    supplier = pd.read_excel("TFC_0_3.xlsx", sheet_name="Supplier")
    supplier_comp = pd.read_excel("TFC_0_3.xlsx", sheet_name="Supplier - Component")
    fin = pd.read_excel("FinanceReport (5).xlsx", sheet_name="Output")
    fin = fin.rename(columns={"Unnamed: 0": "Metric"}).set_index("Metric")
    return sales, comp, product, bottling, wh, supplier, supplier_comp, fin

sales_df, comp_df, product_df, bottling_df, warehouse_df, supplier_df, supplier_comp_df, finance_df = load_data()
rounds_all = sorted(list(set(sales_df["Round"])))

st.sidebar.header("Filters")
rmin, rmax = st.sidebar.select_slider("Round range (week)", options=rounds_all, value=(min(rounds_all), max(rounds_all)))
show_tables = st.sidebar.checkbox("Show underlying tables", value=False)

st.title("Supply Chain Simulation: KPI Navigator (v4)")
tabs = st.tabs(["Finance", "Sales", "Supply Chain", "Operations", "Purchasing"])

# Finance
with tabs[0]:
    def get_fin_metric(key):
        if key in finance_df.index:
            s = pd.Series(finance_df.loc[key].values, name="value")
            return pd.DataFrame({"Round": list(range(len(s))), "value": pd.to_numeric(s, errors="coerce")})
        return pd.DataFrame({"Round": [], "value": []})
    fin_rev = get_fin_metric("Realized revenue - Contracted sales revenue")
    fin_roi = get_fin_metric("ROI")
    fin_gm = get_fin_metric("Gross margin")
    fin_oh = get_fin_metric("Operating profit - Indirect cost - Overhead costs")
    np_df = None
    if not fin_gm.empty and not fin_oh.empty:
        m = min(len(fin_gm), len(fin_oh))
        np_df = pd.DataFrame({
            "Round": list(range(m)),
            "value": pd.to_numeric(fin_gm["value"][:m].values, errors="coerce") - pd.to_numeric(fin_oh["value"][:m].values, errors="coerce")
        })
    for df_ in [fin_rev, fin_roi, fin_gm, fin_oh]:
        df_.query("@rmin <= Round <= @rmax", inplace=True)
    if np_df is not None:
        np_df.query("@rmin <= Round <= @rmax", inplace=True)
    cols = st.columns(2)
    with cols[0]:
        st.altair_chart(line(fin_roi, "Round:O", "value:Q", "ROI", "ROI"), use_container_width=True)
        st.caption("Comment: Negative ROI indicates costs outweigh returns; improve efficiency and cost control.")
    with cols[1]:
        st.altair_chart(line(fin_rev, "Round:O", "value:Q", "Revenue", "Currency"), use_container_width=True)
        st.caption("Comment: Revenue trend vs cost lines shows if growth can overcome rising costs.")
    cols = st.columns(2)
    with cols[0]:
        st.altair_chart(line(fin_gm, "Round:O", "value:Q", "Gross Margin", "Currency"), use_container_width=True)
        st.caption("Comment: GM reflects price × mix minus direct costs; watch compression from COGS or discounts.")
    with cols[1]:
        st.altair_chart(line(fin_oh, "Round:O", "value:Q", "Operating Expense (Overheads)", "Currency"), use_container_width=True)
        st.caption("Comment: Overheads tracking close to revenue can suppress profitability.")
    if np_df is not None:
        st.altair_chart(line(np_df, "Round:O", "value:Q", "Net Profit (Gross Margin − Overheads)", "Currency"), use_container_width=True)
        st.caption("Comment: Practical proxy for net result; turn positive via mix, pricing, and cost discipline.")
    if show_tables: st.dataframe(finance_df)

# Sales
with tabs[1]:
    sales_g = group_mean(sales_df, cols=["Service level (pieces)", "Attained shelf life", "OSA", "Demand per week", "Gross margin"])
    sales_g = filter_rounds(sales_g, rmin, rmax)
    c = st.columns(2)
    with c[0]:
        st.altair_chart(line(sales_g, "Round:O", "`Service level (pieces)`:Q", "Service Level (Pieces)", "Service Level"), use_container_width=True)
        st.caption("Comment: % delivered vs demanded. Drops flag stockouts or late fulfillment.")
    with c[1]:
        st.altair_chart(line(sales_g, "Round:O", "`Attained shelf life`:Q", "Attained Shelf Life", "Shelf life"), use_container_width=True)
        st.caption("Comment: Freshness; rising line = better rotation, fewer expiries.")
    c = st.columns(2)
    with c[0]:
        st.altair_chart(line(sales_g, "Round:O", "OSA:Q", "Order Service Accuracy (OSA)", "OSA"), use_container_width=True)
        st.caption("Comment: Accuracy/completeness; declines hint picking errors or plan mismatch.")
    with c[1]:
        st.altair_chart(line(sales_g, "Round:O", "`Demand per week`:Q", "Demand per Week", "Units"), use_container_width=True)
        st.caption("Comment: Market pull; impacted by service level and OSA.")
    st.altair_chart(line(sales_g, "Round:O", "`Gross margin`:Q", "Gross Margin (Sales)", "Currency"), use_container_width=True)
    st.caption("Comment: Pricing power & mix; compression indicates promos or rising COGS.")
    if show_tables: st.dataframe(filter_rounds(sales_df, rmin, rmax))

# Supply Chain
with tabs[2]:
    lot = group_mean(comp_df, cols=["Order size"]); lot = filter_rounds(lot, rmin, rmax)
    ssrm = group_mean(comp_df, cols=["Stock (weeks)"]); ssrm = filter_rounds(ssrm, rmin, rmax)
    fzn = group_mean(bottling_df, cols=["Production plan adherence (%)"]); fzn = filter_rounds(fzn, rmin, rmax)
    pi = group_mean(product_df, cols=["Production batches previous round"]); pi = filter_rounds(pi, rmin, rmax)
    ssfg = group_mean(product_df, cols=["Stock (weeks)"]); ssfg = filter_rounds(ssfg, rmin, rmax)
    c = st.columns(2)
    with c[0]:
        st.altair_chart(line(lot, "Round:O", "`Order size`:Q", "Lot Size (Raw Material)", "Units per order"), use_container_width=True)
        st.caption("Comment: Larger lots lower transport cost/order frequency but raise holding risk.")
    with c[1]:
        st.altair_chart(line(ssrm, "Round:O", "`Stock (weeks)`:Q", "Safety Stock (Raw Material)", "Weeks of stock"), use_container_width=True)
        st.caption("Comment: Buffer vs supplier variability; too low → stockouts, too high → tied-up cash.")
    c = st.columns(2)
    with c[0]:
        st.altair_chart(line(fzn, "Round:O", "`Production plan adherence (%)`:Q", "Frozen Period (Proxy)", "Plan adherence %"), use_container_width=True)
        st.caption("Comment: Higher adherence = longer frozen horizon (stability) but less flexibility.")
    with c[1]:
        st.altair_chart(line(pi, "Round:O", "`Production batches previous round`:Q", "Production Interval (Finished Goods)", "Batches"), use_container_width=True)
        st.caption("Comment: More batches = shorter intervals; watch changeovers.")
    st.altair_chart(line(ssfg, "Round:O", "`Stock (weeks)`:Q", "Safety Stock (Finished Goods)", "Weeks of stock"), use_container_width=True)
    st.caption("Comment: Balance service vs inventory cost & shelf-life risk.")
    if show_tables:
        st.dataframe(filter_rounds(comp_df, rmin, rmax))
        st.dataframe(filter_rounds(product_df, rmin, rmax))
        st.dataframe(filter_rounds(bottling_df, rmin, rmax))

# Operations
with tabs[3]:
    inbound = warehouse_df[warehouse_df["Warehouse"]=="Raw materials warehouse"]
    inbound = group_mean(inbound, cols=["Capacity"]); inbound = filter_rounds(inbound, rmin, rmax)
    outbound = warehouse_df[~warehouse_df["Warehouse"].isin(["Raw materials warehouse","Tank yard"])]
    outbound = group_mean(outbound, cols=["Capacity"]); outbound = filter_rounds(outbound, rmin, rmax)
    shifts = group_mean(bottling_df, cols=["Run time per week (hours)"]); shifts = filter_rounds(shifts, rmin, rmax)
    smed = group_mean(bottling_df, cols=["Changeover time per week (hours)"]); smed = filter_rounds(smed, rmin, rmax)
    brk = group_mean(bottling_df, cols=["Breakdown time per week (hours)"]); brk = filter_rounds(brk, rmin, rmax)
    c = st.columns(2)
    with c[0]:
        st.altair_chart(line(inbound, "Round:O", "Capacity:Q", "Inbound Warehouse Capacity", "Capacity"), use_container_width=True)
        st.caption("Comment: Higher inbound capacity supports larger procurement lots; fewer receiving bottlenecks.")
    with c[1]:
        st.altair_chart(line(outbound, "Round:O", "Capacity:Q", "Outbound Warehouse Capacity", "Capacity"), use_container_width=True)
        st.caption("Comment: More outbound capacity cushions FG variability and sustains service levels.")
    c = st.columns(2)
    with c[0]:
        tmp = shifts.copy(); tmp["Shifts"] = pd.to_numeric(tmp["Run time per week (hours)"], errors="coerce")/8.0
        st.altair_chart(line(tmp, "Round:O", "Shifts:Q", "Shifts (Run time / 8h)", "Shifts"), use_container_width=True)
        st.caption("Comment: Throughput proxy. More shifts raise output and costs—watch utilization.")
    with c[1]:
        st.altair_chart(line(smed, "Round:O", "`Changeover time per week (hours)`:Q", "SMED (Changeover Time)", "Hours"), use_container_width=True)
        st.caption("Comment: Lower changeover hours = better SMED; unlocks short, flexible runs.")
    st.altair_chart(line(brk, "Round:O", "`Breakdown time per week (hours)`:Q", "Breakdown Trailing (Downtime)", "Hours"), use_container_width=True)
    st.caption("Comment: Rising downtime = reliability gap; schedule preventive maintenance.")
    if show_tables:
        st.dataframe(filter_rounds(warehouse_df, rmin, rmax))
        st.dataframe(filter_rounds(bottling_df, rmin, rmax))

# Purchasing
with tabs[4]:
    lead = group_mean(supplier_df, cols=["Delivery reliability (%)"]); lead = filter_rounds(lead, rmin, rmax)
    trade = group_mean(supplier_comp_df, cols=["Order size"]); trade = filter_rounds(trade, rmin, rmax)
    rej_col = "Rejection  (%)" if "Rejection  (%)" in supplier_df.columns else "Rejection (%)"
    rej = group_mean(supplier_df, cols=[rej_col]); rej = filter_rounds(rej, rmin, rmax)
    c = st.columns(2)
    with c[0]:
        st.altair_chart(line(lead, "Round:O", "`Delivery reliability (%)`:Q", "Lead Time (Proxy: Delivery Reliability %)", "Reliability %"), use_container_width=True)
        st.caption("Comment: Lower reliability extends effective lead time; buffer stocks/dual sourcing help.")
    with c[1]:
        st.altair_chart(line(trade, "Round:O", "`Order size`:Q", "Trade Unit (Order Size)", "Units"), use_container_width=True)
        st.caption("Comment: Align order size with warehouse & production cadence to avoid peaks.")
    st.altair_chart(line(rej, "Round:O", f"`{rej_col}`:Q", "Delivery Window (Proxy: Rejection %)", "Rejection %"), use_container_width=True)
    st.caption("Comment: Fewer rejections suggest on-time, in-spec deliveries; high rejections disrupt production.")
    mix = supplier_df.groupby("Supplier")["Purchase  value previous round"].mean().reset_index()
    mix["Purchase  value previous round"] = pd.to_numeric(mix["Purchase  value previous round"], errors="coerce")
    bar = (alt.Chart(mix, title="Supplier Mix (Avg Purchase Value)")
           .mark_bar()
           .encode(x=alt.X("Supplier:N", sort="-y", title="Supplier"),
                   y=alt.Y("Purchase  value previous round:Q", title="Avg Purchase Value"),
                   tooltip=["Supplier","Purchase  value previous round"])
           .properties(height=300))
    st.altair_chart(bar, use_container_width=True)
    st.caption("Comment: Over-concentration on a single supplier increases risk; consider dual-sourcing.")
    if show_tables:
        st.dataframe(filter_rounds(supplier_df, rmin, rmax))
        st.dataframe(filter_rounds(supplier_comp_df, rmin, rmax))
