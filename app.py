
import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(page_title="Supply Chain Simulation Dashboard", layout="wide")

@st.cache_data
def load_data():
    sales_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Salesarea - Customer - Product")
    comp_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Component")
    product_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Product")
    bottling_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Bottling line")
    warehouse_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Warehouse, Salesarea")
    supplier_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Supplier")
    supplier_comp_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Supplier - Component")
    finance_df = pd.read_excel("FinanceReport (5).xlsx", sheet_name="Output")
    finance_df = finance_df.rename(columns={"Unnamed: 0": "Metric"}).set_index("Metric")
    return sales_df, comp_df, product_df, bottling_df, warehouse_df, supplier_df, supplier_comp_df, finance_df

def line_altair(df, x, y, title, ytitle):
    return (alt.Chart(df, title=title)
            .mark_line(point=True)
            .encode(
                x=alt.X(x, title="Round (week)"),
                y=alt.Y(y, title=ytitle, scale=alt.Scale(zero=False)),
                tooltip=[x, y]
            )
            .properties(height=280))

def f(df, rmin, rmax):
    return df[(df["Round"]>=rmin) & (df["Round"]<=rmax)]

sales_df, comp_df, product_df, bottling_df, warehouse_df, supplier_df, supplier_comp_df, finance_df = load_data()

st.title("Supply Chain Simulation: KPI Navigator")

rounds_all = sorted(list(set(sales_df["Round"])))
rmin, rmax = st.sidebar.select_slider("Round range (week)", options=rounds_all, value=(min(rounds_all), max(rounds_all)))
show_tables = st.sidebar.checkbox("Show underlying tables", value=False)

tabs = st.tabs(["Sales", "Supply Chain", "Operations", "Purchasing"])

# Sales
with tabs[0]:
    s = f(sales_df.groupby("Round").mean(numeric_only=True).reset_index(), rmin, rmax)
    cols = st.columns(2)
    with cols[0]:
        st.altair_chart(line_altair(s, "Round:O", "Service level (pieces):Q", "Service Level (Pieces)", "Service Level"), use_container_width=True)
        st.caption("Comment: % of pieces delivered vs demanded. Falling lines signal stockouts/late fulfillment.")
    with cols[1]:
        st.altair_chart(line_altair(s, "Round:O", "Attained shelf life:Q", "Attained Shelf Life", "Shelf Life"), use_container_width=True)
        st.caption("Comment: Freshness at delivery. Upward trend shows better rotation and lower expiry losses.")
    cols = st.columns(2)
    with cols[0]:
        st.altair_chart(line_altair(s, "Round:O", "OSA:Q", "Order Service Accuracy (OSA)", "OSA"), use_container_width=True)
        st.caption("Comment: Accuracy/completeness of orders. Drop = picking or supply alignment issues.")
    with cols[1]:
        st.altair_chart(line_altair(s, "Round:O", "Demand per week:Q", "Demand per Week", "Units"), use_container_width=True)
        st.caption("Comment: Market pull. Sustained dips often follow poor service levels.")
    st.altair_chart(line_altair(s, "Round:O", "Gross margin:Q", "Gross Margin (Sales)", "Currency"), use_container_width=True)
    st.caption("Comment: Pricing power & mix. Compression hints at promos or rising COGS.")
    if show_tables: st.dataframe(f(sales_df, rmin, rmax))

# Supply Chain
with tabs[1]:
    lot_size = f(comp_df.groupby("Round")["Order size"].mean().reset_index(), rmin, rmax)
    ss_rm = f(comp_df.groupby("Round")["Stock (weeks)"].mean().reset_index(), rmin, rmax)
    fzn = f(bottling_df.groupby("Round")["Production plan adherence (%)"].mean().reset_index(), rmin, rmax)
    pi = f(product_df.groupby("Round")["Production batches previous round"].mean().reset_index(), rmin, rmax)
    ss_fg = f(product_df.groupby("Round")["Stock (weeks)"].mean().reset_index(), rmin, rmax)
    cols = st.columns(2)
    with cols[0]:
        st.altair_chart(line_altair(lot_size, "Round:O", "Order size:Q", "Lot Size (Raw Material)", "Units per order"), use_container_width=True)
        st.caption("Comment: Bigger lots cut transport cost/order frequency but increase holding costs & risk.")
    with cols[1]:
        st.altair_chart(line_altair(ss_rm, "Round:O", "Stock (weeks):Q", "Safety Stock (Raw Material)", "Weeks of stock"), use_container_width=True)
        st.caption("Comment: Buffer vs supplier variability. Too high ties cash; too low causes stockouts.")
    cols = st.columns(2)
    with cols[0]:
        st.altair_chart(line_altair(fzn, "Round:O", "`Production plan adherence (%)`:Q", "Frozen Period (Proxy)", "Plan adherence %"), use_container_width=True)
        st.caption("Comment: Higher adherence = longer frozen horizon (stability) but less flexibility.")
    with cols[1]:
        st.altair_chart(line_altair(pi, "Round:O", "`Production batches previous round`:Q", "Production Interval (Finished Goods)", "Batches"), use_container_width=True)
        st.caption("Comment: More batches = shorter intervals/responsiveness; but more changeovers.")
    st.altair_chart(line_altair(ss_fg, "Round:O", "Stock (weeks):Q", "Safety Stock (Finished Goods)", "Weeks of stock"), use_container_width=True)
    st.caption("Comment: Balance service level vs inventory cost & shelf-life risk.")

# Operations
with tabs[2]:
    inbound = f(warehouse_df[warehouse_df["Warehouse"]=="Raw materials warehouse"].groupby("Round")["Capacity"].mean().reset_index(), rmin, rmax)
    outbound = f(warehouse_df[~warehouse_df["Warehouse"].isin(["Raw materials warehouse","Tank yard"])].groupby("Round")["Capacity"].mean().reset_index(), rmin, rmax)
    shifts = f(bottling_df.groupby("Round")["Run time per week (hours)"].mean().reset_index(), rmin, rmax)
    smed = f(bottling_df.groupby("Round")["Changeover time per week (hours)"].mean().reset_index(), rmin, rmax)
    brk = f(bottling_df.groupby("Round")["Breakdown time per week (hours)"].mean().reset_index(), rmin, rmax)
    cols = st.columns(2)
    with cols[0]:
        st.altair_chart(line_altair(inbound, "Round:O", "Capacity:Q", "Inbound Warehouse Capacity", "Capacity"), use_container_width=True)
        st.caption("Comment: Larger inbound capacity supports higher procurement lots without dock congestion.")
    with cols[1]:
        st.altair_chart(line_altair(outbound, "Round:O", "Capacity:Q", "Outbound Warehouse Capacity", "Capacity"), use_container_width=True)
        st.caption("Comment: More outbound capacity cushions FG variability and sustains service levels.")
    cols = st.columns(2)
    with cols[0]:
        tmp = shifts.copy(); tmp["Shifts"] = tmp["Run time per week (hours)"]/8.0
        st.altair_chart(line_altair(tmp, "Round:O", "Shifts:Q", "Shifts (Run time / 8h)", "Shifts"), use_container_width=True)
        st.caption("Comment: Throughput proxy. More shifts raise output and costs—watch utilization.")
    with cols[1]:
        st.altair_chart(line_altair(smed, "Round:O", "`Changeover time per week (hours)`:Q", "SMED (Changeover Time)", "Hours"), use_container_width=True)
        st.caption("Comment: Lower changeover hours = better SMED; unlocks short, flexible runs.")
    st.altair_chart(line_altair(brk, "Round:O", "`Breakdown time per week (hours)`:Q", "Breakdown Trailing (Downtime)", "Hours"), use_container_width=True)
    st.caption("Comment: Rising downtime = reliability gap; schedule preventive maintenance.")

# Purchasing
with tabs[3]:
    lead = f(supplier_df.groupby("Round")["Delivery reliability (%)"].mean().reset_index(), rmin, rmax)
    trade = f(supplier_comp_df.groupby("Round")["Order size"].mean().reset_index(), rmin, rmax)
    rej = f(supplier_df.groupby("Round")["Rejection  (%)"].mean().reset_index(), rmin, rmax)
    mix = supplier_df.groupby("Supplier")["Purchase  value previous round"].mean().reset_index()
    cols = st.columns(2)
    with cols[0]:
        st.altair_chart(line_altair(lead, "Round:O", "`Delivery reliability (%)`:Q", "Lead Time (Proxy: Delivery Reliability %)", "Reliability %"), use_container_width=True)
        st.caption("Comment: Lower reliability extends effective lead time; buffer stocks/dual sourcing help.")
    with cols[1]:
        st.altair_chart(line_altair(trade, "Round:O", "`Order size`:Q", "Trade Unit (Order Size)", "Units"), use_container_width=True)
        st.caption("Comment: Align order size with warehouse & production cadence to avoid peaks.")
    st.altair_chart(line_altair(rej, "Round:O", "`Rejection  (%)`:Q", "Delivery Window (Proxy: Rejection %)", "Rejection %"), use_container_width=True)
    st.caption("Comment: Fewer rejections suggest on-time, in-spec deliveries; high rejections disrupt production.")
    bar = (alt.Chart(mix, title="Supplier Mix (Avg Purchase Value)")
           .mark_bar()
           .encode(x=alt.X("Supplier:N", sort="-y", title="Supplier"),
                   y=alt.Y("Purchase  value previous round:Q", title="Avg Purchase Value"),
                   tooltip=["Supplier","Purchase  value previous round"])
           .properties(height=300))
    st.altair_chart(bar, use_container_width=True)
    st.caption("Comment: Over-concentration on a single supplier increases risk; consider dual-sourcing.")
