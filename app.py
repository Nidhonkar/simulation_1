import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

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

def line_fig(x, y, title, xlabel, ylabel):
    fig = plt.figure(figsize=(6.0, 3.6))
    plt.plot(x, y, marker="o")
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.grid(True, alpha=0.3)
    st.pyplot(fig, clear_figure=True)

sales_df, comp_df, product_df, bottling_df, warehouse_df, supplier_df, supplier_comp_df, finance_df = load_data()

st.title("Supply Chain Simulation: KPI Dashboard")
st.markdown("Use the tabs below to view **separate KPI visuals** for each function. Each chart includes a short interpretation comment.")

tabs = st.tabs(["Sales", "Supply Chain", "Operations", "Purchasing"])

with tabs[0]:
    st.header("Sales KPIs — Separate Visuals")
    sales_rounds = sales_df.groupby("Round").mean(numeric_only=True).reset_index()
    x = sales_rounds["Round"]

    st.subheader("Service Level (Pieces)")
    line_fig(x, sales_rounds["Service level (pieces)"], "Service Level (Pieces) across Rounds", "Round (Week 0–3)", "Service Level")
    st.caption("Comment: Measures the percentage of demanded pieces delivered. A decline indicates stockouts or fulfillment delays affecting customer satisfaction.")

    st.subheader("Attained Shelf Life")
    line_fig(x, sales_rounds["Attained shelf life"], "Attained Shelf Life across Rounds", "Round (Week 0–3)", "Shelf Life")
    st.caption("Comment: Higher values imply fresher products at delivery. Improvements reflect better rotation and inventory freshness.")

    st.subheader("Order Service Accuracy (OSA)")
    line_fig(x, sales_rounds["OSA"], "Order Service Accuracy (OSA) across Rounds", "Round (Week 0–3)", "OSA")
    st.caption("Comment: Indicates accuracy and completeness of orders. Drops signal picking/stock issues or late deliveries.")

    st.subheader("Demand per Week")
    line_fig(x, sales_rounds["Demand per week"], "Demand per Week across Rounds", "Round (Week 0–3)", "Units per Week")
    st.caption("Comment: Tracks market pull. Declines can follow poor service levels as customers switch or reduce orders.")

    st.subheader("Gross Margin (Sales)")
    line_fig(x, sales_rounds["Gross margin"], "Gross Margin across Rounds", "Round (Week 0–3)", "Gross Margin")
    st.caption("Comment: Reflects pricing power and cost-to-serve. Margin compression hints at promos, mix changes, or higher COGS.")

with tabs[1]:
    st.header("Supply Chain KPIs — Separate Visuals")

    st.subheader("Lot Size (Raw Material)")
    lot_size = comp_df.groupby("Round")["Order size"].mean().reset_index()
    line_fig(lot_size["Round"], lot_size["Order size"], "Lot Size (Raw Material)", "Round (Week 0–3)", "Units per Order")
    st.caption("Comment: Larger lot sizes reduce order frequency and transport cost per unit, but increase holding risk.")

    st.subheader("Safety Stock (Raw Material) — Weeks of Cover")
    ss_rm = comp_df.groupby("Round")["Stock (weeks)"].mean().reset_index()
    line_fig(ss_rm["Round"], ss_rm["Stock (weeks)"], "Safety Stock (Raw Material)", "Round (Week 0–3)", "Weeks of Stock")
    st.caption("Comment: Buffer against supplier variability. Higher cover reduces stockouts but ties up cash.")

    st.subheader("Frozen Period (Proxy: Production Plan Adherence %)")
    fzn = bottling_df.groupby("Round")["Production plan adherence (%)"].mean().reset_index()
    line_fig(fzn["Round"], fzn["Production plan adherence (%)"], "Frozen Period (Proxy)", "Round (Week 0–3)", "Plan Adherence %")
    st.caption("Comment: Higher adherence implies a longer frozen horizon (less schedule flexibility but more stability).")

    st.subheader("Production Interval (Finished Goods) — Batches")
    pi = product_df.groupby("Round")["Production batches previous round"].mean().reset_index()
    line_fig(pi["Round"], pi["Production batches previous round"], "Production Interval (Finished Goods)", "Round (Week 0–3)", "Batches")
    st.caption("Comment: More batches mean shorter intervals and higher responsiveness, but more changeovers.")

    st.subheader("Safety Stock (Finished Goods) — Weeks of Cover")
    ss_fg = product_df.groupby("Round")["Stock (weeks)"].mean().reset_index()
    line_fig(ss_fg["Round"], ss_fg["Stock (weeks)"], "Safety Stock (Finished Goods)", "Round (Week 0–3)", "Weeks of Stock")
    st.caption("Comment: Balances service level vs. inventory cost and obsolescence risk (shelf-life).")

with tabs[2]:
    st.header("Operations KPIs — Separate Visuals")

    inbound = warehouse_df[warehouse_df["Warehouse"]=="Raw materials warehouse"].groupby("Round")["Capacity"].mean().reset_index()
    st.subheader("Size of Inbound Warehouse")
    line_fig(inbound["Round"], inbound["Capacity"], "Inbound Warehouse Capacity", "Round (Week 0–3)", "Capacity")
    st.caption("Comment: More inbound capacity supports larger procurement lots and reduces receiving bottlenecks.")

    outbound = warehouse_df[~warehouse_df["Warehouse"].isin(["Raw materials warehouse","Tank yard"])].groupby("Round")["Capacity"].mean().reset_index()
    st.subheader("Size of Outbound Warehouse")
    line_fig(outbound["Round"], outbound["Capacity"], "Outbound Warehouse Capacity", "Round (Week 0–3)", "Capacity")
    st.caption("Comment: Higher outbound storage cushions finished goods variability and supports high service levels.")

    shifts = bottling_df.groupby("Round")["Run time per week (hours)"].mean().reset_index()
    st.subheader("Number of Shifts in Bottling Plant (Approx)")
    line_fig(shifts["Round"], shifts["Run time per week (hours)"]/8.0, "Shifts (Run Time / 8h)", "Round (Week 0–3)", "Shifts")
    st.caption("Comment: Indicates labor/capacity utilization. More shifts increase throughput at higher operational cost.")

    smed = bottling_df.groupby("Round")["Changeover time per week (hours)"].mean().reset_index()
    st.subheader("SMED (Changeover Time)")
    line_fig(smed["Round"], smed["Changeover time per week (hours)"], "SMED (Changeover Time per Week)", "Round (Week 0–3)", "Hours")
    st.caption("Comment: Lower changeover hours indicate better SMED practices, enabling flexible short runs.")

    breakdown = bottling_df.groupby("Round")["Breakdown time per week (hours)"].mean().reset_index()
    st.subheader("Breakdown Trailing (Downtime)")
    line_fig(breakdown["Round"], breakdown["Breakdown time per week (hours)"], "Breakdown Time per Week", "Round (Week 0–3)", "Hours")
    st.caption("Comment: Rising breakdown time signals reliability issues; plan preventive maintenance to stabilize capacity.")

with tabs[3]:
    st.header("Purchasing KPIs — Separate Visuals")

    lead = supplier_df.groupby("Round")["Delivery reliability (%)"].mean().reset_index()
    st.subheader("Lead Time (Proxy: Delivery Reliability %)")
    line_fig(lead["Round"], lead["Delivery reliability (%)"], "Delivery Reliability %", "Round (Week 0–3)", "Reliability %")
    st.caption("Comment: Lower reliability lengthens effective lead time and increases buffer inventory needs.")

    trade = supplier_comp_df.groupby("Round")["Order size"].mean().reset_index()
    st.subheader("Trade Unit (Order Size)")
    line_fig(trade["Round"], trade["Order size"], "Average Order Size", "Round (Week 0–3)", "Units")
    st.caption("Comment: Stable order sizes support predictable inbound flow; very large orders create peaks and holding costs.")

    rej = supplier_df.groupby("Round")["Rejection  (%)"].mean().reset_index()
    st.subheader("Delivery Window (Proxy: Rejection %)")
    line_fig(rej["Round"], rej["Rejection  (%)"], "Rejection Rate %", "Round (Week 0–3)", "Rejection %")
    st.caption("Comment: Fewer rejections suggest deliveries meet quality/specs within window; high rejections disrupt production.")

    st.subheader("Supplier Mix (Purchase Value Distribution)")
    mix = supplier_df.groupby("Supplier")["Purchase  value previous round"].mean().reset_index()
    fig = plt.figure(figsize=(6.0, 3.6))
    plt.bar(mix["Supplier"], mix["Purchase  value previous round"])
    plt.title("Supplier Purchase Value (Avg)")
    plt.xlabel("Supplier")
    plt.ylabel("Average Purchase Value")
    plt.xticks(rotation=45, ha="right")
    plt.grid(axis="y", alpha=0.3)
    st.pyplot(fig, clear_figure=True)
    st.caption("Comment: Concentration on a few suppliers raises risk; consider dual-sourcing for resilience.")
