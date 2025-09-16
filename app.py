
import streamlit as st
import pandas as pd

# Load your datasets
sales_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Salesarea - Customer - Product")
comp_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Component")
product_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Product")
bottling_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Bottling line")
supplier_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Supplier")
supplier_comp_df = pd.read_excel("TFC_0_3.xlsx", sheet_name="Supplier - Component")
finance_df = pd.read_excel("FinanceReport (5).xlsx", sheet_name="Output")

# --- Clean Finance ---
finance_df = finance_df.rename(columns={"Unnamed: 0": "Metric"}).set_index("Metric")

# Tabs
tabs = st.tabs(["Finance", "Sales", "Supply Chain", "Operations", "Purchasing"])

# --- Finance Tab ---
with tabs[0]:
    st.header("Finance KPIs")
    roi = finance_df.loc["ROI"].values.astype(float)
    revenue = finance_df.loc["Realized revenue - Contracted sales revenue"].values.astype(float)
    gross_margin = finance_df.loc["Gross margin"].values.astype(float)
    operating_costs = finance_df.loc["Operating profit - Indirect cost - Overhead costs"].values.astype(float)
    net_profit = operating_costs * -1

    st.line_chart(roi, y_label="ROI")
    st.line_chart(revenue, y_label="Revenue")
    st.line_chart(gross_margin, y_label="Gross Margin")
    st.line_chart(pd.DataFrame({"Revenue": revenue, "Operating Costs": operating_costs}))
    st.line_chart(net_profit, y_label="Net Profit")

# --- Sales Tab ---
with tabs[1]:
    st.header("Sales KPIs")
    kpis = ["Service level (pieces)", "Attained shelf life", "OSA", "Demand per week", "Gross margin"]
    sales_kpi_rounds = sales_df.groupby("Round")[kpis].mean()
    st.line_chart(sales_kpi_rounds)

# --- Supply Chain Tab ---
with tabs[2]:
    st.header("Supply Chain KPIs")
    lot_size = comp_df.groupby("Round")["Order size"].mean()
    safety_stock_rm = comp_df.groupby("Round")["Stock (weeks)"].mean()
    frozen = bottling_df.groupby("Round")["Production plan adherence (%)"].mean()
    prod_interval = product_df.groupby("Round")["Production batches previous round"].mean()
    safety_stock_fg = product_df.groupby("Round")["Stock (weeks)"].mean()
    sc_df = pd.DataFrame({"Lot Size (Raw)": lot_size, "Safety Stock RM": safety_stock_rm,
                          "Frozen Period": frozen, "Prod Interval": prod_interval,
                          "Safety Stock FG": safety_stock_fg})
    st.line_chart(sc_df)

# --- Operations Tab ---
with tabs[3]:
    st.header("Operations KPIs")
    inbound = pd.read_excel("TFC_0_3.xlsx", sheet_name="Warehouse, Salesarea")
    inbound_capacity = inbound[inbound["Warehouse"]=="Raw materials warehouse"].groupby("Round")["Capacity"].mean()
    outbound_capacity = inbound[~inbound["Warehouse"].isin(["Raw materials warehouse","Tank yard"])].groupby("Round")["Capacity"].mean()
    shifts = bottling_df.groupby("Round")["Run time per week (hours)"].mean() / 8
    smed = bottling_df.groupby("Round")["Changeover time per week (hours)"].mean()
    breakdown = bottling_df.groupby("Round")["Breakdown time per week (hours)"].mean()
    ops_df = pd.DataFrame({"Inbound WH": inbound_capacity, "Outbound WH": outbound_capacity,
                           "Shifts": shifts, "SMED": smed, "Breakdown": breakdown})
    st.line_chart(ops_df)

# --- Purchasing Tab ---
with tabs[4]:
    st.header("Purchasing KPIs")
    lead_time = supplier_df.groupby("Round")["Delivery reliability (%)"].mean()
    trade_unit = supplier_comp_df.groupby("Round")["Order size"].mean()
    delivery_window = supplier_df.groupby("Round")["Rejection  (%)"].mean()
    supplier_purchase = supplier_df.groupby("Supplier")["Purchase  value previous round"].mean()
    purch_df = pd.DataFrame({"Lead Time": lead_time, "Trade Unit": trade_unit, "Delivery Window": delivery_window})
    st.line_chart(purch_df)
    st.bar_chart(supplier_purchase)
