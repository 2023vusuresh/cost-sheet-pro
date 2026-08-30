
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.formatting.rule import CellIsRule

st.set_page_config(
    page_title="Cost Sheet Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# STANDARD COST LIBRARY
# =========================
COST_LIBRARY = {
    "Direct Material": [
        "Opening Raw Material Stock", "Raw Material Purchases",
        "Closing Raw Material Stock", "Raw Materials", "Basic Materials", "Components", "Sub-assemblies",
        "Purchased Parts", "Packing Material - Direct", "Consumables - Direct",
        "Stores - Direct", "Chemicals - Direct", "Fuel - Direct",
        "Lubricants - Direct", "Tools - Direct", "Dies / Moulds - Direct",
        "Subcontracting / Job Work - Direct", "Royalties - Direct",
        "Design / Drawing - Direct", "Special Material",
        "Other Direct Material", "Custom / Other",
    ],
    "Direct Labour": [
        "Direct Wages", "Production Salaries", "Machine Operators",
        "Assembly Labour", "Skilled Labour", "Semi-skilled Labour",
        "Unskilled Labour", "Contract Labour - Direct", "Overtime Wages",
        "Production Incentives", "Piece-rate Wages", "Shift Allowance",
        "Bonus - Direct Labour", "Employee Benefits - Direct Labour",
        "Other Direct Labour", "Custom / Other",
    ],
    "Direct Expenses": [
        "Special Machine Hire", "Special Equipment Hire",
        "Special Design Charges", "Testing Charges - Direct",
        "Inspection Charges - Direct", "Pattern / Tooling Charges",
        "Special Licence / Royalty", "Job Work Charges - Direct",
        "Freight Inward - Direct", "Other Direct Expense", "Custom / Other",
    ],
    "Factory Overhead": [
        "Indirect Materials", "Indirect Stores", "Indirect Consumables",
        "Indirect Tools", "Indirect Wages", "Factory Salaries",
        "Supervisors' Salaries", "Quality Control Salaries",
        "Factory Overtime", "Factory Employee Benefits", "Power", "Electricity",
        "Fuel", "Gas", "Steam", "Water", "Diesel", "Generator Expenses",
        "Machine Running Expenses", "Repairs & Maintenance - Machinery",
        "Repairs & Maintenance - Building", "Repairs & Maintenance - Electrical",
        "Factory Cleaning", "Factory Security", "Factory Insurance",
        "Factory Rent", "Factory Rates & Taxes", "Factory Depreciation",
        "Plant & Machinery Depreciation", "Factory Building Depreciation",
        "Factory Canteen", "Factory Safety Expenses", "Factory PPE",
        "Factory Calibration", "Factory Testing", "Factory Laboratory",
        "Factory IT / Software", "Factory Telephone / Internet",
        "Material Handling", "Internal Transport", "Forklift Expenses",
        "Factory Consumables", "Factory Scrap Handling", "Waste Disposal",
        "Production Planning", "Other Factory Overhead", "Custom / Other",
    ],
    "Administration Overhead": [
        "Office Salaries", "Management Salaries", "HR Salaries",
        "Finance & Accounts Salaries", "Administration Wages",
        "Employee Benefits - Administration", "Office Rent",
        "Office Electricity", "Office Water", "Office Repairs & Maintenance",
        "Office Insurance", "Office Depreciation", "Furniture Depreciation",
        "Computer Depreciation", "Printing & Stationery", "Postage & Courier",
        "Telephone", "Internet", "Software & Subscriptions", "IT Support",
        "Professional Fees", "Audit Fees", "Legal Fees", "Consultancy Fees",
        "Accounting Fees", "Bank Charges", "Office Travel", "Staff Welfare",
        "Training & Development", "Recruitment Expenses", "Office Security",
        "Office Cleaning", "Rates & Taxes - Administration",
        "Licences & Registration", "Corporate Expenses",
        "Other Administration Overhead", "Custom / Other",
    ],
    "Selling & Distribution": [
        "Sales Salaries", "Sales Commission", "Sales Incentives",
        "Marketing Salaries", "Advertising", "Digital Marketing",
        "Promotion Expenses", "Sales Promotion", "Exhibition & Events",
        "Samples", "Catalogue / Brochure", "Market Research",
        "Customer Relationship Expenses", "Freight Outward",
        "Transportation Outward", "Delivery Charges", "Loading & Unloading",
        "Warehousing", "Warehouse Rent", "Warehouse Salaries",
        "Warehouse Electricity", "Warehouse Insurance", "Packing - Selling",
        "Distribution Expenses", "Dealer Commission", "Agent Commission",
        "Bad Debts", "Sales Returns Handling", "After-sales Service",
        "Warranty Expenses", "Customer Support", "Sales Travel",
        "Sales Office Expenses", "Other Selling Expense",
        "Other Distribution Expense", "Custom / Other",
    ],
}

COST_LIBRARY["Cost Sheet Adjustments"] = [
    "Opening Work-in-Progress (WIP)",
    "Closing Work-in-Progress (WIP)",
    "Opening Finished Goods",
    "Closing Finished Goods",
]
CATEGORIES = list(COST_LIBRARY.keys())
CAT_SHORT = {
    "Direct Material": "Direct Material",
    "Direct Labour": "Direct Labour",
    "Direct Expenses": "Direct Expenses",
    "Factory Overhead": "Factory Overhead",
    "Administration Overhead": "Administration Overhead",
    "Selling & Distribution": "Selling & Distribution",
}

# =========================
# STYLING
# =========================
st.markdown("""
<style>
    .block-container {padding-top: 1.25rem; padding-bottom: 2rem; max-width: 1500px;}
    .hero {
        padding: 28px 32px; border-radius: 20px; margin-bottom: 18px;
        background: linear-gradient(135deg, #102A43 0%, #1F5A85 100%);
        color: white; box-shadow: 0 8px 28px rgba(16,42,67,.14);
    }
    .hero h1 {font-size: 2.25rem; margin: 0; font-weight: 800;}
    .hero p {margin: 7px 0 0; color: #D9EAF7; font-size: 1rem;}
    .section-title {
        font-size: 1.2rem; font-weight: 750; color: #102A43;
        margin: 18px 0 8px;
    }
    .info-box {
        border: 1px solid #D9E2EC; border-radius: 14px; padding: 12px 15px;
        background: #F7FAFC; margin-bottom: 10px;
    }
    .small-note {color:#52606D; font-size:.86rem;}
    div[data-testid="stMetric"] {
        border: 1px solid #D9E2EC; border-radius: 14px; padding: 14px 14px;
        background: white; min-height: 92px; box-sizing: border-box;
    }
    div[data-testid="stMetricLabel"] {font-size: .88rem; line-height: 1.2;}
    div[data-testid="stMetricValue"] {font-size: 1.55rem; line-height: 1.15; white-space: nowrap;}
    button[kind="primary"] {font-weight: 700;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>📊 Cost Sheet Pro</h1>
  <p>Professional, interactive standard cost-sheet generator • component-level selection • automated costing • Excel export</p>
</div>
""", unsafe_allow_html=True)

# =========================
# SESSION STATE
# =========================
if "cost_lines" not in st.session_state:
    st.session_state.cost_lines = [
        {"Category": "Direct Labour", "Cost Component": "Direct Wages", "Qty / Units": 1000.0, "Rate": 12.0},
        {"Category": "Factory Overhead", "Cost Component": "Power", "Qty / Units": 1.0, "Rate": 8000.0},
    ]

# =========================
# SIDEBAR: PROJECT INPUTS
# =========================
with st.sidebar:
    st.header("⚙️ Cost Sheet Setup")
    company = st.text_input("Company / Organisation", "ABC Industries Ltd.")
    product = st.text_input("Product / Service", "Product / Service")
    period = st.text_input("Costing Period", "FY 2026-27")
    currency = st.text_input("Currency Symbol", "₹", max_chars=5)
    st.divider()

    st.subheader("Production")
    units = st.number_input("Units Produced", min_value=0.0, value=1000.0, step=1.0)
    st.caption("Opening/closing stock and raw-material purchases are selected as individual cost-sheet lines in the Direct Material / Cost Sheet Adjustments categories below.")


    st.divider()
    st.subheader("🎯 Profit & Pricing")

    profit_basis = st.radio(
        "Choose Profit Margin Basis",
        options=["Profit % on Cost", "Profit % on Sales"],
        horizontal=True,
        index=1,
        help="Choose exactly how the profit percentage should be interpreted."
    )

    profit_pct = st.number_input(
        "Enter Profit Margin (%)",
        min_value=0.0,
        max_value=99.99,
        value=20.0,
        step=0.5,
        format="%.2f",
        help="Example: 20 means 20%. The calculation changes automatically based on the selected basis."
    )

    if profit_basis == "Profit % on Cost":
        st.info(f"Selected: **{profit_pct:.2f}% profit on Cost of Sales**")
    else:
        st.info(f"Selected: **{profit_pct:.2f}% profit on Sales**")

# =========================
# DASHBOARD
# =========================
st.markdown('<div class="section-title">3. Live Cost Dashboard</div>', unsafe_allow_html=True)

m1 = st.columns(3)
metrics_1 = [
    ("Prime Cost", prime_cost),
    ("Works Cost", works_cost),
    ("Cost of Production", cost_of_production),
]
for col, (label, value) in zip(m1, metrics_1):
    col.metric(label, f"{currency}{value:,.2f}")

m2 = st.columns(3)
metrics_2 = [
    ("Cost of Sales", cost_of_sales),
    ("Profit", profit),
    ("Sales Value", sales_value_before_tax),
]
for col, (label, value) in zip(m2, metrics_2):
    col.metric(label, f"{currency}{value:,.2f}")

st.markdown("**Unit Economics**")
u = st.columns(3)
unit_cost = cost_of_sales / units if units else 0
unit_price = sales_value_before_tax / units if units else 0
unit_# Profit is explicitly calculated using the user's selected basis.
if profit_basis == "Profit % on Cost":
    profit = cost_of_sales * profit_pct / 100.0
    sales_value_before_tax = cost_of_sales + profit
else:
    sales_value_before_tax = cost_of_sales / (1.0 - profit_pct / 100.0)
    profit = sales_value_before_tax - cost_of_sales


