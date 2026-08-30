
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# COST SHEET PRO
# Standard Cost Sheet Generator
# ============================================================

st.set_page_config(
    page_title="Cost Sheet Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Cost component library
# -----------------------------
COST_LIBRARY = {
    "Direct Material": [
        "Opening Raw Material Stock",
        "Raw Material Purchases",
        "Closing Raw Material Stock",
        "Raw Materials",
        "Basic Materials",
        "Components",
        "Sub-assemblies",
        "Purchased Parts",
        "Packing Material - Direct",
        "Consumables - Direct",
        "Stores - Direct",
        "Chemicals - Direct",
        "Fuel - Direct",
        "Lubricants - Direct",
        "Tools - Direct",
        "Dies / Moulds - Direct",
        "Subcontracting / Job Work - Direct",
        "Royalties - Direct",
        "Design / Drawing - Direct",
        "Special Material",
        "Other Direct Material",
    ],
    "Direct Labour": [
        "Direct Wages",
        "Production Salaries",
        "Machine Operators",
        "Assembly Labour",
        "Skilled Labour",
        "Semi-skilled Labour",
        "Unskilled Labour",
        "Contract Labour - Direct",
        "Overtime Wages",
        "Production Incentives",
        "Piece-rate Wages",
        "Shift Allowance",
        "Bonus - Direct Labour",
        "Employee Benefits - Direct Labour",
        "Other Direct Labour",
    ],
    "Direct Expenses": [
        "Special Machine Hire",
        "Special Equipment Hire",
        "Special Design Charges",
        "Testing Charges - Direct",
        "Inspection Charges - Direct",
        "Pattern / Tooling Charges",
        "Special Licence / Royalty",
        "Job Work Charges - Direct",
        "Other Direct Expense",
    ],
    "Factory Overhead": [
        "Indirect Materials",
        "Indirect Stores",
        "Indirect Consumables",
        "Indirect Tools",
        "Indirect Wages",
        "Factory Salaries",
        "Supervisors' Salaries",
        "Quality Control Salaries",
        "Factory Overtime",
        "Factory Employee Benefits",
        "Power",
        "Electricity",
        "Fuel",
        "Gas",
        "Steam",
        "Water",
        "Diesel",
        "Generator Expenses",
        "Machine Running Expenses",
        "Repairs & Maintenance - Machinery",
        "Repairs & Maintenance - Building",
        "Repairs & Maintenance - Electrical",
        "Factory Cleaning",
        "Factory Security",
        "Factory Insurance",
        "Factory Rent",
        "Factory Rates & Taxes",
        "Factory Depreciation",
        "Plant & Machinery Depreciation",
        "Factory Building Depreciation",
        "Factory Canteen",
        "Factory Safety Expenses",
        "Factory PPE",
        "Factory Calibration",
        "Factory Testing",
        "Factory Laboratory",
        "Factory IT / Software",
        "Factory Telephone / Internet",
        "Material Handling",
        "Internal Transport",
        "Forklift Expenses",
        "Factory Consumables",
        "Factory Scrap Handling",
        "Waste Disposal",
        "Production Planning",
        "Other Factory Overhead",
    ],
    "Administration Overhead": [
        "Office Salaries",
        "Management Salaries",
        "HR Salaries",
        "Finance & Accounts Salaries",
        "Administration Wages",
        "Employee Benefits - Administration",
        "Office Rent",
        "Office Electricity",
        "Office Water",
        "Office Repairs & Maintenance",
        "Office Insurance",
        "Office Depreciation",
        "Furniture Depreciation",
        "Computer Depreciation",
        "Printing & Stationery",
        "Postage & Courier",
        "Telephone",
        "Internet",
        "Software & Subscriptions",
        "IT Support",
        "Professional Fees",
        "Audit Fees",
        "Legal Fees",
        "Consultancy Fees",
        "Accounting Fees",
        "Bank Charges",
        "Office Travel",
        "Staff Welfare",
        "Training & Development",
        "Recruitment Expenses",
        "Office Security",
        "Office Cleaning",
        "Rates & Taxes - Administration",
        "Licences & Registration",
        "Corporate Expenses",
        "Other Administration Overhead",
    ],
    "Selling & Distribution": [
        "Sales Salaries",
        "Sales Commission",
        "Sales Incentives",
        "Marketing Salaries",
        "Advertising",
        "Digital Marketing",
        "Promotion Expenses",
        "Sales Promotion",
        "Exhibition & Events",
        "Samples",
        "Catalogue / Brochure",
        "Market Research",
        "Customer Relationship Expenses",
        "Freight Outward",
        "Transportation Outward",
        "Delivery Charges",
        "Loading & Unloading",
        "Warehousing",
        "Warehouse Rent",
        "Warehouse Salaries",
        "Warehouse Electricity",
        "Warehouse Insurance",
        "Packing - Selling",
        "Distribution Expenses",
        "Dealer Commission",
        "Agent Commission",
        "Bad Debts",
        "Sales Returns Handling",
        "After-sales Service",
        "Warranty Expenses",
        "Customer Support",
        "Sales Travel",
        "Sales Office Expenses",
        "Other Selling Expense",
        "Other Distribution Expense",
    ],
    "Cost Sheet Adjustments": [
        "Opening Work-in-Progress (WIP)",
        "Closing Work-in-Progress (WIP)",
        "Opening Finished Goods",
        "Closing Finished Goods",
    ],
}

CATEGORIES = list(COST_LIBRARY.keys())

# -----------------------------
# Styling
# -----------------------------
st.markdown(
    """
    <style>
    .block-container {
        max-width: 1450px;
        padding-top: 1.15rem;
        padding-bottom: 3rem;
    }

    .hero {
        background: linear-gradient(135deg, #0B2545 0%, #174E73 55%, #287D9E 100%);
        color: white;
        border-radius: 22px;
        padding: 28px 32px;
        margin-bottom: 22px;
        box-shadow: 0 10px 30px rgba(11,37,69,.14);
    }

    .hero h1 {
        margin: 0;
        font-size: 2.55rem;
        font-weight: 800;
        letter-spacing: -.03em;
    }

    .hero p {
        margin: 7px 0 0;
        color: #E9F4FA;
        font-size: 1rem;
    }

    .section-title {
        color: #102A43;
        font-size: 1.25rem;
        font-weight: 800;
        margin: 25px 0 9px;
    }

    .info-box {
        border: 1px solid #D9E2EC;
        background: #F8FBFD;
        border-radius: 14px;
        padding: 13px 16px;
        margin-bottom: 16px;
        color: #334E68;
    }

    .entry-card {
        border: 1px solid #D9E2EC;
        background: #FFFFFF;
        border-radius: 17px;
        padding: 16px 16px 12px;
        box-shadow: 0 4px 16px rgba(16,42,67,.055);
        margin-bottom: 14px;
    }

    div[data-testid="stMetric"] {
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        background: #FFFFFF;
        padding: 14px 16px;
        min-height: 94px;
        box-shadow: 0 3px 12px rgba(16,42,67,.045);
    }

    div[data-testid="stMetricLabel"] {
        white-space: normal !important;
    }

    div[data-testid="stMetricValue"] {
        font-size: 1.65rem !important;
        line-height: 1.2 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }

    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div {
        border-radius: 10px;
    }

    .stButton button {
        border-radius: 10px;
        min-height: 42px;
        font-weight: 700;
    }

    [data-testid="stSidebar"] {
        border-right: 1px solid #E1E8EF;
    }

    .summary-note {
        color: #52606D;
        font-size: .88rem;
    }

    @media (max-width: 1050px) {
        .hero h1 {font-size: 2.05rem;}
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1>📊 Cost Sheet Pro</h1>
        <p>Professional standard cost-sheet generator • component-level input • automatic rate conversion • Excel export</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Session state
# -----------------------------
if "cost_lines" not in st.session_state:
    st.session_state.cost_lines = []

# -----------------------------
# Sidebar: only genuine setup
# No duplicate inventory/cost inputs
# -----------------------------
with st.sidebar:
    st.header("⚙️ Cost Sheet Setup")
    st.caption("Basic details and profit assumptions.")

    company = st.text_input("Company / Organisation", "Suresh Inc", key="company")
    product = st.text_input("Product / Service", "Premium Office Chairs", key="product")
    period = st.text_input("Costing Period", "May 2026", key="period")
    currency = st.text_input("Currency Symbol", "₹", max_chars=5, key="currency")

    st.divider()

    units = st.number_input(
        "Units Produced",
        min_value=0.0,
        value=1000.0,
        step=1.0,
        format="%.2f",
        key="units",
    )

    st.subheader("Profit Assumption")
    profit_basis = st.selectbox(
        "Profit calculated as",
        ["Profit % on Sales", "Profit % on Cost of Sales"],
        key="profit_basis",
    )
    profit_margin = st.number_input(
        "Profit Margin %",
        min_value=0.0,
        max_value=99.99,
        value=20.0,
        step=0.5,
        format="%.2f",
        key="profit_margin",
    )

# -----------------------------
# Helper functions
# -----------------------------
def money(value):
    return f"{currency}{value:,.2f}"

def safe_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0

def aggregate_components(lines):
    totals = {
        "Direct Material": 0.0,
        "Direct Labour": 0.0,
        "Direct Expenses": 0.0,
        "Factory Overhead": 0.0,
        "Administration Overhead": 0.0,
        "Selling & Distribution": 0.0,
        "Opening RM": 0.0,
        "RM Purchases": 0.0,
        "Closing RM": 0.0,
        "Opening WIP": 0.0,
        "Closing WIP": 0.0,
        "Opening FG": 0.0,
        "Closing FG": 0.0,
    }

    for line in lines:
        category = line["Category"]
        component = line["Cost Component"]
        amount = safe_float(line["Total Amount"])

        if category == "Direct Material":
            if component == "Opening Raw Material Stock":
                totals["Opening RM"] += amount
            elif component == "Raw Material Purchases":
                totals["RM Purchases"] += amount
            elif component == "Closing Raw Material Stock":
                totals["Closing RM"] += amount
            else:
                totals["Direct Material"] += amount

        elif category == "Direct Labour":
            totals["Direct Labour"] += amount

        elif category == "Direct Expenses":
            totals["Direct Expenses"] += amount

        elif category == "Factory Overhead":
            totals["Factory Overhead"] += amount

        elif category == "Administration Overhead":
            totals["Administration Overhead"] += amount

        elif category == "Selling & Distribution":
            totals["Selling & Distribution"] += amount

        elif category == "Cost Sheet Adjustments":
            if component == "Opening Work-in-Progress (WIP)":
                totals["Opening WIP"] += amount
            elif component == "Closing Work-in-Progress (WIP)":
                totals["Closing WIP"] += amount
            elif component == "Opening Finished Goods":
                totals["Opening FG"] += amount
            elif component == "Closing Finished Goods":
                totals["Closing FG"] += amount

    raw_material_consumed = (
        totals["Opening RM"] + totals["RM Purchases"] - totals["Closing RM"]
    )
    direct_material = raw_material_consumed + totals["Direct Material"]
    prime_cost = (
        direct_material + totals["Direct Labour"] + totals["Direct Expenses"]
    )
    works_cost = (
        prime_cost
        + totals["Factory Overhead"]
        + totals["Opening WIP"]
        - totals["Closing WIP"]
    )
    cost_of_production = works_cost + totals["Administration Overhead"]
    cost_of_goods_sold = (
        cost_of_production + totals["Opening FG"] - totals["Closing FG"]
    )
    cost_of_sales = cost_of_goods_sold + totals["Selling & Distribution"]

    if profit_basis == "Profit % on Sales":
        # Profit = margin * Sales
        # Sales = Cost / (1 - margin)
        margin = profit_margin / 100.0
        if margin >= 1:
            profit = 0.0
            sales_value = cost_of_sales
        else:
            sales_value = cost_of_sales / (1.0 - margin)
            profit = sales_value - cost_of_sales
    else:
        # Profit = margin * Cost of Sales
        profit = cost_of_sales * (profit_margin / 100.0)
        sales_value = cost_of_sales + profit

    return {
        **totals,
        "Raw Material Consumed": raw_material_consumed,
        "Direct Material": direct_material,
        "Prime Cost": prime_cost,
        "Works Cost": works_cost,
        "Cost of Production": cost_of_production,
        "Cost of Goods Sold": cost_of_goods_sold,
        "Cost of Sales": cost_of_sales,
        "Profit": profit,
        "Sales Value": sales_value,
    }

# -----------------------------
# Add individual component
# -----------------------------
st.markdown(
    '<div class="section-title">1. Add Individual Cost Components</div>',
    unsafe_allow_html=True,
)
st.markdown(
    """
    <div class="info-box">
        <b>Enter the figures exactly as given in the question.</b>
        Choose <b>Total Amount</b> when the question gives a total amount, or
        <b>Rate / Unit</b> when it gives a per-unit rate. The other value is
        calculated automatically.
    </div>
    """,
    unsafe_allow_html=True,
)

with st.container():
    entry_cols = st.columns([1.05, 2.05, 0.95, 1.15], gap="medium")

    with entry_cols[0]:
        add_category = st.selectbox(
            "Category",
            CATEGORIES,
            key="add_category",
        )

    with entry_cols[1]:
        add_component = st.selectbox(
            "Cost Component",
            COST_LIBRARY[add_category],
            key=f"add_component_{add_category}",
        )

    with entry_cols[2]:
        add_qty = st.number_input(
            "Total Qty / Units",
            min_value=0.0,
            value=1.0,
            step=1.0,
            format="%.2f",
            key="add_qty",
        )

    with entry_cols[3]:
        add_basis = st.selectbox(
            "Enter By",
            ["Total Amount", "Rate / Unit"],
            key="add_basis",
        )

# Separate calculation row: deliberately not squeezed into the first row.
calc_cols = st.columns([1.35, 1.35, 1.35, 1.0], gap="medium")

if add_basis == "Total Amount":
    with calc_cols[0]:
        entered_amount = st.number_input(
            "Total Amount",
            min_value=0.0,
            value=0.0,
            step=100.0,
            format="%.2f",
            key="entered_total",
        )
    add_total = safe_float(entered_amount)
    add_rate = add_total / add_qty if add_qty > 0 else 0.0

    with calc_cols[1]:
        st.text_input(
            "Rate / Unit (calculated)",
            value=f"{add_rate:,.2f}",
            disabled=True,
            key="display_rate",
        )
else:
    with calc_cols[0]:
        entered_rate = st.number_input(
            "Rate / Unit",
            min_value=0.0,
            value=0.0,
            step=1.0,
            format="%.2f",
            key="entered_rate",
        )
    add_rate = safe_float(entered_rate)
    add_total = add_qty * add_rate

    with calc_cols[1]:
        st.text_input(
            "Total Amount (calculated)",
            value=f"{add_total:,.2f}",
            disabled=True,
            key="display_total",
        )

with calc_cols[2]:
    st.metric("Amount to Add", money(add_total))

with calc_cols[3]:
    st.write("")
    if st.button("➕ Add Cost", type="primary", use_container_width=True):
        if add_qty <= 0:
            st.error("Total Qty / Units must be greater than 0.")
        elif add_total < 0:
            st.error("Total Amount cannot be negative.")
        else:
            st.session_state.cost_lines.append(
                {
                    "Category": add_category,
                    "Cost Component": add_component,
                    "Qty / Units": float(add_qty),
                    "Rate / Unit": float(add_rate),
                    "Total Amount": float(add_total),
                    "Input Basis": add_basis,
                }
            )
            st.rerun()

# -----------------------------
# Custom component
# -----------------------------
with st.expander("✏️ Add a Custom Cost Component", expanded=False):
    c1, c2, c3, c4 = st.columns([1.05, 2.05, 0.95, 1.15], gap="medium")

    with c1:
        custom_category = st.selectbox(
            "Category",
            CATEGORIES,
            key="custom_category",
        )
    with c2:
        custom_name = st.text_input(
            "Custom Cost Component",
            placeholder="Enter component name",
            key="custom_name",
        )
    with c3:
        custom_qty = st.number_input(
            "Total Qty / Units",
            min_value=0.0,
            value=1.0,
            step=1.0,
            format="%.2f",
            key="custom_qty",
        )
    with c4:
        custom_basis = st.selectbox(
            "Enter By",
            ["Total Amount", "Rate / Unit"],
            key="custom_basis",
        )

    cc1, cc2, cc3 = st.columns([1.35, 1.35, 1.0], gap="medium")

    if custom_basis == "Total Amount":
        with cc1:
            custom_entered = st.number_input(
                "Total Amount",
                min_value=0.0,
                value=0.0,
                step=100.0,
                format="%.2f",
                key="custom_entered_total",
            )
        custom_total = safe_float(custom_entered)
        custom_rate = custom_total / custom_qty if custom_qty > 0 else 0.0
        with cc2:
            st.text_input(
                "Rate / Unit (calculated)",
                value=f"{custom_rate:,.2f}",
                disabled=True,
                key="custom_rate_display",
            )
    else:
        with cc1:
            custom_entered = st.number_input(
                "Rate / Unit",
                min_value=0.0,
                value=0.0,
                step=1.0,
                format="%.2f",
                key="custom_entered_rate",
            )
        custom_rate = safe_float(custom_entered)
        custom_total = custom_qty * custom_rate
        with cc2:
            st.text_input(
                "Total Amount (calculated)",
                value=f"{custom_total:,.2f}",
                disabled=True,
                key="custom_total_display",
            )

    with cc3:
        if st.button("Add Custom Cost", use_container_width=True):
            if not custom_name.strip():
                st.error("Enter a custom component name.")
            elif custom_qty <= 0:
                st.error("Total Qty / Units must be greater than 0.")
            else:
                st.session_state.cost_lines.append(
                    {
                        "Category": custom_category,
                        "Cost Component": custom_name.strip(),
                        "Qty / Units": float(custom_qty),
                        "Rate / Unit": float(custom_rate),
                        "Total Amount": float(custom_total),
                        "Input Basis": custom_basis,
                    }
                )
                st.rerun()

# -----------------------------
# Current entries
# -----------------------------
st.markdown(
    '<div class="section-title">2. Cost Components Entered</div>',
    unsafe_allow_html=True,
)

if st.session_state.cost_lines:
    display_rows = []
    for i, line in enumerate(st.session_state.cost_lines, start=1):
        display_rows.append(
            {
                "No.": i,
                "Category": line["Category"],
                "Cost Component": line["Cost Component"],
                "Qty / Units": line["Qty / Units"],
                "Rate / Unit": line["Rate / Unit"],
                "Total Amount": line["Total Amount"],
                "Input Basis": line["Input Basis"],
            }
        )

    st.dataframe(
        display_rows,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Qty / Units": st.column_config.NumberColumn(format="%.2f"),
            "Rate / Unit": st.column_config.NumberColumn(format="%.2f"),
            "Total Amount": st.column_config.NumberColumn(format="%.2f"),
        },
    )

    remove_cols = st.columns([3, 1, 1])
    with remove_cols[1]:
        remove_no = st.number_input(
            "Remove No.",
            min_value=1,
            max_value=len(st.session_state.cost_lines),
            value=1,
            step=1,
            key="remove_no",
        )
    with remove_cols[2]:
        st.write("")
        if st.button("🗑️ Remove", use_container_width=True):
            st.session_state.cost_lines.pop(int(remove_no) - 1)
            st.rerun()
else:
    st.info("No cost components have been added yet.")

# -----------------------------
# Live standard cost summary
# -----------------------------
results = aggregate_components(st.session_state.cost_lines)

st.markdown(
    '<div class="section-title">3. Standard Cost Summary</div>',
    unsafe_allow_html=True,
)

summary_rows = [
    ("OPENING RAW MATERIAL STOCK", results["Opening RM"], "normal"),
    ("ADD: RAW MATERIAL PURCHASES", results["RM Purchases"], "normal"),
    ("LESS: CLOSING RAW MATERIAL STOCK", -results["Closing RM"], "normal"),
    ("RAW MATERIAL CONSUMED", results["Raw Material Consumed"], "subtotal"),
    ("OTHER DIRECT MATERIAL", results["Direct Material"], "normal"),
    ("DIRECT MATERIAL", results["Direct Material"], "subtotal"),
    ("DIRECT LABOUR", results["Direct Labour"], "normal"),
    ("DIRECT EXPENSES", results["Direct Expenses"], "normal"),
    ("PRIME COST", results["Prime Cost"], "total"),
    ("FACTORY / WORKS OVERHEAD", results["Factory Overhead"], "normal"),
    ("ADD: OPENING WIP", results["Opening WIP"], "normal"),
    ("LESS: CLOSING WIP", -results["Closing WIP"], "normal"),
    ("WORKS / FACTORY COST", results["Works Cost"], "subtotal"),
    ("ADMINISTRATION OVERHEAD", results["Administration Overhead"], "normal"),
    ("COST OF PRODUCTION", results["Cost of Production"], "subtotal"),
    ("ADD: OPENING FINISHED GOODS", results["Opening FG"], "normal"),
    ("LESS: CLOSING FINISHED GOODS", -results["Closing FG"], "normal"),
    ("COST OF GOODS SOLD", results["Cost of Goods Sold"], "subtotal"),
    ("SELLING & DISTRIBUTION OVERHEAD", results["Selling & Distribution"], "normal"),
    ("COST OF SALES", results["Cost of Sales"], "total"),
]

sales_value = results["Sales Value"]
cost_of_sales = results["Cost of Sales"]
profit = results["Profit"]

summary_display = []
for label, amount, kind in summary_rows:
    pct_sales = (amount / sales_value) if sales_value else 0.0
    summary_display.append(
        {
            "Cost Component": label,
            "Amount": amount,
            "% of Sales": pct_sales,
        }
    )

summary_display.append(
    {
        "Cost Component": (
            f"ADD: PROFIT — {profit_margin:.2f}% ON "
            + ("SALES" if profit_basis == "Profit % on Sales" else "COST OF SALES")
        ),
        "Amount": profit,
        "% of Sales": (profit / sales_value) if sales_value else 0.0,
    }
)
summary_display.append(
    {
        "Cost Component": "SALES VALUE",
        "Amount": sales_value,
        "% of Sales": 1.0 if sales_value else 0.0,
    }
)

st.dataframe(
    summary_display,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Amount": st.column_config.NumberColumn(format="%.2f"),
        "% of Sales": st.column_config.NumberColumn(format="%.1f%%"),
    },
)

m1, m2, m3, m4, m5 = st.columns(5, gap="medium")
with m1:
    st.metric("Prime Cost", money(results["Prime Cost"]))
with m2:
    st.metric("Works Cost", money(results["Works Cost"]))
with m3:
    st.metric("Cost of Production", money(results["Cost of Production"]))
with m4:
    st.metric("Cost of Sales", money(results["Cost of Sales"]))
with m5:
    st.metric("Sales Value", money(results["Sales Value"]))

p1, p2, p3 = st.columns(3, gap="medium")
with p1:
    st.metric("Profit", money(profit))
with p2:
    st.metric("Cost / Unit", money(cost_of_sales / units if units else 0.0))
with p3:
    st.metric("Selling Price / Unit", money(sales_value / units if units else 0.0))

st.caption(
    f"Profit basis: {profit_basis} • Margin selected: {profit_margin:.2f}% • "
    f"Units produced: {units:,.2f}"
)

# -----------------------------
# Excel export
# -----------------------------
def build_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Cost Sheet"

    navy = "0B2545"
    blue = "174E73"
    light_blue = "EAF2F8"
    very_light = "F7FAFC"
    white = "FFFFFF"
    dark = "102A43"
    border_color = "B8C7D9"
    total_fill = "DCEAF4"
    profit_fill = "E7F6EA"

    thin = Side(style="thin", color=border_color)
    medium = Side(style="medium", color=navy)
    thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title
    ws.merge_cells("A1:F1")
    ws["A1"] = company
    ws["A1"].font = Font(bold=True, size=16, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:F2")
    ws["A2"] = "STANDARD COST SHEET"
    ws["A2"].font = Font(bold=True, size=13, color=white)
    ws["A2"].fill = PatternFill("solid", fgColor=blue)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A3:F3")
    ws["A3"] = f"{product}  |  {period}"
    ws["A3"].font = Font(bold=True, size=11, color=dark)
    ws["A3"].fill = PatternFill("solid", fgColor=light_blue)
    ws["A3"].alignment = Alignment(horizontal="left", vertical="center")

    headers = ["No.", "Cost Component", "Qty / Units", "Rate / Unit", "Amount", "% of Sales"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=5, column=col, value=header)
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    row = 6
    line_no = 1

    # Detail components
    for line in st.session_state.cost_lines:
        vals = [
            line_no,
            line["Cost Component"],
            line["Qty / Units"],
            line["Rate / Unit"],
            line["Total Amount"],
            (line["Total Amount"] / sales_value) if sales_value else 0.0,
        ]
        for col, value in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=value)
            cell.border = thin_border
            cell.alignment = Alignment(
                horizontal="right" if col in (1, 3, 4, 5, 6) else "left",
                vertical="center",
            )
        line_no += 1
        row += 1

    # Summary section
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    ws.cell(row=row, column=1, value="STANDARD COST SUMMARY")
    ws.cell(row=row, column=1).font = Font(bold=True, color=white)
    ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor=blue)
    ws.cell(row=row, column=1).alignment = Alignment(horizontal="left")
    row += 1

    summary_start = row
    summary_excel = [
        ("OPENING RAW MATERIAL STOCK", results["Opening RM"], "normal"),
        ("ADD: RAW MATERIAL PURCHASES", results["RM Purchases"], "normal"),
        ("LESS: CLOSING RAW MATERIAL STOCK", -results["Closing RM"], "normal"),
        ("RAW MATERIAL CONSUMED", results["Raw Material Consumed"], "subtotal"),
        ("OTHER DIRECT MATERIAL", results["Direct Material"], "normal"),
        ("DIRECT MATERIAL", results["Direct Material"], "subtotal"),
        ("DIRECT LABOUR", results["Direct Labour"], "normal"),
        ("DIRECT EXPENSES", results["Direct Expenses"], "normal"),
        ("PRIME COST", results["Prime Cost"], "total"),
        ("FACTORY / WORKS OVERHEAD", results["Factory Overhead"], "normal"),
        ("ADD: OPENING WIP", results["Opening WIP"], "normal"),
        ("LESS: CLOSING WIP", -results["Closing WIP"], "normal"),
        ("WORKS / FACTORY COST", results["Works Cost"], "subtotal"),
        ("ADMINISTRATION OVERHEAD", results["Administration Overhead"], "normal"),
        ("COST OF PRODUCTION", results["Cost of Production"], "subtotal"),
        ("ADD: OPENING FINISHED GOODS", results["Opening FG"], "normal"),
        ("LESS: CLOSING FINISHED GOODS", -results["Closing FG"], "normal"),
        ("COST OF GOODS SOLD", results["Cost of Goods Sold"], "subtotal"),
        ("SELLING & DISTRIBUTION OVERHEAD", results["Selling & Distribution"], "normal"),
        ("COST OF SALES", results["Cost of Sales"], "total"),
        (
            f"ADD: PROFIT — {profit_margin:.2f}% ON "
            + ("SALES" if profit_basis == "Profit % on Sales" else "COST OF SALES"),
            profit,
            "profit",
        ),
        ("SALES VALUE", sales_value, "grand"),
    ]

    for label, amount, kind in summary_excel:
        ws.cell(row=row, column=2, value=label)
        ws.cell(row=row, column=5, value=amount)
        ws.cell(row=row, column=6, value=(amount / sales_value) if sales_value else 0.0)

        for col in range(1, 7):
            ws.cell(row=row, column=col).border = thin_border

        if kind in ("subtotal", "total", "grand"):
            fill = total_fill if kind != "grand" else light_blue
            for col in range(1, 7):
                ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=fill)
                ws.cell(row=row, column=col).font = Font(bold=True, color=dark)
        elif kind == "profit":
            for col in range(1, 7):
                ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=profit_fill)
                ws.cell(row=row, column=col).font = Font(bold=True, color=dark)

        row += 1

    # Input/assumption information is kept in the Cost Sheet header area,
    # not as a separate Inputs & Assumptions worksheet.
    row += 2
    ws.cell(row=row, column=1, value="Production Units")
    ws.cell(row=row, column=2, value=units)
    ws.cell(row=row, column=3, value="Profit Basis")
    ws.cell(row=row, column=4, value=profit_basis)
    ws.cell(row=row, column=5, value="Profit Margin %")
    ws.cell(row=row, column=6, value=profit_margin / 100.0)

    for col in range(1, 7):
        ws.cell(row=row, column=col).border = thin_border
        ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=very_light)

    # Number formats
    for r in range(6, row + 1):
        ws.cell(r, 3).number_format = '#,##0.00'
        ws.cell(r, 4).number_format = '#,##0.00'
        ws.cell(r, 5).number_format = '#,##0.00'
        ws.cell(r, 6).number_format = '0.0%'

    # Column widths / layout
    widths = {
        "A": 9,
        "B": 42,
        "C": 16,
        "D": 16,
        "E": 19,
        "F": 14,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A6"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"A5:F{max(5, summary_start - 2)}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:5"

    # Cost Components detail sheet
    detail = wb.create_sheet("Cost Components")
    detail.merge_cells("A1:G1")
    detail["A1"] = "COST COMPONENTS DETAIL"
    detail["A1"].font = Font(bold=True, size=14, color=white)
    detail["A1"].fill = PatternFill("solid", fgColor=navy)
    detail["A1"].alignment = Alignment(horizontal="left")
    detail.row_dimensions[1].height = 26

    detail_headers = [
        "No.", "Category", "Cost Component", "Qty / Units",
        "Rate / Unit", "Total Amount", "Input Basis"
    ]
    for col, header in enumerate(detail_headers, 1):
        cell = detail.cell(row=3, column=col, value=header)
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=blue)
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    for i, line in enumerate(st.session_state.cost_lines, start=1):
        vals = [
            i,
            line["Category"],
            line["Cost Component"],
            line["Qty / Units"],
            line["Rate / Unit"],
            line["Total Amount"],
            line["Input Basis"],
        ]
        for col, value in enumerate(vals, 1):
            cell = detail.cell(row=i + 3, column=col, value=value)
            cell.border = thin_border
            cell.alignment = Alignment(
                horizontal="right" if col in (1, 4, 5, 6) else "left",
                vertical="center",
            )
        detail.cell(i + 3, 4).number_format = '#,##0.00'
        detail.cell(i + 3, 5).number_format = '#,##0.00'
        detail.cell(i + 3, 6).number_format = '#,##0.00'

    detail_widths = {
        "A": 8,
        "B": 25,
        "C": 42,
        "D": 15,
        "E": 15,
        "F": 18,
        "G": 18,
    }
    for col, width in detail_widths.items():
        detail.column_dimensions[col].width = width
    detail.freeze_panes = "A4"
    detail.sheet_view.showGridLines = False
    if st.session_state.cost_lines:
        detail.auto_filter.ref = f"A3:G{len(st.session_state.cost_lines) + 3}"
    detail.page_setup.orientation = "landscape"
    detail.page_setup.fitToWidth = 1
    detail.sheet_properties.pageSetUpPr.fitToPage = True

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# -----------------------------
# Export
# -----------------------------
st.markdown(
    '<div class="section-title">4. Export</div>',
    unsafe_allow_html=True,
)
st.info(
    "The Excel workbook contains only the formatted Cost Sheet and Cost Components detail sheet. "
    "No separate Inputs & Assumptions or Unit Economics worksheet is created."
)

if st.session_state.cost_lines:
    excel_bytes = build_excel()
    st.download_button(
        "⬇️ Download Professional Excel Cost Sheet",
        data=excel_bytes,
        file_name="Suresh_Inc_Standard_Cost_Sheet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
else:
    st.warning("Add at least one cost component before exporting.")
