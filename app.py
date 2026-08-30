
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.chart.label import DataLabelList

st.set_page_config(page_title="Cost Sheet Pro", page_icon="📊", layout="wide")

# ---------- CSS ----------
st.markdown("""
<style>
.block-container{padding-top:1.2rem;padding-bottom:2rem}
.hero{padding:26px 30px;border-radius:20px;background:linear-gradient(135deg,#102A43,#1F5A85);color:white;margin-bottom:18px}
.hero h1{font-size:38px;margin:0;font-weight:800}.hero p{margin:6px 0 0;color:#D9EAF7;font-size:16px}
.kpi{padding:18px;border:1px solid #D9E2EC;border-radius:16px;background:#fff;box-shadow:0 3px 12px rgba(16,42,67,.06)}
.kpi-label{font-size:12px;color:#627D98;text-transform:uppercase;font-weight:700}.kpi-value{font-size:24px;font-weight:800;color:#102A43;margin-top:5px}
.section{font-size:19px;font-weight:800;color:#102A43;margin:20px 0 8px}
.small{font-size:12px;color:#627D98}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>📊 Cost Sheet Pro</h1>
<p>Professional cost-sheet builder • Automatic costing • Profit & pricing analysis • Excel export</p>
</div>
""", unsafe_allow_html=True)

# ---------- Sidebar ----------
with st.sidebar:
    st.header("⚙️ Company & Period")
    company=st.text_input("Company / Organisation","ABC Industries Ltd.")
    product=st.text_input("Product / Service","Product / Service")
    period=st.text_input("Period","For the period ended __________")
    currency=st.text_input("Currency symbol","₹")
    st.divider()
    st.header("📦 Production")
    units=st.number_input("Units produced",min_value=0.0,value=1000.0,step=1.0)
    opening_wip=st.number_input("Opening WIP",min_value=0.0,value=0.0)
    closing_wip=st.number_input("Closing WIP",min_value=0.0,value=0.0)
    opening_fg=st.number_input("Opening finished goods",min_value=0.0,value=0.0)
    closing_fg=st.number_input("Closing finished goods",min_value=0.0,value=0.0)
    st.divider()
    st.header("💰 Pricing")
    target_profit_pct=st.number_input("Target profit % on cost",min_value=0.0,value=15.0,step=0.5)
    gst_pct=st.number_input("GST / tax %",min_value=0.0,value=18.0,step=0.5)

default=[
["Direct Material","Raw materials / components",1000,50.0],
["Direct Material","Consumables",1000,3.0],
["Direct Material","Packing material (direct)",1000,2.0],
["Direct Material","Less: Scrap / material recovery",1,-500.0],
["Direct Labour","Direct wages",1000,12.0],
["Direct Labour","Overtime / production incentives",1,2000.0],
["Direct Expenses","Direct expenses",1,5000.0],
["Factory Overhead","Indirect wages",1,5000.0],
["Factory Overhead","Power & fuel",1,8000.0],
["Factory Overhead","Repairs & maintenance",1,2500.0],
["Factory Overhead","Factory depreciation",1,3000.0],
["Factory Overhead","Factory rent / utilities",1,2500.0],
["Factory Overhead","Other factory overheads",1,1000.0],
["Administration Overhead","Office & administration",1,2500.0],
["Administration Overhead","Office salaries",1,4000.0],
["Administration Overhead","IT / professional / audit",1,1500.0],
["Administration Overhead","Other administration overheads",1,1000.0],
["Selling & Distribution","Selling expenses",1,1500.0],
["Selling & Distribution","Distribution / freight outward",1,2000.0],
["Selling & Distribution","Advertising & promotion",1,2500.0],
["Selling & Distribution","Other selling & distribution",1,500.0],
]

if "data" not in st.session_state:
    st.session_state.data=pd.DataFrame(default,columns=["Category","Component","Qty / Units","Rate"])

st.markdown('<div class="section">1. Enter Cost Components</div>',unsafe_allow_html=True)
edited=st.data_editor(
    st.session_state.data,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Category":st.column_config.SelectboxColumn("Cost Category",options=[
            "Direct Material","Direct Labour","Direct Expenses","Factory Overhead",
            "Administration Overhead","Selling & Distribution"
        ],required=True),
        "Component":st.column_config.TextColumn("Component",required=True),
        "Qty / Units":st.column_config.NumberColumn("Qty / Units",min_value=0.0,format="%.2f"),
        "Rate":st.column_config.NumberColumn("Rate",format="%.2f")
    }
)
st.session_state.data=edited.copy()

df=edited.copy()
df["Amount"]=df["Qty / Units"]*df["Rate"]
# A quantity of 1 is used naturally for lump-sum expenses.
df.loc[df["Qty / Units"]==0,"Amount"]=df.loc[df["Qty / Units"]==0,"Rate"]

def total(cat): return float(df.loc[df["Category"]==cat,"Amount"].sum())

material=total("Direct Material")
labour=total("Direct Labour")
direct_exp=total("Direct Expenses")
factory=total("Factory Overhead")
admin=total("Administration Overhead")
selling=total("Selling & Distribution")

prime=material+labour+direct_exp
works=prime+factory+opening_wip-closing_wip
production=works+admin
fg_adjusted=production+opening_fg-closing_fg
sales_cost=fg_adjusted+selling
profit=sales_cost*target_profit_pct/100
sales_before_tax=sales_cost+profit
tax=sales_before_tax*gst_pct/100
invoice_value=sales_before_tax+tax
cost_per_unit=sales_cost/units if units else 0
selling_per_unit=sales_before_tax/units if units else 0
profit_per_unit=profit/units if units else 0

# ---------- KPIs ----------
st.markdown('<div class="section">2. Executive Cost Dashboard</div>',unsafe_allow_html=True)
kpis=[
("Prime Cost",prime),("Works Cost",works),("Production Cost",production),
("Cost of Sales",sales_cost),("Profit",profit),("Invoice Value",invoice_value)
]
cols=st.columns(6)
for col,(lab,val) in zip(cols,kpis):
    with col:
        st.markdown(f'<div class="kpi"><div class="kpi-label">{lab}</div><div class="kpi-value">{currency}{val:,.0f}</div></div>',unsafe_allow_html=True)

st.markdown('<div class="section">3. Unit Economics</div>',unsafe_allow_html=True)
u1,u2,u3,u4=st.columns(4)
u1.metric("Cost / Unit",f"{currency}{cost_per_unit:,.2f}")
u2.metric("Selling Price / Unit",f"{currency}{selling_per_unit:,.2f}")
u3.metric("Profit / Unit",f"{currency}{profit_per_unit:,.2f}")
u4.metric("Tax / GST",f"{currency}{tax:,.2f}")

# ---------- Standard cost sheet ----------
st.markdown('<div class="section">4. Standard Cost Sheet</div>',unsafe_allow_html=True)
rows=[
["DIRECT MATERIAL",material],
["DIRECT LABOUR",labour],
["DIRECT EXPENSES",direct_exp],
["PRIME COST",prime],
["FACTORY / WORKS OVERHEAD",factory],
["OPENING WIP",opening_wip],
["LESS: CLOSING WIP",-closing_wip],
["WORKS / FACTORY COST",works],
["ADMINISTRATION OVERHEAD",admin],
["COST OF PRODUCTION",production],
["OPENING FINISHED GOODS",opening_fg],
["LESS: CLOSING FINISHED GOODS",-closing_fg],
["ADJUSTED PRODUCTION COST",fg_adjusted],
["SELLING & DISTRIBUTION OVERHEAD",selling],
["COST OF SALES",sales_cost],
["PROFIT",profit],
["SALES VALUE BEFORE TAX",sales_before_tax],
["GST / TAX",tax],
["INVOICE / FINAL SALES VALUE",invoice_value]
]
sheet=pd.DataFrame(rows,columns=["Cost Head","Amount"])
sheet["% of Sales"]=sheet["Amount"].apply(lambda x:x/sales_before_tax if sales_before_tax else 0)
st.dataframe(
    sheet.style.format({"Amount":f"{currency}{{:,.2f}}","% of Sales":"{:.1%}"}),
    use_container_width=True,hide_index=True
)

# ---------- Charts ----------
st.markdown('<div class="section">5. Cost Structure</div>',unsafe_allow_html=True)
chart_df=pd.DataFrame({
"Cost Head":["Material","Labour","Direct Exp.","Factory OH","Admin OH","Selling OH"],
"Amount":[material,labour,direct_exp,factory,admin,selling]
})
st.bar_chart(chart_df.set_index("Cost Head"))

# ---------- Excel export ----------
def make_excel():
    wb=Workbook()
    ws=wb.active; ws.title="Cost Sheet"
    navy="102A43"; blue="1F5A85"; teal="0F6B78"; light="D9EAF7"; green="E2F0D9"; orange="FCE4D6"; white="FFFFFF"
    side=Side(style="thin",color="B7C4CE")
    for c,w in zip(range(1,7),[7,43,15,16,19,16]):
        ws.column_dimensions[get_column_letter(c)].width=w
    ws.merge_cells("A1:F1"); ws["A1"]=company
    ws["A1"].font=Font(size=21,bold=True,color=white); ws["A1"].fill=PatternFill("solid",fgColor=navy); ws["A1"].alignment=Alignment(horizontal="center")
    ws.merge_cells("A2:F2"); ws["A2"]="STANDARD COST SHEET"
    ws["A2"].font=Font(size=16,bold=True,color=white); ws["A2"].fill=PatternFill("solid",fgColor=blue); ws["A2"].alignment=Alignment(horizontal="center")
    ws.merge_cells("A3:F3"); ws["A3"]=f"{product}  |  {period}"
    ws["A3"].alignment=Alignment(horizontal="center"); ws["A3"].font=Font(italic=True)
    headers=["No.","Cost Component","Qty / Units","Rate","Amount", "% of Sales"]
    for j,h in enumerate(headers,1):
        x=ws.cell(5,j,h); x.font=Font(bold=True,color=white); x.fill=PatternFill("solid",fgColor=teal); x.alignment=Alignment(horizontal="center")
        x.border=Border(top=side,bottom=side,left=side,right=side)

    r=6;n=1
    cat_names=["Direct Material","Direct Labour","Direct Expenses","Factory Overhead","Administration Overhead","Selling & Distribution"]
    display={"Direct Material":"DIRECT MATERIAL","Direct Labour":"DIRECT LABOUR","Direct Expenses":"DIRECT EXPENSES",
             "Factory Overhead":"FACTORY / WORKS OVERHEAD","Administration Overhead":"ADMINISTRATION OVERHEAD",
             "Selling & Distribution":"SELLING & DISTRIBUTION OVERHEAD"}
    for cat in cat_names:
        ws.cell(r,2,display[cat])
        for j in range(1,7):
            ws.cell(r,j).fill=PatternFill("solid",fgColor=light); ws.cell(r,j).font=Font(bold=True,color=navy); ws.cell(r,j).border=Border(top=side,bottom=side,left=side,right=side)
        ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6); r+=1
        for _,item in df[df["Category"]==cat].iterrows():
            ws.cell(r,1,n); ws.cell(r,2,str(item["Component"])); ws.cell(r,3,float(item["Qty / Units"])); ws.cell(r,4,float(item["Rate"])); ws.cell(r,5,float(item["Amount"]))
            ws.cell(r,6,f'=IFERROR(E{r}/E{r+len(df)+30},0)')
            for j in range(1,7): ws.cell(r,j).border=Border(top=side,bottom=side,left=side,right=side)
            for j in [3,4,5]: ws.cell(r,j).number_format='#,##0.00'
            ws.cell(r,6).number_format='0.0%'; r+=1;n+=1

    r+=1
    ws.cell(r,2,"STANDARD COST SUMMARY")
    for j in range(1,7):
        ws.cell(r,j).fill=PatternFill("solid",fgColor=navy); ws.cell(r,j).font=Font(bold=True,color=white); ws.cell(r,j).border=Border(top=side,bottom=side,left=side,right=side)
    r+=1
    summary_rows=[
        ("Direct Material",material),("Direct Labour",labour),("Direct Expenses",direct_exp),("PRIME COST",prime),
        ("Factory / Works Overhead",factory),("Opening WIP",opening_wip),("Less: Closing WIP",-closing_wip),("WORKS / FACTORY COST",works),
        ("Administration Overhead",admin),("COST OF PRODUCTION",production),("Opening Finished Goods",opening_fg),("Less: Closing Finished Goods",-closing_fg),
        ("ADJUSTED PRODUCTION COST",fg_adjusted),("Selling & Distribution Overhead",selling),("COST OF SALES",sales_cost),
        ("PROFIT",profit),("SALES VALUE BEFORE TAX",sales_before_tax),("GST / TAX",tax),("INVOICE / FINAL SALES VALUE",invoice_value)
    ]
    summary_start=r
    for label,val in summary_rows:
        ws.cell(r,2,label); ws.cell(r,5,val); ws.cell(r,5).number_format='#,##0.00'
        fill=green if label in ["PRIME COST","WORKS / FACTORY COST","COST OF PRODUCTION","ADJUSTED PRODUCTION COST","COST OF SALES","PROFIT"] else orange if "SALES" in label or "INVOICE" in label else light
        for j in range(1,7):
            ws.cell(r,j).fill=PatternFill("solid",fgColor=fill); ws.cell(r,j).border=Border(top=side,bottom=side,left=side,right=side)
        if label in ["PRIME COST","WORKS / FACTORY COST","COST OF PRODUCTION","ADJUSTED PRODUCTION COST","COST OF SALES","INVOICE / FINAL SALES VALUE"]:
            ws.cell(r,2).font=Font(bold=True); ws.cell(r,5).font=Font(bold=True)
        r+=1

    r+=2
    ws.cell(r,2,"UNIT ECONOMICS"); ws.cell(r,2).font=Font(size=13,bold=True,color=navy)
    for label,val in [("Units Produced",units),("Cost / Unit",cost_per_unit),("Selling Price / Unit",selling_per_unit),("Profit / Unit",profit_per_unit),("Target Profit %",target_profit_pct),("GST / Tax %",gst_pct)]:
        r+=1; ws.cell(r,2,label); ws.cell(r,5,val)
        ws.cell(r,5).number_format='0.0%' if "%" in label else '#,##0.00'

    # Chart
    chart=BarChart(); chart.title="Cost Structure"; chart.y_axis.title="Amount"; chart.x_axis.title="Cost Head"
    chart_data=Reference(ws,min_col=5,min_row=summary_start,max_row=summary_start+5)
    cats=Reference(ws,min_col=2,min_row=summary_start,max_row=summary_start+5)
    chart.add_data(chart_data,titles_from_data=False); chart.set_categories(cats); chart.height=7; chart.width=12
    ws.add_chart(chart,"H5")

    ws.freeze_panes="A6"; ws.sheet_view.showGridLines=False
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.page_setup.orientation="portrait"
    ws.print_title_rows="1:5"
    out=BytesIO(); wb.save(out); out.seek(0); return out

st.markdown('<div class="section">6. Download Your Excel Cost Sheet</div>',unsafe_allow_html=True)
st.markdown('<div class="small">The exported workbook includes the standard cost hierarchy, detailed components, cost percentages, unit economics and a cost-structure chart.</div>',unsafe_allow_html=True)
xlsx=make_excel()
st.download_button("⬇️ Download Professional Cost Sheet — Excel",xlsx,"Professional_Cost_Sheet.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
