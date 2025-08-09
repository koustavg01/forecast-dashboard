
# app_full_forecast.py
# End-to-end Forecast + HR + Attrition + Excel-export Streamlit app
# Author: ChatGPT (provided as requested). Save and run with `streamlit run app_full_forecast.py`

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import plotly.graph_objects as go
import plotly.express as px
import xlsxwriter
import openpyxl
import math
try:
    import numpy_financial as nf
except Exception:
    nf = None

st.set_page_config(layout="wide", page_title="Full Forecasting Model (Revenue + HR + Financials)")

# -------------------------
# Utility / Defaults
DEFAULT_INPUT_PATH = "/mnt/data/ClimateX_5yr_Forecast_India_with_HR.xlsx"
DEFAULT_PDF_PATH = "/mnt/data/PwC Challenge 7.0 Round 1 Case.pdf"

def get_default_assumptions():
    """Return comprehensive default assumptions. Replace with exhibit values as needed."""
    return {
        # scenario
        "Scenario Name": "India Expansion - Detailed",
        "Start Year": 2025,
        "Projection Years": 5,
        "Monthly Phasing": False,  # if True, model runs on months (Start Year + months)
        # Revenue & Pricing
        "DAC price per ton (USD)": 600.0,
        "CCS price per ton (USD)": 500.0,
        "Services price per unit (USD/year)": 10000.0,
        "VeriScope starting ARR (USD)": 2_000_000.0,
        "VeriScope ARR growth p.a. (%)": 40.0,
        "Offset marketplace starting (USD)": 600_000.0,
        # Capacity & units
        "DAC capacity per unit (t/yr)": 100,
        "CCS capacity per unit (t/yr)": 8000,
        # CapEx & OpEx by tech (regionally can be adjusted if user inputs exhibits)
        "DAC CapEx per unit (USD)": 1_250_000.0,
        "DAC OpEx per ton (USD)": 780.0,
        "CCS CapEx per unit (USD)": 16_250_000.0,
        "CCS OpEx per ton (USD)": 950.0,
        # Deployment ramp (default)
        "DAC units schedule": [5, 12, 25, 40, 60],  # per year for projection years
        "CCS units schedule": [0, 1, 2, 3, 5],
        # Financials
        "SG&A as % of Revenue": 0.20,
        "R&D as % of Revenue": 0.10,
        "Depreciation life years": 7,
        "Tax rate (%)": 25,
        "Working capital % of revenue": 0.10,
        "Discount rate (%)": 12,
        # HR & Attrition (role-level)
        "Overall Attrition (%)": 18.0,
        "Engineering Attrition (%)": 15.0,
        "Operations Attrition (%)": 20.0,
        "Sales Attrition (%)": 25.0,
        "Support Attrition (%)": 22.0,
        "Install teams required per DAC unit": 0.02,
        "Install teams required per CCS unit": 0.25,
        "Operations FTE per DAC unit": 0.05,
        "Operations FTE per CCS unit": 0.02,
        "Dev/Engineering FTE per $1m ARR": 2.0,
        "Hiring cost per hire (USD)": 5000.0,
        "Onboarding cost per hire (USD)": 2000.0,
        "Average months to full productivity": 6,
        "Recruitment lead time months": 2,
        "install_team_share_of_headcount": 0.30,
        "starting_headcount": 10,
        # Scenario presets
        "Price up pct (best)": 1.10,
        "Price down pct (worst)": 0.90,
    }

# -------------------------
# Core Model Computation
def compute_forecast(assumps):
    """
    Compute annual forecasts (optionally monthly if Monthly Phasing True).
    Returns dictionary of DataFrames.
    """
    start_year = int(assumps["Start Year"])
    n_years = int(assumps["Projection Years"])
    monthly = bool(assumps.get("Monthly Phasing", False))
    # Build timeline
    if monthly:
        n_months = n_years * 12
        timeline = pd.date_range(start=f"{start_year}-01-01", periods=n_months, freq='MS')
    else:
        timeline = [start_year + i for i in range(n_years)]

    # Planned units schedule (expand to match timeline)
    dac_sched = list(assumps["DAC units schedule"])
    ccs_sched = list(assumps["CCS units schedule"])
    if len(dac_sched) < n_years:
        dac_sched += [0] * (n_years - len(dac_sched))
    if len(ccs_sched) < n_years:
        ccs_sched += [0] * (n_years - len(ccs_sched))

    if monthly:
        # simple distribute annual planned units equally across 12 months in year i
        planned_dac = []
        planned_ccs = []
        for val in dac_sched:
            planned_dac += [val / 12.0] * 12
        for val in ccs_sched:
            planned_ccs += [val / 12.0] * 12
    else:
        planned_dac = dac_sched[:n_years]
        planned_ccs = ccs_sched[:n_years]

    # Cumulative planned for capacity calculations
    cumulative_planned_dac = np.cumsum(planned_dac)
    cumulative_planned_ccs = np.cumsum(planned_ccs)

    # Required install teams and FTEs
    req_install = cumulative_planned_dac * assumps["Install teams required per DAC unit"] + cumulative_planned_ccs * assumps["Install teams required per CCS unit"]
    req_ops = cumulative_planned_dac * assumps["Operations FTE per DAC unit"] + cumulative_planned_ccs * assumps["Operations FTE per CCS unit"]
    # VeriScope ARR per year / month
    veri_arr = []
    for i in range(len(timeline)):
        # index i in years corresponds to floor(i/12) if monthly, else i
        idx = math.floor(i/12) if monthly else i
        arr = assumps["VeriScope starting ARR (USD)"] * ((1 + assumps["VeriScope ARR growth p.a. (%)"]/100.0) ** idx)
        veri_arr.append(arr if not monthly else arr / 12.0)

    req_eng = [(arr/1_000_000.0) * assumps["Dev/Engineering FTE per $1m ARR"] for arr in veri_arr]

    # Headcount simulation with attrition
    start_hc = float(assumps["starting_headcount"])
    attr_rate = assumps["Overall Attrition (%)"] / 100.0
    hc_start = []
    hires = []
    hire_costs = []
    onboard_costs = []
    hc_end = []
    prev_end = start_hc
    for i in range(len(timeline)):
        required_total = req_install[i] + req_ops[i] + req_eng[i]
        hc_start.append(prev_end)
        attr_loss = prev_end * attr_rate
        hires_needed = max(0.0, required_total - (prev_end - attr_loss))
        hires.append(hires_needed)
        hire_costs.append(hires_needed * assumps["Hiring cost per hire (USD)"])
        onboard_costs.append(hires_needed * assumps["Onboarding cost per hire (USD)"])
        end_headcount = prev_end - attr_loss + hires_needed
        hc_end.append(end_headcount)
        prev_end = end_headcount

    # Available install teams (fraction share of headcount)
    available_install_teams = np.array(hc_end) * assumps["install_team_share_of_headcount"]
    # Deployment capacity factor
    deployment_capacity_factor = np.ones(len(timeline))
    for i in range(len(timeline)):
        if req_install[i] > 0:
            deployment_capacity_factor[i] = min(1.0, float(available_install_teams[i]) / float(req_install[i]))
        else:
            deployment_capacity_factor[i] = 1.0

    # Realized units (floor to whole units per period)
    realized_dac = np.floor(np.array(planned_dac) * deployment_capacity_factor).astype(int)
    realized_ccs = np.floor(np.array(planned_ccs) * deployment_capacity_factor).astype(int)
    cumulative_realized_dac = np.cumsum(realized_dac)
    cumulative_realized_ccs = np.cumsum(realized_ccs)

    # Tons captured
    dac_tons = cumulative_realized_dac * assumps["DAC capacity per unit (t/yr)"] / (12.0 if monthly else 1.0 if not monthly else 1.0)
    ccs_tons = cumulative_realized_ccs * assumps["CCS capacity per unit (t/yr)"] / (12.0 if monthly else 1.0)

    # Revenue streams
    dac_revenue = dac_tons * assumps["DAC price per ton (USD)"]
    ccs_revenue = ccs_tons * assumps["CCS price per ton (USD)"]
    veriscope_revenue = np.array(veri_arr)
    offset_revenue = np.array([assumps["Offset marketplace starting (USD)"] * ((1+0.10) ** (math.floor(i/12) if monthly else i)) for i in range(len(timeline))])
    services_revenue = np.array(cumulative_realized_dac + cumulative_realized_ccs) * assumps["Services price per unit (USD/year)"] / (12.0 if monthly else 1.0)

    total_revenue = dac_revenue + ccs_revenue + veriscope_revenue + offset_revenue + services_revenue

    # COGS
    dac_cogs = dac_tons * assumps["DAC OpEx per ton (USD)"]
    ccs_cogs = ccs_tons * assumps["CCS OpEx per ton (USD)"]
    total_cogs = dac_cogs + ccs_cogs

    # CapEx: assume CapEx charged at planned deployment (not realized) to reflect ordering & buildup
    capex_each = np.array(planned_dac) * assumps["DAC CapEx per unit (USD)"] + np.array(planned_ccs) * assumps["CCS CapEx per unit (USD)"]
    cumulative_capex = np.cumsum(capex_each)

    # P&L items
    gross_profit = total_revenue - total_cogs
    sgna_base = total_revenue * assumps["SG&A as % of Revenue"]
    sgna = sgna_base + np.array(hire_costs) + np.array(onboard_costs)
    rnd = total_revenue * assumps["R&D as % of Revenue"]
    dep = cumulative_capex / assumps["Depreciation life years"]
    ebit = gross_profit - sgna - rnd - dep
    tax = np.where(ebit > 0, ebit * (assumps["Tax rate (%)"]/100.0), 0.0)
    net_income = ebit - tax

    # Cashflow
    operating_cf = net_income + dep
    wc = total_revenue * assumps["Working capital % of revenue"]
    # Change in WC: delta from prior period; assume prior WC = 0 at t0
    d_wc = []
    prev_wc = 0.0
    for w in wc:
        d_wc.append(w - prev_wc)
        prev_wc = w
    d_wc = np.array(d_wc)
    free_cf = operating_cf - capex_each - d_wc

    # KPI calculations
    discount = assumps["Discount rate (%)"]/100.0
    # discount as annual periods; if monthly, discount periodic rate = (1+r)^(1/12)-1
    if monthly:
        periodic_r = (1+discount)**(1/12.0) - 1.0
        npv = sum([free_cf[i] / ((1+periodic_r)**(i+1)) for i in range(len(free_cf))])
        try:
            irr = nf.irr(free_cf) if nf else np.irr(free_cf)
        except Exception:
            irr = None
    else:
        npv = sum([free_cf[i] / ((1+discount)**(i+1)) for i in range(len(free_cf))])
        try:
            irr = nf.irr(free_cf) if nf else np.irr(free_cf)
        except Exception:
            irr = None

    cum_fcf = np.cumsum(free_cf)
    payback_period = None
    for i, val in enumerate(cum_fcf):
        if val >= 0:
            payback_period = (timeline[i] if not monthly else f"{timeline[i].year}-{timeline[i].month:02d}")
            break

    # Build DataFrames
    headcount_df = pd.DataFrame({
        "Period": timeline,
        "Planned DAC units": planned_dac,
        "Planned CCS units": planned_ccs,
        "Required Install Teams": np.round(req_install,4),
        "Available Install Teams": np.round(available_install_teams,4),
        "Deployment Capacity Factor": np.round(deployment_capacity_factor,4),
        "Realized DAC units": realized_dac,
        "Cumulative Realized DAC units": cumulative_realized_dac,
        "Realized CCS units": realized_ccs,
        "Cumulative Realized CCS units": cumulative_realized_ccs,
        "Required Ops FTE": np.round(req_ops,4),
        "Required Eng FTE": np.round(req_eng,4),
        "Starting Headcount": np.round(hc_start,3),
        "Attrition (%)": assumps["Overall Attrition (%)"],
        "Attrition (FTE loss)": np.round(np.array(hc_start) * attr_rate,4),
        "Hires required (gross)": np.round(hires,3),
        "Hiring Cost (USD)": np.round(hire_costs,2),
        "Onboarding Cost (USD)": np.round(onboard_costs,2),
        "Net Headcount End": np.round(hc_end,3)
    })

    revenue_df = pd.DataFrame({
        "Period": timeline,
        "DAC Tons": np.round(dac_tons,2),
        "DAC Revenue": np.round(dac_revenue,2),
        "DAC COGS": np.round(dac_cogs,2),
        "CCS Tons": np.round(ccs_tons,2),
        "CCS Revenue": np.round(ccs_revenue,2),
        "CCS COGS": np.round(ccs_cogs,2),
        "VeriScope Revenue": np.round(veriscope_revenue,2),
        "Offset Revenue": np.round(offset_revenue,2),
        "Services Revenue": np.round(services_revenue,2),
        "Total Revenue": np.round(total_revenue,2),
        "Total COGS": np.round(total_cogs,2)
    })

    capex_df = pd.DataFrame({
        "Period": timeline,
        "CapEx Investment": np.round(capex_each,2),
        "Cumulative CapEx": np.round(cumulative_capex,2)
    })

    pl_df = pd.DataFrame({
        "Period": timeline,
        "Total Revenue": np.round(total_revenue,2),
        "COGS": np.round(total_cogs,2),
        "Gross Profit": np.round(gross_profit,2),
        "SG&A (incl hires/onboard)": np.round(sgna,2),
        "Hiring Cost": np.round(hire_costs,2),
        "Onboarding Cost": np.round(onboard_costs,2),
        "R&D": np.round(rnd,2),
        "Depreciation": np.round(dep,2),
        "EBIT": np.round(ebit,2),
        "Tax": np.round(tax,2),
        "Net Income": np.round(net_income,2),
        "CapEx Investment": np.round(capex_each,2)
    })
    pl_df["EBITDA"] = pl_df["EBIT"] + pl_df["Depreciation"]
    pl_df["EBITDA Margin"] = pl_df["EBITDA"] / pl_df["Total Revenue"]

    cf_df = pd.DataFrame({
        "Period": timeline,
        "Net Income": np.round(net_income,2),
        "Depreciation": np.round(dep,2),
        "Operating Cash Flow": np.round(operating_cf,2),
        "CapEx (Outflow)": np.round(capex_each,2),
        "Change in WC": np.round(d_wc,2),
        "Free Cash Flow": np.round(free_cf,2),
        "Cumulative FCF": np.round(np.cumsum(free_cf),2)
    })

    kpis = pd.DataFrame([
        {"Metric": f"NPV (discount {assumps['Discount rate (%)']}%)", "Value": npv},
        {"Metric": "IRR", "Value": irr},
        {"Metric": "Payback Period", "Value": payback_period},
        {"Metric": "Final Period EBITDA Margin", "Value": float(pl_df.iloc[-1]["EBITDA Margin"] if len(pl_df)>0 else 0.0)}
    ])

    return {
        "Headcount_Plan": headcount_df,
        "Revenue_Forecast": revenue_df,
        "Capex": capex_df,
        "P&L": pl_df,
        "Cashflow": cf_df,
        "KPIs": kpis,
        "Assumptions": assumps
    }

# -------------------------
# Excel Export with formulas
def build_excel_bytes(results, filename="full_forecast.xlsx"):
    """Write a detailed Excel workbook including some Excel formulas for KPIs."""
    output = io.BytesIO()
    assumps = results["Assumptions"]
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Write Inputs sheet
        inp_df = pd.DataFrame(list(assumps.items()), columns=["Assumption","Value"])
        inp_df.to_excel(writer, sheet_name="Inputs", index=False)
        # Write Rationale
        rationale = [
            ["Rationale: mapping assumptions to model logic."],
            ["See 'Headcount_Plan' for link between install-teams and realized deployments."],
            ["CapEx charged on planned deployments (order/installation costs)."],
            ["Revenue recognized on realized units (constrained by install teams); VeriScope recognized as ARR growth assumptions; Offset grows at 10% p.a. by default."],
            ["Hiring & Onboarding costs are added to SG&A; attrition reduces starting headcount and increases hires required."],
            ["Depreciation is straight-line on cumulative CapEx across configured depreciation life."]
        ]
        pd.DataFrame(rationale).to_excel(writer, sheet_name="RATIONALE", index=False, header=False)

        # Data sheets
        for sheet_name in ["Headcount_Plan","Revenue_Forecast","Capex","P&L","Cashflow","KPIs"]:
            results[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

        workbook = writer.book
        # add a KPIs formula sheet that links FCF and computes NPV/IRR as excel formulas (if possible)
        # Try to put Excel formulas for NPV and IRR on a 'Formulas' sheet referencing Cashflow sheet
        fmt_sheet = workbook.add_worksheet("Formulas")
        cashflow_sheet = "Cashflow"
        # Write labels
        fmt_sheet.write(0,0,"Label")
        fmt_sheet.write(0,1,"Excel Formula / Value")
        fmt_sheet.write(1,0,"Discount Rate")
        fmt_sheet.write(1,1, assumps["Discount rate (%)"]/100.0)
        # NPV formula in Excel expects values: =NPV(rate, range) + initial (if initial is prior cashflow). We'll assume Cashflow.Free Cash Flow starts at row 2 in Cashflow sheet.
        # Find number of rows
        nrows = len(results["Cashflow"])
        if nrows > 0:
            # Cashflow FCF range e.g. 'Cashflow'!G2:G6 (1-based in Excel)
            first_row = 2
            last_row = 1 + nrows
            fcf_range = f"'{cashflow_sheet}'!G{first_row}:G{last_row}"
            # Excel NPV formula:
            fmt_sheet.write(3,0,"NPV (Excel)")
            try:
                # Place a formula referencing the discount rate cell in Inputs sheet if present; find its row
                # For simplicity insert a manual NPV formula using the discount number
                fmt_sheet.write_formula(3,1, f"=NPV({assumps['Discount rate (%)']/100.0},{fcf_range})")
                fmt_sheet.write(4,0,"IRR (Excel)")
                fmt_sheet.write_formula(4,1, f"=IRR({fcf_range})")
            except Exception:
                # if formula insertion fails, write text
                fmt_sheet.write(3,1, "NPV formula omitted due to writer limitations")
                fmt_sheet.write(4,1, "IRR formula omitted due to writer limitations")
        # Close and get bytes
        writer.save()
        output.seek(0)
    return output.getvalue()

# -------------------------
# Streamlit UI
st.title("Comprehensive Forecasting Dashboard — Revenue + HR + Financials")
st.markdown("Edit assumptions in the sidebar or upload an exhibits/inputs Excel to seed values. Use the controls to toggle monthly phasing, run scenarios, and download a packaged workbook.")

# Load an inputs file (optional)
uploaded = st.sidebar.file_uploader("Upload exhibits / prior inputs Excel (optional)", type=["xlsx","xls"])
if uploaded:
    try:
        xls = pd.ExcelFile(uploaded)
        if "Inputs" in xls.sheet_names:
            raw_inputs = xls.parse("Inputs")
            assumps = get_default_assumptions()
            # Map values from two-column Inputs sheet if present
            for _, r in raw_inputs.iterrows():
                key = str(r.iloc[0]).strip()
                val = r.iloc[1]
                if key in assumps:
                    assumps[key] = val
        else:
            st.sidebar.warning("Uploaded file does not contain 'Inputs' sheet; using defaults.")
            assumps = get_default_assumptions()
    except Exception as e:
        st.sidebar.error(f"Error reading upload: {e}")
        assumps = get_default_assumptions()
else:
    # attempt to load default saved file (if exists)
    try:
        assumps = get_default_assumptions()
        # attempt to load defaults from default input path if file exists and has Inputs
        try:
            xls = pd.ExcelFile(DEFAULT_INPUT_PATH)
            if "Inputs" in xls.sheet_names:
                raw_inputs = xls.parse("Inputs")
                for _, r in raw_inputs.iterrows():
                    key = str(r.iloc[0]).strip()
                    val = r.iloc[1]
                    if key in assumps:
                        assumps[key] = val
        except Exception:
            pass
    except Exception:
        assumps = get_default_assumptions()

# Sidebar editing of the comprehensive assumptions
st.sidebar.header("Core settings")
assumps["Scenario Name"] = st.sidebar.text_input("Scenario Name", assumps["Scenario Name"])
assumps["Start Year"] = int(st.sidebar.number_input("Start Year", min_value=2000, max_value=2100, value=int(assumps["Start Year"])))
assumps["Projection Years"] = int(st.sidebar.number_input("Projection Years", min_value=1, max_value=20, value=int(assumps["Projection Years"])))
assumps["Monthly Phasing"] = st.sidebar.checkbox("Monthly phasing (more granular)", value=assumps.get("Monthly Phasing", False))

st.sidebar.markdown("---")
st.sidebar.header("Pricing & Capacity")
assumps["DAC price per ton (USD)"] = st.sidebar.number_input("DAC $/ton", value=float(assumps["DAC price per ton (USD)"]))
assumps["CCS price per ton (USD)"] = st.sidebar.number_input("CCS $/ton", value=float(assumps["CCS price per ton (USD)"]))
assumps["DAC capacity per unit (t/yr)"] = st.sidebar.number_input("DAC capacity per unit (t/yr)", value=float(assumps["DAC capacity per unit (t/yr)"]))
assumps["CCS capacity per unit (t/yr)"] = st.sidebar.number_input("CCS capacity per unit (t/yr)", value=float(assumps["CCS capacity per unit (t/yr)"]))

st.sidebar.markdown("---")
st.sidebar.header("CapEx & OpEx")
assumps["DAC CapEx per unit (USD)"] = st.sidebar.number_input("DAC CapEx / unit", value=float(assumps["DAC CapEx per unit (USD)"]))
assumps["DAC OpEx per ton (USD)"] = st.sidebar.number_input("DAC OpEx / ton", value=float(assumps["DAC OpEx per ton (USD)"]))
assumps["CCS CapEx per unit (USD)"] = st.sidebar.number_input("CCS CapEx / unit", value=float(assumps["CCS CapEx per unit (USD)"]))
assumps["CCS OpEx per ton (USD)"] = st.sidebar.number_input("CCS OpEx / ton", value=float(assumps["CCS OpEx per ton (USD)"]))

st.sidebar.markdown("---")
st.sidebar.header("Deployment schedule (per year)")
n = assumps["Projection Years"]
cols = st.sidebar.columns(2)
dac_sched = []
ccs_sched = []
for i in range(n):
    d = int(cols[0].number_input(f"DAC Y{i+1}", value=int(assumps.get("DAC units schedule", [0]*n)[i] if i < len(assumps.get("DAC units schedule", [])) else 0), key=f"dac{i}"))
    c = int(cols[1].number_input(f"CCS Y{i+1}", value=int(assumps.get("CCS units schedule", [0]*n)[i] if i < len(assumps.get("CCS units schedule", [])) else 0), key=f"ccs{i}"))
    dac_sched.append(d)
    ccs_sched.append(c)
assumps["DAC units schedule"] = dac_sched
assumps["CCS units schedule"] = ccs_sched

st.sidebar.markdown("---")
st.sidebar.header("HR & Attrition (role-level)")
assumps["Overall Attrition (%)"] = float(st.sidebar.number_input("Overall Attrition (%)", value=float(assumps["Overall Attrition (%)"])))
assumps["Hiring cost per hire (USD)"] = float(st.sidebar.number_input("Hiring cost / hire (USD)", value=float(assumps["Hiring cost per hire (USD)"])))
assumps["Onboarding cost per hire (USD)"] = float(st.sidebar.number_input("Onboarding cost / hire (USD)", value=float(assumps["Onboarding cost per hire (USD)"])))
assumps["install_team_share_of_headcount"] = float(st.sidebar.slider("Install team share of headcount", 0.0, 1.0, float(assumps["install_team_share_of_headcount"])))
assumps["starting_headcount"] = int(st.sidebar.number_input("Starting headcount", value=int(assumps["starting_headcount"])))

st.sidebar.markdown("---")
st.sidebar.header("Financials")
assumps["SG&A as % of Revenue"] = float(st.sidebar.number_input("SG&A %", value=float(assumps["SG&A as % of Revenue"])))
assumps["R&D as % of Revenue"] = float(st.sidebar.number_input("R&D %", value=float(assumps["R&D as % of Revenue"])))
assumps["Depreciation life years"] = int(st.sidebar.number_input("Depreciation life (years)", value=int(assumps["Depreciation life years"])))
assumps["Tax rate (%)"] = float(st.sidebar.number_input("Tax rate (%)", value=float(assumps["Tax rate (%)"])))
assumps["Working capital % of revenue"] = float(st.sidebar.number_input("Working capital % of revenue", value=float(assumps["Working capital % of revenue"])))
assumps["Discount rate (%)"] = float(st.sidebar.number_input("Discount rate (%)", value=float(assumps["Discount rate (%)"])))

if st.sidebar.button("Reset to defaults"):
    assumps = get_default_assumptions()

# Compute model
results = compute_forecast(assumps)

# Present KPIs
k1, k2, k3, k4 = st.columns(4)
kpis = results["KPIs"]
k1.metric("NPV", f"{kpis.loc[0,'Value']:.0f}")
k2.metric("IRR", f"{(kpis.loc[1,'Value']*100):.1f}%" if kpis.loc[1,'Value'] is not None else "N/A")
k3.metric("Payback", kpis.loc[2,'Value'])
k4.metric("Final EBITDA Margin", f"{kpis.loc[3,'Value']*100:.1f}%")

# Charts
st.markdown("### Revenue waterfall (stacked)")
rev = results["Revenue_Forecast"]
fig = go.Figure()
fig.add_trace(go.Bar(name="DAC", x=rev["Period"].astype(str), y=rev["DAC Revenue"]))
fig.add_trace(go.Bar(name="CCS", x=rev["Period"].astype(str), y=rev["CCS Revenue"]))
fig.add_trace(go.Bar(name="VeriScope", x=rev["Period"].astype(str), y=rev["VeriScope Revenue"]))
fig.add_trace(go.Bar(name="Offset", x=rev["Period"].astype(str), y=rev["Offset Revenue"]))
fig.add_trace(go.Bar(name="Services", x=rev["Period"].astype(str), y=rev["Services Revenue"]))
fig.update_layout(barmode='stack', xaxis_title="Period", yaxis_title="Revenue (USD)")
st.plotly_chart(fig, use_container_width=True)

st.markdown("### Headcount & hires")
hc = results["Headcount_Plan"]
fig2 = px.line(hc, x="Period", y=["Starting Headcount","Net Headcount End"], markers=True)
fig2.update_layout(yaxis_title="Headcount")
st.plotly_chart(fig2, use_container_width=True)

# P&L and Cashflow tables
st.markdown("### P&L")
st.dataframe(results["P&L"].style.format("{:,.0f}"))

st.markdown("### Cashflow")
st.dataframe(results["Cashflow"].style.format("{:,.0f}"))

# Sensitivity: Price-per-ton vs Attrition (heatmap)
st.markdown("### Sensitivity: DAC price vs Attrition (NPV heatmap)")
# build small grid
prices = np.linspace(assumps["DAC price per ton (USD)"]*0.8, assumps["DAC price per ton (USD)"]*1.2, 7)
attrs = np.linspace(max(0.01, assumps["Overall Attrition (%)"]*0.5/100.0), assumps["Overall Attrition (%)"]*1.5/100.0, 7)
npv_grid = np.zeros((len(attrs), len(prices)))
for i_a,a in enumerate(attrs):
    for j_p,p in enumerate(prices):
        test_assumps = dict(assumps)
        test_assumps["DAC price per ton (USD)"] = float(p)
        test_assumps["Overall Attrition (%)"] = float(a*100.0)
        res = compute_forecast(test_assumps)
        npv_grid[i_a,j_p] = res["KPIs"].loc[0,"Value"]
heat_df = pd.DataFrame(npv_grid, index=[f"{a*100:.1f}%" for a in attrs], columns=[f"${p:.0f}" for p in prices])
fig3 = px.imshow(heat_df, labels=dict(x="DAC price", y="Attrition", color="NPV"), aspect="auto")
st.plotly_chart(fig3, use_container_width=True)

# Download Excel
st.markdown("---")
st.markdown("Download the full workbook (includes input sheet, rationale, headcount plan, revenue, P&L, cashflow, KPIs).")
excel_bytes = build_excel_bytes(results, filename=f"forecast_{assumps['Scenario Name']}.xlsx")
st.download_button("Download Excel workbook", data=excel_bytes, file_name=f"forecast_{assumps['Scenario Name'].replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# RATIONALE quick view
if st.checkbox("Show rationale & mapping (short)"):
    st.markdown("**Rationale mapping (short)**")
    st.write("- CapEx charged on planned units (reflects procurement/installation costs).")
    st.write("- Realized revenue uses realized units constrained by available install teams derived from headcount and attrition.")
    st.write("- Hiring & onboarding costs included in SG&A.")
    st.write("- VeriScope treated as ARR growth; recognized evenly across period (monthly or yearly depending on phasing).")
    st.write("- Working capital is applied as % of revenue; change in WC deducts from FCF.")
    st.write("- Depreciation is straight-line on cumulative CapEx across configured depreciation life.")

st.markdown("Done — edit inputs, run scenarios, and download. If you want I can (pick one):\n\n- Add role-level hiring rules (senior/junior) and different attrition rates per role with separate cost profiles.\n- Convert the Excel export to include more Excel formulas (e.g., explicit NPV formula referencing the Cashflow sheet) at many cells for full transparency.\n- Produce a PowerPoint summary of the base-case model and recommendations.")
