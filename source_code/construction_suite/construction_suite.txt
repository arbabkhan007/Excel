#!/usr/bin/env python3
"""
Comprehensive Construction Spreadsheet Suite – Excel Generator
================================================================
Generates a single Excel workbook containing 6 fully-integrated systems:
  1. Construction Estimate (10 sheets)
  2. Contractor Bid Template (10 sheets)
  3. Job Costing (12 sheets)
  4. Construction Management (12 sheets)
  5. AIA Invoicing (10 sheets)
  6. Project Timeline / Gantt (12 sheets)
  + Master Dashboard + Config (Hidden)
  = 69 sheets total

Version: 1.0
"""

import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule

# ── Palette ──────────────────────────────────────────────────────────────
DARK_BLUE   = "1F3864"
MID_BLUE    = "2E75B6"
LIGHT_BLUE  = "D6E4F0"
WHITE       = "FFFFFF"
LIGHT_GREEN = "C6EFCE"
LIGHT_RED   = "FFC7CE"
YELLOW_FILL = "FFEB9C"
GREY_BG     = "F2F2F2"
LIGHT_ORANGE= "FCE4D6"

# ── Reusable Styles ─────────────────────────────────────────────────────
thin_border = Border(left=Side('thin'), right=Side('thin'),
                     top=Side('thin'), bottom=Side('thin'))
hdr_font  = Font(name='Calibri', bold=True, color=WHITE, size=11)
hdr_fill  = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type='solid')
title_font = Font(name='Calibri', bold=True, color=WHITE, size=16)
title_fill = PatternFill(start_color=MID_BLUE, end_color=MID_BLUE, fill_type='solid')
ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)
grn = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type='solid')
red = PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')
yel = PatternFill(start_color=YELLOW_FILL, end_color=YELLOW_FILL, fill_type='solid')
gry = PatternFill(start_color=GREY_BG, end_color=GREY_BG, fill_type='solid')
blu = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')
sec_font = Font(name='Calibri', bold=True, size=13, color=DARK_BLUE)
sec_fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')
kpi_font = Font(name='Calibri', bold=True, size=24, color=DARK_BLUE)
kpi_lbl  = Font(name='Calibri', bold=True, size=10, color=MID_BLUE)

# ── Helpers ──────────────────────────────────────────────────────────────
def style_hdr(ws, row, mc):
    for c in range(1, mc+1):
        cl = ws.cell(row=row, column=c)
        cl.font = hdr_font; cl.fill = hdr_fill; cl.alignment = ctr; cl.border = thin_border

def wh(ws, headers, row=2):
    for i, h in enumerate(headers, 1):
        ws.cell(row=row, column=i, value=h)
    style_hdr(ws, row, len(headers))
    return len(headers)

def title(ws, txt, mc):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=mc)
    c = ws.cell(row=1, column=1, value=txt)
    c.font = title_font; c.fill = title_fill
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 36

def dv(ws, rng, opts):
    d = DataValidation(type="list", formula1=f'"{opts}"', allow_blank=True)
    ws.add_data_validation(d); d.add(rng)

def ff(ws, fr, mc):
    ws.freeze_panes = ws.cell(row=fr+1, column=1)
    ws.auto_filter.ref = f"A{fr}:{get_column_letter(mc)}{fr+500}"

def aw(ws, mc, mn=10, mx=32):
    for c in range(1, mc+1):
        ml = mn
        for row in ws.iter_rows(min_col=c, max_col=c, values_only=False):
            for cl in row:
                if cl.value: ml = max(ml, min(len(str(cl.value))+4, mx))
        ws.column_dimensions[get_column_letter(c)].width = ml

# ═════════════════════════════════════════════════════════════════════════
# MAIN BUILDER
# ═════════════════════════════════════════════════════════════════════════
def build():
    wb = Workbook(); dflt = wb.active

    # CONFIG
    _config(wb.create_sheet("Config"))

    # ── SYSTEM 1: CONSTRUCTION ESTIMATE ──
    _est_cover(wb.create_sheet("Est Cover Sheet"))
    _est_csi(wb.create_sheet("Est CSI Division Breakdown"))
    _est_line(wb.create_sheet("Est Detailed Line Items"))
    _est_labor(wb.create_sheet("Est Labor Rates"))
    _est_material(wb.create_sheet("Est Material Prices"))
    _est_sub(wb.create_sheet("Est Sub Quotes"))
    _est_indirect(wb.create_sheet("Est Indirect Costs"))
    _est_markup(wb.create_sheet("Est Markup & Profit"))
    _est_compare(wb.create_sheet("Est Comparison"))
    _est_dash(wb.create_sheet("Est Dashboard"))

    # ── SYSTEM 2: CONTRACTOR BID ──
    _bid_cover(wb.create_sheet("Bid Cover Page"))
    _bid_scope(wb.create_sheet("Bid Scope of Work"))
    _bid_item(wb.create_sheet("Bid Itemization"))
    _bid_takeoff(wb.create_sheet("Bid Material Takeoff"))
    _bid_tab(wb.create_sheet("Bid Sub Tabulation"))
    _bid_price(wb.create_sheet("Bid Pricing Summary"))
    _bid_alt(wb.create_sheet("Bid Alternates & Units"))
    _bid_sched(wb.create_sheet("Bid Schedule"))
    _bid_bond(wb.create_sheet("Bid Bonds & Insurance"))
    _bid_track(wb.create_sheet("Bid Tracking Log"))

    # ── SYSTEM 3: JOB COSTING ──
    _jc_setup(wb.create_sheet("JC Job Setup"))
    _jc_budget(wb.create_sheet("JC Budget Cost Codes"))
    _jc_entry(wb.create_sheet("JC Cost Entry"))
    _jc_summary(wb.create_sheet("JC Cost Summary"))
    _jc_labor(wb.create_sheet("JC Labor Detail"))
    _jc_material(wb.create_sheet("JC Material Detail"))
    _jc_sub(wb.create_sheet("JC Subcontractor Detail"))
    _jc_equip(wb.create_sheet("JC Equipment Detail"))
    _jc_commit(wb.create_sheet("JC Committed Costs"))
    _jc_revenue(wb.create_sheet("JC Revenue vs Cost"))
    _jc_wip(wb.create_sheet("JC WIP Report"))
    _jc_dash(wb.create_sheet("JC Dashboard"))

    # ── SYSTEM 4: CONSTRUCTION MANAGEMENT ──
    _cm_master(wb.create_sheet("CM Project Master"))
    _cm_submittal(wb.create_sheet("CM Submittals Log"))
    _cm_rfi(wb.create_sheet("CM RFI Log"))
    _cm_co(wb.create_sheet("CM Change Order Log"))
    _cm_drawing(wb.create_sheet("CM Drawing Log"))
    _cm_meeting(wb.create_sheet("CM Meeting Minutes"))
    _cm_procure(wb.create_sheet("CM Procurement Log"))
    _cm_insp(wb.create_sheet("CM Inspection Log"))
    _cm_submgmt(wb.create_sheet("CM Sub Management"))
    _cm_daily(wb.create_sheet("CM Daily Site Report"))
    _cm_punch(wb.create_sheet("CM Punch List"))
    _cm_close(wb.create_sheet("CM Closeout Checklist"))

    # ── SYSTEM 5: AIA INVOICING ──
    _aia_contract(wb.create_sheet("AIA Contract Info"))
    _aia_sov(wb.create_sheet("AIA Schedule of Values"))
    _aia_co(wb.create_sheet("AIA Change Order Summary"))
    _aia_payhist(wb.create_sheet("AIA Payment History"))
    _aia_stored(wb.create_sheet("AIA Stored Materials"))
    _aia_lien(wb.create_sheet("AIA Lien Waiver Tracker"))
    _aia_retain(wb.create_sheet("AIA Retainage Tracker"))
    _aia_recon(wb.create_sheet("AIA Certified vs Paid"))
    _aia_g702(wb.create_sheet("AIA G702 Form"))
    _aia_dash(wb.create_sheet("AIA Dashboard"))

    # ── SYSTEM 6: PROJECT TIMELINE ──
    _tl_cal(wb.create_sheet("TL Calendar Setup"))
    _tl_wbs(wb.create_sheet("TL WBS"))
    _tl_task(wb.create_sheet("TL Task Detail"))
    _tl_gantt(wb.create_sheet("TL Gantt Chart"))
    _tl_mile(wb.create_sheet("TL Milestone Tracker"))
    _tl_res(wb.create_sheet("TL Resource Loading"))
    _tl_base(wb.create_sheet("TL Baseline Tracking"))
    _tl_look(wb.create_sheet("TL Look-Ahead Schedule"))
    _tl_risk(wb.create_sheet("TL Risk Register"))
    _tl_crit(wb.create_sheet("TL Critical Path"))
    _tl_prog(wb.create_sheet("TL Progress Reporting"))
    _tl_dash(wb.create_sheet("TL Dashboard"))

    # ── MASTER DASHBOARD ──
    _master_dash(wb.create_sheet("MASTER DASHBOARD", 1), wb)

    wb.remove(dflt)
    wb["Config"].sheet_state = 'hidden'
    return wb

# ═════════════════════════════════════════════════════════════════════════
# CONFIG
# ═════════════════════════════════════════════════════════════════════════
def _config(ws):
    title(ws, "⚙️ CONFIG – Master Dropdowns, CSI Codes, Cost Codes, Rates & Settings", 10)
    lists = {
        "ProjectStatus": ["Planning","Active","On Hold","Complete","Cancelled"],
        "ProjectType": ["Residential","Commercial","Industrial","Infrastructure","Renovation","New Build"],
        "TaskStatus": ["Not Started","In Progress","Complete","Blocked","Delayed"],
        "Priority": ["Critical","High","Medium","Low"],
        "InspectionResult": ["Pass","Fail","Conditional","Pending"],
        "PaymentStatus": ["Paid","Pending","Partial","Overdue","Disputed"],
        "COStatus": ["Pending","Submitted","Under Review","Approved","Rejected","Disputed"],
        "RFIStatus": ["Open","Answered","Closed","Voided"],
        "SubmittalStatus": ["Pending","Approved","Approved as Noted","Revise & Resubmit","Rejected"],
        "RiskLevel": ["Critical","High","Medium","Low"],
        "ContractType": ["Fixed Price","Cost Plus","Time & Material","Unit Price"],
        "EstStatus": ["Draft","Under Review","Submitted","Approved","Rejected","Revised"],
        "BidResult": ["Won","Lost","No Award","Pending"],
        "JobStatus": ["Active","Complete","On Hold","Cancelled"],
        "CostCategory": ["Labor","Material","Equipment","Subcontract","Other"],
        "UOM": ["LS","EA","CY","SF","SY","LF","TON","KG","HR","DAY","MO","GAL","BD","CF"],
        "Dependency": ["FS","SS","FF","SF"],
        "MeetingType": ["OAC","Subcontractor","Safety","Design","Progress"],
        "DocDiscipline": ["Architectural","Structural","Mechanical","Electrical","Civil","Landscape"],
        "DocStatus": ["Current","Superseded","Cancelled"],
        "Weather": ["Sunny","Cloudy","Rainy","Windy","Snowy","Extreme Heat"],
        "InspectionType": ["Concrete Pour","Weld","Soil","Framing","Waterproofing","Electrical","Plumbing","Final"],
        "ProcureStatus": ["Not Started","RFQ Sent","PO Issued","In Production","Shipped","Delivered","Installed"],
        "PunchStatus": ["Open","In Progress","Complete","Verified"],
        "MilestoneStatus": ["Not Started","On Track","At Risk","Delayed","Achieved"],
        "LienStatus": ["Pending","Received","Outstanding"],
        "BondType": ["Bid","Performance","Payment"],
        "ScheduleStatus": ["Ahead","On Track","Slightly Delayed","Significantly Delayed","Critical"],
        "WBSLevel": ["1-Phase","2-Stage","3-Task","4-Sub-task"],
        "ResourceType": ["Labor","Equipment","Material","Subcontractor"],
        "CSI_Divisions": "01-General Requirements,02-Existing Conditions,03-Concrete,04-Masonry,05-Metals,06-Wood/Plastics,07-Thermal/Moisture,08-Openings,09-Finishes,10-Specialties,11-Equipment,12-Furnishings,13-Special Construction,14-Conveying,21-Fire Suppression,22-Plumbing,23-HVAC,26-Electrical,27-Communications,28-Elec Safety,31-Earthwork,32-Exterior Improvements,33-Utilities",
    }
    row = 3
    for key, vals in lists.items():
        ws.cell(row=row, column=1, value=key).font = Font(bold=True, color=DARK_BLUE)
        if isinstance(vals, str):
            ws.cell(row=row, column=2, value=vals)
        else:
            for i, v in enumerate(vals):
                ws.cell(row=row, column=i+2, value=v)
        row += 1

    row += 2
    ws.cell(row=row, column=1, value="SYSTEM SETTINGS").font = Font(bold=True, size=14, color=DARK_BLUE)
    row += 1
    for k, v in [("Company Name","Your Construction Co."),("Address","123 Builder Ave"),
                 ("Phone","+1-XXX-XXX-XXXX"),("Email","info@yourcompany.com"),
                 ("License No.","LIC-XXXXX"),("Tax Rate %",8),("Retainage %",10),
                 ("Overhead %",12),("Profit Margin %",15),("Contingency %",5),
                 ("Working Days/Week",5),("Working Hours/Day",8),("Currency","$")]:
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=v); row += 1

    row += 2
    ws.cell(row=row, column=1, value="STANDARD LABOR RATES").font = Font(bold=True, size=14, color=DARK_BLUE)
    row += 1
    for trade, rate in [("General Laborer",35),("Carpenter",55),("Electrician",65),
                        ("Plumber",60),("Ironworker",70),("Mason",55),("Painter",45),
                        ("Equipment Operator",55),("Foreman",75),("Superintendent",95),
                        ("Project Manager",105),("Safety Officer",70)]:
        ws.cell(row=row, column=1, value=trade)
        ws.cell(row=row, column=2, value=rate); row += 1
    aw(ws, 10)


# ═════════════════════════════════════════════════════════════════════════
# SYSTEM 1: CONSTRUCTION ESTIMATE
# ═════════════════════════════════════════════════════════════════════════
def _est_cover(ws):
    h = ["Estimate ID","Date","Valid Until","Rev #","Project Name","Project Number",
         "Project Type","Address","Client Name","Client Contact","Architect/Engineer",
         "Estimator","Reviewed By","Start Date (Est)","Duration (Weeks)",
         "End Date (Auto)","Status","Total Direct Costs","Total Indirect Costs",
         "Overhead","Profit Margin","Contingency","Grand Total","Notes"]
    mc = wh(ws, h); title(ws, "🏗️ ESTIMATE – COVER SHEET", mc)
    for r in range(3,203):
        ws.cell(row=r,column=1).value = f'=IF(E{r}="","",TEXT(ROW()-2,"EST-0000"))'
        ws.cell(row=r,column=16).value = f'=IF(OR(N{r}="",O{r}=""),"",N{r}+O{r}*7)'
        ws.cell(row=r,column=18).value = f'=IF(E{r}="","",SUMPRODUCT((\'Est Detailed Line Items\'!B3:B10002=E{r})*\'Est Detailed Line Items\'!S3:S10002))'
        ws.cell(row=r,column=23).value = f'=IF(E{r}="","",R{r}+S{r}+T{r}+U{r}+V{r})'
    dv(ws,"G3:G202","Residential,Commercial,Industrial,Infrastructure,Renovation,New Build")
    dv(ws,"Q3:Q202","Draft,Under Review,Submitted,Approved,Rejected,Revised")
    ws.conditional_formatting.add("Q3:Q202",CellIsRule(operator='equal',formula=['"Approved"'],fill=grn))
    ws.conditional_formatting.add("Q3:Q202",CellIsRule(operator='equal',formula=['"Rejected"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _est_csi(ws):
    h = ["Division #","Division Name","Category","Line Item Description",
         "Quantity","Unit","Material Cost","Labor Cost","Equipment Cost",
         "Subcontractor","Other","Total (Auto)","% of Grand Total (Auto)","Notes"]
    mc = wh(ws, h); title(ws, "📊 ESTIMATE – CSI DIVISION BREAKDOWN", mc)
    # Pre-fill all CSI divisions
    divs = [("01","General Requirements","Project Management"),
            ("01","General Requirements","Site Supervision"),
            ("01","General Requirements","Temporary Facilities"),
            ("01","General Requirements","Safety Equipment"),("01","General Requirements","Permits & Fees"),
            ("02","Existing Conditions","Demolition"),("02","Existing Conditions","Site Investigation"),
            ("03","Concrete","Formwork"),("03","Concrete","Reinforcing Steel"),("03","Concrete","Cast-in-Place"),
            ("04","Masonry","Brick Masonry"),("04","Masonry","Block Masonry"),
            ("05","Metals","Structural Steel"),("05","Metals","Metal Decking"),
            ("06","Wood/Plastics","Rough Carpentry"),("06","Wood/Plastics","Finish Carpentry"),
            ("07","Thermal/Moisture","Waterproofing"),("07","Thermal/Moisture","Insulation"),("07","Thermal/Moisture","Roofing"),
            ("08","Openings","Doors"),("08","Openings","Windows"),("08","Openings","Curtain Wall"),
            ("09","Finishes","Drywall"),("09","Finishes","Tiling"),("09","Finishes","Flooring"),("09","Finishes","Painting"),
            ("10","Specialties","Toilet Partitions"),("10","Specialties","Signage"),
            ("11","Equipment","Kitchen Equipment"),("12","Furnishings","Furniture"),
            ("13","Special Construction","Pools"),("14","Conveying","Elevators"),
            ("21","Fire Suppression","Sprinkler Systems"),("22","Plumbing","Plumbing Piping"),
            ("22","Plumbing","Fixtures"),("23","HVAC","Ductwork"),("23","HVAC","Air Handling"),
            ("26","Electrical","Service & Distribution"),("26","Electrical","Lighting"),
            ("27","Communications","Structured Cabling"),("28","Elec Safety","Access Control"),
            ("31","Earthwork","Excavation"),("31","Earthwork","Backfill"),
            ("32","Exterior Improvements","Paving"),("32","Exterior Improvements","Landscaping"),
            ("33","Utilities","Water Supply"),("33","Utilities","Sanitary Sewer")]
    for i,(d,n,cat) in enumerate(divs):
        r = 3+i
        ws.cell(row=r,column=1,value=d); ws.cell(row=r,column=2,value=n); ws.cell(row=r,column=3,value=cat)
        ws.cell(row=r,column=12).value = f'=SUM(G{r}:K{r})'
    r_total = 3 + len(divs) + 1
    ws.cell(row=r_total,column=3,value="GRAND TOTAL:").font = Font(bold=True,size=12,color=DARK_BLUE)
    ws.cell(row=r_total,column=12).value = f'=SUM(L3:L{r_total-1})'
    ws.cell(row=r_total,column=12).font = Font(bold=True,size=14,color=DARK_BLUE)
    ff(ws,2,mc); aw(ws,mc)

def _est_line(ws):
    h = ["Item #","CSI Division","Work Description","Spec Reference","Quantity","UOM",
         "Mat $/Unit","Mat Subtotal (Auto)","Labor Hrs/Unit","Total Labor Hrs (Auto)",
         "Labor Rate $/Hr","Labor $/Unit","Labor Subtotal (Auto)","Equip $/Unit",
         "Equip Subtotal (Auto)","Sub Cost","Other Costs","Total Direct (Auto)",
         "Waste Factor %","Adjusted Total (Auto)","Notes"]
    mc = wh(ws,h); title(ws,"📝 ESTIMATE – DETAILED LINE ITEMS",mc)
    for r in range(3,10003):
        ws.cell(row=r,column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"00000"))'
        ws.cell(row=r,column=8).value = f'=IF(OR(E{r}="",G{r}=""),"",E{r}*G{r})'
        ws.cell(row=r,column=10).value = f'=IF(OR(E{r}="",I{r}=""),"",E{r}*I{r})'
        ws.cell(row=r,column=12).value = f'=IF(K{r}="","",I{r}*K{r})'
        ws.cell(row=r,column=13).value = f'=IF(J{r}="","",J{r}*L{r})'
        ws.cell(row=r,column=15).value = f'=IF(OR(E{r}="",N{r}=""),"",E{r}*N{r})'
        ws.cell(row=r,column=18).value = f'=IF(H{r}="","",H{r}+M{r}+O{r}+IF(P{r}="",0,P{r})+IF(Q{r}="",0,Q{r}))'
        ws.cell(row=r,column=20).value = f'=IF(R{r}="","",ROUND(R{r}*(1+IF(S{r}="",0,S{r})/100),2))'
    dv(ws,"F3:F10002","LS,EA,CY,SF,SY,LF,TON,KG,HR,DAY,MO,GAL,BD,CF")
    ff(ws,2,mc); aw(ws,mc)

def _est_labor(ws):
    h = ["Trade","Base Wage $/Hr","Fringe $/Hr","Insurance & Tax %",
         "Total Burden (Auto)","All-in Rate (Auto)","Overtime Rate (Auto)",
         "Productivity Factor","Effective Rate (Auto)","Location Factor"]
    mc = wh(ws,h); title(ws,"💵 ESTIMATE – LABOR RATES",mc)
    for r in range(3,103):
        ws.cell(row=r,column=5).value = f'=IF(B{r}="","",(B{r}+C{r})*D{r}/100)'
        ws.cell(row=r,column=6).value = f'=IF(B{r}="","",B{r}+C{r}+E{r})'
        ws.cell(row=r,column=7).value = f'=IF(F{r}="","",F{r}*1.5)'
        ws.cell(row=r,column=9).value = f'=IF(F{r}="","",F{r}*IF(H{r}="",1,H{r}))'
    ff(ws,2,mc); aw(ws,mc)

def _est_material(ws):
    h = ["Material Code","Description","Specification","Supplier","Unit",
         "Current Price","Price Date","Price Source","Expected Price",
         "Escalation %","Adjusted Price (Auto)","Lead Time (Days)",
         "Preferred Vendor","Alt Vendor","Notes"]
    mc = wh(ws,h); title(ws,"🧱 ESTIMATE – MATERIAL PRICES",mc)
    for r in range(3,2003):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"MAT0000"))'
        ws.cell(row=r,column=11).value = f'=IF(F{r}="","",ROUND(F{r}*(1+J{r}/100),2))'
    ff(ws,2,mc); aw(ws,mc)

def _est_sub(ws):
    h = ["Quote ID","Trade/Scope","Subcontractor Name","Contact Person","Phone",
         "Email","Date Received","Quote Amount","Inclusions","Exclusions",
         "Clarifications","Duration (Days)","Insurance Confirmed","License Confirmed",
         "Validity Date","Score (1-5)","Status","Notes"]
    mc = wh(ws,h); title(ws,"📄 ESTIMATE – SUBCONTRACTOR QUOTES",mc)
    for r in range(3,1003):
        ws.cell(row=r,column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"SQ0000"))'
    dv(ws,"M3:M1002","Yes,No"); dv(ws,"N3:N1002","Yes,No")
    dv(ws,"Q3:Q1002","Received,Shortlisted,Selected,Rejected")
    ws.conditional_formatting.add("Q3:Q1002",CellIsRule(operator='equal',formula=['"Selected"'],fill=grn))
    ws.conditional_formatting.add("Q3:Q1002",CellIsRule(operator='equal',formula=['"Rejected"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _est_indirect(ws):
    h = ["Category","Description","Basis (Fixed/% of Direct/Per Month)",
         "Amount","% of Direct Cost (Auto)","Notes"]
    mc = wh(ws,h); title(ws,"💰 ESTIMATE – INDIRECT COSTS / OVERHEAD",mc)
    cats = ["Project Management","Home Office Overhead","General Liability Insurance",
            "Builder's Risk Insurance","Workers Comp","Bid Bond","Performance Bond",
            "Payment Bond","Financing Costs","Legal & Professional","Temporary Power",
            "Temporary Water","Equipment Mobilization","Travel & Accommodation","Printing"]
    for i,c in enumerate(cats):
        ws.cell(row=3+i,column=1,value=c)
        ws.cell(row=3+i,column=5).value = f'=IF(D{3+i}="","",D{3+i})'
    ff(ws,2,mc); aw(ws,mc)

def _est_markup(ws):
    h = ["Line Item","Amount","%","Notes"]
    mc = wh(ws,h); title(ws,"📊 ESTIMATE – MARKUP & PROFIT SUMMARY",mc)
    items = ["Total Direct Costs","Total Indirect Costs","Contingency",
             "Profit Margin","Escalation Allowance","Tax","GRAND TOTAL",
             "","Cost per Sq.Ft","Estimate Confidence Level"]
    for i,item in enumerate(items):
        r = 3+i
        ws.cell(row=r,column=1,value=item)
        if i == 6:
            ws.cell(row=r,column=2).value = '=SUM(B3:B8)'
            ws.cell(row=r,column=2).font = Font(bold=True,size=14,color=DARK_BLUE)
    ff(ws,2,mc); aw(ws,mc)

def _est_compare(ws):
    h = ["Category","This Estimate","Previous Estimate","Industry Benchmark",
         "Variance vs Previous (Auto)","Variance vs Benchmark (Auto)",
         "% Variance (Auto)","Notes"]
    mc = wh(ws,h); title(ws,"🔄 ESTIMATE – COMPARISON",mc)
    for r in range(3,52):
        ws.cell(row=r,column=5).value = f'=IF(B{r}="","",B{r}-C{r})'
        ws.cell(row=r,column=6).value = f'=IF(B{r}="","",B{r}-D{r})'
        ws.cell(row=r,column=7).value = f'=IF(OR(B{r}=0,C{r}=""),"",ROUND((B{r}-C{r})/C{r}*100,1))'
    ff(ws,2,mc); aw(ws,mc)

def _est_dash(ws):
    title(ws,"📊 ESTIMATE DASHBOARD",10)
    for lc,lb,vc,fm in [
        ("B3","Grand Total Estimate","B4","='Est Markup & Profit'!B9"),
        ("D3","Cost per Sq.Ft","D4","='Est Markup & Profit'!B11"),
        ("F3","Total Line Items","F4","=COUNTA('Est Detailed Line Items'!C3:C10002)"),
        ("H3","Sub Quotes Received","H4","=COUNTA('Est Sub Quotes'!C3:C1002)")]:
        ws[lc]=lb; ws[lc].font=kpi_lbl; ws[vc]=fm; ws[vc].font=kpi_font
    aw(ws,10)

# ═════════════════════════════════════════════════════════════════════════
# SYSTEM 2: CONTRACTOR BID
# ═════════════════════════════════════════════════════════════════════════
def _bid_cover(ws):
    h = ["Bid Number","Bid Date","Valid Until","Project Name","Project Number",
         "Project Location","Owner Name","Owner Contact","Architect/Engineer",
         "Bid Bond Included","Bond Amount","Base Bid Amount","Alternate Amounts",
         "Total Bid Amount","License Number","Insurance Cert #","Notes"]
    mc = wh(ws,h); title(ws,"📋 BID – COVER PAGE",mc)
    for r in range(3,203):
        ws.cell(row=r,column=1).value = f'=IF(D{r}="","",TEXT(ROW()-2,"BID-0000"))'
        ws.cell(row=r,column=14).value = f'=IF(L{r}="","",L{r}+IF(M{r}="",0,M{r}))'
    dv(ws,"J3:J202","Yes,No")
    ff(ws,2,mc); aw(ws,mc)

def _bid_scope(ws):
    h = ["Section","Item #","Description","Included (Yes/No)","Notes"]
    mc = wh(ws,h); title(ws,"📋 BID – SCOPE OF WORK",mc)
    sections = ["INCLUSIONS","EXCLUSIONS","ASSUMPTIONS","ALTERNATES",
                "UNIT PRICES","ALLOWANCES","OWNER-FURNISHED","PHASING","SPECIAL CONDITIONS"]
    for i,s in enumerate(sections):
        ws.cell(row=3+i,column=1,value=s).font = Font(bold=True,color=DARK_BLUE)
    dv(ws,"D3:D100","Yes,No")
    ff(ws,2,mc); aw(ws,mc)

def _bid_item(ws):
    h = ["Item #","Description","Quantity","Unit","Unit Price",
         "Extension (Auto)","Labor Amount","Material Amount","Sub Amount","Notes"]
    mc = wh(ws,h); title(ws,"📋 BID – ITEMIZATION",mc)
    for r in range(3,5003):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"000"))'
        ws.cell(row=r,column=6).value = f'=IF(OR(C{r}="",E{r}=""),"",C{r}*E{r})'
    ff(ws,2,mc); aw(ws,mc)

def _bid_takeoff(ws):
    h = ["Item","Spec Section","Location/Area","Quantity","Unit","Unit Cost",
         "Total Material Cost (Auto)","Supplier","Lead Time","Availability Risk","Notes"]
    mc = wh(ws,h); title(ws,"📋 BID – MATERIAL TAKEOFF",mc)
    for r in range(3,5003):
        ws.cell(row=r,column=7).value = f'=IF(OR(D{r}="",F{r}=""),"",D{r}*F{r})'
    dv(ws,"J3:J5002","High,Medium,Low")
    ff(ws,2,mc); aw(ws,mc)

def _bid_tab(ws):
    h = ["Trade","Sub 1 Name","Sub 1 Quote","Sub 2 Name","Sub 2 Quote",
         "Sub 3 Name","Sub 3 Quote","Lowest Quote (Auto)","Selected Quote",
         "Diff from Low (Auto)","Scope Verified","Insurance Verified","Selected Sub","Notes"]
    mc = wh(ws,h); title(ws,"📋 BID – SUBCONTRACTOR BID TABULATION",mc)
    for r in range(3,503):
        ws.cell(row=r,column=8).value = f'=IF(C{r}="","",MIN(C{r},IF(E{r}="",999999999,E{r}),IF(G{r}="",999999999,G{r})))'
        ws.cell(row=r,column=10).value = f'=IF(H{r}="","",I{r}-H{r})'
    dv(ws,"K3:K502","Yes,No"); dv(ws,"L3:L502","Yes,No")
    ff(ws,2,mc); aw(ws,mc)

def _bid_price(ws):
    h = ["Division/Category","Material Cost","Labor Cost","Equipment Cost",
         "Subcontractor Cost","Subtotal Direct","General Conditions","Overhead %",
         "Profit %","Contingency %","Bond %","Insurance %","Tax %",
         "Total Bid Price","Cost per SF"]
    mc = wh(ws,h); title(ws,"📋 BID – PRICING SUMMARY",mc)
    for r in range(3,52):
        ws.cell(row=r,column=6).value = f'=IF(B{r}="","",SUM(B{r}:E{r}))'
        ws.cell(row=r,column=14).value = f'=IF(F{r}="","",F{r}+G{r}+F{r}*H{r}/100+F{r}*I{r}/100+F{r}*J{r}/100)'
    r_tot = 53
    ws.cell(row=r_tot,column=1,value="TOTAL").font = Font(bold=True,size=12)
    for c in range(2,15):
        ws.cell(row=r_tot,column=c).value = f'=SUM({get_column_letter(c)}3:{get_column_letter(c)}52)'
        ws.cell(row=r_tot,column=c).font = Font(bold=True)
    ff(ws,2,mc); aw(ws,mc)

def _bid_alt(ws):
    h = ["Alternate #","Description","Add/Deduct","Amount",
         "Duration Impact (Days)","Notes"]
    mc = wh(ws,h); title(ws,"📋 BID – ALTERNATES & UNIT PRICES",mc)
    dv(ws,"C3:C202","Add,Deduct")
    # Unit price schedule section
    r = 210
    ws.cell(row=r,column=1,value="UNIT PRICE SCHEDULE").font = sec_font
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6); ws.cell(row=r,column=1).fill=sec_fill
    r += 1
    for ci,h2 in enumerate(["Item","Unit","Unit Price","Min Qty","Max Qty","Notes"]):
        ws.cell(row=r,column=ci+1,value=h2)
    style_hdr(ws,r,6)
    ff(ws,2,mc); aw(ws,mc)

def _bid_sched(ws):
    h = ["Milestone","Planned Start","Planned Completion","Duration (Auto)",
         "Notes","Contractual Deadline","Liquidated Damages $/Day"]
    mc = wh(ws,h); title(ws,"📋 BID – SCHEDULE",mc)
    for r in range(3,103):
        ws.cell(row=r,column=4).value = f'=IF(OR(B{r}="",C{r}=""),"",C{r}-B{r})'
    ff(ws,2,mc); aw(ws,mc)

def _bid_bond(ws):
    h = ["Bond/Insurance Type","Type","Amount","Company/Agent","Rate %",
         "Cost (Auto)","Premium","Expiry Date","Policy Number","Notes"]
    mc = wh(ws,h); title(ws,"📋 BID – BONDS & INSURANCE",mc)
    for r in range(3,103):
        ws.cell(row=r,column=6).value = f'=IF(OR(C{r}="",E{r}=""),"",C{r}*E{r}/100)'
    dv(ws,"B3:B102","Bid,Performance,Payment,General Liability,Builder's Risk,Workers Comp,Auto Liability,Umbrella")
    ff(ws,2,mc); aw(ws,mc)

def _bid_track(ws):
    h = ["Bid Number","Project Name","Owner/GC","Bid Due Date","Bid Amount",
         "Competitors","Result","Win/Loss Reason","Margin %","Follow-up","Notes"]
    mc = wh(ws,h); title(ws,"📋 BID – TRACKING LOG",mc)
    for r in range(3,1003):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"BID-0000"))'
    dv(ws,"G3:G1002","Won,Lost,No Award,Pending")
    ws.conditional_formatting.add("G3:G1002",CellIsRule(operator='equal',formula=['"Won"'],fill=grn))
    ws.conditional_formatting.add("G3:G1002",CellIsRule(operator='equal',formula=['"Lost"'],fill=red))
    ws.conditional_formatting.add("G3:G1002",CellIsRule(operator='equal',formula=['"Pending"'],fill=yel))
    ff(ws,2,mc); aw(ws,mc)

# ═════════════════════════════════════════════════════════════════════════
# SYSTEM 3: JOB COSTING
# ═════════════════════════════════════════════════════════════════════════
def _jc_setup(ws):
    h = ["Job Number","Job Name","Client Name","Contract Type","Original Contract",
         "Approved Change Orders","Revised Contract (Auto)","Original Start","Original End",
         "Revised Start","Revised End","Project Manager","Superintendent","Status",
         "Job Address","Bonded","Retainage %","Notes"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – JOB SETUP",mc)
    for r in range(3,503):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"JOB-0000"))'
        ws.cell(row=r,column=7).value = f'=IF(E{r}="","",E{r}+F{r})'
    dv(ws,"D3:D502","Fixed Price,Cost Plus,Time & Material,Unit Price")
    dv(ws,"N3:N502","Active,Complete,On Hold,Cancelled")
    dv(ws,"P3:P502","Yes,No")
    ws.conditional_formatting.add("N3:N502",CellIsRule(operator='equal',formula=['"Active"'],fill=grn))
    ws.conditional_formatting.add("N3:N502",CellIsRule(operator='equal',formula=['"Complete"'],fill=blu))
    ws.conditional_formatting.add("N3:N502",CellIsRule(operator='equal',formula=['"Cancelled"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _jc_budget(ws):
    h = ["Cost Code #","Description","Category","Phase","Original Budget",
         "Approved COs to Code","Revised Budget (Auto)","Notes"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – BUDGET / COST CODES",mc)
    for r in range(3,1003):
        ws.cell(row=r,column=7).value = f'=IF(E{r}="","",E{r}+F{r})'
    dv(ws,"C3:C1002","Labor,Material,Equipment,Subcontract,Other")
    ff(ws,2,mc); aw(ws,mc)

def _jc_entry(ws):
    h = ["Transaction ID","Date","Job Number","Cost Code","Category (L/M/E/S/O)",
         "Description","Vendor/Employee","Invoice/Check #","Quantity","Unit",
         "Unit Cost","Amount","Committed/Actual","Status","Posted By","Notes"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – COST ENTRY",mc)
    for r in range(3,50003):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"TXN000000"))'
        ws.cell(row=r,column=12).value = f'=IF(OR(I{r}="",K{r}=""),"",I{r}*K{r})'
    dv(ws,"E3:E50002","L,M,E,S,O")
    dv(ws,"N3:N50002","Posted,Pending")
    ff(ws,2,mc); aw(ws,mc)

def _jc_summary(ws):
    h = ["Cost Code","Description","Original Budget","Change Orders","Revised Budget",
         "Committed Costs","Actual to Date","Total Costs (Auto)","Budget Remaining (Auto)",
         "% Budget Used (Auto)","Est Cost at Completion","Projected Over/Under (Auto)"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – COST SUMMARY",mc)
    for r in range(3,1003):
        ws.cell(row=r,column=5).value = f'=IF(C{r}="","",C{r}+D{r})'
        ws.cell(row=r,column=8).value = f'=IF(C{r}="","",F{r}+G{r})'
        ws.cell(row=r,column=9).value = f'=IF(E{r}="","",E{r}-H{r})'
        ws.cell(row=r,column=10).value = f'=IF(E{r}="","",ROUND(H{r}/E{r}*100,1))'
        ws.cell(row=r,column=12).value = f'=IF(E{r}="","",E{r}-H{r})'
    ws.conditional_formatting.add("J3:J1002",CellIsRule(operator='greaterThan',formula=['100'],fill=red))
    ws.conditional_formatting.add("J3:J1002",CellIsRule(operator='lessThanOrEqual',formula=['100'],fill=grn))
    ws.conditional_formatting.add("L3:L1002",CellIsRule(operator='lessThan',formula=['0'],fill=red))
    ws.conditional_formatting.add("L3:L1002",CellIsRule(operator='greaterThanOrEqual',formula=['0'],fill=grn))
    ff(ws,2,mc); aw(ws,mc)

def _jc_labor(ws):
    h = ["Week Ending","Employee Name/ID","Trade","Job Number","Cost Code",
         "Regular Hours","OT Hours","DT Hours","Regular Rate","OT Rate","DT Rate",
         "Regular Pay (Auto)","OT Pay (Auto)","DT Pay (Auto)","Burden %",
         "Burdened Cost (Auto)","Total Hours (Auto)","Total Cost (Auto)",
         "Productivity Units","Unit Labor Cost (Auto)"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – LABOR DETAIL",mc)
    for r in range(3,20003):
        ws.cell(row=r,column=12).value = f'=IF(F{r}="","",F{r}*I{r})'
        ws.cell(row=r,column=13).value = f'=IF(G{r}="","",G{r}*J{r})'
        ws.cell(row=r,column=14).value = f'=IF(H{r}="","",H{r}*K{r})'
        ws.cell(row=r,column=16).value = f'=IF(L{r}="","",(L{r}+M{r}+N{r})*(1+O{r}/100))'
        ws.cell(row=r,column=17).value = f'=IF(F{r}="","",F{r}+G{r}+H{r})'
        ws.cell(row=r,column=18).value = f'=IF(L{r}="","",L{r}+M{r}+N{r}+P{r})'
        ws.cell(row=r,column=20).value = f'=IF(OR(S{r}=0,R{r}=""),"",R{r}/S{r})'
    ff(ws,2,mc); aw(ws,mc)

def _jc_material(ws):
    h = ["Date","PO Number","Vendor","Invoice #","Cost Code","Description",
         "Qty Ordered","Qty Received","Unit","Unit Price","Total Amount",
         "Tax","Freight","Grand Total (Auto)","Payment Status","Payment Date","Notes"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – MATERIAL DETAIL",mc)
    for r in range(3,10003):
        ws.cell(row=r,column=11).value = f'=IF(OR(H{r}="",J{r}=""),"",H{r}*J{r})'
        ws.cell(row=r,column=14).value = f'=IF(K{r}="","",K{r}+L{r}+M{r})'
    dv(ws,"O3:O10002","Paid,Pending,Partial")
    ff(ws,2,mc); aw(ws,mc)

def _jc_sub(ws):
    h = ["Sub Contract #","Subcontractor Name","Trade/Scope","Cost Code",
         "Original Sub Value","Approved COs","Revised Contract (Auto)",
         "Billed to Date","Paid to Date","Retainage Withheld (Auto)",
         "Balance to Complete (Auto)","% Complete","Projected Final Cost","Notes"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – SUBCONTRACTOR DETAIL",mc)
    for r in range(3,2003):
        ws.cell(row=r,column=7).value = f'=IF(E{r}="","",E{r}+F{r})'
        ws.cell(row=r,column=10).value = f'=IF(G{r}="","",G{r}*0.10)'
        ws.cell(row=r,column=11).value = f'=IF(G{r}="","",G{r}-H{r})'
    ff(ws,2,mc); aw(ws,mc)

def _jc_equip(ws):
    h = ["Date","Equipment ID/Name","Job/Cost Code","Hours/Days Used",
         "Internal Rate","External Rate","Cost (Auto)","Fuel Cost",
         "Operator Cost","Total Equip Cost (Auto)","Notes"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – EQUIPMENT DETAIL",mc)
    for r in range(3,5003):
        ws.cell(row=r,column=7).value = f'=IF(D{r}="","",D{r}*E{r})'
        ws.cell(row=r,column=10).value = f'=IF(D{r}="","",G{r}+H{r}+I{r})'
    ff(ws,2,mc); aw(ws,mc)

def _jc_commit(ws):
    h = ["PO Number","Vendor","Date Issued","Cost Code","PO Amount",
         "Received to Date","Invoiced to Date","Remaining (Auto)","Expected Delivery","Notes"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – COMMITTED COSTS (OPEN POs)",mc)
    for r in range(3,5003):
        ws.cell(row=r,column=8).value = f'=IF(E{r}="","",E{r}-F{r})'
    ff(ws,2,mc); aw(ws,mc)

def _jc_revenue(ws):
    h = ["Cost Code","Revenue Allocated","Actual Cost","Gross Profit (Auto)",
         "Gross Margin % (Auto)","Est Revenue at Complete","Est Cost at Complete",
         "Est Profit at Complete","Earned Revenue JTD","Over/Under Billing (Auto)"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – REVENUE VS COST",mc)
    for r in range(3,1003):
        ws.cell(row=r,column=4).value = f'=IF(B{r}="","",B{r}-C{r})'
        ws.cell(row=r,column=5).value = f'=IF(B{r}="","",ROUND(D{r}/B{r}*100,1))'
        ws.cell(row=r,column=10).value = f'=IF(B{r}="","",I{r}-B{r})'
    ff(ws,2,mc); aw(ws,mc)

def _jc_wip(ws):
    h = ["Job Number","Job Name","Contract Amount","Est Cost at Completion",
         "Gross Profit Expected (Auto)","Gross Margin % (Auto)","% Complete",
         "Earned Revenue (Auto)","Billings to Date","Over/Under Billing (Auto)",
         "Cost Incurred JTD","Remaining Cost (Auto)"]
    mc = wh(ws,h); title(ws,"💼 JOB COSTING – WIP REPORT",mc)
    for r in range(3,103):
        ws.cell(row=r,column=5).value = f'=IF(C{r}="","",C{r}-D{r})'
        ws.cell(row=r,column=6).value = f'=IF(C{r}="","",ROUND(E{r}/C{r}*100,1))'
        ws.cell(row=r,column=8).value = f'=IF(C{r}="","",C{r}*G{r}/100)'
        ws.cell(row=r,column=10).value = f'=IF(H{r}="","",H{r}-I{r})'
        ws.cell(row=r,column=12).value = f'=IF(D{r}="","",D{r}-K{r})'
    ff(ws,2,mc); aw(ws,mc)

def _jc_dash(ws):
    title(ws,"💼 JOB COSTING DASHBOARD",10)
    for lc,lb,vc,fm in [
        ("B3","Active Jobs","B4",'=COUNTIF(\'JC Job Setup\'!N3:N502,"Active")'),
        ("D3","Total Contract Value","D4","=SUM('JC Job Setup'!G3:G502)"),
        ("F3","Total Costs JTD","F4","=SUM('JC Cost Summary'!H3:H1002)"),
        ("H3","Avg Margin %","H4","=IFERROR(ROUND(AVERAGE('JC WIP Report'!F3:F102),1),0)"),
        ("J3","Jobs Over Budget","J4","=COUNTIF('JC Cost Summary'!J3:J1002,\">100\")")]:
        ws[lc]=lb; ws[lc].font=kpi_lbl; ws[vc]=fm; ws[vc].font=kpi_font
    aw(ws,10)

# ═════════════════════════════════════════════════════════════════════════
# SYSTEM 4: CONSTRUCTION MANAGEMENT
# ═════════════════════════════════════════════════════════════════════════
def _cm_master(ws):
    h = ["Project ID","Project Name","Project Type","Owner","Architect","Engineer",
         "GC","Contract Type","Contract Value","Start Date (Plan)","End Date (Plan)",
         "Start Date (Act)","End Date (Act)","Current Phase","% Complete (Auto)",
         "Budget Status","Schedule Status","Safety Status","Quality Status","Overall RAG"]
    mc = wh(ws,h); title(ws,"🏢 CM – PROJECT MASTER CONTROL",mc)
    for r in range(3,203):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"PRJ-000"))'
        ws.cell(row=r,column=15).value = f'=IF(B{r}="","",IFERROR(ROUND(COUNTIFS(\'TL Task Detail\'!B:B,B{r},\'TL Task Detail\'!N:N,"Complete")/MAX(COUNTIF(\'TL Task Detail\'!B:B,B{r}),1)*100,1),0))'
    dv(ws,"C3:C202","Residential,Commercial,Industrial,Infrastructure,Renovation,New Build")
    dv(ws,"T3:T202","Red,Amber,Green")
    ws.conditional_formatting.add("T3:T202",CellIsRule(operator='equal',formula=['"Green"'],fill=grn))
    ws.conditional_formatting.add("T3:T202",CellIsRule(operator='equal',formula=['"Amber"'],fill=yel))
    ws.conditional_formatting.add("T3:T202",CellIsRule(operator='equal',formula=['"Red"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _cm_submittal(ws):
    h = ["Submittal #","Spec Section","Description","Subcontractor/Supplier",
         "Required By","Submitted to GC","Submitted to Arch","Review Days",
         "Date Returned","Status","Revision #","Days to Respond (Auto)",
         "Schedule Impact","Notes"]
    mc = wh(ws,h); title(ws,"🏢 CM – SUBMITTALS LOG",mc)
    for r in range(3,2003):
        ws.cell(row=r,column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"SUB-00000"))'
        ws.cell(row=r,column=12).value = f'=IF(OR(F{r}="",I{r}=""),"",I{r}-F{r})'
    dv(ws,"J3:J2002","Pending,Approved,Approved as Noted,Revise & Resubmit,Rejected")
    ws.conditional_formatting.add("J3:J2002",CellIsRule(operator='equal',formula=['"Approved"'],fill=grn))
    ws.conditional_formatting.add("J3:J2002",CellIsRule(operator='equal',formula=['"Rejected"'],fill=red))
    ws.conditional_formatting.add("J3:J2002",CellIsRule(operator='equal',formula=['"Pending"'],fill=yel))
    ff(ws,2,mc); aw(ws,mc)

def _cm_rfi(ws):
    h = ["RFI #","Date Initiated","Initiated By","Directed To","Subject",
         "Spec Reference","Drawing Ref","Priority","Required Response Date",
         "Date Answered","Days Open (Auto)","Answer Summary",
         "Cost Impact","Schedule Impact","Status","Notes"]
    mc = wh(ws,h); title(ws,"🏢 CM – RFI LOG",mc)
    for r in range(3,5003):
        ws.cell(row=r,column=1).value = f'=IF(E{r}="","",TEXT(ROW()-2,"RFI-00000"))'
        ws.cell(row=r,column=11).value = f'=IF(O{r}="Closed","",IF(I{r}="","",TODAY()-I{r}))'
    dv(ws,"H3:H5002","Routine,Urgent,Critical")
    dv(ws,"O3:O5002","Open,Answered,Closed,Voided")
    ws.conditional_formatting.add("O3:O5002",CellIsRule(operator='equal',formula=['"Closed"'],fill=grn))
    ws.conditional_formatting.add("O3:O5002",CellIsRule(operator='equal',formula=['"Open"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _cm_co(ws):
    h = ["CO Number","Date","Initiated By","Description","Reason",
         "Contractor Proposal","Owner Negotiated","Final Approved","Schedule Impact (Days)",
         "Cost Code Affected","Status","Date Approved","Running Total (Auto)","Notes"]
    mc = wh(ws,h); title(ws,"🏢 CM – CHANGE ORDER LOG",mc)
    for r in range(3,2003):
        ws.cell(row=r,column=1).value = f'=IF(D{r}="","",TEXT(ROW()-2,"CO-00000"))'
        ws.cell(row=r,column=13).value = f'=IF(H{r}="","",SUM($H$3:H{r}))'
    dv(ws,"E3:E2002","Owner Request,Design Error,Differing Conditions,Unforeseen,Code Requirement")
    dv(ws,"K3:K2002","Pending,Submitted,Under Review,Approved,Rejected,Disputed")
    ws.conditional_formatting.add("K3:K2002",CellIsRule(operator='equal',formula=['"Approved"'],fill=grn))
    ws.conditional_formatting.add("K3:K2002",CellIsRule(operator='equal',formula=['"Rejected"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _cm_drawing(ws):
    h = ["Document #","Revision","Title/Description","Discipline","Date Issued",
         "Issued For","Issued To","File Link","Superseded By","Status","Notes"]
    mc = wh(ws,h); title(ws,"🏢 CM – DRAWING & DOCUMENT LOG",mc)
    for r in range(3,5003):
        ws.cell(row=r,column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"DWG-00000"))'
    dv(ws,"D3:D5002","Architectural,Structural,Mechanical,Electrical,Civil,Landscape")
    dv(ws,"J3:J5002","Current,Superseded,Cancelled")
    ws.conditional_formatting.add("J3:J5002",CellIsRule(operator='equal',formula=['"Current"'],fill=grn))
    ws.conditional_formatting.add("J3:J5002",CellIsRule(operator='equal',formula=['"Superseded"'],fill=yel))
    ff(ws,2,mc); aw(ws,mc)

def _cm_meeting(ws):
    h = ["Meeting #","Type","Date","Time","Location","Attendees",
         "Agenda","Decisions","Action Items","Responsible/Due","Next Meeting",
         "Prepared By","Distributed To","Status"]
    mc = wh(ws,h); title(ws,"🏢 CM – MEETING MINUTES LOG",mc)
    for r in range(3,1003):
        ws.cell(row=r,column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"MTG-0000"))'
    dv(ws,"B3:B1002","OAC,Subcontractor,Safety,Design,Progress")
    dv(ws,"N3:N1002","Draft,Distributed,Approved")
    ff(ws,2,mc); aw(ws,mc)

def _cm_procure(ws):
    h = ["Item #","Description","Spec Section","Required on Site","Lead Time (Days)",
         "Order By Date (Auto)","Supplier","PO #","PO Date","PO Amount",
         "Expected Delivery","Actual Delivery","Status","Notes"]
    mc = wh(ws,h); title(ws,"🏢 CM – PROCUREMENT LOG",mc)
    for r in range(3,2003):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"PRC-00000"))'
        ws.cell(row=r,column=6).value = f'=IF(OR(D{r}="",E{r}=""),"",D{r}-E{r})'
    dv(ws,"M3:M2002","Not Started,RFQ Sent,PO Issued,In Production,Shipped,Delivered,Installed")
    ws.conditional_formatting.add("M3:M2002",CellIsRule(operator='equal',formula=['"Delivered"'],fill=grn))
    ws.conditional_formatting.add("M3:M2002",CellIsRule(operator='equal',formula=['"Installed"'],fill=blu))
    ff(ws,2,mc); aw(ws,mc)

def _cm_insp(ws):
    h = ["Inspection ID","Date","Type","Location/Element","Inspector/Agency",
         "Spec Reference","Result","Report #","Report Link","Retest Required",
         "Retest Date","Notes"]
    mc = wh(ws,h); title(ws,"🏢 CM – INSPECTION & TESTING LOG",mc)
    for r in range(3,5003):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"INSP-00000"))'
    dv(ws,"C3:C5002","Concrete Pour,Weld,Soil,Framing,Waterproofing,Electrical,Plumbing,Final")
    dv(ws,"G3:G5002","Pass,Fail,Conditional,Pending")
    dv(ws,"J3:J5002","Yes,No")
    ws.conditional_formatting.add("G3:G5002",CellIsRule(operator='equal',formula=['"Pass"'],fill=grn))
    ws.conditional_formatting.add("G3:G5002",CellIsRule(operator='equal',formula=['"Fail"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _cm_submgmt(ws):
    h = ["Sub Name","Trade","Contract Amount","Change Orders","Revised Contract (Auto)",
         "Billed to Date","% Billed (Auto)","% Work Complete","Over/Under Billing (Auto)",
         "Retainage Withheld","Insurance Expiry","License Expiry","Performance (1-5)","Notes"]
    mc = wh(ws,h); title(ws,"🏢 CM – SUBCONTRACTOR MANAGEMENT",mc)
    for r in range(3,503):
        ws.cell(row=r,column=5).value = f'=IF(C{r}="","",C{r}+D{r})'
        ws.cell(row=r,column=7).value = f'=IF(E{r}="","",ROUND(F{r}/E{r}*100,1))'
        ws.cell(row=r,column=9).value = f'=IF(F{r}="","",F{r}-E{r}*H{r}/100)'
    ff(ws,2,mc); aw(ws,mc)

def _cm_daily(ws):
    h = ["Date","Project Name","Weather AM","Temp AM","Weather PM","Temp PM",
         "Work in Progress","Manpower (Trade/Count)","Equipment on Site",
         "Materials Received","Visitors","Inspections Today","Safety Issues",
         "Delays/Problems","Work Completed","Work Planned Tomorrow","Photos (Links)","Prepared By"]
    mc = wh(ws,h); title(ws,"🏢 CM – DAILY SITE REPORT",mc)
    ff(ws,2,mc); aw(ws,mc)

def _cm_punch(ws):
    h = ["Item #","Area/Location","Description","Trade Responsible","Date Identified",
         "Due Date","Status","Completion Date","Verified By","Re-inspection","Priority","Notes"]
    mc = wh(ws,h); title(ws,"🏢 CM – PUNCH LIST",mc)
    for r in range(3,5003):
        ws.cell(row=r,column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"PL-00000"))'
    dv(ws,"G3:G5002","Open,In Progress,Complete,Verified")
    dv(ws,"K3:K5002","Critical,High,Medium,Low")
    ws.conditional_formatting.add("G3:G5002",CellIsRule(operator='equal',formula=['"Verified"'],fill=grn))
    ws.conditional_formatting.add("G3:G5002",CellIsRule(operator='equal',formula=['"Complete"'],fill=blu))
    ws.conditional_formatting.add("G3:G5002",CellIsRule(operator='equal',formula=['"Open"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _cm_close(ws):
    h = ["Item #","Category","Description","Responsible Party","Required Date",
         "Completion Date","Status","Notes"]
    mc = wh(ws,h); title(ws,"🏢 CM – CLOSEOUT CHECKLIST",mc)
    cats = ["Punch List Complete","Final Inspections","Certificate of Occupancy",
            "As-Built Drawings","O&M Manuals","Warranties & Guarantees",
            "Spare Parts","Keys/Access Cards","Training","Final Cleaning",
            "Lien Waivers (all subs)","Final Payment Applications",
            "Release of Retainage","Final Accounting"]
    for i,c in enumerate(cats):
        r = 3+i
        ws.cell(row=r,column=1,value=i+1)
        ws.cell(row=r,column=2,value=c)
    dv(ws,"G3:G20","Not Started,In Progress,Complete")
    ws.conditional_formatting.add("G3:G20",CellIsRule(operator='equal',formula=['"Complete"'],fill=grn))
    ff(ws,2,mc); aw(ws,mc)

# ═════════════════════════════════════════════════════════════════════════
# SYSTEM 5: AIA INVOICING
# ═════════════════════════════════════════════════════════════════════════
def _aia_contract(ws):
    h = ["App #","App Date","Period From","Period To","Project Name","Project #",
         "Location","Owner Name","Owner Address","Architect Name","Architect Address",
         "Contractor Name","Contractor Address","Contract #","Contract Date",
         "Original Contract Sum","Net Change by COs","Contract Sum to Date (Auto)",
         "Total Completed & Stored","Retainage %","Total Retainage (Auto)",
         "Earned Less Retainage (Auto)","Less Previous Certificates (Auto)",
         "Current Payment Due (Auto)","Balance to Finish (Auto)"]
    mc = wh(ws,h); title(ws,"💰 AIA – PROJECT & CONTRACT INFO",mc)
    for r in range(3,203):
        ws.cell(row=r,column=1).value = f'=IF(E{r}="","",TEXT(ROW()-2,"G702-000"))'
        ws.cell(row=r,column=18).value = f'=IF(P{r}="","",P{r}+Q{r})'
        ws.cell(row=r,column=21).value = f'=IF(S{r}="","",S{r}*T{r}/100)'
        ws.cell(row=r,column=22).value = f'=IF(S{r}="","",S{r}-U{r})'
        ws.cell(row=r,column=23).value = f'=IF(r>3,V{r-1},0)'
        ws.cell(row=r,column=24).value = f'=IF(V{r}="","",V{r}-IF(W{r}="",0,W{r}))'
        ws.cell(row=r,column=25).value = f'=IF(R{r}="","",R{r}-S{r})'
    ff(ws,2,mc); aw(ws,mc)

def _aia_sov(ws):
    h = ["Item #","Description","Scheduled Value","CO Add/Deduct",
         "Revised Value (Auto)","Completed Prev App","Completed This Period",
         "Materials Stored","Total Completed & Stored (Auto)","% Complete (Auto)",
         "Balance to Finish (Auto)","Retainage %","Retainage Amount (Auto)",
         "Net Due This Item (Auto)","Notes"]
    mc = wh(ws,h); title(ws,"💰 AIA – SCHEDULE OF VALUES (G703)",mc)
    for r in range(3,203):
        ws.cell(row=r,column=5).value = f'=IF(C{r}="","",C{r}+D{r})'
        ws.cell(row=r,column=9).value = f'=IF(F{r}="","",F{r}+G{r}+H{r})'
        ws.cell(row=r,column=10).value = f'=IF(E{r}="","",ROUND(I{r}/E{r}*100,1))'
        ws.cell(row=r,column=11).value = f'=IF(E{r}="","",E{r}-I{r})'
        ws.cell(row=r,column=13).value = f'=IF(I{r}="","",I{r}*L{r}/100)'
        ws.cell(row=r,column=14).value = f'=IF(I{r}="","",I{r}-M{r})'
    # Totals row
    r=204
    ws.cell(row=r,column=1,value="TOTALS").font = Font(bold=True,size=12,color=DARK_BLUE)
    for c in [3,4,5,6,7,8,9,11,13,14]:
        ws.cell(row=r,column=c).value = f'=SUM({get_column_letter(c)}3:{get_column_letter(c)}203)'
        ws.cell(row=r,column=c).font = Font(bold=True)
    ff(ws,2,mc); aw(ws,mc)

def _aia_co(ws):
    h = ["CO Number","Description","Date Approved","Amount","Running Total (Auto)",
         "Status","Schedule Impact","Notes"]
    mc = wh(ws,h); title(ws,"💰 AIA – CHANGE ORDER SUMMARY (G701)",mc)
    for r in range(3,503):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"CO-00000"))'
        ws.cell(row=r,column=5).value = f'=IF(D{r}="","",SUM($D$3:D{r}))'
    dv(ws,"F3:F502","Approved,Pending,Disputed")
    ff(ws,2,mc); aw(ws,mc)

def _aia_payhist(ws):
    h = ["App Number","Period","Contract Sum","Scheduled Value",
         "Total Completed & Stored","Retainage Held","Amount Certified",
         "Amount Paid","Date Paid","Running Balance (Auto)","Notes"]
    mc = wh(ws,h); title(ws,"💰 AIA – PAYMENT HISTORY",mc)
    for r in range(3,203):
        ws.cell(row=r,column=10).value = f'=IF(E{r}="","",E{r}-IF(H{r}="",0,H{r}))'
    ff(ws,2,mc); aw(ws,mc)

def _aia_stored(ws):
    h = ["Item Description","Location","Supplier/Invoice","Purchase Date",
         "Value","Verified By","Insurance Confirmed","Added to App Date",
         "Incorporated Date","Status"]
    mc = wh(ws,h); title(ws,"💰 AIA – STORED MATERIALS LOG",mc)
    dv(ws,"B3:B502","On-site,Off-site,In Transit")
    dv(ws,"G3:G502","Yes,No")
    dv(ws,"J3:J502","Stored,Installed,Released")
    ff(ws,2,mc); aw(ws,mc)

def _aia_lien(ws):
    h = ["Tier","Company Name","Invoice Amount","Conditional Waiver Issued",
         "Conditional Waiver Received","Payment Made (Date)","Payment Amount",
         "Unconditional Waiver Issued","Unconditional Waiver Received",
         "Status","Notes"]
    mc = wh(ws,h); title(ws,"💰 AIA – LIEN WAIVER TRACKER",mc)
    dv(ws,"A3:A502","GC,Sub,Supplier")
    dv(ws,"J3:J502","Pending,Received,Outstanding")
    ws.conditional_formatting.add("J3:J502",CellIsRule(operator='equal',formula=['"Received"'],fill=grn))
    ws.conditional_formatting.add("J3:J502",CellIsRule(operator='equal',formula=['"Outstanding"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _aia_retain(ws):
    h = ["App Number","Period","Gross Billing","Retainage %","Retainage Withheld",
         "Cumulative Retainage (Auto)","Retainage Released","Balance (Auto)",
         "Release Conditions Met","Notes"]
    mc = wh(ws,h); title(ws,"💰 AIA – RETAINAGE TRACKER",mc)
    for r in range(3,203):
        ws.cell(row=r,column=5).value = f'=IF(C{r}="","",C{r}*D{r}/100)'
        ws.cell(row=r,column=6).value = f'=IF(E{r}="","",IF(r>3,F{r-1},0)+E{r})'
        ws.cell(row=r,column=8).value = f'=IF(F{r}="","",F{r}-IF(G{r}="",0,G{r}))'
    dv(ws,"I3:I202","Yes,No")
    ff(ws,2,mc); aw(ws,mc)

def _aia_recon(ws):
    h = ["App Number","Amount Requested","Amount Certified","Diff Certified vs Req (Auto)",
         "Amount Paid","Diff Paid vs Certified (Auto)","Date Paid",
         "Days to Pay (Auto)","Contract Days","Compliance","Interest Due","Notes"]
    mc = wh(ws,h); title(ws,"💰 AIA – CERTIFIED VS PAID RECONCILIATION",mc)
    for r in range(3,203):
        ws.cell(row=r,column=4).value = f'=IF(B{r}="","",B{r}-C{r})'
        ws.cell(row=r,column=6).value = f'=IF(C{r}="","",C{r}-E{r})'
        ws.cell(row=r,column=8).value = f'=IF(G{r}="","",G{r}-F{r})'
        ws.cell(row=r,column=10).value = f'=IF(H{r}="","",IF(H{r}<=I{r},"On Time","Late"))'
    ws.conditional_formatting.add("J3:J202",CellIsRule(operator='equal',formula=['"On Time"'],fill=grn))
    ws.conditional_formatting.add("J3:J202",CellIsRule(operator='equal',formula=['"Late"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _aia_g702(ws):
    h = ["Line","Description","Completed This Period","Completed Previously",
         "Total Completed","Materials Stored","Total Completed & Stored",
         "% Complete","Balance to Finish"]
    mc = wh(ws,h); title(ws,"💰 AIA – G702 PRINT-READY FORM",mc)
    for r in range(3,53):
        ws.cell(row=r,column=5).value = f'=IF(C{r}="","",C{r}+D{r})'
        ws.cell(row=r,column=7).value = f'=IF(E{r}="","",E{r}+F{r})'
        ws.cell(row=r,column=8).value = f'=IF(G{r}="","",ROUND(G{r}/(G{r}+I{r})*100,1))'
    # Summary section
    r = 55
    labels = ["Original Contract Sum","Net Change by COs","Contract Sum to Date",
              "Total Completed & Stored","Retainage","Earned Less Retainage",
              "Less Previous Certificates","CURRENT PAYMENT DUE"]
    for i,lb in enumerate(labels):
        ws.cell(row=r+i,column=2,value=lb).font = Font(bold=True)
        for c in range(2,5):
            ws.cell(row=r+i,column=c).border = thin_border
    ff(ws,2,mc); aw(ws,mc)

def _aia_dash(ws):
    title(ws,"💰 AIA INVOICING DASHBOARD",10)
    for lc,lb,vc,fm in [
        ("B3","Cumulative Billed","B4","=SUM('AIA Schedule of Values'!I204:I204)"),
        ("D3","Retainage Held","D4","=SUM('AIA Retainage Tracker'!H3:H202)"),
        ("F3","Payments Received","F4","=SUM('AIA Payment History'!H3:H202)"),
        ("H3","Outstanding Balance","H4","=B4-F4"),
        ("J3","Apps Submitted","J4","=COUNTA('AIA Payment History'!A3:A202)")]:
        ws[lc]=lb; ws[lc].font=kpi_lbl; ws[vc]=fm; ws[vc].font=kpi_font
    aw(ws,10)

# ═════════════════════════════════════════════════════════════════════════
# SYSTEM 6: PROJECT TIMELINE
# ═════════════════════════════════════════════════════════════════════════
def _tl_cal(ws):
    h = ["Setting","Value","Notes"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – CALENDAR SETUP",mc)
    settings = [("Project Start Date",""),("Project End Date",""),
                ("Working Days/Week",5),("Working Hours/Day",8),
                ("Holidays List",""),("Reporting Period","Weekly"),
                ("Time Scale","Weeks"),("Baseline Date","")]
    for i,(k,v) in enumerate(settings):
        ws.cell(row=3+i,column=1,value=k).font = Font(bold=True)
        ws.cell(row=3+i,column=2,value=v)
    ff(ws,2,mc); aw(ws,mc)

def _tl_wbs(ws):
    h = ["WBS Code","Task ID","Task Name","WBS Level","Summary Task",
         "Milestone","Task Type","Deliverable","Responsible Party","Notes"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – WORK BREAKDOWN STRUCTURE",mc)
    dv(ws,"D3:D502","1-Phase,2-Stage,3-Task,4-Sub-task")
    dv(ws,"E3:E502","Yes,No"); dv(ws,"F3:F502","Yes,No")
    dv(ws,"G3:G502","Fixed Duration,Fixed Units,Fixed Work")
    ff(ws,2,mc); aw(ws,mc)

def _tl_task(ws):
    h = ["Task ID","WBS Code","Task Name","Phase","Description",
         "Assigned To","Resource Type","Planned Start","Planned End",
         "Planned Duration (Auto)","Actual Start","Actual End",
         "Actual Duration (Auto)","% Complete","Remaining Days (Auto)",
         "Est Completion (Auto)","Baseline Start","Baseline End",
         "Schedule Variance (Auto)","Predecessors","Dependency Type",
         "Lag/Lead Days","Priority","Budget","Cost to Date","Notes"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – TASK DETAIL",mc)
    for r in range(3,5003):
        ws.cell(row=r,column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"T-00000"))'
        ws.cell(row=r,column=10).value = f'=IF(OR(H{r}="",I{r}=""),"",I{r}-H{r})'
        ws.cell(row=r,column=13).value = f'=IF(OR(K{r}="",L{r}=""),"",L{r}-K{r})'
        ws.cell(row=r,column=15).value = f'=IF(OR(J{r}="",N{r}=""),"",ROUND(J{r}*(1-N{r}/100),0))'
        ws.cell(row=r,column=16).value = f'=IF(K{r}="","",K{r}+O{r})'
        ws.cell(row=r,column=19).value = f'=IF(OR(I{r}="",L{r}=""),"",I{r}-L{r})'
    dv(ws,"G3:G5002","Labor,Equipment,Subcontractor")
    dv(ws,"U3:U5002","FS,SS,FF,SF")
    dv(ws,"W3:W5002","Critical,High,Medium,Low")
    ff(ws,2,mc); aw(ws,mc)

def _tl_gantt(ws):
    h = ["Task ID","Task Name","Start Date","End Date","Duration","% Complete",
         "Assigned To","Predecessor"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – VISUAL GANTT CHART",mc)
    for r in range(3,103):
        ws.cell(row=r,column=5).value = f'=IF(OR(C{r}="",D{r}=""),"",D{r}-C{r})'
    # Week columns (26 half-year weeks for readability)
    base = datetime.date.today()
    for i in range(26):
        col = 9 + i
        d = base + datetime.timedelta(weeks=i)
        ws.cell(row=2,column=col,value=d.strftime("%d-%b"))
        ws.cell(row=2,column=col).font = Font(bold=True,size=8,color=WHITE)
        ws.cell(row=2,column=col).fill = hdr_fill
        ws.cell(row=2,column=col).alignment = ctr
        ws.column_dimensions[get_column_letter(col)].width = 5
    ff(ws,2,mc); aw(ws,mc)

def _tl_mile(ws):
    h = ["Milestone ID","Name","Description","Target Date","Actual/Forecast",
         "Variance (Auto)","Status","Owner","Dependencies","Completion Criteria","Notes"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – MILESTONE TRACKER",mc)
    for r in range(3,103):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"MS-000"))'
        ws.cell(row=r,column=6).value = f'=IF(OR(D{r}="",E{r}=""),"",E{r}-D{r})'
    dv(ws,"G3:G102","Not Started,On Track,At Risk,Delayed,Achieved")
    ws.conditional_formatting.add("G3:G102",CellIsRule(operator='equal',formula=['"Achieved"'],fill=grn))
    ws.conditional_formatting.add("G3:G102",CellIsRule(operator='equal',formula=['"Delayed"'],fill=red))
    ws.conditional_formatting.add("G3:G102",CellIsRule(operator='equal',formula=['"At Risk"'],fill=yel))
    ws.conditional_formatting.add("G3:G102",CellIsRule(operator='equal',formula=['"On Track"'],fill=blu))
    ff(ws,2,mc); aw(ws,mc)

def _tl_res(ws):
    h = ["Resource Name","Type","Role/Trade","Availability (Hrs/Day)",
         "Week 1 Hrs","Week 2 Hrs","Week 3 Hrs","Week 4 Hrs",
         "Total Hrs (Auto)","Capacity/Week","Utilization % (Auto)",
         "Over-allocated","Cost/Hr","Total Cost (Auto)"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – RESOURCE LOADING",mc)
    for r in range(3,203):
        ws.cell(row=r,column=9).value = f'=IF(E{r}="","",SUM(E{r}:H{r}))'
        ws.cell(row=r,column=11).value = f'=IF(OR(I{r}="",J{r}=""),"",ROUND(I{r}/J{r}*100,1))'
        ws.cell(row=r,column=12).value = f'=IF(K{r}="","",IF(K{r}>100,"⚠ YES","No"))'
        ws.cell(row=r,column=14).value = f'=IF(I{r}="","",I{r}*M{r})'
    dv(ws,"B3:B202","Labor,Equipment,Material,Subcontractor")
    ws.conditional_formatting.add("L3:L202",CellIsRule(operator='containsText',formula=['"YES"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _tl_base(ws):
    h = ["Task ID","Task Name","Baseline Start","Baseline End","Baseline Duration",
         "Current Start","Current End","Current Duration",
         "Actual Start","Actual End","Schedule Variance (Auto)",
         "SPI (Auto)","Status","Recovery Plan","Notes"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – BASELINE & TRACKING",mc)
    for r in range(3,2003):
        ws.cell(row=r,column=11).value = f'=IF(OR(E{r}="",H{r}=""),"",H{r}-E{r})'
        ws.cell(row=r,column=12).value = f'=IF(OR(L{r}="",E{r}=""),"","N/A")'
    dv(ws,"M3:M2002","Ahead,On Track,Slightly Delayed,Significantly Delayed,Critical")
    ws.conditional_formatting.add("M3:M2002",CellIsRule(operator='equal',formula=['"On Track"'],fill=grn))
    ws.conditional_formatting.add("M3:M2002",CellIsRule(operator='equal',formula=['"Critical"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _tl_look(ws):
    h = ["Week","Task Name","Responsible Party","% Complete Start",
         "% Complete End","Resources Needed","Materials Needed",
         "Equipment Needed","Constraints/Issues","PPC Target","PPC Actual",
         "Non-completion Reasons"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – LOOK-AHEAD SCHEDULE",mc)
    ff(ws,2,mc); aw(ws,mc)

def _tl_risk(ws):
    h = ["Risk ID","Description","Tasks Affected","Probability %",
         "Impact (Days)","Risk Score (Auto)","Risk Level",
         "Early Warning","Mitigation","Contingency","Owner","Status","Notes"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – SCHEDULE RISK REGISTER",mc)
    for r in range(3,503):
        ws.cell(row=r,column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"RSK-000"))'
        ws.cell(row=r,column=6).value = f'=IF(OR(D{r}="",E{r}=""),"",D{r}*E{r}/100)'
        ws.cell(row=r,column=7).value = f'=IF(F{r}="","",IF(F{r}>20,"Critical",IF(F{r}>10,"High",IF(F{r}>5,"Medium","Low"))))'
    dv(ws,"L3:L502","Open,Monitoring,Mitigated,Occurred,Closed")
    ws.conditional_formatting.add("G3:G502",CellIsRule(operator='equal',formula=['"Critical"'],fill=red))
    ws.conditional_formatting.add("G3:G502",CellIsRule(operator='equal',formula=['"High"'],fill=yel))
    ff(ws,2,mc); aw(ws,mc)

def _tl_crit(ws):
    h = ["Task ID","Task Name","Early Start","Early Finish","Late Start","Late Finish",
         "Total Float (Auto)","Free Float","Critical Path","Near-Critical","Notes"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – CRITICAL PATH ANALYSIS",mc)
    for r in range(3,2003):
        ws.cell(row=r,column=7).value = f'=IF(OR(E{r}="",C{r}=""),"",E{r}-C{r})'
        ws.cell(row=r,column=9).value = f'=IF(G{r}="","",IF(G{r}=0,"YES","No"))'
        ws.cell(row=r,column=10).value = f'=IF(G{r}="","",IF(AND(G{r}>0,G{r}<=3),"YES","No"))'
    ws.conditional_formatting.add("I3:I2002",CellIsRule(operator='equal',formula=['"YES"'],fill=red))
    ff(ws,2,mc); aw(ws,mc)

def _tl_prog(ws):
    h = ["Period","Date Reported","Target % Complete","Actual % Complete",
         "Cumulative Target","Cumulative Actual","Schedule Variance %",
         "Activities Completed","Activities In Progress","Not Started (Should Have)",
         "Milestones Achieved","Issues & Risks","Corrective Actions",
         "Forecast Completion"]
    mc = wh(ws,h); title(ws,"📅 TIMELINE – PROGRESS REPORTING",mc)
    for r in range(3,203):
        ws.cell(row=r,column=7).value = f'=IF(OR(E{r}="",F{r}=""),"",F{r}-E{r})'
    ff(ws,2,mc); aw(ws,mc)

def _tl_dash(ws):
    title(ws,"📅 TIMELINE DASHBOARD",10)
    for lc,lb,vc,fm in [
        ("B3","Total Tasks","B4","=COUNTA('TL Task Detail'!C3:C5002)"),
        ("D3","Tasks Complete","D4",'=COUNTIF(\'TL Task Detail\'!N3:N5002,"100")'),
        ("F3","Overall % Complete","F4","=IFERROR(ROUND(AVERAGE('TL Task Detail'!N3:N5002),1),0)"),
        ("H3","Milestones Achieved","H4",'=COUNTIF(\'TL Milestone Tracker\'!G3:G102,"Achieved")'),
        ("J3","Delayed Milestones","J4",'=COUNTIF(\'TL Milestone Tracker\'!G3:G102,"Delayed")'),
        ("L3","Open Risks","L4",'=COUNTIF(\'TL Risk Register\'!L3:L502,"Open")')]:
        ws[lc]=lb; ws[lc].font=kpi_lbl; ws[vc]=fm; ws[vc].font=kpi_font
    aw(ws,10)

# ═════════════════════════════════════════════════════════════════════════
# MASTER DASHBOARD
# ═════════════════════════════════════════════════════════════════════════
def _master_dash(ws, wb):
    mc = 12
    ws.merge_cells("A1:L1")
    ws["A1"] = "🏗️  COMPREHENSIVE CONSTRUCTION SPREADSHEET SUITE – MASTER DASHBOARD"
    ws["A1"].font = Font(bold=True,color=WHITE,size=18)
    ws["A1"].fill = PatternFill(start_color="1F3864",end_color="1F3864",fill_type='solid')
    ws["A1"].alignment = Alignment(horizontal='center',vertical='center')
    ws.row_dimensions[1].height = 45

    ws.merge_cells("A2:L2")
    ws["A2"] = "Company: Your Construction Co.    |    6 Integrated Systems    |    69 Sheets Total"
    ws["A2"].font = Font(italic=True,color=MID_BLUE,size=11)
    ws["A2"].alignment = Alignment(horizontal='center')

    # KPI Row
    kpis = [
        ("B4","Active Projects","B5",'=COUNTIF(\'CM Project Master\'!T3:T202,"Green")+COUNTIF(\'CM Project Master\'!T3:T202,"Amber")+COUNTIF(\'CM Project Master\'!T3:T202,"Red")'),
        ("D4","Active Jobs (JC)","D5",'=COUNTIF(\'JC Job Setup\'!N3:N502,"Active")'),
        ("F4","Open RFIs","F4_v",'=COUNTIF(\'CM RFI Log\'!O3:O5002,"Open")'),
        ("H4","Pending COs","H5",'=COUNTIF(\'CM Change Order Log\'!K3:K2002,"Pending")'),
        ("J4","Estimate Value","J5","='Est Markup & Profit'!B9"),
        ("L4","Tasks On Track","L4_v",'=COUNTIF(\'TL Baseline Tracking\'!M3:M2002,"On Track")'),
    ]
    # Fix tuples
    kpi_items = [
        ("B4","Active Projects","B5",'=COUNTA(\'CM Project Master\'!B3:B202)'),
        ("D4","Active Jobs","D5",'=COUNTIF(\'JC Job Setup\'!N3:N502,"Active")'),
        ("F4","Open RFIs","F5",'=COUNTIF(\'CM RFI Log\'!O3:O5002,"Open")'),
        ("H4","Pending COs","H5",'=COUNTIF(\'CM Change Order Log\'!K3:K2002,"Pending")'),
        ("J4","Estimate Total","J5","='Est Markup & Profit'!B9"),
        ("L4","Tasks Complete","L5",'=COUNTIF(\'TL Task Detail\'!N3:N5002,">=90")'),
    ]
    for lc,lb,vc,fm in kpi_items:
        ws[lc]=lb; ws[lc].font=kpi_lbl; ws[lc].alignment=Alignment(horizontal='center')
        ws[lc].fill=PatternFill(start_color=LIGHT_BLUE,end_color=LIGHT_BLUE,fill_type='solid')
        ws[vc]=fm; ws[vc].font=kpi_font; ws[vc].alignment=Alignment(horizontal='center')

    # Module links
    r = 8
    ws.cell(row=r,column=2,value="📋 SYSTEM QUICK LINKS").font = sec_font
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6)
    ws.cell(row=r,column=2).fill = sec_fill
    r += 1

    systems = [
        ("📊 1. Construction Estimate","Est Cover Sheet","Est CSI Division Breakdown","Est Detailed Line Items","Est Dashboard"),
        ("📋 2. Contractor Bid","Bid Cover Page","Bid Itemization","Bid Pricing Summary","Bid Tracking Log"),
        ("💼 3. Job Costing","JC Job Setup","JC Cost Entry","JC Cost Summary","JC Dashboard"),
        ("🏢 4. Construction Mgmt","CM Project Master","CM RFI Log","CM Change Order Log","CM Punch List"),
        ("💰 5. AIA Invoicing","AIA Contract Info","AIA Schedule of Values","AIA G702 Form","AIA Dashboard"),
        ("📅 6. Project Timeline","TL Calendar Setup","TL Task Detail","TL Gantt Chart","TL Dashboard"),
    ]
    for sys in systems:
        ws.cell(row=r,column=2,value=sys[0]).font = Font(bold=True,color=DARK_BLUE)
        ws.cell(row=r,column=2).fill = gry
        for ci,name in enumerate(sys[1:],3):
            ws.cell(row=r,column=ci,value=f"→ {name}")
            ws.cell(row=r,column=ci).font = Font(color=MID_BLUE,underline='single')
        r += 1

    # Financial summary
    r += 1
    ws.cell(row=r,column=2,value="💰 FINANCIAL SUMMARY").font = sec_font
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    ws.cell(row=r,column=2).fill = sec_fill
    r += 1
    items = [
        ("Total Estimate Value","='Est Markup & Profit'!B9"),
        ("Total Contract Value (All Jobs)","=SUM('JC Job Setup'!G3:G502)"),
        ("Total Costs JTD","=SUM('JC Cost Summary'!H3:H1002)"),
        ("Cumulative Billed (AIA)","=SUM('AIA Schedule of Values'!I3:I203)"),
    ]
    for lb,fm in items:
        ws.cell(row=r,column=2,value=lb).font = Font(bold=True)
        ws.cell(row=r,column=3).value = fm
        ws.cell(row=r,column=3).number_format = '#,##0'
        for c in range(2,4): ws.cell(row=r,column=c).border = thin_border
        r += 1

    # Schedule summary
    r += 1
    ws.cell(row=r,column=2,value="📅 SCHEDULE SUMMARY").font = sec_font
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    ws.cell(row=r,column=2).fill = sec_fill
    r += 1
    items2 = [
        ("Total Tasks",'=COUNTA(\'TL Task Detail\'!C3:C5002)'),
        ("Tasks Complete",'=COUNTIF(\'TL Task Detail\'!N3:N5002,">=90")'),
        ("Milestones Achieved",'=COUNTIF(\'TL Milestone Tracker\'!G3:G102,"Achieved")'),
        ("Delayed Milestones",'=COUNTIF(\'TL Milestone Tracker\'!G3:G102,"Delayed")'),
        ("Open Risks",'=COUNTIF(\'TL Risk Register\'!L3:L502,"Open")'),
    ]
    for lb,fm in items2:
        ws.cell(row=r,column=2,value=lb).font = Font(bold=True)
        ws.cell(row=r,column=3).value = fm
        for c in range(2,4): ws.cell(row=r,column=c).border = thin_border
        r += 1

    for col in range(1,13):
        ws.column_dimensions[get_column_letter(col)].width = 20
    ws.freeze_panes = "A3"

# ═════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("🏗️ Building Comprehensive Construction Spreadsheet Suite...")
    wb = build()
    out = "Construction_Management_Suite.xlsx"
    wb.save(out)
    print(f"✅ Created: {out}")
    print(f"   Sheets: {len(wb.sheetnames)}")
    print(f"   {', '.join(wb.sheetnames)}")
