#!/usr/bin/env python3
"""
Comprehensive Home Renovation Management System – Excel Generator
==================================================================
Creates a fully-structured multi-sheet Excel workbook covering:
  1.  Project Management (Master, Phases, Tasks, Gantt)
  2.  Client Management (Master, Communication, Satisfaction)
  3.  Budget & Cost Management (Planning, Tracking, Dashboard, Change Orders)
  4.  Contractor & Vendor Management (Contractors, Vendors, Performance, Payments)
  5.  Materials Management (Master, Estimation, POs, Inventory, Delivery Log)
  6.  Labor Management (Workers, Attendance, Cost Calc, Productivity)
  7.  Room/Area Tracking (Room Master, Checklist, Cost Summary)
  8.  Design & Planning (Requirements, Measurements, Material Selection)
  9.  Equipment & Tool Management (Inventory, Usage, Maintenance)
  10. Quality Control & Inspection (Standards, Log, Defect/Snagging)
  11. Safety & Compliance (Checklist, Incidents, Permits)
  12. Quotation & Invoicing (Quotations, Invoices, Receipts)
  13. Financial Management (Income, Expense, P&L, Cash Flow, Tax)
  14. Document Management (Document Register, Contract Register)
  15. Warranty & After-Sales (Warranty Register, Service Requests)
  16. Reporting & Analytics (Project Status, Weekly Report, KPI Dashboard)
  +   Main Dashboard (Home)
  +   Config (Hidden – dropdown lists, system settings)

Total: 56 data sheets + 1 dashboard + 1 config = 58 sheets

Version: 1.0
"""

import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule

# ── Colour Palette ──────────────────────────────────────────────────────
DARK_BLUE   = "1F3864"
MID_BLUE    = "2E75B6"
LIGHT_BLUE  = "D6E4F0"
WHITE       = "FFFFFF"
DARK_GREEN  = "375623"
LIGHT_GREEN = "C6EFCE"
RED         = "FF0000"
LIGHT_RED   = "FFC7CE"
YELLOW_FILL = "FFEB9C"
ORANGE      = "ED7D31"
GREY_BG     = "F2F2F2"
HEADER_GREY = "D9E2F3"
PURPLE      = "7030A0"
TEAL        = "20B2AA"

# ── Reusable Styles ─────────────────────────────────────────────────────
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

header_font = Font(name='Calibri', bold=True, color=WHITE, size=11)
header_fill = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type='solid')
title_font  = Font(name='Calibri', bold=True, color=WHITE, size=16)
title_fill  = PatternFill(start_color=MID_BLUE, end_color=MID_BLUE, fill_type='solid')
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

green_fill  = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type='solid')
red_fill    = PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')
yellow_cond = PatternFill(start_color=YELLOW_FILL, end_color=YELLOW_FILL, fill_type='solid')
grey_cond   = PatternFill(start_color=GREY_BG, end_color=GREY_BG, fill_type='solid')
section_font = Font(name='Calibri', bold=True, size=13, color=DARK_BLUE)
section_fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')
kpi_font     = Font(name='Calibri', bold=True, size=26, color=DARK_BLUE)
kpi_label    = Font(name='Calibri', bold=True, size=10, color=MID_BLUE)


# ═════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═════════════════════════════════════════════════════════════════════════

def style_header_row(ws, row, max_col, font=None, fill=None):
    f = font or header_font
    fi = fill or header_fill
    for col in range(1, max_col + 1):
        c = ws.cell(row=row, column=col)
        c.font = f; c.fill = fi; c.alignment = center_align; c.border = thin_border


def auto_width(ws, max_col, mn=10, mx=35):
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        ml = mn
        for row in ws.iter_rows(min_col=col, max_col=col, values_only=False):
            for cell in row:
                if cell.value:
                    ml = max(ml, min(len(str(cell.value)) + 4, mx))
        ws.column_dimensions[letter].width = ml


def add_title_row(ws, title, max_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    c = ws.cell(row=1, column=1, value=title)
    c.font = title_font; c.fill = title_fill
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 36


def add_dv(ws, cell_range, formula_list):
    """Add dropdown data validation."""
    dv = DataValidation(type="list", formula1=f'"{formula_list}"', allow_blank=True)
    dv.error = "Please select a valid option"
    dv.errorTitle = "Invalid Entry"
    ws.add_data_validation(dv)
    dv.add(cell_range)


def write_headers(ws, headers, row=2):
    for ci, h in enumerate(headers, 1):
        ws.cell(row=row, column=ci, value=h)
    style_header_row(ws, row, len(headers))
    return len(headers)


def freeze_and_filter(ws, freeze_row, max_col):
    ws.freeze_panes = ws.cell(row=freeze_row + 1, column=1)
    ws.auto_filter.ref = f"A{freeze_row}:{get_column_letter(max_col)}{freeze_row + 500}"


# ═════════════════════════════════════════════════════════════════════════
# MAIN BUILDER
# ═════════════════════════════════════════════════════════════════════════

def build_workbook():
    wb = Workbook()
    default = wb.active

    # ── 0. CONFIG ───────────────────────────────────────────────────────
    _config(wb.create_sheet("Config_Dropdowns"))

    # ── 1. PROJECT MODULE ───────────────────────────────────────────────
    _project_master(wb.create_sheet("Project Master Data"))
    _project_phases(wb.create_sheet("Project Phases"))
    _task_mgmt(wb.create_sheet("Task Management"))
    _gantt(wb.create_sheet("Gantt Chart"))

    # ── 2. CLIENT MODULE ────────────────────────────────────────────────
    _client_master(wb.create_sheet("Client Master Data"))
    _client_comm(wb.create_sheet("Client Communication Log"))
    _client_sat(wb.create_sheet("Client Satisfaction"))

    # ── 3. BUDGET & COST ────────────────────────────────────────────────
    _budget_planning(wb.create_sheet("Budget Planning"))
    _cost_tracking(wb.create_sheet("Cost Tracking"))
    _budget_dashboard(wb.create_sheet("Budget Dashboard"))
    _change_orders(wb.create_sheet("Change Orders"))

    # ── 4. CONTRACTOR & VENDOR ──────────────────────────────────────────
    _contractor_master(wb.create_sheet("Contractor Master Data"))
    _vendor_master(wb.create_sheet("Vendor Master Data"))
    _contractor_perf(wb.create_sheet("Contractor Performance"))
    _contractor_pay(wb.create_sheet("Contractor Payments"))

    # ── 5. MATERIALS ────────────────────────────────────────────────────
    _material_master(wb.create_sheet("Material Master List"))
    _material_est(wb.create_sheet("Material Estimation"))
    _purchase_orders(wb.create_sheet("Purchase Orders"))
    _material_inventory(wb.create_sheet("Material Inventory"))
    _delivery_log(wb.create_sheet("Material Delivery Log"))

    # ── 6. LABOR ────────────────────────────────────────────────────────
    _worker_master(wb.create_sheet("Worker Master Data"))
    _worker_attendance(wb.create_sheet("Worker Attendance"))
    _labor_cost(wb.create_sheet("Labor Cost Calculation"))
    _productivity(wb.create_sheet("Productivity Tracker"))

    # ── 7. ROOM-WISE TRACKING ───────────────────────────────────────────
    _room_master(wb.create_sheet("Room Area Master"))
    _room_checklist(wb.create_sheet("Room Work Checklist"))
    _room_cost(wb.create_sheet("Room Cost Summary"))

    # ── 8. DESIGN & PLANNING ────────────────────────────────────────────
    _design_req(wb.create_sheet("Design Requirements"))
    _measurements(wb.create_sheet("Measurements Specs"))
    _material_selection(wb.create_sheet("Material Selection Board"))

    # ── 9. EQUIPMENT ────────────────────────────────────────────────────
    _equip_inventory(wb.create_sheet("Equipment Inventory"))
    _equip_usage(wb.create_sheet("Equipment Usage Log"))
    _equip_maint(wb.create_sheet("Equipment Maintenance"))

    # ── 10. QUALITY CONTROL ─────────────────────────────────────────────
    _quality_standards(wb.create_sheet("Quality Standards"))
    _inspection_log(wb.create_sheet("Inspection Log"))
    _snagging(wb.create_sheet("Defect Snagging List"))

    # ── 11. SAFETY & COMPLIANCE ─────────────────────────────────────────
    _safety_checklist(wb.create_sheet("Safety Checklist"))
    _incident_log(wb.create_sheet("Incident Accident Log"))
    _permits(wb.create_sheet("Permits Compliance"))

    # ── 12. QUOTATION & INVOICING ───────────────────────────────────────
    _quotation(wb.create_sheet("Quotation Builder"))
    _invoice_mgmt(wb.create_sheet("Invoice Management"))
    _payment_receipts(wb.create_sheet("Payment Receipts"))

    # ── 13. FINANCIAL MANAGEMENT ────────────────────────────────────────
    _income_tracker(wb.create_sheet("Income Tracker"))
    _expense_tracker(wb.create_sheet("Expense Tracker"))
    _pnl_project(wb.create_sheet("Profit Loss Project"))
    _cash_flow(wb.create_sheet("Cash Flow Tracker"))
    _tax_compliance(wb.create_sheet("Tax Compliance"))

    # ── 14. DOCUMENT MANAGEMENT ─────────────────────────────────────────
    _doc_register(wb.create_sheet("Document Register"))
    _contract_register(wb.create_sheet("Contract Register"))

    # ── 15. WARRANTY & AFTER-SALES ──────────────────────────────────────
    _warranty(wb.create_sheet("Warranty Register"))
    _after_sales(wb.create_sheet("After-Sales Service"))

    # ── 16. REPORTING & ANALYTICS ───────────────────────────────────────
    _project_status(wb.create_sheet("Project Status Report"))
    _weekly_progress(wb.create_sheet("Weekly Progress Report"))
    _kpi_dashboard(wb.create_sheet("KPI Dashboard"))

    # ── MAIN DASHBOARD (insert at position 1) ──────────────────────────
    _main_dashboard(wb.create_sheet("DASHBOARD", 1), wb)

    wb.remove(default)
    wb["Config_Dropdowns"].sheet_state = 'hidden'
    return wb


# ═════════════════════════════════════════════════════════════════════════
# 0. CONFIG
# ═════════════════════════════════════════════════════════════════════════

def _config(ws):
    add_title_row(ws, "⚙️ CONFIG – Dropdown Lists, Cost Rates, Tax & System Settings", 10)

    lists = {
        "ProjectType": ["Full Renovation","Partial","Room-specific","Extension","Repair"],
        "PropertyType": ["Apartment","Villa","Office","Commercial","Bungalow","Studio","Townhouse"],
        "ProjectStatus": ["Planning","In Progress","On Hold","Completed","Cancelled"],
        "Priority": ["High","Medium","Low"],
        "PhaseName": ["Demolition","Foundation","Structural","Plumbing","Electrical",
                       "Flooring","Painting","Carpentry","Tiling","Finishing","Landscaping"],
        "PhaseStatus": ["Not Started","In Progress","Completed","Delayed"],
        "TaskStatus": ["Not Started","In Progress","Completed","Blocked","Cancelled"],
        "ClientType": ["Individual","Family","Corporate","Real Estate Company"],
        "Gender": ["Male","Female","Other"],
        "CommPref": ["Phone","Email","WhatsApp","In Person"],
        "ClientStatus": ["Active","Inactive","Prospect","Blacklisted"],
        "CommType": ["Call","Email","WhatsApp","Meeting","Site Visit"],
        "LogStatus": ["Open","Closed","Pending"],
        "BudgetCategory": ["Labor","Materials","Equipment","Permits","Design","Contingency","Miscellaneous"],
        "ExpenseCategory": ["Labor","Material","Equipment","Transport","Permit","Design","Marketing","Office","Utilities","Misc"],
        "PaymentMode": ["Cash","Cheque","Bank Transfer","Online","UPI"],
        "PaymentStatus": ["Paid","Pending","Partial","Disputed","Overdue"],
        "ContractorType": ["General Contractor","Electrician","Plumber","Carpenter","Painter",
                           "Mason","Tiler","HVAC","Interior Designer","Architect","Landscaper"],
        "ContractorStatus": ["Active","Inactive","Blacklisted","On Hold"],
        "VendorType": ["Material Supplier","Equipment Rental","Tool Supplier","Furniture",
                       "Appliances","Paint","Tile","Electrical Supplies","Plumbing Supplies"],
        "PaymentTerms": ["Net 15","Net 30","Advance","COD","Milestone Based"],
        "MaterialCategory": ["Cement","Steel","Sand","Bricks","Paint","Tiles","Wood","Glass",
                             "Electrical","Plumbing","Fixtures","Hardware","Insulation","Adhesives"],
        "UnitOfMeasure": ["Bags","Kg","Tons","Liters","Sq.Ft","Sq.M","Pieces","Meters","Rolls","Boxes","Sets"],
        "QualityGrade": ["A","B","C","Premium","Standard","Economy"],
        "ProcureStatus": ["Not Started","Ordered","In Transit","Received","Cancelled"],
        "POStatus": ["Draft","Sent","Confirmed","Partially Delivered","Fully Delivered","Cancelled"],
        "WorkerTrade": ["Mason","Carpenter","Painter","Electrician","Plumber","Helper",
                        "Supervisor","Welder","Tiler","HVAC Tech","Landscaper"],
        "SkillLevel": ["Unskilled","Semi-skilled","Skilled","Expert"],
        "RoomType": ["Living Room","Bedroom 1","Bedroom 2","Bedroom 3","Kitchen","Bathroom 1",
                     "Bathroom 2","Dining Room","Garage","Garden","Terrace","Basement",
                     "Office Room","Balcony","Hallway","Utility Room"],
        "RenoType": ["Full","Partial","Cosmetic"],
        "DesignStyle": ["Modern","Contemporary","Traditional","Minimalist","Industrial","Rustic","Scandinavian"],
        "DesignStatus": ["Concept","Draft","Revised","Final Approved"],
        "EquipCategory": ["Owned","Rented"],
        "EquipType": ["Power Tool","Heavy Machinery","Hand Tool","Safety Equipment",
                      "Scaffolding","Generator","Mixer","Compressor"],
        "EquipCondition": ["New","Good","Fair","Needs Repair","Retired"],
        "EquipStatus": ["Available","In Use","Under Repair","Lost"],
        "MaintType": ["Routine","Repair","Replacement"],
        "MaintStatus": ["Scheduled","Completed","Overdue"],
        "InspectionType": ["In-progress","Stage Completion","Final","Client","Municipal","Safety"],
        "InspectionResult": ["Pass","Fail","Conditional Pass"],
        "DefectSeverity": ["Critical","Major","Minor","Cosmetic"],
        "SnagStatus": ["Open","In Progress","Fixed","Closed","Rejected"],
        "SafetyCompliance": ["Compliant","Non-compliant","N/A"],
        "IncidentType": ["Near Miss","First Aid","Medical Treatment","Lost Time","Fatal"],
        "IncidentCause": ["Human Error","Equipment Failure","Environmental","Negligence"],
        "PermitType": ["Building Permit","Electrical Permit","Plumbing Permit","Demolition Permit",
                       "Fire Safety","Environmental Clearance","HOA Approval"],
        "PermitStatus": ["Pending","Approved","Rejected","Expired","Renewed"],
        "QuotationStatus": ["Draft","Sent","Accepted","Rejected","Revised"],
        "InvoiceStatus": ["Draft","Sent","Paid","Partially Paid","Overdue","Disputed"],
        "IncomeSource": ["Advance","Progress Payment","Final Payment","Retention Release","Variation Payment"],
        "ContractType": ["Main Contract","Subcontract","Supplier Agreement","Design Contract","Labor Contract"],
        "ContractStatus": ["Active","Completed","Terminated","Disputed"],
        "DocType": ["Contract","Drawing","Permit","Invoice","Photo","Report","Quotation",
                    "Warranty","Insurance","Specification","BOQ"],
        "DocStatus": ["Draft","Active","Superseded","Archived"],
        "AccessLevel": ["Public","Internal","Confidential"],
        "WarrantyCategory": ["Roofing","Plumbing","Electrical","Painting","Flooring","Appliances","Structure"],
        "WarrantyType": ["Workmanship","Material","Manufacturer"],
        "ServiceCategory": ["Repair","Replacement","Inspection","Warranty Claim"],
        "TrafficLight": ["Red","Amber","Green"],
        "YesNo": ["Yes","No"],
        "QualityCheckResult": ["Pass","Fail","Rework Needed"],
        "RoomCheckStatus": ["Pending","In Progress","Done","Skipped"],
    }

    row = 3
    for key, values in lists.items():
        ws.cell(row=row, column=1, value=key).font = Font(bold=True, color=DARK_BLUE)
        for i, v in enumerate(values):
            ws.cell(row=row, column=i + 2, value=v)
        row += 1

    # System settings
    row += 2
    ws.cell(row=row, column=1, value="SYSTEM SETTINGS").font = Font(bold=True, size=14, color=DARK_BLUE)
    row += 1
    settings = [
        ("Company Name", "Your Renovation Company"),
        ("Company Address", "123 Builder Street, City, State"),
        ("Company Phone", "+91-XXXXXXXXXX"),
        ("Company Email", "info@yourcompany.com"),
        ("Tax Rate (%)", 18),
        ("Standard Payment Terms", "Net 30"),
        ("Default Warranty Period (Months)", 12),
        ("Safety Inspection Frequency (Days)", 7),
        ("Material Reorder Alert Days", 5),
        ("Currency Symbol", "₹"),
        ("Working Hours Per Day", 8),
        ("Overtime Multiplier", 1.5),
    ]
    for k, v in settings:
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=v)
        row += 1

    # Tax Rates table
    row += 2
    ws.cell(row=row, column=1, value="TAX RATES").font = Font(bold=True, size=14, color=DARK_BLUE)
    row += 1
    tax_data = [("Labor", 0), ("Materials (Standard)", 18), ("Materials (Essential)", 5),
                ("Equipment Rental", 18), ("Design Services", 18), ("Permits/Fees", 0),
                ("Transport", 5), ("Miscellaneous", 18)]
    for cat, rate in tax_data:
        ws.cell(row=row, column=1, value=cat)
        ws.cell(row=row, column=2, value=f"{rate}%")
        row += 1

    # Cost rates reference
    row += 2
    ws.cell(row=row, column=1, value="STANDARD COST RATES (Reference)").font = Font(bold=True, size=14, color=DARK_BLUE)
    row += 1
    rates = [
        ("Mason (Daily)", 1500), ("Carpenter (Daily)", 1800), ("Painter (Daily)", 1200),
        ("Electrician (Daily)", 1500), ("Plumber (Daily)", 1400), ("Helper (Daily)", 800),
        ("Supervisor (Daily)", 2500), ("Welder (Daily)", 1600), ("Tiler (Daily)", 1500),
        ("Cement (per bag 50kg)", 400), ("Steel (per kg)", 70), ("Sand (per cu.ft)", 80),
        ("Bricks (per 1000)", 8000), ("Paint - Interior (per liter)", 350),
        ("Paint - Exterior (per liter)", 450), ("Vitrified Tiles (per sq.ft)", 60),
        ("Wood - Teak (per cu.ft)", 3500), ("Electrical Wiring (per point)", 800),
    ]
    for item, rate in rates:
        ws.cell(row=row, column=1, value=item)
        ws.cell(row=row, column=2, value=rate)
        ws.cell(row=row, column=2).number_format = '#,##0'
        row += 1

    auto_width(ws, 10)


# ═════════════════════════════════════════════════════════════════════════
# 1. PROJECT MODULE
# ═════════════════════════════════════════════════════════════════════════

def _project_master(ws):
    headers = [
        "Project ID", "Project Name", "Project Type", "Property Address",
        "Property Type", "Property Size (Sq.Ft)", "No. of Rooms", "No. of Floors",
        "Start Date", "Planned End Date", "Actual End Date", "Duration (Days)",
        "Status", "Project Manager", "Client ID", "Total Budget",
        "Total Spent (Auto)", "Completion % (Auto)", "Priority", "Notes / Description"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🏠 PROJECT MASTER DATA", mc)

    for r in range(3, 503):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"PRJ0000"))'
        # Duration
        ws.cell(row=r, column=12).value = f'=IF(OR(I{r}="",J{r}=""),"",J{r}-I{r})'
        # Total Spent (sum of cost tracking for this project)
        ws.cell(row=r, column=17).value = (
            f'=IF(A{r}="","",SUMIFS(\'Cost Tracking\'!M:M,\'Cost Tracking\'!B:B,A{r}))'
        )
        # Completion %
        ws.cell(row=r, column=18).value = (
            f'=IF(A{r}="","",IFERROR(ROUND(AVERAGEIF(\'Project Phases\'!B:B,A{r},\'Project Phases\'!N:N),1),0))'
        )

    add_dv(ws, "C3:C502", "Full Renovation,Partial,Room-specific,Extension,Repair")
    add_dv(ws, "E3:E502", "Apartment,Villa,Office,Commercial,Bungalow,Studio,Townhouse")
    add_dv(ws, "M3:M502", "Planning,In Progress,On Hold,Completed,Cancelled")
    add_dv(ws, "S3:S502", "High,Medium,Low")

    ws.conditional_formatting.add("M3:M502",
        CellIsRule(operator='equal', formula=['"Completed"'], fill=green_fill))
    ws.conditional_formatting.add("M3:M502",
        CellIsRule(operator='equal', formula=['"In Progress"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))
    ws.conditional_formatting.add("M3:M502",
        CellIsRule(operator='equal', formula=['"On Hold"'], fill=yellow_cond))
    ws.conditional_formatting.add("M3:M502",
        CellIsRule(operator='equal', formula=['"Cancelled"'], fill=red_fill))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _project_phases(ws):
    headers = [
        "Phase ID", "Project ID", "Phase Name", "Phase Description",
        "Planned Start", "Planned End", "Actual Start", "Actual End",
        "Planned Duration", "Actual Duration", "Status",
        "Assigned Contractor", "Contractor ID",
        "Phase Budget", "Phase Actual Cost", "Completion %",
        "Dependencies", "Remarks"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📋 PROJECT PHASES", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"PHS0000"))'
        ws.cell(row=r, column=9).value = f'=IF(OR(E{r}="",F{r}=""),"",F{r}-E{r})'
        ws.cell(row=r, column=10).value = f'=IF(OR(G{r}="",H{r}=""),"",H{r}-G{r})'
        # Completion % (auto from task completion)
        ws.cell(row=r, column=16).value = (
            f'=IF(A{r}="","",IFERROR(ROUND(COUNTIFS(\'Task Management\'!C:C,A{r},\'Task Management\'!N:N,"Completed")/'
            f'MAX(COUNTIF(\'Task Management\'!C:C,A{r}),1)*100,1),0))'
        )

    add_dv(ws, "C3:C2002", "Demolition,Foundation,Structural,Plumbing,Electrical,Flooring,Painting,Carpentry,Tiling,Finishing,Landscaping")
    add_dv(ws, "K3:K2002", "Not Started,In Progress,Completed,Delayed")

    ws.conditional_formatting.add("K3:K2002",
        CellIsRule(operator='equal', formula=['"Completed"'], fill=green_fill))
    ws.conditional_formatting.add("K3:K2002",
        CellIsRule(operator='equal', formula=['"Delayed"'], fill=red_fill))
    ws.conditional_formatting.add("K3:K2002",
        CellIsRule(operator='equal', formula=['"In Progress"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _task_mgmt(ws):
    headers = [
        "Task ID", "Project ID", "Phase ID", "Task Name", "Task Description",
        "Assigned To", "Worker/Contractor ID", "Priority",
        "Planned Start", "Planned End", "Actual Start", "Actual End",
        "Hours Estimated", "Hours Actual", "Status",
        "Blocker / Issue", "% Complete", "Remarks"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "✅ TASK MANAGEMENT", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(D{r}="","",TEXT(ROW()-2,"TSK00000"))'

    add_dv(ws, "H3:H10002", "High,Medium,Low")
    add_dv(ws, "O3:O10002", "Not Started,In Progress,Completed,Blocked,Cancelled")

    ws.conditional_formatting.add("O3:O10002",
        CellIsRule(operator='equal', formula=['"Completed"'], fill=green_fill))
    ws.conditional_formatting.add("O3:O10002",
        CellIsRule(operator='equal', formula=['"Blocked"'], fill=red_fill))
    ws.conditional_formatting.add("O3:O10002",
        CellIsRule(operator='equal', formula=['"In Progress"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))
    ws.conditional_formatting.add("O3:O10002",
        CellIsRule(operator='equal', formula=['"Cancelled"'], fill=grey_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _gantt(ws):
    """Visual Gantt chart using conditional formatting bars across date columns."""
    headers = ["Task ID", "Task Name", "Start Date", "End Date",
               "Duration", "Status", "Assigned To", "Milestone"]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📊 GANTT CHART", mc)

    # Pre-fill formulas
    for r in range(3, 1003):
        ws.cell(row=r, column=5).value = f'=IF(OR(C{r}="",D{r}=""),"",D{r}-C{r})'
        ws.cell(row=r, column=8).value = f'=IF(D{r}="","",IF(D{r}=C{r},"⭐ MILESTONE",""))'

    # Generate date columns (columns I onward) for visual bars
    # We put week start dates as column headers
    from datetime import timedelta
    base_date = datetime.date.today()
    for i in range(52):  # 52 weeks
        col = 9 + i
        d = base_date + timedelta(weeks=i)
        ws.cell(row=2, column=col, value=d.strftime("%d-%b"))
        ws.cell(row=2, column=col).font = Font(bold=True, size=8, color=WHITE)
        ws.cell(row=2, column=col).fill = header_fill
        ws.cell(row=2, column=col).alignment = center_align
        ws.column_dimensions[get_column_letter(col)].width = 5

    # Conditional formatting rules for Gantt bars (In Progress = blue, Completed = green)
    for r in range(3, 53):  # Show first 50 tasks
        for i in range(52):
            col = 9 + i
            cl = get_column_letter(col)
            # In Progress bar (blue)
            ws.conditional_formatting.add(f"{cl}{r}:{cl}{r}",
                FormulaRule(formula=[f'AND($F{r}="In Progress",$C{r}<>{cl}$2+($C{r}-$C{r}),$C{r}<={cl}$2+7,$D{r}>={cl}$2)'],
                           fill=PatternFill(start_color="4472C4", end_color="4472C4", fill_type='solid')))
            # Completed bar (green)
            ws.conditional_formatting.add(f"{cl}{r}:{cl}{r}",
                FormulaRule(formula=[f'AND($F{r}="Completed",$C{r}<>{cl}$2+($C{r}-$C{r}),$C{r}<={cl}$2+7,$D{r}>={cl}$2)'],
                           fill=PatternFill(start_color="70AD47", end_color="70AD47", fill_type='solid')))

    add_dv(ws, "F3:F1002", "Not Started,In Progress,Completed,Blocked,Cancelled")
    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 2. CLIENT MODULE
# ═════════════════════════════════════════════════════════════════════════

def _client_master(ws):
    headers = [
        "Client ID", "Client Type", "First Name", "Last Name", "Company Name",
        "Date of Birth", "Gender", "Primary Phone", "Secondary Phone",
        "Email", "WhatsApp Number", "Preferred Communication",
        "Current Address", "Property Address", "National ID / Passport",
        "Client Since", "Referred By", "Client Status",
        "Credit Rating", "Payment History", "Notes / Special Instructions"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "👤 CLIENT MASTER DATA", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"CLT0000"))'

    add_dv(ws, "B3:B2002", "Individual,Family,Corporate,Real Estate Company")
    add_dv(ws, "G3:G2002", "Male,Female,Other")
    add_dv(ws, "L3:L2002", "Phone,Email,WhatsApp,In Person")
    add_dv(ws, "R3:R2002", "Active,Inactive,Prospect,Blacklisted")

    ws.conditional_formatting.add("R3:R2002",
        CellIsRule(operator='equal', formula=['"Active"'], fill=green_fill))
    ws.conditional_formatting.add("R3:R2002",
        CellIsRule(operator='equal', formula=['"Blacklisted"'], fill=red_fill))
    ws.conditional_formatting.add("R3:R2002",
        CellIsRule(operator='equal', formula=['"Prospect"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _client_comm(ws):
    headers = [
        "Log ID", "Client ID", "Client Name", "Date / Time",
        "Communication Type", "Subject / Topic", "Summary / Notes",
        "Action Items", "Follow-up Date", "Handled By", "Status"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📞 CLIENT COMMUNICATION LOG", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"LOG00000"))'

    add_dv(ws, "E3:E10002", "Call,Email,WhatsApp,Meeting,Site Visit")
    add_dv(ws, "K3:K10002", "Open,Closed,Pending")

    ws.conditional_formatting.add("K3:K10002",
        CellIsRule(operator='equal', formula=['"Closed"'], fill=green_fill))
    ws.conditional_formatting.add("K3:K10002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=yellow_cond))
    ws.conditional_formatting.add("K3:K10002",
        CellIsRule(operator='equal', formula=['"Open"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _client_sat(ws):
    headers = [
        "Survey ID", "Project ID", "Client Name", "Survey Date",
        "Quality of Work (1-10)", "Timeliness (1-10)", "Communication (1-10)",
        "Budget Adherence (1-10)", "Cleanliness & Safety (1-10)",
        "Overall Satisfaction (1-10)", "Overall Score (Auto)",
        "Written Feedback", "Would Recommend", "Issues Reported",
        "Action Taken", "Follow-up Date"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "⭐ CLIENT SATISFACTION SURVEY", mc)

    for r in range(3, 503):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"SAT000"))'
        # Overall Score average
        ws.cell(row=r, column=11).value = (
            f'=IF(E{r}="","",ROUND(AVERAGE(E{r}:J{r}),1))'
        )

    add_dv(ws, "M3:M502", "Yes,No")

    # Score validations
    for col in "EFGHIJ":
        add_dv(ws, f"{col}3:{col}502", "1,2,3,4,5,6,7,8,9,10")

    # Color the overall score
    ws.conditional_formatting.add("K3:K502",
        CellIsRule(operator='greaterThanOrEqual', formula=['8'], fill=green_fill))
    ws.conditional_formatting.add("K3:K502",
        CellIsRule(operator='between', formula=['5', '7.9'], fill=yellow_cond))
    ws.conditional_formatting.add("K3:K502",
        CellIsRule(operator='lessThan', formula=['5'], fill=red_fill))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 3. BUDGET & COST
# ═════════════════════════════════════════════════════════════════════════

def _budget_planning(ws):
    headers = [
        "Project ID", "Project Name", "Category", "Sub-category",
        "Item Description", "Unit", "Qty Estimated", "Unit Cost Est.",
        "Total Estimated (Auto)", "Qty Actual", "Unit Cost Actual",
        "Total Actual (Auto)", "Variance (Auto)", "Variance % (Auto)",
        "Budget Status", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "💰 BUDGET PLANNING", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=9).value = f'=IF(OR(G{r}="",H{r}=""),"",G{r}*H{r})'
        ws.cell(row=r, column=12).value = f'=IF(OR(J{r}="",K{r}=""),"",J{r}*K{r})'
        ws.cell(row=r, column=13).value = f'=IF(OR(I{r}="",L{r}=""),"",I{r}-L{r})'
        ws.cell(row=r, column=14).value = f'=IF(OR(I{r}=0,I{r}=""),"",ROUND(M{r}/I{r}*100,1))'
        ws.cell(row=r, column=15).value = (
            f'=IF(M{r}="","",IF(M{r}<0,"Over Budget",IF(M{r}>0,"Under Budget","Within Budget")))'
        )

    add_dv(ws, "C3:C5002", "Labor,Materials,Equipment,Permits,Design,Contingency,Miscellaneous")

    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Within Budget"'], fill=green_fill))
    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Over Budget"'], fill=red_fill))
    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Under Budget"'], fill=yellow_cond))

    # Totals row
    r = 5005
    ws.cell(row=r, column=4, value="TOTALS:").font = Font(bold=True, size=12, color=DARK_BLUE)
    ws.cell(row=r, column=9).value = '=SUM(I3:I5002)'
    ws.cell(row=r, column=9).font = Font(bold=True, size=14, color=DARK_BLUE)
    ws.cell(row=r, column=12).value = '=SUM(L3:L5002)'
    ws.cell(row=r, column=12).font = Font(bold=True, size=14, color=RED)
    ws.cell(row=r, column=13).value = '=I5005-L5005'
    ws.cell(row=r, column=13).font = Font(bold=True, size=14, color=MID_BLUE)

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _cost_tracking(ws):
    headers = [
        "Expense ID", "Project ID", "Phase", "Date",
        "Category", "Sub-category", "Item / Service Description",
        "Vendor / Worker Name", "Quantity", "Unit Price",
        "Total Amount (Auto)", "Tax Amount", "Grand Total (Auto)",
        "Payment Mode", "Payment Status", "Invoice Number",
        "Receipt Attached", "Approved By", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "💸 COST TRACKING", mc)

    for r in range(3, 50003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"EXP000000"))'
        ws.cell(row=r, column=11).value = f'=IF(OR(I{r}="",J{r}=""),"",I{r}*J{r})'
        ws.cell(row=r, column=13).value = f'=IF(K{r}="","",K{r}+L{r})'

    add_dv(ws, "E3:E50002", "Labor,Material,Equipment,Transport,Permit,Design,Marketing,Office,Utilities,Misc")
    add_dv(ws, "N3:N50002", "Cash,Cheque,Bank Transfer,Online,UPI")
    add_dv(ws, "O3:O50002", "Paid,Pending,Partial,Disputed,Overdue")
    add_dv(ws, "Q3:Q50002", "Yes,No")

    ws.conditional_formatting.add("O3:O50002",
        CellIsRule(operator='equal', formula=['"Paid"'], fill=green_fill))
    ws.conditional_formatting.add("O3:O50002",
        CellIsRule(operator='equal', formula=['"Overdue"'], fill=red_fill))
    ws.conditional_formatting.add("O3:O50002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _budget_dashboard(ws):
    add_title_row(ws, "📊 BUDGET DASHBOARD", 12)

    kpis = [
        ("B3", "Total Budget", "B4", "=SUM('Budget Planning'!I3:I5002)"),
        ("D3", "Total Spent", "D4", "=SUM('Budget Planning'!L3:L5002)"),
        ("F3", "Remaining Budget", "F4", "=B4-D4"),
        ("H3", "Utilization %", "H4", '=IFERROR(ROUND(D4/B4*100,1)&"%","N/A")'),
        ("J3", "Over Budget Items", "J4", '=COUNTIF(\'Budget Planning\'!O3:O5002,"Over Budget")'),
    ]
    for lc, label, vc, formula in kpis:
        ws[lc] = label; ws[lc].font = kpi_label
        ws[vc] = formula; ws[vc].font = kpi_font

    # Category-wise breakdown
    r = 7
    ws.cell(row=r, column=2, value="CATEGORY-WISE BREAKDOWN").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1
    for ci, h in enumerate(["Category", "Budgeted", "Spent", "Variance", "Variance %", "Status"]):
        ws.cell(row=r, column=2+ci, value=h)
    style_header_row(ws, r, 6)

    cats = ["Labor","Materials","Equipment","Permits","Design","Contingency","Miscellaneous"]
    for i, cat in enumerate(cats):
        row = r + 1 + i
        ws.cell(row=row, column=2, value=cat)
        ws.cell(row=row, column=3).value = f'=SUMIF(\'Budget Planning\'!C:C,B{row},\'Budget Planning\'!I:I)'
        ws.cell(row=row, column=4).value = f'=SUMIF(\'Budget Planning\'!C:C,B{row},\'Budget Planning\'!L:L)'
        ws.cell(row=row, column=5).value = f'=C{row}-D{row}'
        ws.cell(row=row, column=6).value = f'=IFERROR(ROUND(E{row}/C{row}*100,1),0)'
        ws.cell(row=row, column=7).value = f'=IF(E{row}<0,"⚠ Over","✓ Within")'

    # Contingency status
    r2 = r + 10
    ws.cell(row=r2, column=2, value="CONTINGENCY FUND STATUS").font = section_font
    ws.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=6)
    r2 += 1
    ws.cell(row=r2, column=2, value="Contingency Budget:")
    ws.cell(row=r2, column=3).value = '=SUMIF(\'Budget Planning\'!C:C,"Contingency",\'Budget Planning\'!I:I)'
    ws.cell(row=r2+1, column=2, value="Contingency Used:")
    ws.cell(row=r2+1, column=3).value = '=SUMIF(\'Budget Planning\'!C:C,"Contingency",\'Budget Planning\'!L:L)'
    ws.cell(row=r2+2, column=2, value="Contingency Remaining:")
    ws.cell(row=r2+2, column=3).value = f'=C{r2}-C{r2+1}'

    auto_width(ws, 12)


def _change_orders(ws):
    headers = [
        "Change Order ID", "Project ID", "Date Requested",
        "Requested By", "Description of Change", "Reason for Change",
        "Impact Type", "Cost Impact Amount", "Timeline Impact (Days)",
        "Status", "Approved By", "Approval Date",
        "Revised Budget (Auto)", "Revised Timeline (Auto)",
        "Documentation Link"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🔄 CHANGE ORDERS / VARIATIONS", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"CO-0000"))'
        ws.cell(row=r, column=13).value = (
            f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},\'Project Master Data\'!A:P,16,0)+H{r},""))'
        )

    add_dv(ws, "D3:D2002", "Client,Contractor,Engineer")
    add_dv(ws, "G3:G2002", "Additional Cost,Cost Reduction")
    add_dv(ws, "J3:J2002", "Pending,Approved,Rejected,In Review")

    ws.conditional_formatting.add("J3:J2002",
        CellIsRule(operator='equal', formula=['"Approved"'], fill=green_fill))
    ws.conditional_formatting.add("J3:J2002",
        CellIsRule(operator='equal', formula=['"Rejected"'], fill=red_fill))
    ws.conditional_formatting.add("J3:J2002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 4. CONTRACTOR & VENDOR
# ═════════════════════════════════════════════════════════════════════════

def _contractor_master(ws):
    headers = [
        "Contractor ID", "Company / Individual Name", "Contractor Type",
        "Specialization", "Contact Person", "Phone", "Email",
        "Office Address", "License Number", "Insurance Type",
        "Insurance Expiry", "Tax ID / VAT", "Bank Account", "Bank Name",
        "Rating (1-5)", "Past Projects Count", "Years of Experience",
        "Payment Terms", "Status", "Notes / Reviews"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "👷 CONTRACTOR MASTER DATA", mc)

    for r in range(3, 1003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"CTR0000"))'

    add_dv(ws, "C3:C1002", "General Contractor,Electrician,Plumber,Carpenter,Painter,Mason,Tiler,HVAC,Interior Designer,Architect,Landscaper")
    add_dv(ws, "S3:S1002", "Active,Inactive,Blacklisted,On Hold")
    add_dv(ws, "R3:R1002", "Net 15,Net 30,Advance,COD,Milestone Based")

    ws.conditional_formatting.add("S3:S1002",
        CellIsRule(operator='equal', formula=['"Active"'], fill=green_fill))
    ws.conditional_formatting.add("S3:S1002",
        CellIsRule(operator='equal', formula=['"Blacklisted"'], fill=red_fill))
    ws.conditional_formatting.add("S3:S1002",
        CellIsRule(operator='equal', formula=['"On Hold"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _vendor_master(ws):
    headers = [
        "Vendor ID", "Company Name", "Vendor Type", "Contact Person",
        "Phone", "Email", "Website", "Address",
        "Payment Terms", "Credit Limit", "Tax ID",
        "Bank Details", "Avg Delivery Time (Days)",
        "Quality Rating (1-5)", "Price Rating (1-5)",
        "Preferred Vendor", "Status", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🏪 VENDOR / SUPPLIER MASTER", mc)

    for r in range(3, 1003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"VND0000"))'

    add_dv(ws, "C3:C1002", "Material Supplier,Equipment Rental,Tool Supplier,Furniture,Appliances,Paint,Tile,Electrical Supplies,Plumbing Supplies")
    add_dv(ws, "I3:I1002", "Net 15,Net 30,Advance,COD,Milestone Based")
    add_dv(ws, "P3:P1002", "Yes,No")
    add_dv(ws, "Q3:Q1002", "Active,Inactive")

    ws.conditional_formatting.add("Q3:Q1002",
        CellIsRule(operator='equal', formula=['"Active"'], fill=green_fill))
    ws.conditional_formatting.add("Q3:Q1002",
        CellIsRule(operator='equal', formula=['"Inactive"'], fill=grey_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _contractor_perf(ws):
    headers = [
        "Evaluation ID", "Contractor ID", "Contractor Name", "Project ID",
        "Task / Phase", "Start Date", "End Date (Planned)", "End Date (Actual)",
        "Quality Score (1-10)", "Timeliness Score (1-10)",
        "Communication Score (1-10)", "Safety Score (1-10)",
        "Budget Adherence Score (1-10)", "Overall Rating (Auto)",
        "Issues Reported", "Would Hire Again", "Comments"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📈 CONTRACTOR PERFORMANCE", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"EVL0000"))'
        ws.cell(row=r, column=14).value = f'=IF(I{r}="","",ROUND(AVERAGE(I{r}:M{r}),1))'

    for col in "IJKLM":
        add_dv(ws, f"{col}3:{col}2002", "1,2,3,4,5,6,7,8,9,10")
    add_dv(ws, "Q3:Q2002", "Yes,No")

    ws.conditional_formatting.add("N3:N2002",
        CellIsRule(operator='greaterThanOrEqual', formula=['8'], fill=green_fill))
    ws.conditional_formatting.add("N3:N2002",
        CellIsRule(operator='between', formula=['5', '7.9'], fill=yellow_cond))
    ws.conditional_formatting.add("N3:N2002",
        CellIsRule(operator='lessThan', formula=['5'], fill=red_fill))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _contractor_pay(ws):
    headers = [
        "Payment ID", "Contractor ID", "Contractor Name", "Project ID",
        "Phase", "Work Description", "Work Completed %",
        "Invoice Number", "Invoice Date", "Invoice Amount",
        "Advance Paid", "Current Payment", "Total Paid to Date (Auto)",
        "Balance Remaining (Auto)", "Payment Date", "Payment Mode",
        "Transaction Ref", "Status", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "💵 CONTRACTOR PAYMENTS", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"PAY00000"))'
        ws.cell(row=r, column=13).value = f'=IF(J{r}="","",K{r}+L{r})'
        ws.cell(row=r, column=14).value = f'=IF(J{r}="","",J{r}-M{r})'

    add_dv(ws, "P3:P10002", "Cash,Cheque,Bank Transfer,Online,UPI")
    add_dv(ws, "R3:R10002", "Paid,Pending,Disputed,Overdue")

    ws.conditional_formatting.add("R3:R10002",
        CellIsRule(operator='equal', formula=['"Paid"'], fill=green_fill))
    ws.conditional_formatting.add("R3:R10002",
        CellIsRule(operator='equal', formula=['"Overdue"'], fill=red_fill))
    ws.conditional_formatting.add("R3:R10002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=yellow_cond))
    ws.conditional_formatting.add("R3:R10002",
        CellIsRule(operator='equal', formula=['"Disputed"'], fill=grey_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 5. MATERIALS
# ═════════════════════════════════════════════════════════════════════════

def _material_master(ws):
    headers = [
        "Material ID", "Material Name", "Category",
        "Brand / Manufacturer", "Unit of Measurement",
        "Standard Unit Cost", "Quality Grade",
        "Typical Lead Time (Days)", "Preferred Vendor", "Specifications / Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🧱 MATERIAL MASTER LIST", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"MAT0000"))'

    add_dv(ws, "C3:C2002", "Cement,Steel,Sand,Bricks,Paint,Tiles,Wood,Glass,Electrical,Plumbing,Fixtures,Hardware,Insulation,Adhesives")
    add_dv(ws, "E3:E2002", "Bags,Kg,Tons,Liters,Sq.Ft,Sq.M,Pieces,Meters,Rolls,Boxes,Sets")
    add_dv(ws, "G3:G2002", "A,B,C,Premium,Standard,Economy")

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _material_est(ws):
    headers = [
        "Estimation ID", "Project ID", "Room / Area",
        "Material ID", "Material Name", "Quantity Required",
        "Unit", "Wastage Factor (%)", "Adjusted Qty (Auto)",
        "Unit Rate", "Estimated Cost (Auto)", "Approved",
        "Procurement Status", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📐 MATERIAL ESTIMATION", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"EST00000"))'
        ws.cell(row=r, column=9).value = f'=IF(OR(F{r}="",H{r}=""),"",ROUND(F{r}*(1+H{r}/100),1))'
        ws.cell(row=r, column=11).value = f'=IF(OR(I{r}="",J{r}=""),"",I{r}*J{r})'

    add_dv(ws, "L3:L10002", "Yes,No")
    add_dv(ws, "M3:M10002", "Not Started,Ordered,In Transit,Received,Cancelled")

    ws.conditional_formatting.add("M3:M10002",
        CellIsRule(operator='equal', formula=['"Received"'], fill=green_fill))
    ws.conditional_formatting.add("M3:M10002",
        CellIsRule(operator='equal', formula=['"Ordered"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _purchase_orders(ws):
    headers = [
        "PO Number", "Project ID", "Date", "Vendor ID", "Vendor Name",
        "Delivery Address", "Expected Delivery Date", "Material ID",
        "Material Name", "Quantity", "Unit", "Unit Price",
        "Subtotal (Auto)", "Tax (Auto)", "Delivery Charges",
        "Grand Total (Auto)", "Payment Terms", "Special Instructions",
        "PO Status", "Approved By"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📦 PURCHASE ORDERS", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"PO-00000"))'
        ws.cell(row=r, column=13).value = f'=IF(OR(J{r}="",L{r}=""),"",J{r}*L{r})'
        ws.cell(row=r, column=14).value = f'=IF(M{r}="","",ROUND(M{r}*0.18,2))'
        ws.cell(row=r, column=16).value = f'=IF(M{r}="","",M{r}+N{r}+IF(O{r}="",0,O{r}))'

    add_dv(ws, "Q3:Q5002", "Net 15,Net 30,Advance,COD,Milestone Based")
    add_dv(ws, "S3:S5002", "Draft,Sent,Confirmed,Partially Delivered,Fully Delivered,Cancelled")

    ws.conditional_formatting.add("S3:S5002",
        CellIsRule(operator='equal', formula=['"Fully Delivered"'], fill=green_fill))
    ws.conditional_formatting.add("S3:S5002",
        CellIsRule(operator='equal', formula=['"Cancelled"'], fill=red_fill))
    ws.conditional_formatting.add("S3:S5002",
        CellIsRule(operator='equal', formula=['"Sent"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _material_inventory(ws):
    headers = [
        "Material ID", "Material Name", "Project ID",
        "Opening Stock", "Received Qty", "Received Date",
        "PO Number", "Quantity Used", "Quantity Returned/Waste",
        "Closing Stock (Auto)", "Reorder Level",
        "Stock Alert (Auto)", "Storage Location", "Condition"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📦 MATERIAL INVENTORY / STOCK", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=10).value = f'=IF(D{r}="","",D{r}+E{r}-H{r}-I{r})'
        ws.cell(row=r, column=12).value = (
            f'=IF(J{r}="","",IF(J{r}<=K{r},"🔴 REORDER NOW",IF(J{r}<=K{r}*2,"🟡 LOW","✓ OK")))'
        )

    add_dv(ws, "N3:N5002", "Good,Damaged,Expired")

    ws.conditional_formatting.add("L3:L5002",
        CellIsRule(operator='containsText', formula=['"REORDER"'],
                   fill=PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')))
    ws.conditional_formatting.add("L3:L5002",
        CellIsRule(operator='containsText', formula=['"OK"'],
                   fill=green_fill))
    ws.conditional_formatting.add("N3:N5002",
        CellIsRule(operator='equal', formula=['"Damaged"'], fill=red_fill))
    ws.conditional_formatting.add("N3:N5002",
        CellIsRule(operator='equal', formula=['"Expired"'], fill=red_fill))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _delivery_log(ws):
    headers = [
        "Delivery ID", "PO Number", "Vendor Name",
        "Delivery Date / Time", "Material ID", "Material Name",
        "Qty Ordered", "Qty Delivered", "Qty Accepted",
        "Qty Rejected", "Rejection Reason",
        "Received By", "Driver Name", "Vehicle Number",
        "Delivery Note No.", "Remarks"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🚚 MATERIAL DELIVERY LOG", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"DLV00000"))'

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 6. LABOR
# ═════════════════════════════════════════════════════════════════════════

def _worker_master(ws):
    headers = [
        "Worker ID", "Full Name", "Trade / Skill",
        "Skill Level", "Contact Number", "Address",
        "Date of Birth", "Age (Auto)", "National ID",
        "Emergency Contact", "Bank Account", "Bank Name",
        "Daily Rate", "Hourly Rate", "Overtime Rate (Auto)",
        "Contractor ID", "Status", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "👷 WORKER MASTER DATA", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"WRK0000"))'
        ws.cell(row=r, column=8).value = f'=IF(G{r}="","",DATEDIF(G{r},TODAY(),"Y"))'
        ws.cell(row=r, column=15).value = f'=IF(M{r}="","",ROUND(M{r}*1.5,0))'

    add_dv(ws, "C3:C2002", "Mason,Carpenter,Painter,Electrician,Plumber,Helper,Supervisor,Welder,Tiler,HVAC Tech,Landscaper")
    add_dv(ws, "D3:D2002", "Unskilled,Semi-skilled,Skilled,Expert")
    add_dv(ws, "Q3:Q2002", "Active,Inactive,On Project")

    ws.conditional_formatting.add("Q3:Q2002",
        CellIsRule(operator='equal', formula=['"Active"'], fill=green_fill))
    ws.conditional_formatting.add("Q3:Q2002",
        CellIsRule(operator='equal', formula=['"Inactive"'], fill=grey_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _worker_attendance(ws):
    headers = [
        "Date", "Project ID", "Site Name", "Worker ID", "Worker Name",
        "Trade", "Check-in Time", "Check-out Time",
        "Hours Worked (Auto)", "Overtime Hours", "Status",
        "Weather Delay", "Task Assigned", "Daily Output",
        "Supervisor Confirmation", "Remarks"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📅 WORKER ATTENDANCE", mc)

    for r in range(3, 50003):
        ws.cell(row=r, column=9).value = f'=IF(OR(G{r}="",H{r}=""),"",ROUND((H{r}-G{r})*24,1))'

    add_dv(ws, "K3:K50002", "Present,Absent,Half-day,Leave")
    add_dv(ws, "L3:L50002", "Yes,No")

    ws.conditional_formatting.add("K3:K50002",
        CellIsRule(operator='equal', formula=['"Present"'], fill=green_fill))
    ws.conditional_formatting.add("K3:K50002",
        CellIsRule(operator='equal', formula=['"Absent"'], fill=red_fill))
    ws.conditional_formatting.add("K3:K50002",
        CellIsRule(operator='equal', formula=['"Half-day"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _labor_cost(ws):
    headers = [
        "Worker ID", "Worker Name", "Period (Week/Month)",
        "Working Days", "Working Hours", "Daily Rate",
        "Basic Wage (Auto)", "Overtime Hours", "OT Amount (Auto)",
        "Bonus / Incentive", "Deductions", "Gross Pay (Auto)",
        "Advance Previously", "Advance Deduction", "Net Pay (Auto)",
        "Payment Date", "Payment Mode", "Acknowledgment"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "💵 LABOR COST CALCULATION", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=7).value = f'=IF(OR(D{r}="",F{r}=""),"",D{r}*F{r})'
        ws.cell(row=r, column=9).value = f'=IF(OR(H{r}="",F{r}=""),"",H{r}*F{r}/8*1.5)'
        ws.cell(row=r, column=12).value = f'=IF(G{r}="","",G{r}+I{r}+J{r}-K{r})'
        ws.cell(row=r, column=15).value = f'=IF(L{r}="","",L{r}-M{r}-N{r})'

    add_dv(ws, "Q3:Q10002", "Cash,Cheque,Bank Transfer,Online,UPI")
    add_dv(ws, "R3:R10002", "Received,Pending")

    ws.protection.sheet = True
    ws.protection.password = "labor2026"

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _productivity(ws):
    headers = [
        "Date", "Worker ID", "Worker Name", "Trade",
        "Task Assigned", "Target (Qty/Area)", "Achieved (Qty/Area)",
        "Efficiency % (Auto)", "Quality Check",
        "Supervisor Notes", "Bonus Eligible"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📊 PRODUCTIVITY TRACKER", mc)

    for r in range(3, 20003):
        ws.cell(row=r, column=8).value = f'=IF(OR(F{r}=0,F{r}=""),"",ROUND(G{r}/F{r}*100,1))'

    add_dv(ws, "I3:I20002", "Pass,Fail,Rework Needed")
    add_dv(ws, "K3:K20002", "Yes,No")

    ws.conditional_formatting.add("H3:H20002",
        CellIsRule(operator='greaterThanOrEqual', formula=['100'], fill=green_fill))
    ws.conditional_formatting.add("H3:H20002",
        CellIsRule(operator='between', formula=['75', '99.9'], fill=yellow_cond))
    ws.conditional_formatting.add("H3:H20002",
        CellIsRule(operator='lessThan', formula=['75'], fill=red_fill))

    ws.conditional_formatting.add("I3:I20002",
        CellIsRule(operator='equal', formula=['"Pass"'], fill=green_fill))
    ws.conditional_formatting.add("I3:I20002",
        CellIsRule(operator='equal', formula=['"Fail"'], fill=red_fill))
    ws.conditional_formatting.add("I3:I20002",
        CellIsRule(operator='equal', formula=['"Rework Needed"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 7. ROOM-WISE TRACKING
# ═════════════════════════════════════════════════════════════════════════

def _room_master(ws):
    headers = [
        "Area ID", "Project ID", "Area Name",
        "Area Size (Sq.Ft)", "Renovation Type",
        "Priority Order", "Assigned Contractor", "Contractor ID",
        "Status"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🏠 ROOM / AREA MASTER", mc)

    for r in range(3, 1003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"ARE0000"))'

    add_dv(ws, "C3:C1002", "Living Room,Bedroom 1,Bedroom 2,Bedroom 3,Kitchen,Bathroom 1,Bathroom 2,Dining Room,Garage,Garden,Terrace,Basement,Office Room,Balcony,Hallway,Utility Room")
    add_dv(ws, "E3:E1002", "Full,Partial,Cosmetic")
    add_dv(ws, "I3:I1002", "Not Started,In Progress,Completed,Snagging")

    ws.conditional_formatting.add("I3:I1002",
        CellIsRule(operator='equal', formula=['"Completed"'], fill=green_fill))
    ws.conditional_formatting.add("I3:I1002",
        CellIsRule(operator='equal', formula=['"In Progress"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))
    ws.conditional_formatting.add("I3:I1002",
        CellIsRule(operator='equal', formula=['"Snagging"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _room_checklist(ws):
    headers = [
        "Checklist ID", "Area Name", "Work Category",
        "Specific Task", "Assigned To", "Planned Date",
        "Completed Date", "Status", "Quality Check",
        "Photo Evidence (Link)", "Inspector Sign-off", "Remarks"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "✅ ROOM-WISE WORK CHECKLIST", mc)

    for r in range(3, 20003):
        ws.cell(row=r, column=1).value = f'=IF(D{r}="","",TEXT(ROW()-2,"CHK000000"))'

    add_dv(ws, "H3:H20002", "Pending,In Progress,Done,Skipped")
    add_dv(ws, "I3:I20002", "Pass,Fail,Rework Needed")

    ws.conditional_formatting.add("H3:H20002",
        CellIsRule(operator='equal', formula=['"Done"'], fill=green_fill))
    ws.conditional_formatting.add("H3:H20002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=yellow_cond))
    ws.conditional_formatting.add("H3:H20002",
        CellIsRule(operator='equal', formula=['"Skipped"'], fill=grey_cond))

    ws.conditional_formatting.add("I3:I20002",
        CellIsRule(operator='equal', formula=['"Pass"'], fill=green_fill))
    ws.conditional_formatting.add("I3:I20002",
        CellIsRule(operator='equal', formula=['"Fail"'], fill=red_fill))
    ws.conditional_formatting.add("I3:I20002",
        CellIsRule(operator='equal', formula=['"Rework Needed"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _room_cost(ws):
    headers = [
        "Area Name", "Labor Cost", "Material Cost",
        "Equipment Cost", "Other Costs", "Total Area Cost (Auto)",
        "Cost per Sq.Ft (Auto)", "Budget Allocated",
        "Variance (Auto)"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "💰 ROOM-WISE COST SUMMARY", mc)

    for r in range(3, 103):
        ws.cell(row=r, column=6).value = f'=IF(B{r}="","",SUM(B{r}:E{r}))'
        ws.cell(row=r, column=7).value = (
            f'=IF(OR(F{r}="",VLOOKUP(B{r},\'Room Area Master\'!C:D,2,0)=0),"",ROUND(F{r}/VLOOKUP(B{r},\'Room Area Master\'!C:D,2,0),0))'
        )
        ws.cell(row=r, column=9).value = f'=IF(H{r}="","",H{r}-F{r})'

    ws.conditional_formatting.add("I3:I102",
        CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))
    ws.conditional_formatting.add("I3:I102",
        CellIsRule(operator='greaterThanOrEqual', formula=['0'], fill=green_fill))

    r = 105
    ws.cell(row=r, column=1, value="GRAND TOTAL").font = Font(bold=True, size=12)
    ws.cell(row=r, column=6).value = '=SUM(F3:F102)'
    ws.cell(row=r, column=6).font = Font(bold=True, size=14, color=DARK_BLUE)

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 8. DESIGN & PLANNING
# ═════════════════════════════════════════════════════════════════════════

def _design_req(ws):
    headers = [
        "Design ID", "Project ID", "Area Name",
        "Design Style Preference", "Color Scheme",
        "Material Preferences", "Fixture Preferences",
        "Inspiration References (Links)", "Client Approved",
        "Design Status", "Designer Name",
        "Revision History", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🎨 DESIGN REQUIREMENTS", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"DSN0000"))'

    add_dv(ws, "D3:D2002", "Modern,Contemporary,Traditional,Minimalist,Industrial,Rustic,Scandinavian")
    add_dv(ws, "I3:I2002", "Yes,No")
    add_dv(ws, "J3:J2002", "Concept,Draft,Revised,Final Approved")

    ws.conditional_formatting.add("J3:J2002",
        CellIsRule(operator='equal', formula=['"Final Approved"'], fill=green_fill))
    ws.conditional_formatting.add("J3:J2002",
        CellIsRule(operator='equal', formula=['"Concept"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _measurements(ws):
    headers = [
        "Area ID", "Area Name", "Length (Ft)", "Width (Ft)",
        "Height (Ft)", "Floor Area (Auto Sq.Ft)",
        "Wall Area (Auto Sq.Ft)", "Ceiling Area (Auto)",
        "Window Count", "Window Dimensions", "Door Count",
        "Door Dimensions", "Special Features",
        "Flooring Needed (Auto Sq.Ft)", "Paint Required (Auto Ltr)",
        "Notes / Sketch Reference"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📏 MEASUREMENTS & SPECIFICATIONS", mc)

    for r in range(3, 1003):
        ws.cell(row=r, column=6).value = f'=IF(OR(C{r}="",D{r}=""),"",C{r}*D{r})'
        # Wall area = perimeter * height - windows - doors (simplified)
        ws.cell(row=r, column=7).value = (
            f'=IF(OR(C{r}="",D{r}="",E{r}=""),"",'
            f'ROUND((2*(C{r}+D{r})*E{r})-(I{r}*3*2)-(K{r}*7*2),1))'
        )
        ws.cell(row=r, column=8).value = f'=IF(F{r}="","",F{r})'
        ws.cell(row=r, column=14).value = f'=IF(F{r}="","",ROUND(F{r}*1.1,1))'  # +10% wastage
        ws.cell(row=r, column=15).value = f'=IF(G{r}="","",ROUND(G{r}/120,1))'  # ~120 sqft/liter coverage

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _material_selection(ws):
    headers = [
        "Area Name", "Item Category",
        "Option 1 Name", "Option 1 Specs", "Option 1 Price", "Option 1 Vendor",
        "Option 2 Name", "Option 2 Specs", "Option 2 Price", "Option 2 Vendor",
        "Option 3 Name", "Option 3 Specs", "Option 3 Price", "Option 3 Vendor",
        "Client Selected", "Approval Status",
        "Sample Provided", "Order Status", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🎯 MATERIAL SELECTION BOARD", mc)

    add_dv(ws, "B3:B2002", "Flooring,Wall,Ceiling,Fixtures,Fittings,Furniture,Lighting")
    add_dv(ws, "O3:O2002", "Option 1,Option 2,Option 3")
    add_dv(ws, "P3:P2002", "Pending,Approved,Rejected")
    add_dv(ws, "Q3:Q2002", "Yes,No")
    add_dv(ws, "R3:R2002", "Not Started,Ordered,In Transit,Received")

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 9. EQUIPMENT
# ═════════════════════════════════════════════════════════════════════════

def _equip_inventory(ws):
    headers = [
        "Equipment ID", "Equipment Name", "Category",
        "Type", "Brand / Model", "Serial Number",
        "Purchase Date", "Purchase Cost", "Rental Company",
        "Daily Rental Rate", "Weekly Rental Rate",
        "Condition", "Last Service Date", "Next Service Date",
        "Current Location", "Assigned Project", "Status"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🔧 EQUIPMENT INVENTORY", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"EQP0000"))'

    add_dv(ws, "C3:C2002", "Owned,Rented")
    add_dv(ws, "D3:D2002", "Power Tool,Heavy Machinery,Hand Tool,Safety Equipment,Scaffolding,Generator,Mixer,Compressor")
    add_dv(ws, "L3:L2002", "New,Good,Fair,Needs Repair,Retired")
    add_dv(ws, "Q3:Q2002", "Available,In Use,Under Repair,Lost")

    ws.conditional_formatting.add("Q3:Q2002",
        CellIsRule(operator='equal', formula=['"Available"'], fill=green_fill))
    ws.conditional_formatting.add("Q3:Q2002",
        CellIsRule(operator='equal', formula=['"In Use"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))
    ws.conditional_formatting.add("Q3:Q2002",
        CellIsRule(operator='equal', formula=['"Under Repair"'], fill=yellow_cond))
    ws.conditional_formatting.add("Q3:Q2002",
        CellIsRule(operator='equal', formula=['"Lost"'], fill=red_fill))

    ws.conditional_formatting.add("L3:L2002",
        CellIsRule(operator='equal', formula=['"Needs Repair"'], fill=yellow_cond))
    ws.conditional_formatting.add("L3:L2002",
        CellIsRule(operator='equal', formula=['"Retired"'], fill=grey_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _equip_usage(ws):
    headers = [
        "Log ID", "Equipment ID", "Equipment Name",
        "Project ID", "Assigned To", "Check-out Date",
        "Return Date", "Days Used (Auto)", "Rental Cost (Auto)",
        "Condition on Return", "Damage Reported",
        "Damage Description", "Repair Cost", "Remarks"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📋 EQUIPMENT USAGE LOG", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"USG00000"))'
        ws.cell(row=r, column=8).value = f'=IF(OR(F{r}="",G{r}=""),"",G{r}-F{r})'
        ws.cell(row=r, column=9).value = f'=IF(H{r}="","",H{r}*100)'  # default 100/day

    add_dv(ws, "J3:J10002", "New,Good,Fair,Needs Repair,Retired")
    add_dv(ws, "K3:K10002", "Yes,No")

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _equip_maint(ws):
    headers = [
        "Maintenance ID", "Equipment ID", "Equipment Name",
        "Maintenance Type", "Scheduled Date", "Actual Date",
        "Service Provider", "Cost", "Parts Replaced",
        "Next Maintenance Due", "Status", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🔩 EQUIPMENT MAINTENANCE", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"MNT0000"))'

    add_dv(ws, "D3:D5002", "Routine,Repair,Replacement")
    add_dv(ws, "K3:K5002", "Scheduled,Completed,Overdue")

    ws.conditional_formatting.add("K3:K5002",
        CellIsRule(operator='equal', formula=['"Completed"'], fill=green_fill))
    ws.conditional_formatting.add("K3:K5002",
        CellIsRule(operator='equal', formula=['"Overdue"'], fill=red_fill))
    ws.conditional_formatting.add("K3:K5002",
        CellIsRule(operator='equal', formula=['"Scheduled"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 10. QUALITY CONTROL
# ═════════════════════════════════════════════════════════════════════════

def _quality_standards(ws):
    headers = [
        "Standard ID", "Work Category", "Quality Criteria",
        "Acceptable Standards", "Testing Method",
        "Pass/Fail Criteria", "Inspector Role"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📋 QUALITY STANDARDS SETUP", mc)

    for r in range(3, 503):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"QST000"))'

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _inspection_log(ws):
    headers = [
        "Inspection ID", "Project ID", "Phase", "Area",
        "Inspection Type", "Inspection Date",
        "Inspector Name", "Inspector Role",
        "Items Inspected", "Result per Item",
        "Overall Result", "Defects Found",
        "Photo Links", "Rework Required",
        "Re-inspection Date", "Sign-off Status"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🔍 INSPECTION LOG", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"INS0000"))'

    add_dv(ws, "E3:E5002", "In-progress,Stage Completion,Final,Client,Municipal,Safety")
    add_dv(ws, "K3:K5002", "Pass,Fail,Conditional Pass")
    add_dv(ws, "N3:N5002", "Yes,No")
    add_dv(ws, "P3:P5002", "Signed,Pending,Rejected")

    ws.conditional_formatting.add("K3:K5002",
        CellIsRule(operator='equal', formula=['"Pass"'], fill=green_fill))
    ws.conditional_formatting.add("K3:K5002",
        CellIsRule(operator='equal', formula=['"Fail"'], fill=red_fill))
    ws.conditional_formatting.add("K3:K5002",
        CellIsRule(operator='equal', formula=['"Conditional Pass"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _snagging(ws):
    headers = [
        "Snag ID", "Project ID", "Area", "Date Identified",
        "Defect Description", "Severity", "Identified By",
        "Assigned To", "Due Date for Fix", "Status",
        "Fix Date", "Verified By", "Photo Before (Link)",
        "Photo After (Link)", "Cost to Rectify",
        "Responsibility"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🔧 DEFECT / SNAGGING LIST", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"SNG0000"))'

    add_dv(ws, "F3:F5002", "Critical,Major,Minor,Cosmetic")
    add_dv(ws, "J3:J5002", "Open,In Progress,Fixed,Closed,Rejected")
    add_dv(ws, "P3:P5002", "Contractor,Supplier,Client Change")

    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='equal', formula=['"Closed"'], fill=green_fill))
    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='equal', formula=['"Fixed"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))
    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='equal', formula=['"Open"'], fill=yellow_cond))
    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='equal', formula=['"Rejected"'], fill=grey_cond))

    ws.conditional_formatting.add("F3:F5002",
        CellIsRule(operator='equal', formula=['"Critical"'], fill=red_fill))
    ws.conditional_formatting.add("F3:F5002",
        CellIsRule(operator='equal', formula=['"Major"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 11. SAFETY & COMPLIANCE
# ═════════════════════════════════════════════════════════════════════════

def _safety_checklist(ws):
    headers = [
        "Checklist ID", "Date", "Project ID", "Site Name",
        "Safety Item", "Status", "Inspector",
        "Non-compliance Action", "Next Check Date"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🦺 SAFETY CHECKLIST", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(E{r}="","",TEXT(ROW()-2,"SAF0000"))'

    add_dv(ws, "E3:E5002", "Helmets Worn,Safety Harness,Fire Extinguisher,First Aid Kit,Proper Scaffolding,Electrical Safety,Hazardous Material Handling,Site Fencing")
    add_dv(ws, "F3:F5002", "Compliant,Non-compliant,N/A")

    ws.conditional_formatting.add("F3:F5002",
        CellIsRule(operator='equal', formula=['"Compliant"'], fill=green_fill))
    ws.conditional_formatting.add("F3:F5002",
        CellIsRule(operator='equal', formula=['"Non-compliant"'], fill=red_fill))
    ws.conditional_formatting.add("F3:F5002",
        CellIsRule(operator='equal', formula=['"N/A"'], fill=grey_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _incident_log(ws):
    headers = [
        "Incident ID", "Date / Time", "Project ID",
        "Location on Site", "Type", "Description",
        "Persons Involved", "Injuries / Damage",
        "Cause", "Immediate Action Taken",
        "Medical Treatment", "Reported to Authorities",
        "Corrective Action", "Preventive Measures",
        "Investigation Status"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "⚠️ INCIDENT & ACCIDENT LOG", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"INC0000"))'

    add_dv(ws, "E3:E2002", "Near Miss,First Aid,Medical Treatment,Lost Time,Fatal")
    add_dv(ws, "I3:I2002", "Human Error,Equipment Failure,Environmental,Negligence")
    add_dv(ws, "K3:K2002", "Yes,No")
    add_dv(ws, "L3:L2002", "Yes,No")
    add_dv(ws, "O3:O2002", "Open,Closed")

    ws.conditional_formatting.add("E3:E2002",
        CellIsRule(operator='equal', formula=['"Fatal"'], fill=red_fill))
    ws.conditional_formatting.add("E3:E2002",
        CellIsRule(operator='equal', formula=['"Medical Treatment"'], fill=yellow_cond))
    ws.conditional_formatting.add("O3:O2002",
        CellIsRule(operator='equal', formula=['"Closed"'], fill=green_fill))
    ws.conditional_formatting.add("O3:O2002",
        CellIsRule(operator='equal', formula=['"Open"'], fill=red_fill))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _permits(ws):
    headers = [
        "Permit ID", "Project ID", "Permit Type",
        "Issuing Authority", "Application Date",
        "Approval Date", "Expiry Date", "Fee Paid",
        "Status", "Document Link", "Renewal Required",
        "Days Until Expiry (Auto)", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📜 PERMITS & COMPLIANCE", mc)

    for r in range(3, 1003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"PRM0000"))'
        ws.cell(row=r, column=12).value = f'=IF(G{r}="","",G{r}-TODAY())'

    add_dv(ws, "C3:C1002", "Building Permit,Electrical Permit,Plumbing Permit,Demolition Permit,Fire Safety,Environmental Clearance,HOA Approval")
    add_dv(ws, "I3:I1002", "Pending,Approved,Rejected,Expired,Renewed")
    add_dv(ws, "K3:K1002", "Yes,No")

    ws.conditional_formatting.add("I3:I1002",
        CellIsRule(operator='equal', formula=['"Approved"'], fill=green_fill))
    ws.conditional_formatting.add("I3:I1002",
        CellIsRule(operator='equal', formula=['"Expired"'], fill=red_fill))
    ws.conditional_formatting.add("I3:I1002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=yellow_cond))
    ws.conditional_formatting.add("I3:I1002",
        CellIsRule(operator='equal', formula=['"Renewed"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))

    # Expiry alert
    ws.conditional_formatting.add("L3:L1002",
        CellIsRule(operator='lessThan', formula=['30'],
                   fill=PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 12. QUOTATION & INVOICING
# ═════════════════════════════════════════════════════════════════════════

def _quotation(ws):
    headers = [
        "Quotation ID", "Client ID", "Client Name",
        "Project Description", "Date Created", "Valid Until",
        "Line Item 1 Desc", "Line 1 Category", "Line 1 Unit",
        "Line 1 Qty", "Line 1 Rate", "Line 1 Amount (Auto)",
        "Line Item 2 Desc", "Line 2 Category", "Line 2 Unit",
        "Line 2 Qty", "Line 2 Rate", "Line 2 Amount (Auto)",
        "Line Item 3 Desc", "Line 3 Category", "Line 3 Unit",
        "Line 3 Qty", "Line 3 Rate", "Line 3 Amount (Auto)",
        "Subtotal (Auto)", "Discount (%)", "Discount Amount (Auto)",
        "Tax (Auto)", "Grand Total (Auto)",
        "Terms & Conditions", "Payment Schedule",
        "Status", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📄 QUOTATION BUILDER", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"QTN0000"))'
        ws.cell(row=r, column=12).value = f'=IF(OR(J{r}="",K{r}=""),"",J{r}*K{r})'
        ws.cell(row=r, column=18).value = f'=IF(OR(P{r}="",Q{r}=""),"",P{r}*Q{r})'
        ws.cell(row=r, column=24).value = f'=IF(OR(V{r}="",W{r}=""),"",V{r}*W{r})'
        ws.cell(row=r, column=25).value = f'=IF(L{r}="","",L{r}+R{r}+X{r})'
        ws.cell(row=r, column=27).value = f'=IF(OR(E{r}="",Y{r}=""),"",E{r}*Y{r}/100)'
        ws.cell(row=r, column=28).value = f'=IF(E{r}="","",ROUND((E{r}-IF(G{r}="",0,G{r}))*0.18,2))'
        ws.cell(row=r, column=29).value = f'=IF(E{r}="","",E{r}-IF(G{r}="",0,G{r})+AB{r})'

    add_dv(ws, "AF3:AF2002", "Draft,Sent,Accepted,Rejected,Revised")

    ws.conditional_formatting.add("AF3:AF2002",
        CellIsRule(operator='equal', formula=['"Accepted"'], fill=green_fill))
    ws.conditional_formatting.add("AF3:AF2002",
        CellIsRule(operator='equal', formula=['"Rejected"'], fill=red_fill))
    ws.conditional_formatting.add("AF3:AF2002",
        CellIsRule(operator='equal', formula=['"Sent"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _invoice_mgmt(ws):
    headers = [
        "Invoice ID", "Client ID", "Project ID",
        "Invoice Date", "Due Date", "Billing Period / Description",
        "Previous Amount Billed", "This Invoice Amount",
        "Total Billed to Date (Auto)", "Subtotal",
        "Tax (Auto)", "Total", "Amount Received",
        "Balance Due (Auto)", "Status",
        "Overdue Alert (Auto)", "Payment Link / Mode"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🧾 INVOICE MANAGEMENT", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"INV00000"))'
        ws.cell(row=r, column=9).value = f'=IF(G{r}="","",G{r}+H{r})'
        ws.cell(row=r, column=11).value = f'=IF(J{r}="","",ROUND(J{r}*0.18,2))'
        ws.cell(row=r, column=12).value = f'=IF(J{r}="","",J{r}+K{r})'
        ws.cell(row=r, column=14).value = f'=IF(L{r}="","",L{r}-IF(M{r}="",0,M{r}))'
        ws.cell(row=r, column=16).value = (
            f'=IF(OR(E{r}="",O{r}="Paid",O{r}=""),"",'
            f'IF(TODAY()>E{r},"⚠ OVERDUE by "&TEXT(TODAY()-E{r},0)&" days","Due in "&TEXT(E{r}-TODAY(),0)&" days"))'
        )

    add_dv(ws, "O3:O5002", "Draft,Sent,Paid,Partially Paid,Overdue,Disputed")

    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Paid"'], fill=green_fill))
    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Overdue"'], fill=red_fill))
    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Partially Paid"'], fill=yellow_cond))
    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Disputed"'], fill=grey_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _payment_receipts(ws):
    headers = [
        "Receipt ID", "Invoice ID", "Client ID",
        "Payment Date", "Amount Received", "Payment Mode",
        "Transaction Reference", "Received By",
        "Balance After Payment (Auto)", "Thank You Note Sent"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🧾 PAYMENT RECEIPTS", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"RCT00000"))'
        ws.cell(row=r, column=9).value = (
            f'=IF(D{r}="","",IFERROR(VLOOKUP(B{r},\'Invoice Management\'!A:N,14,0),0)-E{r})'
        )

    add_dv(ws, "F3:F10002", "Cash,Cheque,Bank Transfer,Online,UPI")
    add_dv(ws, "J3:J10002", "Yes,No")

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 13. FINANCIAL MANAGEMENT
# ═════════════════════════════════════════════════════════════════════════

def _income_tracker(ws):
    headers = [
        "Entry ID", "Date", "Client ID", "Project ID",
        "Source", "Amount", "Payment Mode",
        "Reference Number", "Received By", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "💵 INCOME TRACKER", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"INC000000"))'

    add_dv(ws, "E3:E10002", "Advance,Progress Payment,Final Payment,Retention Release,Variation Payment")
    add_dv(ws, "G3:G10002", "Cash,Cheque,Bank Transfer,Online,UPI")

    r = 10005
    ws.cell(row=r, column=3, value="TOTAL INCOME:").font = Font(bold=True, size=12, color=DARK_GREEN)
    ws.cell(row=r, column=6).value = '=SUM(F3:F10002)'
    ws.cell(row=r, column=6).font = Font(bold=True, size=14, color=DARK_GREEN)
    ws.cell(row=r, column=6).number_format = '#,##0'

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _expense_tracker(ws):
    headers = [
        "Entry ID", "Date", "Project ID",
        "Category", "Description", "Vendor / Payee",
        "Amount", "Tax", "Total (Auto)",
        "Payment Mode", "Invoice Reference",
        "Approved By", "Status"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "💸 EXPENSE TRACKER", mc)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"EXP000000"))'
        ws.cell(row=r, column=9).value = f'=IF(G{r}="","",G{r}+H{r})'

    add_dv(ws, "D3:D10002", "Labor,Material,Equipment,Transport,Permits,Design,Marketing,Office,Utilities,Misc")
    add_dv(ws, "J3:J10002", "Cash,Cheque,Bank Transfer,Online,UPI")
    add_dv(ws, "M3:M10002", "Paid,Pending")

    r = 10005
    ws.cell(row=r, column=3, value="TOTAL EXPENSES:").font = Font(bold=True, size=12, color=RED)
    ws.cell(row=r, column=9).value = '=SUM(I3:I10002)'
    ws.cell(row=r, column=9).font = Font(bold=True, size=14, color=RED)
    ws.cell(row=r, column=9).number_format = '#,##0'

    ws.conditional_formatting.add("M3:M10002",
        CellIsRule(operator='equal', formula=['"Paid"'], fill=green_fill))
    ws.conditional_formatting.add("M3:M10002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _pnl_project(ws):
    headers = [
        "Project ID", "Project Name",
        "Total Revenue (Invoiced)", "Labor Costs",
        "Material Costs", "Equipment Costs",
        "Subcontractor Costs", "Overhead Allocation",
        "Total Costs (Auto)", "Gross Profit (Auto)",
        "Gross Margin % (Auto)", "Net Profit (Auto)",
        "Net Margin % (Auto)", "Project ROI % (Auto)"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📊 PROFIT & LOSS PER PROJECT", mc)

    for r in range(3, 503):
        ws.cell(row=r, column=9).value = f'=IF(C{r}="","",SUM(D{r}:H{r}))'
        ws.cell(row=r, column=10).value = f'=IF(C{r}="","",C{r}-I{r})'
        ws.cell(row=r, column=11).value = f'=IF(C{r}="","",ROUND(J{r}/C{r}*100,1))'
        ws.cell(row=r, column=12).value = f'=IF(C{r}="","",J{r})'
        ws.cell(row=r, column=13).value = f'=IF(C{r}="","",ROUND(L{r}/C{r}*100,1))'
        ws.cell(row=r, column=14).value = f'=IF(I{r}="","",ROUND(J{r}/I{r}*100,1))'

    ws.conditional_formatting.add("K3:K502",
        CellIsRule(operator='greaterThanOrEqual', formula=['20'], fill=green_fill))
    ws.conditional_formatting.add("K3:K502",
        CellIsRule(operator='between', formula=['10', '19.9'], fill=yellow_cond))
    ws.conditional_formatting.add("K3:K502",
        CellIsRule(operator='lessThan', formula=['10'], fill=red_fill))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _cash_flow(ws):
    headers = [
        "Period (Week/Month)", "Opening Balance",
        "Income Expected", "Income Received",
        "Expenses Expected", "Expenses Paid",
        "Closing Balance (Auto)", "Cash Flow Status",
        "Forecast vs Actual", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "💹 CASH FLOW TRACKER", mc)

    for r in range(3, 203):
        ws.cell(row=r, column=7).value = f'=IF(B{r}="","",B{r}+D{r}-F{r})'
        ws.cell(row=r, column=8).value = f'=IF(G{r}="","",IF(G{r}>B{r},"✅ Positive","⚠ Negative"))'
        ws.cell(row=r, column=9).value = f'=IF(C{r}="","",C{r}-D{r})'

    ws.conditional_formatting.add("H3:H202",
        CellIsRule(operator='containsText', formula=['"Positive"'], fill=green_fill))
    ws.conditional_formatting.add("H3:H202",
        CellIsRule(operator='containsText', formula=['"Negative"'], fill=red_fill))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _tax_compliance(ws):
    headers = [
        "Tax Period", "Revenue (Taxable)",
        "Input Tax (on Purchases)", "Output Tax (on Sales)",
        "Net Tax Payable (Auto)", "Filing Due Date",
        "Filing Status", "Payment Reference"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🏛️ TAX & COMPLIANCE", mc)

    for r in range(3, 203):
        ws.cell(row=r, column=5).value = f'=IF(D{r}="","",D{r}-C{r})'

    add_dv(ws, "G3:G202", "Pending,Filed,Paid,Overdue")

    ws.conditional_formatting.add("G3:G202",
        CellIsRule(operator='equal', formula=['"Paid"'], fill=green_fill))
    ws.conditional_formatting.add("G3:G202",
        CellIsRule(operator='equal', formula=['"Overdue"'], fill=red_fill))
    ws.conditional_formatting.add("G3:G202",
        CellIsRule(operator='equal', formula=['"Filed"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 14. DOCUMENT MANAGEMENT
# ═════════════════════════════════════════════════════════════════════════

def _doc_register(ws):
    headers = [
        "Document ID", "Project ID", "Document Type",
        "Document Title / Description", "Version Number",
        "Created By", "Created Date", "Approved By",
        "Approval Date", "File Location / Link",
        "Expiry Date", "Status", "Access Level", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📁 DOCUMENT REGISTER", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"DOC00000"))'

    add_dv(ws, "C3:C5002", "Contract,Drawing,Permit,Invoice,Photo,Report,Quotation,Warranty,Insurance,Specification,BOQ")
    add_dv(ws, "L3:L5002", "Draft,Active,Superseded,Archived")
    add_dv(ws, "M3:M5002", "Public,Internal,Confidential")

    ws.conditional_formatting.add("L3:L5002",
        CellIsRule(operator='equal', formula=['"Active"'], fill=green_fill))
    ws.conditional_formatting.add("L3:L5002",
        CellIsRule(operator='equal', formula=['"Superseded"'], fill=yellow_cond))
    ws.conditional_formatting.add("L3:L5002",
        CellIsRule(operator='equal', formula=['"Archived"'], fill=grey_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _contract_register(ws):
    headers = [
        "Contract ID", "Project ID", "Contract Type",
        "Parties Involved", "Contract Value",
        "Start Date", "End Date", "Key Terms Summary",
        "Payment Terms", "Penalty Clauses",
        "Variation Allowed", "Retention Amount (%)",
        "Contract Status", "Document Link", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📑 CONTRACT REGISTER", mc)

    for r in range(3, 1003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"CTR0000"))'

    add_dv(ws, "C3:C1002", "Main Contract,Subcontract,Supplier Agreement,Design Contract,Labor Contract")
    add_dv(ws, "I3:I1002", "Net 15,Net 30,Advance,COD,Milestone Based")
    add_dv(ws, "K3:K1002", "Yes,No")
    add_dv(ws, "M3:M1002", "Active,Completed,Terminated,Disputed")

    ws.conditional_formatting.add("M3:M1002",
        CellIsRule(operator='equal', formula=['"Active"'], fill=green_fill))
    ws.conditional_formatting.add("M3:M1002",
        CellIsRule(operator='equal', formula=['"Completed"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))
    ws.conditional_formatting.add("M3:M1002",
        CellIsRule(operator='equal', formula=['"Terminated"'], fill=red_fill))
    ws.conditional_formatting.add("M3:M1002",
        CellIsRule(operator='equal', formula=['"Disputed"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 15. WARRANTY & AFTER-SALES
# ═════════════════════════════════════════════════════════════════════════

def _warranty(ws):
    headers = [
        "Warranty ID", "Project ID", "Client ID",
        "Item / Work Category", "Warranty Type",
        "Warranty Period (Months)", "Start Date",
        "Expiry Date", "Provider (Contractor/Vendor/Mfr)",
        "Provider Contact", "Document Link",
        "Status", "Days Until Expiry (Auto)", "Reminder Alert"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🛡️ WARRANTY REGISTER", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"WAR0000"))'
        ws.cell(row=r, column=13).value = f'=IF(H{r}="","",H{r}-TODAY())'
        ws.cell(row=r, column=14).value = (
            f'=IF(OR(M{r}="",L{r}<>"Active"),"",'
            f'IF(M{r}<=0,"🔴 EXPIRED",IF(M{r}<=30,"🟠 EXPIRING IN 30 DAYS",IF(M{r}<=60,"🟡 EXPIRING IN 60 DAYS","✓ Active"))))'
        )

    add_dv(ws, "D3:D2002", "Roofing,Plumbing,Electrical,Painting,Flooring,Appliances,Structure")
    add_dv(ws, "E3:E2002", "Workmanship,Material,Manufacturer")
    add_dv(ws, "L3:L2002", "Active,Expired,Claimed")

    ws.conditional_formatting.add("L3:L2002",
        CellIsRule(operator='equal', formula=['"Active"'], fill=green_fill))
    ws.conditional_formatting.add("L3:L2002",
        CellIsRule(operator='equal', formula=['"Expired"'], fill=red_fill))
    ws.conditional_formatting.add("L3:L2002",
        CellIsRule(operator='equal', formula=['"Claimed"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _after_sales(ws):
    headers = [
        "Service Request ID", "Client ID", "Project ID",
        "Date of Request", "Issue Description",
        "Category", "Priority", "Assigned To",
        "Scheduled Date", "Completion Date",
        "Cost (if outside warranty)", "Status",
        "Client Satisfaction (1-5)", "Notes"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "🔧 AFTER-SALES SERVICE / MAINTENANCE", mc)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"SRV0000"))'

    add_dv(ws, "F3:F5002", "Repair,Replacement,Inspection,Warranty Claim")
    add_dv(ws, "G3:G5002", "High,Medium,Low")
    add_dv(ws, "L3:L5002", "Open,Scheduled,In Progress,Completed,Closed")
    add_dv(ws, "M3:M5002", "1,2,3,4,5")

    ws.conditional_formatting.add("L3:L5002",
        CellIsRule(operator='equal', formula=['"Completed"'], fill=green_fill))
    ws.conditional_formatting.add("L3:L5002",
        CellIsRule(operator='equal', formula=['"Closed"'],
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type='solid')))
    ws.conditional_formatting.add("L3:L5002",
        CellIsRule(operator='equal', formula=['"Open"'], fill=yellow_cond))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


# ═════════════════════════════════════════════════════════════════════════
# 16. REPORTING & ANALYTICS
# ═════════════════════════════════════════════════════════════════════════

def _project_status(ws):
    headers = [
        "Project ID", "Project Name", "Current Phase",
        "Overall % Complete (Auto)", "Budget Status",
        "Timeline Status", "Open Issues Count",
        "Pending Approvals", "Key Milestones Status",
        "Traffic Light Status"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📊 PROJECT STATUS REPORT", mc)

    for r in range(3, 503):
        ws.cell(row=r, column=4).value = (
            f'=IF(A{r}="","",IFERROR(VLOOKUP(A{r},\'Project Master Data\'!A:R,18,0),0))'
        )
        ws.cell(row=r, column=5).value = (
            f'=IF(A{r}="","",IFERROR(IF(VLOOKUP(A{r},\'Project Master Data\'!A:Q,17,0)>VLOOKUP(A{r},\'Project Master Data\'!A:P,16,0),"Over Budget","Within Budget"),""))'
        )

    add_dv(ws, "J3:J502", "Red,Amber,Green")

    ws.conditional_formatting.add("J3:J502",
        CellIsRule(operator='equal', formula=['"Green"'], fill=green_fill))
    ws.conditional_formatting.add("J3:J502",
        CellIsRule(operator='equal', formula=['"Amber"'], fill=yellow_cond))
    ws.conditional_formatting.add("J3:J502",
        CellIsRule(operator='equal', formula=['"Red"'], fill=red_fill))

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _weekly_progress(ws):
    headers = [
        "Report ID", "Week Number", "Date Range",
        "Project ID", "Work Completed This Week",
        "Work Planned Next Week", "Issues / Risks",
        "Resources on Site", "Materials Received",
        "Inspections Done", "Visitors to Site",
        "Weather Conditions", "Photo Links",
        "Report Prepared By"
    ]
    mc = write_headers(ws, headers)
    add_title_row(ws, "📝 WEEKLY PROGRESS REPORT", mc)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(D{r}="","",TEXT(ROW()-2,"WPR0000"))'

    freeze_and_filter(ws, 2, mc)
    auto_width(ws, mc)


def _kpi_dashboard(ws):
    add_title_row(ws, "📈 KPI DASHBOARD", 12)

    # KPI cards
    kpis = [
        ("B3", "Active Projects", "B4",
         '=COUNTIF(\'Project Master Data\'!M3:M502,"In Progress")'),
        ("D3", "Completed Projects", "D4",
         '=COUNTIF(\'Project Master Data\'!M3:M502,"Completed")'),
        ("F3", "Delayed Phases", "F4",
         '=COUNTIF(\'Project Phases\'!K3:K2002,"Delayed")'),
        ("H3", "Avg Client Satisfaction", "H4",
         '=IFERROR(ROUND(AVERAGE(\'Client Satisfaction\'!K3:K502),1),"N/A")'),
        ("J3", "Open Snags", "J4",
         '=COUNTIF(\'Defect Snagging List\'!J3:J5002,"Open")'),
        ("L3", "Safety Incidents (YTD)", "L4",
         '=COUNTA(\'Incident Accident Log\'!A3:A2002)'),
    ]
    for lc, label, vc, formula in kpis:
        ws[lc] = label; ws[lc].font = kpi_label
        ws[vc] = formula; ws[vc].font = kpi_font

    # Contractor rankings
    r = 7
    ws.cell(row=r, column=2, value="TOP CONTRACTOR PERFORMANCE RANKINGS").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1
    for ci, h in enumerate(["Rank", "Contractor ID", "Contractor Name", "Overall Rating", "Projects Done", "Would Hire Again %"]):
        ws.cell(row=r, column=2+ci, value=h)
    style_header_row(ws, r, 6)

    # Revenue summary
    r2 = r + 15
    ws.cell(row=r2, column=2, value="FINANCIAL KPIs").font = section_font
    ws.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=5)
    ws.cell(row=r2, column=2).fill = section_fill
    r2 += 1
    fin_kpis = [
        ("Total Revenue (YTD)", "='Income Tracker'!F10005"),
        ("Total Expenses (YTD)", "='Expense Tracker'!I10005"),
        ("Net Profit (YTD)", "=B{0}-B{1}".format(r2+2, r2+3)),
        ("Profit Margin %", "=IFERROR(ROUND(B{0}/B{1}*100,1),0)".format(r2+2, r2+2)),
    ]
    for i, (label, formula) in enumerate(fin_kpis):
        row = r2 + i
        ws.cell(row=row, column=2, value=label).font = Font(bold=True)
        ws.cell(row=row, column=3).value = formula
        ws.cell(row=row, column=3).number_format = '#,##0'

    auto_width(ws, 12)


# ═════════════════════════════════════════════════════════════════════════
# MAIN DASHBOARD
# ═════════════════════════════════════════════════════════════════════════

def _main_dashboard(ws, wb):
    max_col = 14
    # Title
    ws.merge_cells("A1:N1")
    ws["A1"] = "🏠  COMPREHENSIVE HOME RENOVATION MANAGEMENT SYSTEM – DASHBOARD"
    ws["A1"].font = Font(name='Calibri', bold=True, color=WHITE, size=18)
    ws["A1"].fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type='solid')
    ws["A1"].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 45

    ws.merge_cells("A2:N2")
    ws["A2"] = "Company: Your Renovation Company    |    Currency: ₹    |    Tax Rate: 18%"
    ws["A2"].font = Font(name='Calibri', italic=True, color=MID_BLUE, size=11)
    ws["A2"].alignment = Alignment(horizontal='center')

    # ── KPI Cards Row 1 ────────────────────────────────────────────────
    kpi1 = [
        ("B4", "Active Projects", "B5",
         '=COUNTIF(\'Project Master Data\'!M3:M502,"In Progress")'),
        ("D4", "Completed Projects", "D5",
         '=COUNTIF(\'Project Master Data\'!M3:M502,"Completed")'),
        ("F4", "Total Revenue (YTD)", "F5", "='Income Tracker'!F10005"),
        ("H4", "Total Expenses (YTD)", "H5", "='Expense Tracker'!I10005"),
        ("J4", "Net Profit (YTD)", "J5", "=F5-H5"),
        ("L4", "Open Snags", "L5", '=COUNTIF(\'Defect Snagging List\'!J3:J5002,"Open")'),
    ]
    for lc, label, vc, formula in kpi1:
        ws[lc] = label; ws[lc].font = kpi_label
        ws[lc].alignment = Alignment(horizontal='center')
        ws[lc].fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')
        ws[vc] = formula
        ws[vc].font = Font(name='Calibri', bold=True, size=20, color=DARK_BLUE)
        ws[vc].alignment = Alignment(horizontal='center')
        if "Revenue" in label or "Expenses" in label or "Profit" in label:
            ws[vc].number_format = '#,##0'

    # ── KPI Cards Row 2 ────────────────────────────────────────────────
    kpi2 = [
        ("B7", "Delayed Phases", "B8",
         '=COUNTIF(\'Project Phases\'!K3:K2002,"Delayed")'),
        ("D7", "Budget Overruns", "D8",
         '=COUNTIF(\'Budget Planning\'!O3:O5002,"Over Budget")'),
        ("F7", "Low Stock Alerts", "F8",
         '=COUNTIF(\'Material Inventory\'!L3:L5002,"*REORDER*")'),
        ("H7", "Pending Payments", "H8",
         '=COUNTIF(\'Contractor Payments\'!R3:R10002,"Pending")'),
        ("J7", "Open Incidents", "J8",
         '=COUNTIF(\'Incident Accident Log\'!O3:O2002,"Open")'),
        ("L7", "Permits Expiring <30d", "L8",
         '=COUNTIF(\'Permits Compliance\'!L3:L1002,"<30")'),
    ]
    for lc, label, vc, formula in kpi2:
        ws[lc] = label; ws[lc].font = kpi_label
        ws[lc].alignment = Alignment(horizontal='center')
        ws[lc].fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')
        ws[vc] = formula
        ws[vc].font = Font(name='Calibri', bold=True, size=20, color=DARK_BLUE)
        ws[vc].alignment = Alignment(horizontal='center')

    # ── Module Quick Links ─────────────────────────────────────────────
    r = 11
    ws.cell(row=r, column=2, value="📋 MODULE QUICK LINKS").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1

    modules = [
        ("🏠 Project Module", "Project Master Data", "Project Phases", "Task Management", "Gantt Chart"),
        ("👤 Client Module", "Client Master Data", "Client Communication Log", "Client Satisfaction", ""),
        ("💰 Budget & Finance", "Budget Planning", "Cost Tracking", "Budget Dashboard", "Change Orders"),
        ("👷 Contractor & Vendor", "Contractor Master Data", "Vendor Master Data", "Contractor Performance", "Contractor Payments"),
        ("🧱 Materials", "Material Master List", "Material Estimation", "Purchase Orders", "Material Inventory"),
        ("👷 Labor", "Worker Master Data", "Worker Attendance", "Labor Cost Calculation", "Productivity Tracker"),
        ("🏠 Room Tracking", "Room Area Master", "Room Work Checklist", "Room Cost Summary", ""),
        ("🎨 Design", "Design Requirements", "Measurements Specs", "Material Selection Board", ""),
        ("🔧 Equipment", "Equipment Inventory", "Equipment Usage Log", "Equipment Maintenance", ""),
        ("🔍 Quality", "Quality Standards", "Inspection Log", "Defect Snagging List", ""),
        ("🦺 Safety", "Safety Checklist", "Incident Accident Log", "Permits Compliance", ""),
        ("📄 Quotation & Invoice", "Quotation Builder", "Invoice Management", "Payment Receipts", ""),
        ("💹 Finance", "Income Tracker", "Expense Tracker", "Profit Loss Project", "Cash Flow Tracker"),
        ("📁 Documents", "Document Register", "Contract Register", "", ""),
        ("🛡️ Warranty", "Warranty Register", "After-Sales Service", "", ""),
        ("📊 Reports", "Project Status Report", "Weekly Progress Report", "KPI Dashboard", ""),
    ]

    for mod in modules:
        ws.cell(row=r, column=2, value=mod[0]).font = Font(bold=True, color=DARK_BLUE)
        ws.cell(row=r, column=2).fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type='solid')
        for ci, name in enumerate(mod[1:], 3):
            if name:
                ws.cell(row=r, column=ci, value=f"→ {name}")
                ws.cell(row=r, column=ci).font = Font(color=MID_BLUE, underline='single')
        r += 1

    # ── Today's Tasks & Upcoming Deadlines ─────────────────────────────
    r += 1
    ws.cell(row=r, column=2, value="📅 TODAY'S TASKS & UPCOMING DEADLINES").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1
    for ci, h in enumerate(["Task ID", "Task Name", "Project", "Assigned To", "Due Date", "Status"]):
        ws.cell(row=r, column=2+ci, value=h)
    style_header_row(ws, r, 6)

    # ── Budget Overview ────────────────────────────────────────────────
    r += 12
    ws.cell(row=r, column=2, value="💰 BUDGET OVERVIEW").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1
    budget_items = [
        ("Total Budget Allocated", "=SUM('Budget Planning'!I3:I5002)"),
        ("Total Spent", "=SUM('Budget Planning'!L3:L5002)"),
        ("Remaining", "=B{0}-B{1}".format(r+2, r+3)),
        ("Utilization %", '=IFERROR(ROUND(B{0}/B{1}*100,1)&"%","N/A")'.format(r+3, r+2)),
    ]
    for i, (label, formula) in enumerate(budget_items):
        row = r + 1 + i
        ws.cell(row=row, column=2, value=label).font = Font(bold=True)
        ws.cell(row=row, column=3).value = formula
        ws.cell(row=row, column=3).number_format = '#,##0'
        for c in range(2, 4):
            ws.cell(row=row, column=c).border = thin_border

    # ── Pending Payments ───────────────────────────────────────────────
    r += 8
    ws.cell(row=r, column=2, value="💵 PENDING PAYMENTS").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1
    ws.cell(row=r, column=2, value="Client Payments Pending:")
    ws.cell(row=r, column=3).value = '=COUNTIF(\'Invoice Management\'!O3:O5002,"Sent")+COUNTIF(\'Invoice Management\'!O3:O5002,"Partially Paid")'
    ws.cell(row=r+1, column=2, value="Contractor Payments Pending:")
    ws.cell(row=r+1, column=3).value = '=COUNTIF(\'Contractor Payments\'!R3:R10002,"Pending")'
    ws.cell(row=r+2, column=2, value="Total Outstanding Amount:")
    ws.cell(row=r+2, column=3).value = '=SUM(\'Invoice Management\'!N3:N5002)'
    ws.cell(row=r+2, column=3).number_format = '#,##0'

    # Column widths
    for col in range(1, 15):
        ws.column_dimensions[get_column_letter(col)].width = 20

    ws.freeze_panes = "A3"


# ═════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("🏠 Building Comprehensive Home Renovation Management System...")
    wb = build_workbook()

    output_path = "Home_Renovation_Management_System.xlsx"
    wb.save(output_path)
    print(f"✅ Successfully created: {output_path}")
    print(f"   Total sheets: {len(wb.sheetnames)}")
    print(f"   Sheets: {', '.join(wb.sheetnames)}")
