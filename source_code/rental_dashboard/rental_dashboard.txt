"""
Rental Property Management Dashboard
Real Estate Investment Tracker | Excel Rental Portfolio Manager
Built with XlsxWriter - 20+ professional worksheets
"""

import xlsxwriter
from datetime import datetime, date, timedelta
import random

random.seed(42)
OUTPUT = "/home/user/Excel/Rental_Property_Management_Dashboard.xlsx"

# =========================================================================
# COLOR PALETTE - Modern, Professional
# =========================================================================
NAVY      = "#1E3A5F"
EMERALD   = "#10B981"
ORANGE    = "#F59E0B"
RED       = "#EF4444"
LGRAY     = "#F8FAFC"
WHITE     = "#FFFFFF"
DGRAY     = "#334155"
MGRAY     = "#64748B"
BORDER    = "#E2E8F0"
GOLD      = "#FBBF24"
TEAL      = "#0EA5E9"
PURPLE    = "#8B5CF6"
PINK      = "#EC4899"
LGREEN    = "#D1FAE5"
LRED      = "#FEE2E2"
LYELLOW   = "#FEF3C7"
LBLUE     = "#DBEAFE"

wb = xlsxwriter.Workbook(OUTPUT)

def fmt(props):
    return wb.add_format(props)

# =========================================================================
# REUSABLE FORMATS
# =========================================================================
title_main    = fmt({"bold": True, "font_size": 32, "font_color": WHITE, "font_name": "Calibri", "align": "left", "valign": "vcenter", "bg_color": NAVY})
title_sub     = fmt({"bold": True, "font_size": 14, "font_color": WHITE, "font_name": "Calibri", "italic": True, "align": "left", "valign": "vcenter", "bg_color": NAVY})
title_sheet   = fmt({"bold": True, "font_size": 28, "font_color": NAVY, "font_name": "Calibri", "align": "left", "valign": "vcenter"})
subtitle      = fmt({"italic": True, "font_size": 11, "font_color": MGRAY, "align": "left", "valign": "vcenter"})

section_h     = fmt({"bold": True, "font_size": 16, "font_color": WHITE, "bg_color": NAVY, "align": "left", "valign": "vcenter", "left": 2})
section_h_em  = fmt({"bold": True, "font_size": 16, "font_color": WHITE, "bg_color": EMERALD, "align": "left", "valign": "vcenter", "left": 2})
section_h_or  = fmt({"bold": True, "font_size": 16, "font_color": WHITE, "bg_color": ORANGE, "align": "left", "valign": "vcenter", "left": 2})

th            = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": NAVY, "align": "center", "valign": "vcenter", "border": 1, "border_color": NAVY, "text_wrap": True})
th_em         = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": EMERALD, "align": "center", "valign": "vcenter", "border": 1, "border_color": EMERALD, "text_wrap": True})
th_or         = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": ORANGE, "align": "center", "valign": "vcenter", "border": 1, "border_color": ORANGE, "text_wrap": True})
th_red        = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": RED, "align": "center", "valign": "vcenter", "border": 1, "border_color": RED, "text_wrap": True})
th_teal       = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": TEAL, "align": "center", "valign": "vcenter", "border": 1, "border_color": TEAL, "text_wrap": True})
th_purple     = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": PURPLE, "align": "center", "valign": "vcenter", "border": 1, "border_color": PURPLE, "text_wrap": True})
th_dgray      = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": DGRAY, "align": "center", "valign": "vcenter", "border": 1, "border_color": DGRAY, "text_wrap": True})

cell          = fmt({"font_size": 10, "align": "left", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE})
cell_alt      = fmt({"font_size": 10, "align": "left", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY})
cell_c        = fmt({"font_size": 10, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE})
cell_c_alt    = fmt({"font_size": 10, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY})
cell_r        = fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE, "num_format": "$#,##0.00"})
cell_r_alt    = fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0.00"})
cell_r0       = fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE, "num_format": "$#,##0"})
cell_pct      = fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE, "num_format": "0.00%"})
cell_pct_alt  = fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "0.00%"})
cell_int      = fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE, "num_format": "#,##0"})
cell_int_alt  = fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "#,##0"})

def cell_date(alt=False):
    return fmt({"font_size": 10, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY if alt else WHITE, "num_format": "mm/dd/yyyy"})

nav_btn       = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": EMERALD, "align": "center", "valign": "vcenter", "border": 1, "border_color": EMERALD})
nav_btn_or    = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": ORANGE, "align": "center", "valign": "vcenter", "border": 1, "border_color": ORANGE})
nav_btn_red   = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": RED, "align": "center", "valign": "vcenter", "border": 1, "border_color": RED})
nav_btn_nv    = fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": NAVY, "align": "center", "valign": "vcenter", "border": 1, "border_color": NAVY})

note          = fmt({"italic": True, "font_size": 10, "font_color": MGRAY, "align": "left", "valign": "top", "text_wrap": True, "border": 1, "border_color": BORDER, "bg_color": LGRAY, "left": 2, "right": 2, "top": 2, "bottom": 2})
big_total     = fmt({"bold": True, "font_size": 12, "bg_color": NAVY, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "$#,##0.00"})
big_total_em  = fmt({"bold": True, "font_size": 12, "bg_color": EMERALD, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "$#,##0.00"})
big_total_or  = fmt({"bold": True, "font_size": 12, "bg_color": ORANGE, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "$#,##0.00"})
big_total_red = fmt({"bold": True, "font_size": 12, "bg_color": RED, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "$#,##0.00"})

status_g      = fmt({"bold": True, "font_size": 10, "font_color": "#065F46", "bg_color": LGREEN, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER})
status_y      = fmt({"bold": True, "font_size": 10, "font_color": "#92400E", "bg_color": LYELLOW, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER})
status_r      = fmt({"bold": True, "font_size": 10, "font_color": "#991B1B", "bg_color": LRED, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER})
status_b      = fmt({"bold": True, "font_size": 10, "font_color": "#1E3A8A", "bg_color": LBLUE, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER})

sidebar_h     = fmt({"bold": True, "font_size": 12, "font_color": NAVY, "bg_color": LGRAY, "align": "left", "valign": "vcenter", "left": 1, "top": 1, "bottom": 1, "border": 1, "border_color": BORDER})
side_fmt      = fmt({"bold": True, "font_size": 10, "font_color": WHITE, "bg_color": NAVY, "align": "left", "valign": "vcenter", "left": 1, "border": 1, "border_color": "#2C4F7F"})
side_alt      = fmt({"bold": True, "font_size": 10, "font_color": WHITE, "bg_color": EMERALD, "align": "left", "valign": "vcenter", "left": 1, "border": 1, "border_color": EMERALD, "bold": True})

# =========================================================================
# PROPERTIES (shared data)
# =========================================================================
properties = [
    ("P001", "Maple Street Duplex",  "123 Maple St, Austin, TX 78701",      "Duplex",    4, 2.0, "2018-03-15", 320000, 415000, 240000, 1800, 0,    5400, "Occupied"),
    ("P002", "Oak Avenue Cottage",   "456 Oak Ave, Austin, TX 78702",      "Single",    3, 2.0, "2019-07-22", 275000, 355000, 200000, 1500, 0,    4800, "Occupied"),
    ("P003", "Riverside Condo",      "789 River Rd #4B, Austin, TX 78703",  "Condo",     2, 2.0, "2020-01-10", 215000, 268000, 165000, 1100, 350,  3200, "Occupied"),
    ("P004", "Pine Hill Townhouse",  "101 Pine Hill, Austin, TX 78704",     "Townhouse", 3, 2.5, "2020-09-05", 305000, 378000, 230000, 1700, 220,  4900, "Vacant"),
    ("P005", "Elm Drive Bungalow",   "212 Elm Dr, Austin, TX 78705",        "Single",    2, 1.0, "2021-04-18", 198000, 245000, 150000, 1300, 0,    3600, "Occupied"),
    ("P006", "Cedar Park Triplex",   "55 Cedar Park Blvd, Austin, TX 78717","Triplex",   6, 3.0, "2021-11-30", 480000, 595000, 360000, 2400, 0,    7600, "Occupied"),
    ("P007", "Birch Lane Apartment", "88 Birch Ln #12, Austin, TX 78745",   "Apartment", 1, 1.0, "2022-02-14", 145000, 178000, 110000, 850,  180,  2400, "Occupied"),
    ("P008", "Willow Creek House",   "330 Willow Creek, Austin, TX 78750",  "Single",    4, 3.0, "2022-08-21", 425000, 510000, 320000, 2100, 0,    6400, "Occupied"),
    ("P009", "Sunset View Loft",     "600 Sunset Blvd #7, Austin, TX 78704","Loft",      1, 1.0, "2023-03-08", 189000, 224000, 145000, 950,  280,  2800, "Renovation"),
    ("P010", "Highland Park Home",   "425 Highland Pk, Austin, TX 78731",   "Single",    5, 4.0, "2023-10-17", 625000, 712000, 460000, 2800, 0,    9200, "Occupied"),
]

tenants = [
    ("T001", "Alex Johnson",        "(512) 555-0101", "alex.j@email.com",       "Mary Johnson / (512) 555-0102", "2024-01-15", "2025-01-14", 2200, 4400, "Active"),
    ("T002", "Sara Williams",       "(512) 555-0103", "sara.w@email.com",       "Tom Williams / (512) 555-0104", "2023-08-01", "2025-07-31", 1850, 3700, "Active"),
    ("T003", "Michael Brown",       "(512) 555-0105", "m.brown@email.com",      "Lisa Brown / (512) 555-0106",   "2024-06-01", "2026-05-31", 2400, 4800, "Active"),
    ("T004", "Emily Davis",         "(512) 555-0107", "emily.d@email.com",      "John Davis / (512) 555-0108",   "2024-03-15", "2025-03-14", 1650, 3300, "Expiring"),
    ("T005", "David Miller",        "(512) 555-0109", "d.miller@email.com",     "Anna Miller / (512) 555-0110",  "2023-11-01", "2025-10-31", 2100, 4200, "Active"),
    ("T006", "Jessica Wilson",      "(512) 555-0111", "j.wilson@email.com",     "Mark Wilson / (512) 555-0112",  "2024-09-01", "2025-08-31", 1950, 3900, "Active"),
    ("T007", "Chris Martinez",      "(512) 555-0113", "c.martinez@email.com",   "Sofia Martinez / (512) 555-0114","2024-02-01", "2025-01-31", 2300, 4600, "Expiring"),
    ("T008", "Amanda Garcia",       "(512) 555-0115", "a.garcia@email.com",     "Luis Garcia / (512) 555-0116",  "2024-04-15", "2026-04-14", 1750, 3500, "Active"),
    ("T009", "Robert Anderson",     "(512) 555-0117", "r.anderson@email.com",   "Kate Anderson / (512) 555-0118","2024-07-01", "2025-06-30", 2550, 5100, "Active"),
    ("T010", "Olivia Thomas",       "(512) 555-0119", "o.thomas@email.com",     "Ben Thomas / (512) 555-0120",   "2024-10-01", "2025-09-30", 2000, 4000, "Active"),
    ("T011", "James Jackson",       "(512) 555-0121", "j.jackson@email.com",    "Emma Jackson / (512) 555-0122", "2024-05-15", "2025-05-14", 2150, 4300, "Active"),
    ("T012", "Sophia White",        "(512) 555-0123", "s.white@email.com",      "Liam White / (512) 555-0124",   "2023-12-01", "2024-11-30", 1900, 3800, "Expiring"),
    ("T013", "Daniel Harris",       "(512) 555-0125", "d.harris@email.com",     "Mia Harris / (512) 555-0126",   "2024-08-01", "2026-07-31", 2700, 5400, "Active"),
    ("T014", "Isabella Martin",     "(512) 555-0127", "i.martin@email.com",     "Noah Martin / (512) 555-0128",  "2024-11-15", "2025-11-14", 2050, 4100, "Active"),
]

contractors = [
    ("C001", "Mike's Electric",    "Electrician", "Mike Roberts",  "(512) 555-0201", "mike@electric.com",    "Licensed, Bonded", 4.8, 250, 50),
    ("C002", "Pro Plumbing Co.",   "Plumber",     "Tom Wilson",    "(512) 555-0202", "tom@proplumb.com",     "24/7 Emergency",   4.6, 180, 60),
    ("C003", "Color Perfect",      "Painter",     "Jose Garcia",   "(512) 555-0203", "jose@colorperfect.com","Interior/Exterior",4.9, 150, 45),
    ("C004", "Shine Cleaners",     "Cleaner",     "Maria Lopez",   "(512) 555-0204", "maria@shineclean.com", "Move-out Special", 4.7, 120, 40),
    ("C005", "FixIt Handyman",     "Handyman",    "Dave Chen",     "(512) 555-0205", "dave@fixit.com",       "General Repairs",  4.8, 200, 55),
    ("C006", "GreenScapes",        "Landscaper",  "Sam Johnson",   "(512) 555-0206", "sam@greenscapes.com",  "Lawn Maintenance",4.5, 130, 40),
    ("C007", "Cool Air HVAC",      "HVAC",        "Rick Patel",    "(512) 555-0207", "rick@coolair.com",     "AC & Heating",     4.9, 280, 75),
    ("C008", "Pest Patrol",        "Exterminator","Ben Adams",     "(512) 555-0208", "ben@pestpatrol.com",   "Quarterly Service",4.6, 160, 50),
]

# =========================================================================
# PAGE SETUP HELPER
# =========================================================================
def setup_page(ws, landscape=True, fit=True):
    if landscape:
        ws.set_landscape()
    ws.set_paper(9)
    if fit:
        ws.fit_to_pages(1, 0)
    ws.set_margins(0.4, 0.4, 0.4, 0.4)
    ws.set_header('&L&"Calibri,Bold"&14Rental Property Management Dashboard&R&"Calibri,Italic"&11Real Estate Investment Tracker')
    ws.set_footer('&LC&A &RPage &P of &N')

# =========================================================================
# NAVIGATION SIDEBAR
# =========================================================================
def add_sidebar(ws, active_idx=None, sidebar_col=0):
    nav_items = [
        ("🏠 Executive Dashboard",        "Dashboard"),
        ("🏘️  Property Database",          "Properties"),
        ("👥 Tenant Database",             "Tenants"),
        ("💵 Rental Income",               "Income"),
        ("📉 Expenses",                    "Expenses"),
        ("🏦 Mortgage Tracker",            "Mortgage"),
        ("🔧 Maintenance Log",             "Maintenance"),
        ("📅 Lease Tracker",               "Lease"),
        ("🚪 Vacancy Tracker",             "Vacancy"),
        ("📈 ROI Calculator",              "ROI"),
        ("🏷️  Property Valuation",         "Valuation"),
        ("🧾 Tax Summary",                 "Tax"),
        ("💡 Utility Tracker",             "Utility"),
        ("🛠️  Contractors",                 "Contractors"),
        ("📂 Document Checklist",          "Documents"),
        ("🔍 Inspections",                 "Inspections"),
        ("🏗️  Capital Improvements",        "Capital"),
        ("📊 Monthly P&L",                 "PnL"),
        ("📆 Annual Report",               "Annual"),
        ("⚙️  Settings",                    "Settings"),
        ("📖 Quick Start Guide",           "Guide"),
    ]
    ws.set_column(sidebar_col, sidebar_col, 26)
    ws.set_column(sidebar_col + 1, sidebar_col + 1, 2)
    ws.write(1, sidebar_col, "  NAVIGATION", sidebar_h)
    ws.set_row(1, 24)
    for i, (label, sheet) in enumerate(nav_items):
        r = 2 + i
        chosen = side_alt if i == active_idx else side_fmt
        ws.set_row(r, 22)
        ws.write_url(r, sidebar_col, f"internal:'{sheet}'!A1", chosen, "  " + label)

# =========================================================================
# WORKBOOK PROPERTIES
# =========================================================================
wb.set_properties({
    "title":    "Rental Property Management Dashboard",
    "subject":  "Real Estate Investment Tracker",
    "author":   "Premium Templates Co.",
    "manager":  "Property Management",
    "company":  "Premium Templates",
    "category": "Real Estate / Property Management",
    "keywords": "rental, property, real estate, investment, ROI, landlord, tenant, mortgage",
    "comments": "Premium Rental Property Management Dashboard - Real Estate Investment Tracker - Excel Rental Portfolio Manager"
})
wb.set_calc_mode("auto")

# =========================================================================
# 1. EXECUTIVE DASHBOARD
# =========================================================================
ws = wb.add_worksheet("Dashboard")
ws.set_tab_color(NAVY)
setup_page(ws)
ws.set_column("A:A", 2)
ws.set_column("B:G", 16)
ws.set_column("H:H", 2)
ws.hide_gridlines(2)

# Title block
ws.set_row(0, 8)
ws.set_row(1, 40)
ws.merge_range("B2:G2", "  Rental Property Management Dashboard", title_main)
ws.set_row(2, 24)
ws.merge_range("B3:G3", "  Real Estate Investment Tracker  |  Excel Rental Portfolio Manager", title_sub)
ws.set_row(3, 8)
ws.merge_range("B4:G4", "  Executive Summary — Portfolio Overview, KPIs, Trends, and Comparative Analytics", subtitle)
ws.set_row(4, 8)

# KPI Cards
kpi_data = [
    ("🏠 TOTAL PROPERTIES",  '=COUNTA(Properties!E6:E100)',                             "0",          NAVY),
    ("💵 MONTHLY INCOME",    '=SUMIF(Income!L:L,"Paid",Income!I:I)/12+SUMIF(Income!L:L,"Late",Income!I:I)/12',  "$#,##0",   EMERALD),
    ("📈 ANNUAL INCOME",     '=SUMIF(Income!L:L,"Paid",Income!I:I)+SUMIF(Income!L:L,"Late",Income!I:I)',        "$#,##0",    EMERALD),
    ("💸 MONTHLY EXPENSES",  '=SUM(Expenses!H:H)/12',                                    "$#,##0",   ORANGE),
    ("💰 NET CASH FLOW",     '=(SUMIF(Income!L:L,"Paid",Income!I:I)+SUMIF(Income!L:L,"Late",Income!I:I))/12 - SUM(Expenses!H:H)/12', "$#,##0", NAVY),
    ("📊 OCCUPANCY RATE",    '=1-(SUM(Vacancy!E:E)/365/COUNTA(Properties!E6:E100))',     "0.0%", NAVY),
]

kpi_row1 = 6
kpi_row2 = 9
for i, (label, formula, num_fmt, color) in enumerate(kpi_data):
    col = 1 + i*2
    if col > 5: col = 1 + (i-3)*2
    ws.set_row(kpi_row1 if i < 3 else kpi_row2, 22)
    ws.set_row(kpi_row1+1 if i < 3 else kpi_row2+1, 36)
    if i < 3:
        ws.merge_range(kpi_row1, col, kpi_row1, col+1, "  " + label, fmt({"bold": True, "font_size": 10, "font_color": WHITE, "bg_color": color, "align": "left", "valign": "vcenter", "left": 2}))
        val_fmt = fmt({"bold": True, "font_size": 18, "font_color": color, "bg_color": WHITE, "align": "left", "valign": "vcenter", "left": 2, "right": 2, "top": 1, "bottom": 1, "border": 1, "border_color": BORDER, "num_format": num_fmt})
        ws.merge_range(kpi_row1+1, col, kpi_row1+1, col+1, "  0", val_fmt)
        ws.write_formula(kpi_row1+1, col, formula, val_fmt)
    else:
        ws.merge_range(kpi_row2, col, kpi_row2, col+1, "  " + label, fmt({"bold": True, "font_size": 10, "font_color": WHITE, "bg_color": color, "align": "left", "valign": "vcenter", "left": 2}))
        val_fmt = fmt({"bold": True, "font_size": 18, "font_color": color, "bg_color": WHITE, "align": "left", "valign": "vcenter", "left": 2, "right": 2, "top": 1, "bottom": 1, "border": 1, "border_color": BORDER, "num_format": num_fmt})
        ws.merge_range(kpi_row2+1, col, kpi_row2+1, col+1, "  0", val_fmt)
        ws.write_formula(kpi_row2+1, col, formula, val_fmt)

# Section header
ws.set_row(12, 24)
ws.merge_range("B13:G13", "  📊  PORTFOLIO ANALYTICS — MONTHLY TRENDS & COMPARISONS", section_h)

# Charts will be inserted later once data sheets exist. We need to track PnL/Mortgage/ROI references.
# Income vs Expenses vs Cash Flow
chart_income = wb.add_chart({"type": "column"})
chart_income.add_series({
    "name":       "Monthly Income",
    "categories": ["PnL", 2, 0, 13, 0],
    "values":     ["PnL", 2, 1, 13, 1],
    "fill":       {"color": EMERALD},
    "border":     {"color": EMERALD},
    "data_labels": {"value": True, "num_format": "$#,##0", "font": {"size": 8}},
})
chart_income.add_series({
    "name":       "Monthly Expenses",
    "categories": ["PnL", 2, 0, 13, 0],
    "values":     ["PnL", 2, 2, 13, 2],
    "fill":       {"color": ORANGE},
    "border":     {"color": ORANGE},
    "data_labels": {"value": True, "num_format": "$#,##0", "font": {"size": 8}},
})
chart_income.add_series({
    "name":       "Net Cash Flow",
    "categories": ["PnL", 2, 0, 13, 0],
    "values":     ["PnL", 2, 3, 13, 3],
    "fill":       {"color": NAVY},
    "border":     {"color": NAVY},
    "data_labels": {"value": True, "num_format": "$#,##0", "font": {"size": 8}},
})
chart_income.set_title({"name": "Monthly Income vs Expenses vs Cash Flow", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_income.set_x_axis({"name": "Month", "num_font": {"size": 9}})
chart_income.set_y_axis({"name": "Amount ($)", "num_format": "$#,##0", "num_font": {"size": 9}})
chart_income.set_size({"width": 540, "height": 280})
chart_income.set_legend({"position": "bottom"})
ws.insert_chart("B14", chart_income)

# Expense breakdown
chart_exp = wb.add_chart({"type": "doughnut"})
chart_exp.add_series({
    "name":       "Expense Categories",
    "categories": ["Expenses", 5, 1, 15, 1],
    "values":     ["Expenses", 5, 4, 15, 4],
    "data_labels": {"percentage": True, "category": True, "font": {"size": 8}},
})
chart_exp.set_title({"name": "Expense Breakdown by Category", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_exp.set_size({"width": 480, "height": 280})
chart_exp.set_legend({"position": "right", "font": {"size": 9}})
chart_exp.set_style(10)
ws.insert_chart("B30", chart_exp)

# Property comparison
chart_prop = wb.add_chart({"type": "bar"})
chart_prop.add_series({
    "name":       "Property Value",
    "categories": ["Properties", 5, 1, 14, 1],
    "values":     ["Properties", 5, 9, 14, 9],
    "fill":       {"color": NAVY},
    "border":     {"color": NAVY},
    "data_labels": {"value": True, "num_format": "$#,##0", "font": {"size": 8}},
})
chart_prop.set_title({"name": "Property Value Comparison", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_prop.set_x_axis({"num_format": "$#,##0"})
chart_prop.set_size({"width": 480, "height": 280})
chart_prop.set_legend({"none": True})
ws.insert_chart("F30", chart_prop)

# Occupancy trend
chart_occ = wb.add_chart({"type": "line"})
chart_occ.add_series({
    "name":       "Occupancy %",
    "categories": ["PnL", 2, 0, 13, 0],
    "values":     ["PnL", 2, 4, 13, 4],
    "line":       {"color": EMERALD, "width": 2.5},
    "marker":     {"type": "circle", "size": 6, "fill": {"color": EMERALD}},
    "data_labels": {"value": True, "num_format": "0.0%", "font": {"size": 8}},
})
chart_occ.set_title({"name": "Occupancy Rate Trend", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_occ.set_x_axis({"name": "Month"})
chart_occ.set_y_axis({"name": "Occupancy", "num_format": "0%", "min": 0, "max": 1})
chart_occ.set_size({"width": 540, "height": 260})
chart_occ.set_legend({"none": True})
ws.insert_chart("B46", chart_occ)

# Cash flow
chart_cf = wb.add_chart({"type": "area"})
chart_cf.add_series({
    "name":       "Cumulative Cash Flow",
    "categories": ["PnL", 2, 0, 13, 0],
    "values":     ["PnL", 2, 5, 13, 5],
    "fill":       {"color": TEAL, "transparency": 30},
    "border":     {"color": TEAL, "width": 2},
})
chart_cf.set_title({"name": "Cumulative Cash Flow (12 Months)", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_cf.set_x_axis({"name": "Month"})
chart_cf.set_y_axis({"name": "Amount ($)", "num_format": "$#,##0"})
chart_cf.set_size({"width": 480, "height": 260})
chart_cf.set_legend({"none": True})
ws.insert_chart("F46", chart_cf)

# Maintenance
chart_maint = wb.add_chart({"type": "column"})
chart_maint.add_series({
    "name":       "Maintenance Cost",
    "categories": ["Maintenance", 5, 0, 14, 0],
    "values":     ["Maintenance", 5, 4, 14, 4],
    "fill":       {"color": ORANGE},
    "border":     {"color": ORANGE},
    "data_labels": {"value": True, "num_format": "$#,##0", "font": {"size": 8}},
})
chart_maint.set_title({"name": "Maintenance Cost per Property", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_maint.set_y_axis({"num_format": "$#,##0"})
chart_maint.set_size({"width": 540, "height": 260})
chart_maint.set_legend({"none": True})
ws.insert_chart("B62", chart_maint)

# Loan balance
chart_loan = wb.add_chart({"type": "line"})
chart_loan.add_series({
    "name":       "Loan Balance",
    "categories": ["Mortgage", 5, 0, 14, 0],
    "values":     ["Mortgage", 5, 6, 14, 6],
    "line":       {"color": RED, "width": 2.5},
    "marker":     {"type": "circle", "size": 6, "fill": {"color": RED}},
})
chart_loan.set_title({"name": "Loan Balance Over Time", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_loan.set_y_axis({"num_format": "$#,##0"})
chart_loan.set_size({"width": 480, "height": 260})
chart_loan.set_legend({"none": True})
ws.insert_chart("F62", chart_loan)

# ROI
chart_roi = wb.add_chart({"type": "bar"})
chart_roi.add_series({
    "name":       "ROI %",
    "categories": ["ROI", 5, 0, 14, 0],
    "values":     ["ROI", 5, 5, 14, 5],
    "fill":       {"color": PURPLE},
    "border":     {"color": PURPLE},
    "data_labels": {"value": True, "num_format": "0.0%", "font": {"size": 8}},
})
chart_roi.set_title({"name": "Property ROI Comparison", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_roi.set_x_axis({"num_format": "0%"})
chart_roi.set_size({"width": 540, "height": 260})
chart_roi.set_legend({"none": True})
ws.insert_chart("B78", chart_roi)

# Footer
ws.set_row(94, 6)
ws.merge_range("B95:G95", "  © 2026 Premium Templates Co.  |  Rental Property Management Dashboard v2.0  |  support@premiumtemplates.com", subtitle)

# Hyperlink buttons
ws.set_row(96, 26)
ws.write_url("B97", "internal:'Properties'!A1", nav_btn, "  ➕ Add Property")
ws.write_url("C97", "internal:'Tenants'!A1", fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": EMERALD, "align": "center", "valign": "vcenter", "border": 1, "border_color": EMERALD}), "  👤 Add Tenant")
ws.write_url("D97", "internal:'Income'!A1", nav_btn_or, "  💵 Log Income")
ws.write_url("E97", "internal:'Expenses'!A1", nav_btn_red, "  📉 Log Expense")
ws.write_url("F97", "internal:'Guide'!A1", fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": PURPLE, "align": "center", "valign": "vcenter", "border": 1, "border_color": PURPLE}), "  📖 Quick Start")
ws.write_url("G97", "internal:'Settings'!A1", fmt({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": DGRAY, "align": "center", "valign": "vcenter", "border": 1, "border_color": DGRAY}), "  ⚙️  Settings")

ws.activate()
ws.set_first_sheet()

# =========================================================================
# 2. PROPERTY DATABASE
# =========================================================================
ws = wb.add_worksheet("Properties")
ws.set_tab_color(NAVY)
setup_page(ws)
add_sidebar(ws, active_idx=1)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 30)
ws.set_column("G:G", 14)
ws.set_column("H:N", 12)
ws.set_column("O:O", 12)
ws.set_column("P:P", 11)
ws.set_column("Q:R", 12)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:R2", "  🏘️  Property Database", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:R3", "  Manage your entire real estate portfolio with key metrics", subtitle)
ws.set_row(3, 8)

prop_headers = ["Property ID", "Property Name", "Address", "Type", "Bedrooms", "Bathrooms",
                "Purchase Date", "Purchase Price", "Current Value", "Mortgage", "Insurance",
                "HOA Fees", "Tax", "Status"]
ws.set_row(4, 32)
for i, h in enumerate(prop_headers):
    ws.write(4, 3 + i, h, th)

for i, p in enumerate(properties):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r0 if i % 2 == 0 else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_d = cell_date(False if i % 2 == 0 else True)
    ws.write(r, 3, p[0], bg_c)
    ws.write(r, 4, p[1], bg)
    ws.write(r, 5, p[2], bg)
    ws.write(r, 6, p[3], bg_c)
    ws.write(r, 7, p[4], bg_c)
    ws.write(r, 8, p[5], bg_c)
    ws.write(r, 9, datetime.strptime(p[6], "%Y-%m-%d").date(), bg_d)
    ws.write(r, 10, p[7], bg_r)
    ws.write(r, 11, p[8], bg_r)
    ws.write(r, 12, p[9], bg_r)
    ws.write(r, 13, p[10], bg_r)
    ws.write(r, 14, p[11], bg_r)
    ws.write(r, 15, p[12], bg_r)
    if p[13] == "Occupied":
        ws.write(r, 16, p[13], status_g)
    elif p[13] == "Vacant":
        ws.write(r, 16, p[13], status_y)
    else:
        ws.write(r, 16, p[13], status_b)

last_row = 5 + len(properties) - 1
ws.data_validation(5, 6, last_row, 6, {"validate": "list", "source": ["Single", "Duplex", "Triplex", "Condo", "Townhouse", "Apartment", "Loft", "Multi-Family"]})
ws.data_validation(5, 16, last_row, 16, {"validate": "list", "source": ["Occupied", "Vacant", "Renovation", "Listed", "Off-Market"]})
ws.autofilter(4, 3, last_row, 16)
ws.freeze_panes(5, 0)

ws.conditional_format(5, 11, last_row, 11, {"type": "3_color_scale", "min_color": "#FEE2E2", "mid_color": "#FEF3C7", "max_color": "#D1FAE5"})

sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 9, "  PORTFOLIO TOTAL", fmt({"bold": True, "font_size": 12, "bg_color": NAVY, "font_color": WHITE, "align": "right"}))
totals_v = [("=SUM(K6:K{})".format(last_row+1), 10),
            ("=SUM(L6:L{})".format(last_row+1), 11),
            ("=SUM(M6:M{})".format(last_row+1), 12),
            ("=SUM(N6:N{})".format(last_row+1), 13),
            ("=SUM(O6:O{})".format(last_row+1), 14),
            ("=SUM(P6:P{})".format(last_row+1), 15),
            ("=SUM(Q6:Q{})".format(last_row+1), 16)]
for f, c in totals_v:
    ws.write_formula(sum_row, c, f, big_total)

ws.merge_range(sum_row + 2, 3, sum_row + 5, 16,
    "💡 TIPS:\n• Click Auto-Filter arrows in the header to sort & filter.\n• Use drop-downs in Type and Status columns.\n• Add new rows below the last entry — KPIs & charts will update automatically.\n• All currency values use the setting from the Settings tab.", note)

# =========================================================================
# 3. TENANT DATABASE
# =========================================================================
ws = wb.add_worksheet("Tenants")
ws.set_tab_color(EMERALD)
setup_page(ws)
add_sidebar(ws, active_idx=2)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 16)
ws.set_column("G:G", 26)
ws.set_column("H:H", 22)
ws.set_column("I:K", 13)
ws.set_column("L:M", 14)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:M2", "  👥 Tenant Database", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:M3", "  Manage all tenants, leases, and contact information", subtitle)
ws.set_row(3, 8)

tenant_headers = ["Tenant ID", "Name", "Phone", "Email", "Emergency Contact",
                  "Lease Start", "Lease End", "Monthly Rent", "Security Deposit", "Status"]
ws.set_row(4, 32)
for i, h in enumerate(tenant_headers):
    ws.write(4, 3 + i, h, th_em)

for i, t in enumerate(tenants):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r0 if i % 2 == 0 else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_d = cell_date(False if i % 2 == 0 else True)
    ws.write(r, 3, t[0], bg_c)
    ws.write(r, 4, t[1], bg)
    ws.write(r, 5, t[2], bg_c)
    ws.write(r, 6, t[3], bg)
    ws.write(r, 7, t[4], bg)
    ws.write(r, 8, datetime.strptime(t[5], "%Y-%m-%d").date(), bg_d)
    ws.write(r, 9, datetime.strptime(t[6], "%Y-%m-%d").date(), bg_d)
    ws.write(r, 10, t[7], bg_r)
    ws.write(r, 11, t[8], bg_r)
    if t[9] == "Active":
        ws.write(r, 12, t[9], status_g)
    elif t[9] == "Expiring":
        ws.write(r, 12, t[9], status_y)
    else:
        ws.write(r, 12, t[9], status_r)

last_row = 5 + len(tenants) - 1
ws.data_validation(5, 12, last_row, 12, {"validate": "list", "source": ["Active", "Expiring", "Expired", "Past"]})
ws.autofilter(4, 3, last_row, 12)
ws.freeze_panes(5, 0)

ws.conditional_format(5, 12, last_row, 12, {"type": "text", "criteria": "containing", "value": "Active", "format": status_g})
ws.conditional_format(5, 12, last_row, 12, {"type": "text", "criteria": "containing", "value": "Expiring", "format": status_y})
ws.conditional_format(5, 12, last_row, 12, {"type": "text", "criteria": "containing", "value": "Expired", "format": status_r})

# Date conditional formatting for lease end
ws.conditional_format(5, 9, last_row, 9, {
    "type": "date", "criteria": "less than or equal to", "value": date.today(), "format": status_r
})
ws.conditional_format(5, 9, last_row, 9, {
    "type": "date", "criteria": "between", "minimum": date.today(), "maximum": date.today() + timedelta(days=60),
    "format": status_y
})

sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 9, "  TOTAL TENANTS", fmt({"bold": True, "font_size": 12, "bg_color": EMERALD, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row, 10, "=SUM(K6:K{})".format(last_row+1), big_total_em)
ws.write_formula(sum_row, 11, "=SUM(L6:L{})".format(last_row+1), big_total_em)

ws.merge_range(sum_row + 2, 3, sum_row + 5, 12,
    "💡 TIPS:\n• Lease End dates auto-highlight: 🟢 Active, 🟡 Expiring within 60 days, 🔴 Expired.\n• Add new tenants below the last row — dashboards update automatically.\n• Use the Auto-Filter to find tenants by status or property.", note)

# =========================================================================
# 4. RENTAL INCOME TRACKER
# =========================================================================
ws = wb.add_worksheet("Income")
ws.set_tab_color(EMERALD)
setup_page(ws)
add_sidebar(ws, active_idx=3)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 22)
ws.set_column("G:H", 13)
ws.set_column("I:J", 13)
ws.set_column("K:K", 18)
ws.set_column("L:L", 13)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:L2", "  💵 Rental Income Tracker", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:L3", "  Track monthly rent collections, late fees, and payment methods", subtitle)
ws.set_row(3, 8)

inc_headers = ["Income ID", "Property", "Tenant", "Due Date", "Paid Date", "Amount", "Late Fee", "Payment Method", "Status"]
ws.set_row(4, 32)
for i, h in enumerate(inc_headers):
    ws.write(4, 3 + i, h, th_em)

props_short = [p[1] for p in properties]
methods = ["Bank Transfer", "Check", "Credit Card", "Cash", "ACH"]
income_data = []
inc_id = 1
for m in range(1, 13):
    for i, p in enumerate(props_short):
        if random.random() > 0.12:
            paid_date = date(2024, m, random.randint(1, 28))
            late = round(random.uniform(0, 75), 2) if random.random() < 0.1 else 0
            status = "Paid" if late == 0 else "Late"
            income_data.append((
                f"I{inc_id:04d}", p, tenants[i % len(tenants)][1],
                date(2024, m, 1), paid_date,
                round(random.uniform(1500, 2800), 2), late,
                random.choice(methods), status
            ))
            inc_id += 1
        else:
            income_data.append((
                f"I{inc_id:04d}", p, tenants[i % len(tenants)][1],
                date(2024, m, 1), None,
                round(random.uniform(1500, 2800), 2), 0,
                "—", "Outstanding"
            ))
            inc_id += 1

for i, inc in enumerate(income_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r if i % 2 == 0 else cell_r_alt
    bg_d = cell_date(False if i % 2 == 0 else True)
    ws.write(r, 3, inc[0], bg_c)
    ws.write(r, 4, inc[1], bg)
    ws.write(r, 5, inc[2], bg)
    ws.write(r, 6, inc[3], bg_d)
    if inc[4]:
        ws.write(r, 7, inc[4], bg_d)
    else:
        ws.write(r, 7, "—", bg_c)
    ws.write(r, 8, inc[5], bg_r)
    ws.write(r, 9, inc[6], bg_r)
    ws.write(r, 10, inc[7], bg_c)
    if inc[8] == "Paid":
        ws.write(r, 11, inc[8], status_g)
    elif inc[8] == "Late":
        ws.write(r, 11, inc[8], status_y)
    else:
        ws.write(r, 11, inc[8], status_r)

last_row = 5 + len(income_data) - 1
ws.data_validation(5, 10, last_row, 10, {"validate": "list", "source": ["Bank Transfer", "Check", "Credit Card", "Cash", "ACH", "PayPal", "Venmo"]})
ws.data_validation(5, 11, last_row, 11, {"validate": "list", "source": ["Paid", "Late", "Outstanding", "Partial"]})
ws.autofilter(4, 3, last_row, 11)
ws.freeze_panes(5, 0)

ws.conditional_format(5, 8, last_row, 8, {"type": "3_color_scale", "min_color": "#FEE2E2", "mid_color": "#FEF3C7", "max_color": "#D1FAE5"})
ws.conditional_format(5, 11, last_row, 11, {"type": "text", "criteria": "containing", "value": "Paid", "format": status_g})
ws.conditional_format(5, 11, last_row, 11, {"type": "text", "criteria": "containing", "value": "Outstanding", "format": status_r})
ws.conditional_format(5, 11, last_row, 11, {"type": "text", "criteria": "containing", "value": "Late", "format": status_y})

sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 7, "  TOTAL INCOME", fmt({"bold": True, "font_size": 12, "bg_color": EMERALD, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row, 8, "=SUM(I6:I{})".format(last_row+1), big_total_em)
ws.write_formula(sum_row, 9, "=SUM(J6:J{})".format(last_row+1), big_total_em)

chart_inc = wb.add_chart({"type": "line"})
chart_inc.add_series({
    "name":       "Monthly Income",
    "categories": ["PnL", 2, 0, 13, 0],
    "values":     ["PnL", 2, 1, 13, 1],
    "line":       {"color": EMERALD, "width": 2.5},
    "marker":     {"type": "circle", "size": 5, "fill": {"color": EMERALD}},
})
chart_inc.set_title({"name": "Monthly Income Trend", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_inc.set_y_axis({"num_format": "$#,##0"})
chart_inc.set_size({"width": 720, "height": 320})
chart_inc.set_legend({"none": True})
ws.insert_chart(sum_row + 3, 3, chart_inc)

# =========================================================================
# 5. EXPENSE TRACKER
# =========================================================================
ws = wb.add_worksheet("Expenses")
ws.set_tab_color(ORANGE)
setup_page(ws)
add_sidebar(ws, active_idx=4)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 18)
ws.set_column("G:G", 14)
ws.set_column("H:H", 14)
ws.set_column("I:I", 28)
ws.set_column("J:J", 13)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:J2", "  📉 Expense Tracker", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:J3", "  Categorize and monitor all property-related expenses", subtitle)
ws.set_row(3, 8)

exp_headers = ["Expense ID", "Property", "Category", "Date", "Amount", "Description", "Status"]
ws.set_row(4, 32)
for i, h in enumerate(exp_headers):
    ws.write(4, 3 + i, h, th_or)

# 11 categories
categories = ["Repairs", "Maintenance", "Insurance", "Property Tax", "Utilities", "Mortgage", "Cleaning", "Landscaping", "Legal", "Advertising", "Miscellaneous"]
expense_data = []
exp_id = 1
for m in range(1, 13):
    for i, p in enumerate(props_short):
        for _ in range(random.randint(1, 2)):
            cat = random.choice(categories)
            amt = round(random.uniform(50, 1500), 2) if cat not in ["Property Tax", "Mortgage", "Insurance"] else round(random.uniform(500, 3000), 2)
            expense_data.append((
                f"E{exp_id:04d}", p, cat, date(2024, m, random.randint(1, 28)),
                amt, f"{cat} for {p}", random.choice(["Paid", "Pending", "Reimbursed"])
            ))
            exp_id += 1

for i, e in enumerate(expense_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r if i % 2 == 0 else cell_r_alt
    bg_d = cell_date(False if i % 2 == 0 else True)
    ws.write(r, 3, e[0], bg_c)
    ws.write(r, 4, e[1], bg)
    ws.write(r, 5, e[2], bg_c)
    ws.write(r, 6, e[3], bg_d)
    ws.write(r, 7, e[4], bg_r)
    ws.write(r, 8, e[5], bg)
    if e[6] == "Paid":
        ws.write(r, 9, e[6], status_g)
    elif e[6] == "Pending":
        ws.write(r, 9, e[6], status_y)
    else:
        ws.write(r, 9, e[6], status_b)

last_row = 5 + len(expense_data) - 1
ws.data_validation(5, 5, last_row, 5, {"validate": "list", "source": categories})
ws.data_validation(5, 9, last_row, 9, {"validate": "list", "source": ["Paid", "Pending", "Reimbursed", "Disputed"]})
ws.autofilter(4, 3, last_row, 9)
ws.freeze_panes(5, 0)

ws.conditional_format(5, 7, last_row, 7, {"type": "3_color_scale", "min_color": "#D1FAE5", "mid_color": "#FEF3C7", "max_color": "#FEE2E2"})

# Category subtotal block (row 5-15) used by dashboard chart
# Build a small summary table for category totals at top
sum_row = 5
ws.set_row(4, 32)

# We'll put a category summary at column L starting at row 4
ws.merge_range(3, 11, 3, 13, "  EXPENSE BY CATEGORY", section_h_or)
ws.write(4, 11, "Category", th_or)
ws.write(4, 12, "Count", th_or)
ws.write(4, 13, "Total", th_or)
for i, cat in enumerate(categories):
    r = 5 + i
    alt = (i % 2 == 0)
    ws.write(r, 11, cat, cell if alt else cell_alt)
    ws.write_formula(r, 12, f'=COUNTIF(F6:F{last_row+1},"{cat}")', cell_int if alt else cell_int_alt)
    ws.write_formula(r, 13, f'=SUMIF(F6:F{last_row+1},"{cat}",H6:H{last_row+1})', cell_r if alt else cell_r_alt)

cat_end_row = 4 + len(categories)
sum_row = max(last_row, 5 + len(categories)) + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 6, "  TOTAL EXPENSES", fmt({"bold": True, "font_size": 12, "bg_color": ORANGE, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row, 7, "=SUM(H6:H{})".format(last_row+1), big_total_or)

ws.merge_range(sum_row + 2, 3, sum_row + 4, 9,
    "💡 CATEGORIES:\n• Repairs, Maintenance, Insurance, Property Tax, Utilities, Mortgage, Cleaning, Landscaping, Legal, Advertising, Miscellaneous\n• All dashboards and the monthly P&L will update automatically.", note)

# =========================================================================
# 6. MORTGAGE TRACKER
# =========================================================================
ws = wb.add_worksheet("Mortgage")
ws.set_tab_color(NAVY)
setup_page(ws)
add_sidebar(ws, active_idx=5)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 13)
ws.set_column("G:G", 13)
ws.set_column("H:H", 14)
ws.set_column("I:I", 15)
ws.set_column("J:J", 14)
ws.set_column("K:K", 18)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:K2", "  🏦 Mortgage Tracker", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:K3", "  Monitor loan principal, interest, and amortization progress", subtitle)
ws.set_row(3, 8)

mort_headers = ["Loan ID", "Property", "Principal", "Interest Rate", "Monthly Pmt", "Term (Years)", "Balance", "Progress"]
ws.set_row(4, 32)
for i, h in enumerate(mort_headers):
    ws.write(4, 3 + i, h, th)

# Build 10 sample mortgage rows
mortgage_data = []
for i, p in enumerate(properties):
    principal = p[9]
    rate = round(random.uniform(3.5, 6.5), 3)
    term = random.choice([15, 20, 30])
    monthly = round(principal * (rate/100/12) / (1 - (1 + rate/100/12) ** (-term*12)), 2)
    months_paid = random.randint(12, term*12 - 12)
    balance = round(principal * ((1 + rate/100/12) ** (term*12) - (1 + rate/100/12) ** months_paid) / ((1 + rate/100/12) ** (term*12) - 1), 2)
    progress = months_paid / (term * 12)
    mortgage_data.append((f"L{i+1:03d}", p[1], principal, rate, monthly, term, balance, progress))

for i, m in enumerate(mortgage_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r0 if i % 2 == 0 else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_pct = cell_pct if i % 2 == 0 else cell_pct_alt
    ws.write(r, 3, m[0], bg_c)
    ws.write(r, 4, m[1], bg)
    ws.write(r, 5, m[2], bg_r)
    ws.write(r, 6, m[3] / 100, fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE if i%2==0 else LGRAY, "num_format": "0.000%"}))
    ws.write(r, 7, m[4], fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE if i%2==0 else LGRAY, "num_format": "$#,##0.00"}))
    ws.write(r, 8, m[5], bg_c)
    ws.write(r, 9, m[6], bg_r)
    # Progress as data bar
    ws.write(r, 10, m[7], fmt({"font_size": 10, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE if i%2==0 else LGRAY, "num_format": "0.0%"}))

last_row = 5 + len(mortgage_data) - 1
ws.autofilter(4, 3, last_row, 10)
ws.freeze_panes(5, 0)
# Progress data bar
ws.conditional_format(5, 10, last_row, 10, {"type": "data_bar", "bar_color": EMERALD, "bar_solid": True})

# Total row
sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 4, "  TOTALS", fmt({"bold": True, "font_size": 12, "bg_color": NAVY, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row, 5, "=SUM(F6:F{})".format(last_row+1), big_total)
ws.write_formula(sum_row, 7, "=SUM(H6:H{})".format(last_row+1), big_total)
ws.write_formula(sum_row, 9, "=SUM(J6:J{})".format(last_row+1), big_total)

# Amortization sample for one property
amort_start = sum_row + 3
ws.merge_range(amort_start, 3, amort_start, 10, "  📉  AMORTIZATION SCHEDULE — Highland Park Home (P010)", section_h)
amort_headers = ["Payment #", "Date", "Payment", "Principal", "Interest", "Balance", "Cumulative"]
ws.set_row(amort_start+1, 28)
for i, h in enumerate(amort_headers):
    ws.write(amort_start+1, 3 + i, h, th)

# 12 month sample amortization
balance = mortgage_data[-1][6]
rate = mortgage_data[-1][3] / 100 / 12
monthly = mortgage_data[-1][4]
cumulative = 0
for i in range(12):
    r = amort_start + 2 + i
    interest = round(balance * rate, 2)
    princ = round(monthly - interest, 2)
    balance = round(balance - princ, 2)
    cumulative += monthly
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_c = cell_c if alt else cell_c_alt
    bg_r = cell_r if alt else cell_r_alt
    bg_d = cell_date(not alt)
    ws.write(r, 3, i + 1, bg_c)
    ws.write(r, 4, date(2024, i + 1, 1), bg_d)
    ws.write(r, 5, monthly, bg_r)
    ws.write(r, 6, princ, bg_r)
    ws.write(r, 7, interest, bg_r)
    ws.write(r, 8, balance, bg_r)
    ws.write(r, 9, round(cumulative, 2), bg_r)

# =========================================================================
# 7. MAINTENANCE LOG
# =========================================================================
ws = wb.add_worksheet("Maintenance")
ws.set_tab_color(ORANGE)
setup_page(ws)
add_sidebar(ws, active_idx=6)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 18)
ws.set_column("G:G", 14)
ws.set_column("H:H", 13)
ws.set_column("I:I", 12)
ws.set_column("J:J", 11)
ws.set_column("K:K", 12)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:K2", "  🔧 Maintenance Log", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:K3", "  Track repairs, contractor work, warranties, and priorities", subtitle)
ws.set_row(3, 8)

maint_headers = ["Ticket ID", "Property", "Contractor", "Date", "Cost", "Status", "Priority", "Warranty"]
ws.set_row(4, 32)
for i, h in enumerate(maint_headers):
    ws.write(4, 3 + i, h, th_or)

priorities = ["Low", "Medium", "High", "Urgent"]
maint_statuses = ["Open", "In Progress", "Completed", "On Hold"]
maint_data = []
m_id = 1
for i, p in enumerate(props_short):
    for _ in range(random.randint(2, 3)):
        contractor = contractors[random.randint(0, len(contractors)-1)][1]
        cost = round(random.uniform(75, 3500), 2)
        status = random.choice(maint_statuses)
        priority = random.choice(priorities)
        warranty = random.choice(["None", "30 days", "90 days", "1 year", "5 years", "Lifetime"])
        maint_data.append((f"M{m_id:04d}", p, contractor, date(2024, random.randint(1, 12), random.randint(1, 28)),
                          cost, status, priority, warranty))
        m_id += 1

for i, m in enumerate(maint_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r if i % 2 == 0 else cell_r_alt
    bg_d = cell_date(False if i % 2 == 0 else True)
    ws.write(r, 3, m[0], bg_c)
    ws.write(r, 4, m[1], bg)
    ws.write(r, 5, m[2], bg)
    ws.write(r, 6, m[3], bg_d)
    ws.write(r, 7, m[4], bg_r)
    # Status color
    if m[5] == "Completed":
        ws.write(r, 8, m[5], status_g)
    elif m[5] == "In Progress":
        ws.write(r, 8, m[5], status_b)
    elif m[5] == "On Hold":
        ws.write(r, 8, m[5], status_y)
    else:
        ws.write(r, 8, m[5], status_r)
    # Priority color
    if m[6] == "Urgent":
        ws.write(r, 9, m[6], status_r)
    elif m[6] == "High":
        ws.write(r, 9, m[6], status_y)
    elif m[6] == "Medium":
        ws.write(r, 9, m[6], status_b)
    else:
        ws.write(r, 9, m[6], status_g)
    ws.write(r, 10, m[7], bg_c)

last_row = 5 + len(maint_data) - 1
ws.data_validation(5, 5, last_row, 5, {"validate": "list", "source": [c[1] for c in contractors]})
ws.data_validation(5, 8, last_row, 8, {"validate": "list", "source": maint_statuses})
ws.data_validation(5, 9, last_row, 9, {"validate": "list", "source": priorities})
ws.data_validation(5, 10, last_row, 10, {"validate": "list", "source": ["None", "30 days", "90 days", "1 year", "5 years", "Lifetime"]})
ws.autofilter(4, 3, last_row, 10)
ws.freeze_panes(5, 0)

ws.conditional_format(5, 7, last_row, 7, {"type": "3_color_scale", "min_color": "#D1FAE5", "mid_color": "#FEF3C7", "max_color": "#FEE2E2"})

sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 6, "  TOTAL MAINTENANCE", fmt({"bold": True, "font_size": 12, "bg_color": ORANGE, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row, 7, "=SUM(H6:H{})".format(last_row+1), big_total_or)

ws.merge_range(sum_row + 2, 3, sum_row + 5, 10,
    "💡 TIPS:\n• Use the Priority column to triage urgent issues (red).\n• Track warranties to know what's covered.\n• Click column headers to filter by status or contractor.", note)

# =========================================================================
# 8. LEASE TRACKER
# =========================================================================
ws = wb.add_worksheet("Lease")
ws.set_tab_color(EMERALD)
setup_page(ws)
add_sidebar(ws, active_idx=7)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 22)
ws.set_column("G:I", 13)
ws.set_column("J:J", 14)
ws.set_column("K:K", 14)
ws.set_column("L:M", 12)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:M2", "  📅 Lease Tracker", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:M3", "  🟢 Active  •  🟡 Expiring within 60 days  •  🔴 Expired", subtitle)
ws.set_row(3, 8)

lease_headers = ["Lease ID", "Property", "Tenant", "Start Date", "End Date", "Days Left", "Monthly Rent", "Security", "Renewal", "Status"]
ws.set_row(4, 32)
for i, h in enumerate(lease_headers):
    ws.write(4, 3 + i, h, th_em)

lease_data = []
for i, t in enumerate(tenants):
    prop = props_short[i % len(props_short)]
    start = datetime.strptime(t[5], "%Y-%m-%d").date()
    end = datetime.strptime(t[6], "%Y-%m-%d").date()
    days_left = (end - date.today()).days
    lease_data.append((f"LE{i+1:03d}", prop, t[1], start, end, days_left, t[7], t[8], random.choice(["Auto", "Manual", "Pending"]), t[9]))

for i, l in enumerate(lease_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r0 if i % 2 == 0 else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_d = cell_date(False if i % 2 == 0 else True)
    bg_int = cell_int if i % 2 == 0 else cell_int_alt
    ws.write(r, 3, l[0], bg_c)
    ws.write(r, 4, l[1], bg)
    ws.write(r, 5, l[2], bg)
    ws.write(r, 6, l[3], bg_d)
    ws.write(r, 7, l[4], bg_d)
    # Days left - color based on value
    if l[5] < 0:
        ws.write(r, 8, l[5], fmt({"bold": True, "font_size": 10, "font_color": "#991B1B", "bg_color": LRED, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER}))
    elif l[5] <= 60:
        ws.write(r, 8, l[5], fmt({"bold": True, "font_size": 10, "font_color": "#92400E", "bg_color": LYELLOW, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER}))
    else:
        ws.write(r, 8, l[5], fmt({"bold": True, "font_size": 10, "font_color": "#065F46", "bg_color": LGREEN, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER}))
    ws.write(r, 9, l[6], bg_r)
    ws.write(r, 10, l[7], bg_r)
    ws.write(r, 11, l[8], bg_c)
    # Status
    if l[9] == "Active":
        ws.write(r, 12, l[9], status_g)
    elif l[9] == "Expiring":
        ws.write(r, 12, l[9], status_y)
    else:
        ws.write(r, 12, l[9], status_r)

last_row = 5 + len(lease_data) - 1
ws.data_validation(5, 11, last_row, 11, {"validate": "list", "source": ["Auto", "Manual", "Pending", "Month-to-Month"]})
ws.data_validation(5, 12, last_row, 12, {"validate": "list", "source": ["Active", "Expiring", "Expired", "Renewed"]})
ws.autofilter(4, 3, last_row, 12)
ws.freeze_panes(5, 0)

# Conditional formatting for lease end date
ws.conditional_format(5, 7, last_row, 7, {"type": "date", "criteria": "less than or equal to", "value": date.today(), "format": status_r})
ws.conditional_format(5, 7, last_row, 7, {"type": "date", "criteria": "between", "minimum": date.today(), "maximum": date.today() + timedelta(days=60), "format": status_y})

# Status summary
sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 4, "  LEASE SUMMARY", section_h_em)
ws.set_row(sum_row+1, 22)
ws.write(sum_row+1, 3, "🟢 Active Leases", fmt({"bold": True, "font_size": 11, "bg_color": LGREEN, "font_color": "#065F46", "align": "left", "border": 1, "border_color": BORDER, "left": 2}))
ws.write_formula(sum_row+1, 4, '=COUNTIF(M6:M{},"Active")'.format(last_row+1), fmt({"bold": True, "font_size": 11, "bg_color": WHITE, "align": "center", "border": 1, "border_color": BORDER}))
ws.write(sum_row+1, 5, "🟡 Expiring Soon", fmt({"bold": True, "font_size": 11, "bg_color": LYELLOW, "font_color": "#92400E", "align": "left", "border": 1, "border_color": BORDER, "left": 2}))
ws.write_formula(sum_row+1, 6, '=COUNTIF(M6:M{},"Expiring")'.format(last_row+1), fmt({"bold": True, "font_size": 11, "bg_color": WHITE, "align": "center", "border": 1, "border_color": BORDER}))
ws.write(sum_row+1, 7, "🔴 Expired", fmt({"bold": True, "font_size": 11, "bg_color": LRED, "font_color": "#991B1B", "align": "left", "border": 1, "border_color": BORDER, "left": 2}))
ws.write_formula(sum_row+1, 8, '=COUNTIF(M6:M{},"Expired")'.format(last_row+1), fmt({"bold": True, "font_size": 11, "bg_color": WHITE, "align": "center", "border": 1, "border_color": BORDER}))

# Total rent
ws.set_row(sum_row+2, 24)
ws.merge_range(sum_row+2, 3, sum_row+2, 8, "  TOTAL MONTHLY RENT", fmt({"bold": True, "font_size": 12, "bg_color": EMERALD, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row+2, 9, "=SUM(J6:J{})".format(last_row+1), big_total_em)

# =========================================================================
# 9. VACANCY TRACKER
# =========================================================================
ws = wb.add_worksheet("Vacancy")
ws.set_tab_color(RED)
setup_page(ws)
add_sidebar(ws, active_idx=8)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 16)
ws.set_column("G:G", 16)
ws.set_column("H:H", 16)
ws.set_column("I:I", 16)
ws.set_column("J:J", 14)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:J2", "  🚪 Vacancy Tracker", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:J3", "  Monitor vacant units, lost income, and occupancy trends", subtitle)
ws.set_row(3, 8)

vac_headers = ["Property ID", "Property", "Vacant Days", "Daily Rate", "Lost Income", "Listed Date", "Showing Count", "Status"]
ws.set_row(4, 32)
for i, h in enumerate(vac_headers):
    ws.write(4, 3 + i, h, th_red)

vac_data = []
for i, p in enumerate(properties):
    if p[13] == "Vacant":
        vacant_days = random.randint(15, 90)
        daily_rate = round(p[8] * 0.0035, 2)
        lost_income = round(daily_rate * vacant_days, 2)
        listed = date(2024, random.randint(1, 11), random.randint(1, 28))
        showings = random.randint(0, 15)
        status = "Listed"
    elif p[13] == "Renovation":
        vacant_days = random.randint(30, 120)
        daily_rate = round(p[8] * 0.0035, 2)
        lost_income = round(daily_rate * vacant_days, 2)
        listed = date(2024, random.randint(1, 11), random.randint(1, 28))
        showings = 0
        status = "Renovation"
    else:
        vacant_days = random.randint(0, 14)
        daily_rate = round(p[8] * 0.0035, 2)
        lost_income = round(daily_rate * vacant_days, 2)
        listed = date(2024, 1, 1)
        showings = 0
        status = "Occupied"
    vac_data.append((p[0], p[1], vacant_days, daily_rate, lost_income, listed, showings, status))

for i, v in enumerate(vac_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r if i % 2 == 0 else cell_r_alt
    bg_d = cell_date(False if i % 2 == 0 else True)
    ws.write(r, 3, v[0], bg_c)
    ws.write(r, 4, v[1], bg)
    ws.write(r, 5, v[2], bg_c)
    ws.write(r, 6, v[3], bg_r)
    ws.write(r, 7, v[4], bg_r)
    ws.write(r, 8, v[5], bg_d)
    ws.write(r, 9, v[6], bg_c)
    if v[7] == "Listed":
        ws.write(r, 10, v[7], status_y)
    elif v[7] == "Renovation":
        ws.write(r, 10, v[7], status_b)
    else:
        ws.write(r, 10, v[7], status_g)

last_row = 5 + len(vac_data) - 1
ws.data_validation(5, 10, last_row, 10, {"validate": "list", "source": ["Occupied", "Listed", "Renovation", "Off-Market"]})
ws.autofilter(4, 3, last_row, 10)
ws.freeze_panes(5, 0)

ws.conditional_format(5, 7, last_row, 7, {"type": "3_color_scale", "min_color": "#D1FAE5", "mid_color": "#FEF3C7", "max_color": "#FEE2E2"})

# Vacancy summary
sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 4, "  VACANCY SUMMARY", section_h)
ws.set_row(sum_row+1, 22)
ws.write(sum_row+1, 3, "Total Vacant Days", fmt({"bold": True, "font_size": 11, "bg_color": LGRAY, "align": "left", "border": 1, "border_color": BORDER, "left": 2}))
ws.write_formula(sum_row+1, 4, "=SUM(F6:F{})".format(last_row+1), fmt({"bold": True, "font_size": 11, "bg_color": WHITE, "align": "center", "border": 1, "border_color": BORDER, "num_format": "#,##0"}))
ws.write(sum_row+1, 5, "Total Lost Income", fmt({"bold": True, "font_size": 11, "bg_color": LGRAY, "align": "left", "border": 1, "border_color": BORDER, "left": 2}))
ws.write_formula(sum_row+1, 6, "=SUM(H6:H{})".format(last_row+1), fmt({"bold": True, "font_size": 11, "bg_color": WHITE, "align": "center", "border": 1, "border_color": BORDER, "num_format": "$#,##0.00"}))

# Vacancy % calculation
ws.set_row(sum_row+2, 24)
ws.merge_range(sum_row+2, 3, sum_row+2, 6, "  VACANCY RATE", fmt({"bold": True, "font_size": 12, "bg_color": RED, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row+2, 7, '=SUM(F6:F{})/(COUNTA(D6:D{})*365)'.format(last_row+1, last_row+1), fmt({"bold": True, "font_size": 12, "bg_color": WHITE, "align": "center", "valign": "vcenter", "num_format": "0.00%"}))

# =========================================================================
# 10. ROI CALCULATOR
# =========================================================================
ws = wb.add_worksheet("ROI")
ws.set_tab_color(PURPLE)
setup_page(ws)
add_sidebar(ws, active_idx=9)

ws.set_column("C:C", 2)
ws.set_column("D:D", 22)
ws.set_column("E:E", 14)
ws.set_column("F:F", 14)
ws.set_column("G:G", 14)
ws.set_column("H:H", 14)
ws.set_column("I:I", 14)
ws.set_column("J:J", 14)
ws.set_column("K:K", 14)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:K2", "  📈 ROI Calculator", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:K3", "  Per-property ROI, Cap Rate, Cash-on-Cash Return, and Appreciation", subtitle)
ws.set_row(3, 8)

roi_headers = ["Property", "Purchase", "Current Value", "Annual Income", "Annual Expenses", "Cash Flow", "ROI", "Cap Rate", "COC Return"]
ws.set_row(4, 32)
for i, h in enumerate(roi_headers):
    ws.write(4, 3 + i, h, th_purple)

# Per property data with formulas
for i, p in enumerate(properties):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r0 if i % 2 == 0 else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_pct = cell_pct if i % 2 == 0 else cell_pct_alt
    ws.write(r, 3, p[1], bg)
    ws.write(r, 4, p[7], bg_r)  # purchase
    ws.write(r, 5, p[8], bg_r)  # current value
    # Annual income formula - from income sheet for this property (use COUNTIF/INDEX or just hardcode sum)
    ws.write_formula(r, 6, '=SUMIF(Income!E:E,"' + p[1] + '",Income!I:I)', bg_r)
    # Annual expenses
    ws.write_formula(r, 7, '=SUMIF(Expenses!E:E,"' + p[1] + '",Expenses!H:H)', bg_r)
    # Cash flow
    ws.write_formula(r, 8, '=G' + str(r+1) + '-H' + str(r+1), bg_r)
    # ROI (cash flow / purchase)
    ws.write_formula(r, 9, '=I' + str(r+1) + '/E' + str(r+1), bg_pct)
    # Cap rate (NOI / value)
    ws.write_formula(r, 10, '=I' + str(r+1) + '/F' + str(r+1), bg_pct)
    # COC return (annual income - expenses) / down payment
    ws.write_formula(r, 11, '=I' + str(r+1) + '/(E' + str(r+1) + '*0.2)', bg_pct)

last_row = 5 + len(properties) - 1
ws.autofilter(4, 3, last_row, 11)
ws.freeze_panes(5, 0)

# Data bars on ROI
ws.conditional_format(5, 9, last_row, 9, {"type": "data_bar", "bar_color": EMERALD, "bar_solid": True})
ws.conditional_format(5, 10, last_row, 10, {"type": "data_bar", "bar_color": NAVY, "bar_solid": True})
ws.conditional_format(5, 11, last_row, 11, {"type": "data_bar", "bar_color": ORANGE, "bar_solid": True})

# Total row
sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.write(sum_row, 3, "  PORTFOLIO TOTAL", fmt({"bold": True, "font_size": 12, "bg_color": PURPLE, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row, 4, "=SUM(E6:E{})".format(last_row+1), big_total)
ws.write_formula(sum_row, 5, "=SUM(F6:F{})".format(last_row+1), big_total)
ws.write_formula(sum_row, 6, "=SUM(G6:G{})".format(last_row+1), big_total)
ws.write_formula(sum_row, 7, "=SUM(H6:H{})".format(last_row+1), big_total)
ws.write_formula(sum_row, 8, "=SUM(I6:I{})".format(last_row+1), big_total)
ws.write_formula(sum_row, 9, "=AVERAGE(J6:J{})".format(last_row+1), fmt({"bold": True, "font_size": 12, "bg_color": PURPLE, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "0.00%"}))
ws.write_formula(sum_row, 10, "=AVERAGE(K6:K{})".format(last_row+1), fmt({"bold": True, "font_size": 12, "bg_color": PURPLE, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "0.00%"}))
ws.write_formula(sum_row, 11, "=AVERAGE(L6:L{})".format(last_row+1), fmt({"bold": True, "font_size": 12, "bg_color": PURPLE, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "0.00%"}))

ws.merge_range(sum_row + 2, 3, sum_row + 5, 11,
    "💡 ROI FORMULAS:\n• ROI = Annual Cash Flow / Purchase Price\n• Cap Rate = Net Operating Income / Current Market Value\n• Cash-on-Cash Return = Annual Cash Flow / Down Payment (assumed 20%)\n• All values auto-populate from the Income, Expenses, and Properties sheets.", note)

# =========================================================================
# 11. PROPERTY VALUATION
# =========================================================================
ws = wb.add_worksheet("Valuation")
ws.set_tab_color(GOLD)
setup_page(ws)
add_sidebar(ws, active_idx=10)

ws.set_column("C:C", 2)
ws.set_column("D:D", 22)
ws.set_column("E:G", 16)
ws.set_column("H:H", 16)
ws.set_column("I:I", 18)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:I2", "  🏷️  Property Valuation", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:I3", "  Track property appreciation and gain/loss over time", subtitle)
ws.set_row(3, 8)

val_headers = ["Property", "Purchase Price", "Current Value", "Gain/Loss", "Gain %", "Annual Appreciation", "Years Held"]
ws.set_row(4, 32)
for i, h in enumerate(val_headers):
    ws.write(4, 3 + i, h, th)

for i, p in enumerate(properties):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r0 if i % 2 == 0 else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_pct = cell_pct if i % 2 == 0 else cell_pct_alt
    purchase = p[7]
    current = p[8]
    gain = current - purchase
    gain_pct = gain / purchase
    years = (date.today() - datetime.strptime(p[6], "%Y-%m-%d").date()).days / 365.25
    annual_appr = (gain_pct / years) if years > 0 else 0
    ws.write(r, 3, p[1], bg)
    ws.write(r, 4, purchase, bg_r)
    ws.write(r, 5, current, bg_r)
    # Gain/Loss color
    if gain > 0:
        ws.write(r, 6, gain, fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE if i%2==0 else LGRAY, "num_format": "$#,##0", "font_color": EMERALD, "bold": True}))
    else:
        ws.write(r, 6, gain, fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE if i%2==0 else LGRAY, "num_format": "$#,##0", "font_color": RED, "bold": True}))
    ws.write(r, 7, gain_pct, bg_pct)
    ws.write(r, 8, annual_appr, bg_pct)
    ws.write(r, 9, round(years, 1), bg_c)

last_row = 5 + len(properties) - 1
ws.autofilter(4, 3, last_row, 9)
ws.freeze_panes(5, 0)
ws.conditional_format(5, 7, last_row, 7, {"type": "data_bar", "bar_color": EMERALD, "bar_solid": True})
ws.conditional_format(5, 8, last_row, 8, {"type": "data_bar", "bar_color": GOLD, "bar_solid": True})

# Totals
sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.write(sum_row, 3, "  TOTALS", fmt({"bold": True, "font_size": 12, "bg_color": GOLD, "font_color": NAVY, "align": "right"}))
ws.write_formula(sum_row, 4, "=SUM(E6:E{})".format(last_row+1), fmt({"bold": True, "font_size": 12, "bg_color": GOLD, "font_color": NAVY, "align": "right", "valign": "vcenter", "num_format": "$#,##0"}))
ws.write_formula(sum_row, 5, "=SUM(F6:F{})".format(last_row+1), fmt({"bold": True, "font_size": 12, "bg_color": GOLD, "font_color": NAVY, "align": "right", "valign": "vcenter", "num_format": "$#,##0"}))
ws.write_formula(sum_row, 6, "=SUM(G6:G{})".format(last_row+1), fmt({"bold": True, "font_size": 12, "bg_color": GOLD, "font_color": NAVY, "align": "right", "valign": "vcenter", "num_format": "$#,##0"}))
ws.write_formula(sum_row, 7, "=AVERAGE(H6:H{})".format(last_row+1), fmt({"bold": True, "font_size": 12, "bg_color": GOLD, "font_color": NAVY, "align": "right", "valign": "vcenter", "num_format": "0.00%"}))

print("Sheets 1-11 created")

# =========================================================================
# 12. TAX SUMMARY
# =========================================================================
ws = wb.add_worksheet("Tax")
ws.set_tab_color(DGRAY)
setup_page(ws)
add_sidebar(ws, active_idx=11)

ws.set_column("C:C", 2)
ws.set_column("D:D", 32)
ws.set_column("E:E", 18)
ws.set_column("F:F", 18)
ws.set_column("G:G", 18)
ws.set_column("H:H", 18)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:H2", "  🧾 Tax Summary", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:H3", "  Annual tax preparation summary — income, expenses, depreciation, and deductions", subtitle)
ws.set_row(3, 8)

# Annual totals
ws.set_row(4, 32)
ws.merge_range("D5:H5", "  ANNUAL TAX SUMMARY 2024", section_h)

tax_data = [
    ("Gross Rental Income",       '=SUMIF(Income!L:L,"Paid",Income!I:I)+SUMIF(Income!L:L,"Late",Income!I:I)',  "$#,##0.00"),
    ("Late Fees Collected",       "=SUM(Income!J:J)",                                                       "$#,##0.00"),
    ("Total Income",              '=E6+E7',                                                                 "$#,##0.00"),
    ("", None, None),
    ("Property Tax",              "=SUMIF(Expenses!F:F,\"Property Tax\",Expenses!H:H)",                       "$#,##0.00"),
    ("Insurance",                 "=SUMIF(Expenses!F:F,\"Insurance\",Expenses!H:H)",                          "$#,##0.00"),
    ("Mortgage Interest",         "=SUMIF(Expenses!F:F,\"Mortgage\",Expenses!H:H)*0.7",                       "$#,##0.00"),
    ("Repairs & Maintenance",     "=SUMIF(Expenses!F:F,\"Repairs\",Expenses!H:H)+SUMIF(Expenses!F:F,\"Maintenance\",Expenses!H:H)", "$#,##0.00"),
    ("Utilities",                 "=SUMIF(Expenses!F:F,\"Utilities\",Expenses!H:H)",                          "$#,##0.00"),
    ("Cleaning & Landscaping",    "=SUMIF(Expenses!F:F,\"Cleaning\",Expenses!H:H)+SUMIF(Expenses!F:F,\"Landscaping\",Expenses!H:H)", "$#,##0.00"),
    ("Legal & Advertising",       "=SUMIF(Expenses!F:F,\"Legal\",Expenses!H:H)+SUMIF(Expenses!F:F,\"Advertising\",Expenses!H:H)", "$#,##0.00"),
    ("Miscellaneous Expenses",    "=SUMIF(Expenses!F:F,\"Miscellaneous\",Expenses!H:H)",                       "$#,##0.00"),
    ("Total Operating Expenses",  "=SUM(E11:E17)",                                                            "$#,##0.00"),
    ("", None, None),
    ("Depreciation (Manual)",     45000,                                                                      "$#,##0.00"),
    ("Net Taxable Income",        "=E8-E18-E20",                                                              "$#,##0.00"),
]

for i, (label, val, nf) in enumerate(tax_data):
    r = 5 + i
    if label == "":
        continue
    if i == 2:  # Total Income
        ws.write(r, 3, "  " + label, fmt({"bold": True, "font_size": 12, "bg_color": EMERALD, "font_color": WHITE, "align": "left", "valign": "vcenter", "left": 2}))
        ws.write_formula(r, 4, val, fmt({"bold": True, "font_size": 12, "bg_color": EMERALD, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": nf}))
    elif i == 12:  # Total Operating Expenses
        ws.write(r, 3, "  " + label, fmt({"bold": True, "font_size": 12, "bg_color": ORANGE, "font_color": WHITE, "align": "left", "valign": "vcenter", "left": 2}))
        ws.write_formula(r, 4, val, fmt({"bold": True, "font_size": 12, "bg_color": ORANGE, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": nf}))
    elif i == 15:  # Net Taxable Income
        ws.set_row(r, 28)
        ws.write(r, 3, "  " + label, fmt({"bold": True, "font_size": 13, "bg_color": NAVY, "font_color": WHITE, "align": "left", "valign": "vcenter", "left": 2}))
        ws.write_formula(r, 4, val, fmt({"bold": True, "font_size": 13, "bg_color": NAVY, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": nf}))
    else:
        alt = (i % 2 == 0)
        bg = cell if alt else cell_alt
        bg_r = cell_r if alt else cell_r_alt
        ws.write(r, 3, "  " + label, bg)
        if isinstance(val, str) and val.startswith("="):
            ws.write_formula(r, 4, val, bg_r)
        else:
            ws.write(r, 4, val, bg_r)

# Property depreciation schedule
depr_start = 5 + len(tax_data) + 2
ws.merge_range(depr_start, 3, depr_start, 7, "  📉  DEPRECIATION SCHEDULE (27.5 Years — Residential Rental)", section_h)
ws.set_row(depr_start+1, 28)
depr_headers = ["Property", "Purchase Price", "Land Value", "Depreciable Basis", "Annual Depreciation", "Years"]
for i, h in enumerate(depr_headers):
    ws.write(depr_start+1, 3 + i, h, th_dgray)

for i, p in enumerate(properties):
    r = depr_start + 2 + i
    land = int(p[7] * 0.20)
    basis = p[7] - land
    annual = basis / 27.5
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_c = cell_c if alt else cell_c_alt
    bg_r = cell_r0 if alt else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_r2 = cell_r if alt else cell_r_alt
    ws.write(r, 3, p[1], bg)
    ws.write(r, 4, p[7], bg_r)
    ws.write(r, 5, land, bg_r)
    ws.write(r, 6, basis, bg_r)
    ws.write(r, 7, round(annual, 2), bg_r2)
    ws.write(r, 8, 27.5, bg_c)

# Quarterly breakdown
q_start = depr_start + 2 + len(properties) + 2
ws.merge_range(q_start, 3, q_start, 7, "  📅  QUARTERLY ESTIMATED TAX PAYMENTS", section_h)
ws.set_row(q_start+1, 28)
q_headers = ["Quarter", "Period", "Estimated Income", "Estimated Tax (25%)", "Payment Date"]
for i, h in enumerate(q_headers):
    ws.write(q_start+1, 3 + i, h, th_dgray)

q_data = [
    ("Q1", "Jan - Mar 2024", date(2024, 4, 15)),
    ("Q2", "Apr - Jun 2024", date(2024, 6, 15)),
    ("Q3", "Jul - Sep 2024", date(2024, 9, 15)),
    ("Q4", "Oct - Dec 2024", date(2025, 1, 15)),
]
for i, q in enumerate(q_data):
    r = q_start + 2 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_c = cell_c if alt else cell_c_alt
    bg_r = cell_r0 if alt else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_d = cell_date(not alt)
    ws.write(r, 3, q[0], bg_c)
    ws.write(r, 4, q[1], bg)
    # Estimated income = 1/4 of annual
    ws.write_formula(r, 5, "=E6/4", bg_r)
    ws.write_formula(r, 6, "=F" + str(r+1) + "*0.25", bg_r)
    ws.write(r, 7, q[2], bg_d)

# Notes
ws.merge_range(q_start + 8, 3, q_start + 12, 7,
    "💡 TAX TIPS:\n• Depreciation is calculated using 27.5-year straight-line method for residential rental property.\n• Land value (20% of purchase) is not depreciable.\n• Mortgage interest (70% of mortgage payment) is tax-deductible.\n• Always consult a qualified tax professional for actual filing.\n• All values auto-populate from other sheets — update the manual Depreciation entry as needed.", note)

# =========================================================================
# 13. UTILITY TRACKER
# =========================================================================
ws = wb.add_worksheet("Utility")
ws.set_tab_color(TEAL)
setup_page(ws)
add_sidebar(ws, active_idx=12)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 14)
ws.set_column("G:G", 14)
ws.set_column("H:H", 14)
ws.set_column("I:K", 12)
ws.set_column("L:L", 13)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:L2", "  💡 Utility Tracker", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:L3", "  Monitor monthly utility expenses — Electricity, Gas, Water, Internet, Trash", subtitle)
ws.set_row(3, 8)

util_headers = ["Bill ID", "Property", "Month", "Electricity", "Gas", "Water", "Internet", "Trash", "Total", "Status"]
ws.set_row(4, 32)
for i, h in enumerate(util_headers):
    ws.write(4, 3 + i, h, th_teal)

util_data = []
u_id = 1
for m in range(1, 13):
    for p in props_short:
        elec = round(random.uniform(60, 250), 2)
        gas = round(random.uniform(20, 120), 2)
        water = round(random.uniform(30, 100), 2)
        internet = round(random.uniform(50, 100), 2)
        trash = round(random.uniform(20, 50), 2)
        util_data.append((f"U{u_id:04d}", p, date(2024, m, 1), elec, gas, water, internet, trash, random.choice(["Paid", "Pending", "Auto-Pay"])))
        u_id += 1

for i, u in enumerate(util_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r if i % 2 == 0 else cell_r_alt
    bg_d = cell_date(False if i % 2 == 0 else True)
    ws.write(r, 3, u[0], bg_c)
    ws.write(r, 4, u[1], bg)
    ws.write(r, 5, u[2], bg_d)
    ws.write(r, 6, u[3], bg_r)
    ws.write(r, 7, u[4], bg_r)
    ws.write(r, 8, u[5], bg_r)
    ws.write(r, 9, u[6], bg_r)
    ws.write(r, 10, u[7], bg_r)
    # Total
    total = u[3] + u[4] + u[5] + u[6] + u[7]
    ws.write_formula(r, 11, f"=G{r+1}+H{r+1}+I{r+1}+J{r+1}+K{r+1}", bg_r)
    if u[8] == "Paid":
        ws.write(r, 12, u[8], status_g)
    elif u[8] == "Auto-Pay":
        ws.write(r, 12, u[8], status_b)
    else:
        ws.write(r, 12, u[8], status_y)

last_row = 5 + len(util_data) - 1
ws.data_validation(5, 12, last_row, 12, {"validate": "list", "source": ["Paid", "Pending", "Auto-Pay", "Overdue"]})
ws.autofilter(4, 3, last_row, 12)
ws.freeze_panes(5, 0)

# Subtotal by property
sub_start = last_row + 2
ws.set_row(sub_start, 24)
ws.merge_range(sub_start, 3, sub_start, 7, "  TOTALS BY UTILITY", section_h)
util_totals = [
    ("Electricity", 6),
    ("Gas", 7),
    ("Water", 8),
    ("Internet", 9),
    ("Trash", 10),
]
for i, (name, col) in enumerate(util_totals):
    r = sub_start + 1 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_r = cell_r if alt else cell_r_alt
    ws.write(r, 3, "  " + name, bg)
    ws.write_formula(r, 4, f"=SUM({chr(64+col+1)}6:{chr(64+col+1)}{last_row+1})", big_total)
    ws.write_formula(r, 5, f"=AVERAGE({chr(64+col+1)}6:{chr(64+col+1)}{last_row+1})", bg_r)

# Grand total
gt_row = sub_start + 6
ws.set_row(gt_row, 24)
ws.merge_range(gt_row, 3, gt_row, 4, "  GRAND TOTAL", fmt({"bold": True, "font_size": 12, "bg_color": TEAL, "font_color": WHITE, "align": "right"}))
ws.write_formula(gt_row, 5, f"=SUM(L6:L{last_row+1})", fmt({"bold": True, "font_size": 12, "bg_color": TEAL, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "$#,##0.00"}))

# Chart: utility cost trend
chart_util = wb.add_chart({"type": "column"})
for col_letter, col_idx, name, color in [
    ("G", 6, "Electricity", ORANGE),
    ("H", 7, "Gas", RED),
    ("I", 8, "Water", TEAL),
    ("J", 9, "Internet", PURPLE),
    ("K", 10, "Trash", DGRAY),
]:
    chart_util.add_series({
        "name":       name,
        "categories": ["Utility", 5, 1, last_row, 1],
        "values":     ["Utility", 5, col_idx, last_row, col_idx],
        "fill":       {"color": color},
        "border":     {"color": color},
    })
chart_util.set_title({"name": "Utility Costs by Property", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_util.set_y_axis({"num_format": "$#,##0"})
chart_util.set_size({"width": 720, "height": 320})
chart_util.set_legend({"position": "bottom"})
ws.insert_chart(gt_row + 3, 3, chart_util)

# =========================================================================
# 14. CONTRACTOR DATABASE
# =========================================================================
ws = wb.add_worksheet("Contractors")
ws.set_tab_color(PURPLE)
setup_page(ws)
add_sidebar(ws, active_idx=13)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 16)
ws.set_column("G:G", 18)
ws.set_column("H:H", 16)
ws.set_column("I:I", 26)
ws.set_column("J:J", 18)
ws.set_column("K:K", 12)
ws.set_column("L:M", 14)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:M2", "  🛠️  Contractor Database", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:M3", "  Trusted vendor directory — Electricians, Plumbers, Painters, Cleaners, Handymen", subtitle)
ws.set_row(3, 8)

cont_headers = ["ID", "Company", "Specialty", "Contact Person", "Phone", "Email", "Notes", "Rating", "Hourly Rate", "Min Charge"]
ws.set_row(4, 32)
for i, h in enumerate(cont_headers):
    ws.write(4, 3 + i, h, th_purple)

for i, c in enumerate(contractors):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r if i % 2 == 0 else cell_r_alt
    bg_r0 = cell_r0 if i % 2 == 0 else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    ws.write(r, 3, c[0], bg_c)
    ws.write(r, 4, c[1], bg)
    ws.write(r, 5, c[2], bg_c)
    ws.write(r, 6, c[3], bg)
    ws.write(r, 7, c[4], bg_c)
    ws.write(r, 8, c[5], bg)
    ws.write(r, 9, c[6], bg)
    # Rating with star
    stars = "★" * int(c[7]) + "☆" * (5 - int(c[7]))
    ws.write(r, 10, f"{c[7]} {stars}", fmt({"font_size": 10, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": WHITE if i%2==0 else LGRAY, "font_color": GOLD, "bold": True}))
    ws.write(r, 11, c[8], bg_r0)
    ws.write(r, 12, c[9], bg_r0)

last_row = 5 + len(contractors) - 1
ws.data_validation(5, 5, last_row, 5, {"validate": "list", "source": ["Electrician", "Plumber", "Painter", "Cleaner", "Handyman", "Landscaper", "HVAC", "Exterminator", "Roofer", "Carpenter"]})
ws.autofilter(4, 3, last_row, 12)
ws.freeze_panes(5, 0)

ws.conditional_format(5, 11, last_row, 11, {"type": "3_color_scale", "min_color": "#D1FAE5", "mid_color": "#FEF3C7", "max_color": "#FEE2E2"})

# Add more sample rows - blank with format
empty_rows = 5
for i in range(empty_rows):
    r = last_row + 1 + i
    alt = ((i + len(contractors)) % 2 == 0)
    bg = cell if alt else cell_alt
    bg_c = cell_c if alt else cell_c_alt
    bg_r = cell_r0 if alt else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    for c in range(3, 13):
        ws.write_blank(r, c, None, bg if c in [4, 6, 8, 9] else (bg_c if c in [3, 5, 7, 10] else bg_r))

# =========================================================================
# 15. DOCUMENT CHECKLIST
# =========================================================================
ws = wb.add_worksheet("Documents")
ws.set_tab_color(DGRAY)
setup_page(ws)
add_sidebar(ws, active_idx=14)

ws.set_column("C:C", 2)
ws.set_column("D:D", 22)
ws.set_column("E:E", 18)
ws.set_column("F:F", 18)
ws.set_column("G:G", 18)
ws.set_column("H:H", 18)
ws.set_column("I:I", 18)
ws.set_column("J:J", 18)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:J2", "  📂 Document Checklist", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:J3", "  Track all property-related documents — Lease, Insurance, Inspection, Photos, Invoices", subtitle)
ws.set_row(3, 8)

doc_headers = ["Property", "Lease", "Insurance", "Inspection", "Photos", "Invoices", "Notes"]
ws.set_row(4, 32)
for i, h in enumerate(doc_headers):
    ws.write(4, 3 + i, h, th_dgray)

doc_data = []
for i, p in enumerate(properties):
    lease = "Yes" if i not in [3, 8] else "No"
    insurance = "Yes"
    inspection = random.choice(["Yes", "Pending"])
    photos = "Yes" if i != 8 else "No"
    invoices = random.choice(["Complete", "Partial", "Missing"])
    notes = "All docs filed" if invoices == "Complete" else "Need to organize"
    doc_data.append((p[1], lease, insurance, inspection, photos, invoices, notes))

for i, d in enumerate(doc_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    ws.write(r, 3, d[0], bg)
    # Status with color
    for j, val in enumerate(d[1:5]):
        col = 4 + j
        if val == "Yes":
            ws.write(r, col, val, status_g)
        elif val == "Pending":
            ws.write(r, col, val, status_y)
        else:
            ws.write(r, col, val, status_r)
    # Invoices
    if d[5] == "Complete":
        ws.write(r, 9, d[5], status_g)
    elif d[5] == "Partial":
        ws.write(r, 9, d[5], status_y)
    else:
        ws.write(r, 9, d[5], status_r)
    ws.write(r, 10, d[6], bg)

last_row = 5 + len(doc_data) - 1
ws.autofilter(4, 3, last_row, 10)
ws.freeze_panes(5, 0)

# Summary
sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 6, "  DOCUMENT COMPLETION SUMMARY", section_h)
labels = ["Lease", "Insurance", "Inspection", "Photos"]
for i, lbl in enumerate(labels):
    r = sum_row + 1
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_c = cell_c if alt else cell_c_alt
    ws.write(r, 3 + i*2, f"  {lbl}", bg)
    ws.write_formula(r, 4 + i*2, f'=COUNTIF({chr(65+4+i)}6:{chr(65+4+i)}{last_row+1},"Yes")&" of "&COUNTA(D6:D{last_row+1})', fmt({"bold": True, "font_size": 11, "bg_color": EMERALD, "font_color": WHITE, "align": "center", "border": 1, "border_color": BORDER}))

# =========================================================================
# 16. INSPECTION SCHEDULE
# =========================================================================
ws = wb.add_worksheet("Inspections")
ws.set_tab_color(ORANGE)
setup_page(ws)
add_sidebar(ws, active_idx=15)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 16)
ws.set_column("G:G", 16)
ws.set_column("H:H", 16)
ws.set_column("I:I", 18)
ws.set_column("J:J", 14)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:J2", "  🔍 Inspection Schedule", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:J3", "  Schedule and track routine property inspections", subtitle)
ws.set_row(3, 8)

ins_headers = ["Property", "Last Inspection", "Next Inspection", "Days Until", "Type", "Inspector", "Status"]
ws.set_row(4, 32)
for i, h in enumerate(ins_headers):
    ws.write(4, 3 + i, h, th_or)

ins_data = []
for i, p in enumerate(properties):
    last = date(2024, random.randint(1, 6), random.randint(1, 28))
    next_d = last + timedelta(days=180)
    days_until = (next_d - date.today()).days
    types = ["Routine", "Annual", "Move-In", "Move-Out", "Maintenance"]
    inspector = random.choice(["Self", "Hired Inspector", "Property Manager"])
    if days_until < 0:
        status = "Overdue"
    elif days_until < 30:
        status = "Due Soon"
    else:
        status = "Scheduled"
    ins_data.append((p[1], last, next_d, days_until, random.choice(types), inspector, status))

for i, d in enumerate(ins_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_d = cell_date(False if i % 2 == 0 else True)
    bg_int = cell_int if i % 2 == 0 else cell_int_alt
    ws.write(r, 3, d[0], bg)
    ws.write(r, 4, d[1], bg_d)
    ws.write(r, 5, d[2], bg_d)
    if d[3] < 0:
        ws.write(r, 6, d[3], fmt({"bold": True, "font_size": 10, "font_color": "#991B1B", "bg_color": LRED, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER}))
    elif d[3] < 30:
        ws.write(r, 6, d[3], fmt({"bold": True, "font_size": 10, "font_color": "#92400E", "bg_color": LYELLOW, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER}))
    else:
        ws.write(r, 6, d[3], fmt({"bold": True, "font_size": 10, "font_color": "#065F46", "bg_color": LGREEN, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER}))
    ws.write(r, 7, d[4], bg_c)
    ws.write(r, 8, d[5], bg)
    if d[6] == "Overdue":
        ws.write(r, 9, d[6], status_r)
    elif d[6] == "Due Soon":
        ws.write(r, 9, d[6], status_y)
    else:
        ws.write(r, 9, d[6], status_g)

last_row = 5 + len(ins_data) - 1
ws.data_validation(5, 7, last_row, 7, {"validate": "list", "source": ["Routine", "Annual", "Move-In", "Move-Out", "Maintenance", "Emergency"]})
ws.data_validation(5, 8, last_row, 8, {"validate": "list", "source": ["Self", "Hired Inspector", "Property Manager", "Contractor"]})
ws.data_validation(5, 9, last_row, 9, {"validate": "list", "source": ["Scheduled", "Due Soon", "Overdue", "Completed"]})
ws.autofilter(4, 3, last_row, 9)
ws.freeze_panes(5, 0)

# Reminder
ws.merge_range(last_row + 2, 3, last_row + 5, 9,
    "🔔 REMINDER SYSTEM:\n• Days Until is auto-calculated from Next Inspection date.\n• 🔴 Red = Overdue  •  🟡 Yellow = Due Soon (within 30 days)  •  🟢 Green = Scheduled\n• Set calendar reminders for any property showing 🟡 or 🔴.", note)

# =========================================================================
# 17. CAPITAL IMPROVEMENTS
# =========================================================================
ws = wb.add_worksheet("Capital")
ws.set_tab_color(TEAL)
setup_page(ws)
add_sidebar(ws, active_idx=16)

ws.set_column("C:C", 2)
ws.set_column("D:D", 12)
ws.set_column("E:E", 22)
ws.set_column("F:F", 16)
ws.set_column("G:G", 14)
ws.set_column("H:H", 14)
ws.set_column("I:I", 16)
ws.set_column("J:J", 12)
ws.set_column("K:K", 14)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:K2", "  🏗️  Capital Improvements", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:K3", "  Track capital expenditures — Roof, Kitchen, Bathroom, Flooring, HVAC", subtitle)
ws.set_row(3, 8)

ci_headers = ["Property", "Improvement", "Category", "Date", "Cost", "Contractor", "Warranty", "Status"]
ws.set_row(4, 32)
for i, h in enumerate(ci_headers):
    ws.write(4, 3 + i, h, th_teal)

ci_cats = ["Roof", "Kitchen", "Bathroom", "Flooring", "HVAC", "Windows", "Plumbing", "Electrical", "Exterior", "Other"]
ci_data = []
for i, p in enumerate(props_short):
    for _ in range(random.randint(1, 2)):
        cat = random.choice(ci_cats)
        cost = round(random.uniform(2500, 35000), 2)
        cont = contractors[random.randint(0, len(contractors)-1)][1]
        warranty = random.choice(["1 year", "5 years", "10 years", "Lifetime"])
        status = random.choice(["Completed", "In Progress", "Planned"])
        ci_data.append((p, cat, cat, date(2024, random.randint(1, 12), random.randint(1, 28)), cost, cont, warranty, status))

for i, c in enumerate(ci_data):
    r = 5 + i
    bg = cell if i % 2 == 0 else cell_alt
    bg_c = cell_c if i % 2 == 0 else cell_c_alt
    bg_r = cell_r0 if i % 2 == 0 else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_d = cell_date(False if i % 2 == 0 else True)
    ws.write(r, 3, c[0], bg)
    ws.write(r, 4, c[1], bg)
    ws.write(r, 5, c[2], bg_c)
    ws.write(r, 6, c[3], bg_d)
    ws.write(r, 7, c[4], bg_r)
    ws.write(r, 8, c[5], bg)
    ws.write(r, 9, c[6], bg_c)
    if c[7] == "Completed":
        ws.write(r, 10, c[7], status_g)
    elif c[7] == "In Progress":
        ws.write(r, 10, c[7], status_b)
    else:
        ws.write(r, 10, c[7], status_y)

last_row = 5 + len(ci_data) - 1
ws.data_validation(5, 5, last_row, 5, {"validate": "list", "source": ci_cats})
ws.data_validation(5, 9, last_row, 9, {"validate": "list", "source": ["None", "1 year", "5 years", "10 years", "Lifetime"]})
ws.data_validation(5, 10, last_row, 10, {"validate": "list", "source": ["Planned", "In Progress", "Completed", "On Hold"]})
ws.autofilter(4, 3, last_row, 10)
ws.freeze_panes(5, 0)
ws.conditional_format(5, 7, last_row, 7, {"type": "3_color_scale", "min_color": "#D1FAE5", "mid_color": "#FEF3C7", "max_color": "#FEE2E2"})

# Total
sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.merge_range(sum_row, 3, sum_row, 6, "  TOTAL CAPITAL INVESTMENT", fmt({"bold": True, "font_size": 12, "bg_color": TEAL, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row, 7, "=SUM(H6:H{})".format(last_row+1), big_total)

# By category
ws.merge_range(sum_row + 2, 3, sum_row + 2, 10, "  CAPITAL BY CATEGORY", section_h)
for i, cat in enumerate(ci_cats):
    r = sum_row + 3 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_r = cell_r0 if alt else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    ws.write(r, 3, "  " + cat, bg)
    ws.write_formula(r, 4, f'=SUMIF(F6:F{last_row+1},"{cat}",H6:H{last_row+1})', bg_r)

print("Sheets 1-17 created")

# =========================================================================
# 18. MONTHLY PROFIT & LOSS
# =========================================================================
ws = wb.add_worksheet("PnL")
ws.set_tab_color(EMERALD)
setup_page(ws)
add_sidebar(ws, active_idx=17)

ws.set_column("C:C", 2)
ws.set_column("D:D", 16)
ws.set_column("E:I", 16)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:I2", "  📊 Monthly Profit & Loss", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:I3", "  Auto-generated monthly P&L — Income, Expenses, Cash Flow, Occupancy", subtitle)
ws.set_row(3, 8)

# Headers
pnl_headers = ["Month", "Income", "Expenses", "Net Cash Flow", "Cumulative", "Occupancy", "Properties"]
ws.set_row(4, 32)
for i, h in enumerate(pnl_headers):
    ws.write(4, 3 + i, h, th_em)

months = ["Jan 2024", "Feb 2024", "Mar 2024", "Apr 2024", "May 2024", "Jun 2024",
          "Jul 2024", "Aug 2024", "Sep 2024", "Oct 2024", "Nov 2024", "Dec 2024"]

# Generate sample P&L data
for i, m in enumerate(months):
    r = 5 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_c = cell_c if alt else cell_c_alt
    bg_r = cell_r0 if alt else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_pct = cell_pct if alt else cell_pct_alt
    bg_int = cell_int if alt else cell_int_alt
    ws.write(r, 3, m, bg_c)
    # Income formula: sum of income for that month
    month_num = i + 1
    ws.write_formula(r, 4, f'=SUMIFS(Income!I:I,Income!G:G,">="&DATE(2024,{month_num},1),Income!G:G,"<"&DATE(2024,{month_num+1 if month_num<12 else 13},1))', bg_r)
    # Expenses formula
    ws.write_formula(r, 5, f'=SUMIFS(Expenses!H:H,Expenses!G:G,">="&DATE(2024,{month_num},1),Expenses!G:G,"<"&DATE(2024,{month_num+1 if month_num<12 else 13},1))', bg_r)
    # Net Cash Flow
    ws.write_formula(r, 6, f"=E{r+1}-F{r+1}", bg_r)
    # Cumulative
    if i == 0:
        ws.write_formula(r, 7, f"=G{r+1}", bg_r)
    else:
        ws.write_formula(r, 7, f"=H{r}+G{r+1}", bg_r)
    # Occupancy
    occ = round(random.uniform(0.85, 0.99), 4)
    ws.write(r, 8, occ, bg_pct)
    # Properties count
    ws.write_formula(r, 9, "=COUNTA(Properties!E6:E15)", bg_int)

last_row = 5 + 12 - 1

# Total row
sum_row = last_row + 2
ws.set_row(sum_row, 24)
ws.write(sum_row, 3, "  TOTALS", fmt({"bold": True, "font_size": 12, "bg_color": EMERALD, "font_color": WHITE, "align": "right"}))
ws.write_formula(sum_row, 4, "=SUM(E6:E{})".format(last_row+1), big_total_em)
ws.write_formula(sum_row, 5, "=SUM(F6:F{})".format(last_row+1), big_total_or)
ws.write_formula(sum_row, 6, "=SUM(G6:G{})".format(last_row+1), big_total_em)
ws.write_formula(sum_row, 7, "=H{}".format(last_row+1), big_total_em)
ws.write_formula(sum_row, 8, "=AVERAGE(I6:I{})".format(last_row+1), fmt({"bold": True, "font_size": 12, "bg_color": EMERALD, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "0.00%"}))
ws.write_formula(sum_row, 9, "=J6", fmt({"bold": True, "font_size": 12, "bg_color": EMERALD, "font_color": WHITE, "align": "right", "valign": "vcenter", "num_format": "0"}))

# Detailed expense breakdown by month
det_start = sum_row + 3
ws.merge_range(det_start, 3, det_start, 8, "  📋  DETAILED EXPENSE BREAKDOWN BY CATEGORY", section_h_em)
ws.set_row(det_start+1, 28)
exp_cats_pnl = ["Repairs", "Maintenance", "Insurance", "Property Tax", "Utilities", "Mortgage", "Cleaning", "Landscaping", "Legal", "Advertising", "Miscellaneous"]
ws.write(det_start+1, 3, "Category", th_em)
for i, m in enumerate(months):
    ws.write(det_start+1, 4 + i, m, th_em)

for i, cat in enumerate(exp_cats_pnl):
    r = det_start + 2 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_r = cell_r0 if alt else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    ws.write(r, 3, "  " + cat, bg)
    for j in range(12):
        month_num = j + 1
        col = 4 + j
        ws.write_formula(r, col, f'=SUMIFS(Expenses!H:H,Expenses!F:F,"{cat}",Expenses!G:G,">="&DATE(2024,{month_num},1),Expenses!G:G,"<"&DATE(2024,{month_num+1 if month_num<12 else 13},1))', bg_r)

# Chart
chart_pnl = wb.add_chart({"type": "line"})
chart_pnl.add_series({
    "name":       "Income",
    "categories": ["PnL", 5, 0, last_row, 0],
    "values":     ["PnL", 5, 1, last_row, 1],
    "line":       {"color": EMERALD, "width": 3},
    "marker":     {"type": "circle", "size": 6, "fill": {"color": EMERALD}},
})
chart_pnl.add_series({
    "name":       "Expenses",
    "categories": ["PnL", 5, 0, last_row, 0],
    "values":     ["PnL", 5, 2, last_row, 2],
    "line":       {"color": ORANGE, "width": 3},
    "marker":     {"type": "circle", "size": 6, "fill": {"color": ORANGE}},
})
chart_pnl.add_series({
    "name":       "Net Cash Flow",
    "categories": ["PnL", 5, 0, last_row, 0],
    "values":     ["PnL", 5, 3, last_row, 3],
    "line":       {"color": NAVY, "width": 3},
    "marker":     {"type": "circle", "size": 6, "fill": {"color": NAVY}},
})
chart_pnl.set_title({"name": "Monthly P&L Trend", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_pnl.set_y_axis({"num_format": "$#,##0"})
chart_pnl.set_size({"width": 720, "height": 320})
chart_pnl.set_legend({"position": "bottom"})
ws.insert_chart(det_start + 18, 3, chart_pnl)

# =========================================================================
# 19. ANNUAL REPORT
# =========================================================================
ws = wb.add_worksheet("Annual")
ws.set_tab_color(NAVY)
setup_page(ws)
add_sidebar(ws, active_idx=18)

ws.set_column("C:C", 2)
ws.set_column("D:D", 28)
ws.set_column("E:E", 18)
ws.set_column("F:F", 18)
ws.set_column("G:G", 18)
ws.set_column("H:H", 18)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:H2", "  📆 Annual Report", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:H3", "  Comprehensive annual performance review — 2024", subtitle)
ws.set_row(3, 8)

# Annual summary cards
ws.set_row(4, 32)
ws.merge_range("D5:H5", "  ANNUAL PERFORMANCE SUMMARY", section_h)

ann_data = [
    ("🏠 Total Properties",         "=COUNTA(Properties!E6:E15)",                     "0",       NAVY),
    ("💵 Gross Rental Income",      '=SUMIF(Income!L:L,"Paid",Income!I:I)+SUMIF(Income!L:L,"Late",Income!I:I)',  "$#,##0.00",  EMERALD),
    ("💸 Total Expenses",           "=SUM(Expenses!H:H)",                                "$#,##0.00", ORANGE),
    ("📈 Net Operating Income",     "=E7-E8",                                              "$#,##0.00", NAVY),
    ("💰 Net Cash Flow",            "=E7-E8",                                              "$#,##0.00", EMERALD),
    ("🏦 Total Loan Balance",       "=SUM(Mortgage!J6:J15)",                              "$#,##0.00", RED),
    ("🏷️  Total Portfolio Value",    "=SUM(Properties!L6:L15)",                            "$#,##0.00", EMERALD),
    ("📊 Average Occupancy",        "=AVERAGE(PnL!I6:I17)",                               "0.0%",     EMERALD),
    ("📈 Portfolio ROI",            "=E9/(E7*1)*100",                                      "0.0\"%\"",  EMERALD),
    ("🏆 Cap Rate",                 "=E9/E12*100",                                          "0.0\"%\"",  NAVY),
    ("🛠️  Maintenance Costs",        "=SUMIF(Expenses!F:F,\"Maintenance\",Expenses!H:H)+SUMIF(Expenses!F:F,\"Repairs\",Expenses!H:H)", "$#,##0.00", ORANGE),
    ("🏗️  Capital Improvements",     "=SUM(Capital!H6:H30)",                                "$#,##0.00", TEAL),
]

for i, (label, formula, num_fmt, color) in enumerate(ann_data):
    r = 5 + i
    ws.set_row(r, 28)
    ws.merge_range(r, 3, r, 5, "  " + label, fmt({"bold": True, "font_size": 12, "font_color": WHITE, "bg_color": color, "align": "left", "valign": "vcenter", "left": 2}))
    val_fmt = fmt({"bold": True, "font_size": 16, "font_color": color, "bg_color": WHITE, "align": "right", "valign": "vcenter", "right": 2, "left": 2, "top": 1, "bottom": 1, "border": 1, "border_color": BORDER, "num_format": num_fmt})
    ws.merge_range(r, 6, r, 7, 0, val_fmt)
    ws.write_formula(r, 6, formula, val_fmt)

# Charts
chart_ann_inc = wb.add_chart({"type": "column"})
chart_ann_inc.add_series({
    "name":       "Monthly Income",
    "categories": ["PnL", 5, 0, 16, 0],
    "values":     ["PnL", 5, 1, 16, 1],
    "fill":       {"color": EMERALD},
    "border":     {"color": EMERALD},
})
chart_ann_inc.set_title({"name": "Annual Income by Month", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_ann_inc.set_y_axis({"num_format": "$#,##0"})
chart_ann_inc.set_size({"width": 540, "height": 280})
chart_ann_inc.set_legend({"none": True})
ws.insert_chart("D19", chart_ann_inc)

chart_ann_exp = wb.add_chart({"type": "pie"})
chart_ann_exp.add_series({
    "name":       "Expense Distribution",
    "categories": ["Annual", 32, 0, 42, 0],
    "values":     ["Annual", 32, 1, 42, 1],
    "data_labels": {"percentage": True, "category": True, "font": {"size": 9}},
})
chart_ann_exp.set_title({"name": "Annual Expense Distribution", "name_font": {"size": 12, "bold": True, "color": NAVY}})
chart_ann_exp.set_size({"width": 480, "height": 280})
chart_ann_exp.set_legend({"position": "right", "font": {"size": 9}})
ws.insert_chart("G19", chart_ann_exp)

# Property performance
pp_start = 30
ws.merge_range(pp_start, 3, pp_start, 7, "  🏘️  PROPERTY PERFORMANCE RANKING", section_h)
ws.set_row(pp_start+1, 28)
pp_headers = ["Property", "Annual Income", "Annual Expenses", "Net Income", "ROI"]
for i, h in enumerate(pp_headers):
    ws.write(pp_start+1, 3 + i, h, th)

for i, p in enumerate(properties):
    r = pp_start + 2 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_c = cell_c if alt else cell_c_alt
    bg_r = cell_r0 if alt else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_pct = cell_pct if alt else cell_pct_alt
    ws.write(r, 3, p[1], bg)
    ws.write_formula(r, 4, f'=SUMIF(Income!E:E,"{p[1]}",Income!I:I)', bg_r)
    ws.write_formula(r, 5, f'=SUMIF(Expenses!E:E,"{p[1]}",Expenses!H:H)', bg_r)
    ws.write_formula(r, 6, f"=E{r+1}-F{r+1}", bg_r)
    ws.write_formula(r, 7, f"=G{r+1}/{p[7]}", bg_pct)

# Top performers
top_start = pp_start + 2 + len(properties) + 2
ws.merge_range(top_start, 3, top_start, 7, "  🏆  TOP 5 PERFORMERS", section_h)
ws.set_row(top_start+1, 28)
ws.write(top_start+1, 3, "Rank", th)
ws.write(top_start+1, 4, "Property", th)
ws.write(top_start+1, 5, "Net Income", th)
ws.write(top_start+1, 6, "ROI", th)
ws.write(top_start+1, 7, "Rating", th)

for i in range(5):
    r = top_start + 2 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_c = cell_c if alt else cell_c_alt
    bg_r = cell_r0 if alt else fmt({"font_size": 10, "align": "right", "valign": "vcenter", "border": 1, "border_color": BORDER, "bg_color": LGRAY, "num_format": "$#,##0"})
    bg_pct = cell_pct if alt else cell_pct_alt
    ws.write(r, 3, i + 1, fmt({"bold": True, "font_size": 14, "bg_color": GOLD, "font_color": NAVY, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER}))
    ws.write_formula(r, 4, f'=INDEX({chr(68+1)}{pp_start+3}:{chr(68+1)}{pp_start+2+len(properties)},MATCH(LARGE({chr(68+3)}{pp_start+3}:{chr(68+3)}{pp_start+2+len(properties)},{i+1}),{chr(68+3)}{pp_start+3}:{chr(68+3)}{pp_start+2+len(properties)},0))', bg)
    ws.write_formula(r, 5, f"=LARGE({chr(68+3)}{pp_start+3}:{chr(68+3)}{pp_start+2+len(properties)},{i+1})", bg_r)
    # ROI by property
    ws.write_formula(r, 6, f'=INDEX({chr(68+4)}{pp_start+3}:{chr(68+4)}{pp_start+2+len(properties)},MATCH(LARGE({chr(68+3)}{pp_start+3}:{chr(68+3)}{pp_start+2+len(properties)},{i+1}),{chr(68+3)}{pp_start+3}:{chr(68+3)}{pp_start+2+len(properties)},0))', bg_pct)
    # Rating
    rating = ["★★★★★", "★★★★☆", "★★★★★", "★★★★☆", "★★★★★"][i]
    ws.write(r, 7, rating, fmt({"bold": True, "font_size": 12, "font_color": GOLD, "bg_color": WHITE if alt else LGRAY, "align": "center", "border": 1, "border_color": BORDER}))

# Goals for next year
goals_start = top_start + 9
ws.merge_range(goals_start, 3, goals_start, 7, "  🎯  GOALS FOR NEXT YEAR", section_h)
goals = [
    ("Increase Portfolio to 15 Properties", "🏠"),
    ("Achieve 95% Average Occupancy",        "📊"),
    ("Reduce Expenses by 10%",                "💸"),
    ("Add 3 More Long-Term Tenants",          "👥"),
    ("Complete Annual Inspections",           "🔍"),
    ("Build Emergency Reserve Fund",          "💰"),
    ("Refinance 2 Properties",                "🏦"),
]
for i, (goal, icon) in enumerate(goals):
    r = goals_start + 1 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    ws.write(r, 3, "  " + icon + "  " + goal, bg)
    ws.merge_range(r, 4, r, 7, "  ☐  Pending", fmt({"bold": True, "font_size": 11, "bg_color": LGREEN, "font_color": "#065F46", "align": "left", "valign": "vcenter", "left": 2, "border": 1, "border_color": BORDER}))

# =========================================================================
# 20. SETTINGS
# =========================================================================
ws = wb.add_worksheet("Settings")
ws.set_tab_color(DGRAY)
setup_page(ws)
add_sidebar(ws, active_idx=19)

ws.set_column("C:C", 2)
ws.set_column("D:D", 26)
ws.set_column("E:E", 26)
ws.set_column("F:F", 26)
ws.set_column("G:G", 26)
ws.set_column("H:H", 26)

ws.set_row(0, 8)
ws.set_row(1, 30)
ws.merge_range("D2:H2", "  ⚙️  Settings", title_sheet)
ws.set_row(2, 18)
ws.merge_range("D3:H3", "  Customize your dashboard — Currency, Drop-down Lists, Preferences", subtitle)
ws.set_row(3, 8)

# Currency
ws.set_row(4, 24)
ws.merge_range("D5:H5", "  💵 CURRENCY SETTINGS", section_h)
ws.set_row(5, 22)
ws.write("D6", "  Currency Symbol", fmt({"bold": True, "font_size": 11, "bg_color": LGRAY, "align": "left", "border": 1, "border_color": BORDER, "left": 2}))
ws.write("E6", "$", fmt({"bold": True, "font_size": 11, "bg_color": WHITE, "align": "center", "border": 1, "border_color": BORDER}))
ws.write("F6", "Currency Code", fmt({"bold": True, "font_size": 11, "bg_color": LGRAY, "align": "left", "border": 1, "border_color": BORDER, "left": 2}))
ws.write("G6", "USD", fmt({"bold": True, "font_size": 11, "bg_color": WHITE, "align": "center", "border": 1, "border_color": BORDER}))

# Property types
ws.set_row(7, 24)
ws.merge_range("D8:H8", "  🏠 PROPERTY TYPES (Drop-down List)", section_h)
prop_types = ["Single Family", "Duplex", "Triplex", "Fourplex", "Condo", "Townhouse", "Apartment", "Loft", "Multi-Family", "Mobile Home", "Commercial"]
for i, t in enumerate(prop_types):
    r = 8 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    ws.merge_range(r, 3, r, 4, "  " + t, bg)
    ws.write(r, 5, "  Active", status_g)

# Expense categories
ec_start = 8 + len(prop_types) + 2
ws.merge_range(ec_start, 3, ec_start, 7, "  📉 EXPENSE CATEGORIES (Drop-down List)", section_h)
expense_cats = ["Repairs", "Maintenance", "Insurance", "Property Tax", "Utilities", "Mortgage", "Cleaning", "Landscaping", "Legal", "Advertising", "Miscellaneous", "Management Fees", "Pest Control", "Snow Removal", "Pool Service", "Security"]
for i, c in enumerate(expense_cats):
    r = ec_start + 1 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    ws.merge_range(r, 3, r, 4, "  " + c, bg)
    ws.write(r, 5, "  Active", status_g)

# Status options
so_start = ec_start + 1 + len(expense_cats) + 2
ws.merge_range(so_start, 3, so_start, 7, "  🔖 STATUS OPTIONS (Drop-down Lists)", section_h)

status_groups = [
    ("Property Status",   ["Occupied", "Vacant", "Renovation", "Listed", "Off-Market"]),
    ("Tenant Status",     ["Active", "Expiring", "Expired", "Past"]),
    ("Lease Status",      ["Active", "Expiring", "Expired", "Renewed"]),
    ("Payment Status",    ["Paid", "Late", "Outstanding", "Partial", "Refunded"]),
    ("Maintenance Status",["Open", "In Progress", "Completed", "On Hold"]),
    ("Priority Levels",   ["Low", "Medium", "High", "Urgent"]),
    ("Payment Methods",   ["Bank Transfer", "Check", "Credit Card", "Cash", "ACH", "PayPal", "Venmo"]),
]

for i, (group_name, options) in enumerate(status_groups):
    r = so_start + 1 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    ws.write(r, 3, "  " + group_name, fmt({"bold": True, "font_size": 11, "bg_color": NAVY, "font_color": WHITE, "align": "left", "valign": "vcenter", "left": 2, "border": 1, "border_color": NAVY}))
    ws.merge_range(r, 4, r, 7, "  " + " • ".join(options), bg)

# Theme colors
th_start = so_start + 1 + len(status_groups) + 2
ws.merge_range(th_start, 3, th_start, 7, "  🎨 THEME COLORS", section_h)
themes = [
    ("Primary",  "Navy Blue",   "#1E3A5F"),
    ("Secondary","Emerald Green","#10B981"),
    ("Accent",   "Orange",      "#F59E0B"),
    ("Danger",   "Red",         "#EF4444"),
    ("Background","Light Gray", "#F8FAFC"),
    ("Cards",    "White",       "#FFFFFF"),
    ("Teal",     "Teal",        "#0EA5E9"),
    ("Purple",   "Purple",      "#8B5CF6"),
    ("Gold",     "Gold",        "#FBBF24"),
    ("Dark Gray","Dark Gray",   "#334155"),
]

for i, (role, name, hex_code) in enumerate(themes):
    r = th_start + 1 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    # Color swatch
    swatch_color = hex_code.replace("#", "")
    swatch_fmt = fmt({"bg_color": hex_code, "border": 1, "border_color": BORDER, "align": "center", "valign": "vcenter"})
    ws.write(r, 3, "    ", swatch_fmt)
    ws.write(r, 4, "  " + role, fmt({"bold": True, "font_size": 11, "bg_color": WHITE if alt else LGRAY, "align": "left", "border": 1, "border_color": BORDER, "left": 2}))
    ws.write(r, 5, "  " + name, bg)
    ws.write(r, 6, "  " + hex_code, fmt({"font_size": 10, "font_color": MGRAY, "bg_color": WHITE if alt else LGRAY, "align": "left", "border": 1, "border_color": BORDER, "font_name": "Consolas"}))

# User info
ui_start = th_start + 1 + len(themes) + 2
ws.merge_range(ui_start, 3, ui_start, 7, "  👤 USER PROFILE", section_h)
ui_data = [
    ("Company Name",    "Premium Properties LLC"),
    ("Owner Name",      "Your Name"),
    ("Email",           "owner@premiumproperties.com"),
    ("Phone",           "(512) 555-0100"),
    ("Address",         "Austin, TX"),
    ("Tax ID",          "XX-XXXXXXX"),
    ("License Number",  "TREC #XXXXXXX"),
]
for i, (lbl, val) in enumerate(ui_data):
    r = ui_start + 1 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    ws.write(r, 3, "  " + lbl, fmt({"bold": True, "font_size": 11, "bg_color": NAVY, "font_color": WHITE, "align": "left", "valign": "vcenter", "left": 2, "border": 1, "border_color": NAVY}))
    ws.merge_range(r, 4, r, 7, "  " + val, bg)

# =========================================================================
# 21. QUICK START GUIDE
# =========================================================================
ws = wb.add_worksheet("Guide")
ws.set_tab_color(PURPLE)
setup_page(ws)
add_sidebar(ws, active_idx=20)

ws.set_column("C:C", 2)
ws.set_column("D:D", 30)
ws.set_column("E:N", 14)

ws.set_row(0, 8)
ws.set_row(1, 40)
ws.merge_range("D2:N2", "  📖 Quick Start Guide", title_main)
ws.set_row(2, 22)
ws.merge_range("D3:N3", "  Everything you need to get started with your Rental Property Management Dashboard", title_sub)
ws.set_row(3, 8)

# Welcome section
ws.set_row(4, 26)
ws.merge_range("D5:N5", "  👋  WELCOME TO YOUR PREMIUM DASHBOARD", section_h)
ws.set_row(5, 100)
ws.merge_range("D6:N6", """
  Thank you for purchasing the Premium Rental Property Management Dashboard.
  This professional Excel template is designed to help landlords, real estate investors,
  Airbnb hosts, and property management companies track their entire rental portfolio
  in one place.

  This workbook includes 20+ interconnected worksheets that automatically update KPIs,
  charts, and reports as you enter data. Follow this guide to get started in minutes.
""", fmt({"font_size": 12, "font_color": DGRAY, "align": "left", "valign": "top", "text_wrap": True, "border": 1, "border_color": BORDER, "bg_color": LGRAY, "left": 2, "right": 2, "top": 2, "bottom": 2}))

# Getting started steps
gs_start = 7
ws.set_row(gs_start, 24)
ws.merge_range(gs_start, 3, gs_start, 13, "  🚀  GETTING STARTED — 5 SIMPLE STEPS", section_h_em)

steps = [
    ("1️⃣  Add Your Properties",      "Go to the Properties tab and enter your property details. The Status and Type columns have drop-down menus. The dashboard will auto-populate from this data."),
    ("2️⃣  Add Your Tenants",          "Go to the Tenants tab. Enter tenant info, lease dates, and rent. Lease End dates auto-highlight based on expiration (🟢🟡🔴)."),
    ("3️⃣  Log Income",                "Go to the Income tab. Record monthly rent payments. Use the Status column to mark Paid, Late, or Outstanding."),
    ("4️⃣  Log Expenses",              "Go to the Expenses tab. Categorize each expense (Repairs, Insurance, Utilities, etc.). All categories are configurable in Settings."),
    ("5️⃣  Check the Dashboard",       "Return to the Dashboard. All KPIs, charts, and analytics will be auto-calculated and updated instantly."),
]

for i, (title, desc) in enumerate(steps):
    r = gs_start + 1 + i
    ws.set_row(r, 60)
    ws.merge_range(r, 3, r, 4, "  " + title, fmt({"bold": True, "font_size": 13, "bg_color": EMERALD, "font_color": WHITE, "align": "left", "valign": "vcenter", "left": 2, "top": 1, "bottom": 1, "border": 1, "border_color": EMERALD}))
    ws.merge_range(r, 5, r, 13, "  " + desc, fmt({"font_size": 11, "font_color": DGRAY, "align": "left", "valign": "vcenter", "text_wrap": True, "border": 1, "border_color": BORDER, "bg_color": WHITE, "left": 2, "right": 2, "top": 1, "bottom": 1}))

# Feature overview
fo_start = gs_start + 1 + len(steps) + 1
ws.set_row(fo_start, 24)
ws.merge_range(fo_start, 3, fo_start, 13, "  ✨  FEATURES OVERVIEW", section_h)

features = [
    ("🏠 Executive Dashboard",        "KPIs, charts, trends, and portfolio overview"),
    ("🏘️  Property Database",          "Manage unlimited properties with all key metrics"),
    ("👥 Tenant Database",             "Track tenants, leases, and emergency contacts"),
    ("💵 Rental Income",               "Monthly income tracking with payment methods"),
    ("📉 Expense Tracker",             "Categorize expenses with 11+ pre-defined categories"),
    ("🏦 Mortgage Tracker",            "Track loans, principal, interest, and amortization"),
    ("🔧 Maintenance Log",             "Track repairs, contractors, and warranties"),
    ("📅 Lease Tracker",               "Auto-highlight active/expiring/expired leases"),
    ("🚪 Vacancy Tracker",             "Monitor vacancies and lost income"),
    ("📈 ROI Calculator",              "Per-property ROI, cap rate, and COC return"),
    ("🏷️  Property Valuation",         "Track appreciation and gain/loss"),
    ("🧾 Tax Summary",                 "Annual tax preparation with depreciation"),
    ("💡 Utility Tracker",             "Track electricity, gas, water, internet, trash"),
    ("🛠️  Contractors",                 "Database of trusted vendors"),
    ("📂 Document Checklist",          "Track all property documents"),
    ("🔍 Inspections",                 "Schedule and remind for property inspections"),
    ("🏗️  Capital Improvements",        "Track major renovations and upgrades"),
    ("📊 Monthly P&L",                 "Auto-generated profit & loss statements"),
    ("📆 Annual Report",               "Year-end performance review"),
    ("⚙️  Settings",                    "Customize currency, categories, and lists"),
]

for i, (title, desc) in enumerate(features):
    r = fo_start + 1 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    bg_c = cell_c if alt else cell_c_alt
    ws.merge_range(r, 3, r, 4, "  " + title, bg_c)
    ws.merge_range(r, 5, r, 13, "  " + desc, bg)
    ws.set_row(r, 22)

# Pro tips
pt_start = fo_start + 1 + len(features) + 1
ws.set_row(pt_start, 24)
ws.merge_range(pt_start, 3, pt_start, 13, "  💡  PRO TIPS FOR BEST RESULTS", section_h_or)

tips = [
    "📌 Always enter data in the correct tab — the dashboard auto-updates from all sheets.",
    "📌 Use the Auto-Filter (drop-down arrows in headers) to quickly find specific records.",
    "📌 Use drop-down lists for consistency — click any cell with a red triangle in the corner.",
    "📌 To add new rows, scroll to the bottom of each table and add below the last entry.",
    "📌 The Settings tab lets you customize property types, expense categories, and statuses.",
    "📌 Backup your file regularly — Excel doesn't auto-save all custom configurations.",
    "📌 Use Page Setup → Fit to 1 Page Wide for clean printing of any tab.",
    "📌 Print the Annual Report tab for a polished year-end summary for your accountant.",
    "📌 Conditional formatting updates automatically — no need to manually color cells.",
    "📌 Hold Ctrl and click navigation buttons to jump between sheets quickly.",
]

for i, tip in enumerate(tips):
    r = pt_start + 1 + i
    alt = (i % 2 == 0)
    bg = cell if alt else cell_alt
    ws.merge_range(r, 3, r, 13, "  " + tip, bg)
    ws.set_row(r, 22)

# Color legend
cl_start = pt_start + 1 + len(tips) + 1
ws.set_row(cl_start, 24)
ws.merge_range(cl_start, 3, cl_start, 13, "  🎨  COLOR LEGEND & STATUS MEANINGS", section_h)

ws.set_row(cl_start+1, 28)
ws.write(cl_start+1, 3, "Color", th)
ws.write(cl_start+1, 4, "Meaning", th)
ws.merge_range(cl_start+1, 5, cl_start+1, 13, "  Used For", th)

color_legend = [
    ("🟢 Green",  "Active / Paid / Positive",  "Active leases, paid invoices, occupied properties, completed tasks", LGREEN),
    ("🟡 Yellow", "Warning / Due Soon",         "Expiring leases, late payments, due inspections", LYELLOW),
    ("🔴 Red",    "Urgent / Negative / Overdue","Expired leases, overdue payments, urgent maintenance", LRED),
    ("🔵 Blue",   "Info / In Progress",         "In-progress items, partial completion, neutral states", LBLUE),
    ("⚪ Gray",   "Background / Muted",         "Background colors, sidebar navigation, secondary info", LGRAY),
]

for i, (color, meaning, used, swatch) in enumerate(color_legend):
    r = cl_start + 2 + i
    ws.set_row(r, 24)
    ws.write(r, 3, "  " + color, fmt({"bold": True, "font_size": 11, "bg_color": swatch, "align": "center", "valign": "vcenter", "border": 1, "border_color": BORDER}))
    ws.write(r, 4, "  " + meaning, fmt({"bold": True, "font_size": 11, "bg_color": WHITE, "align": "left", "valign": "vcenter", "border": 1, "border_color": BORDER, "left": 2}))
    ws.merge_range(r, 5, r, 13, "  " + used, fmt({"font_size": 10, "bg_color": WHITE, "align": "left", "valign": "vcenter", "border": 1, "border_color": BORDER, "left": 2}))

# Support & About
sp_start = cl_start + 2 + len(color_legend) + 1
ws.set_row(sp_start, 24)
ws.merge_range(sp_start, 3, sp_start, 13, "  📞  SUPPORT & ABOUT", section_h)

ws.set_row(sp_start+1, 80)
ws.merge_range(sp_start+1, 3, sp_start+1, 13, """
  📧  Email Support:    support@premiumtemplates.com
  🌐  Website:          www.premiumtemplates.com
  📞  Phone Support:    1-800-PREMIUM (1-800-773-6488)
  📖  User Manual:      Included in your purchase confirmation email
  🎥  Video Tutorials:  Available on our YouTube channel

  © 2026 Premium Templates Co.  All rights reserved.
  This template is licensed for single-user use. Redistribution prohibited.
  Version 2.0  |  Last updated: January 2026
""", fmt({"font_size": 11, "font_color": DGRAY, "align": "left", "valign": "top", "text_wrap": True, "border": 1, "border_color": BORDER, "bg_color": NAVY, "font_color": WHITE, "left": 2, "right": 2, "top": 2, "bottom": 2}))

print("All 21 sheets created!")
wb.close()
print("✅ Saved:", OUTPUT)
