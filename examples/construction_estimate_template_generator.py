#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Construction Estimate Template Generator (Excel / Google Sheets Compatible)
===========================================================================

Title:
“Construction Estimate Template Excel | Contractor Bid & Job Costing Spreadsheet, Markup Margin Calculator, Proposal Generator Google Sheets”

Author: novality store
Formula Protection Password: premium
Compatibility: Microsoft Excel (.xlsx) & Google Sheets (no macros required)

Features Included:
1.  Instructions Sheet
2.  Project Info (Client details, contractor details, global tax/discount/contingency rates)
3.  Estimate (Construction Estimate Sheet & Contractor Bid Template)
4.  Materials (Materials Cost Sheet with quantity, unit cost, supplier, tax/shipping allowance)
5.  Labor (Labor Cost Sheet with task description, crew size, hours, rates, overtime)
6.  Equipment (Equipment Cost Sheet with rental/ownership cost, daily/hourly rates, usage)
7.  Subcontractors (Subcontractor Cost Sheet with scope, quoted price, markup %, billed amount)
8.  Markup & Margin Calculator (Separate markups by category, markup vs. margin comparison, target margin calculator)
9.  Job Costing & Actuals (Estimated vs actual cost tracking, variance analysis, status indicator)
10. Summary Dashboard (Executive KPI cards, cell-based visual dashboard, profitability summary)
11. Client Proposal (Clean printable proposal page, scope of work, pricing summary, acceptance sign-off)
12. Terms & Conditions (Professional contractor terms, payment schedule, change order policy, warranty)
13. Material Price DB (Material price database with SKU, supplier, standard cost) [Optional Advanced Feature]
14. Labor Rate DB (Labor rate database with standard trade hourly rates and overtime) [Optional Advanced Feature]
15. Estimate Versions (Side-by-side comparison of Good / Better / Best estimate versions) [Optional Advanced Feature]
16. Change Order Tracker (Change order log with scope, add/deduct amount, client approval status) [Optional Advanced Feature]
17. Payment Schedule (Milestone payment schedule with % of contract, scheduled amount, status) [Optional Advanced Feature]
18. Invoice Summary (Progress invoice generator linking contract amount, change orders, retainage) [Optional Advanced Feature]
19. Project Timeline (12-week construction schedule tracker with phase status) [Optional Advanced Feature]
"""

import os
import sys
from xlsxwriter.workbook import Workbook


def generate_construction_estimate_template(output_path="construction_estimate_template.xlsx"):
    """
    Generates the complete 19-tab commercial Construction Estimate & Contractor Bid Excel Workbook.
    All formula cells are locked and protected with password 'premium'.
    Author name is set to 'novality store'.
    """
    wb = Workbook(output_path)

    # -------------------------------------------------------------------------
    # DOCUMENT PROPERTIES (Author: novality store)
    # -------------------------------------------------------------------------
    wb.set_properties({
        'title': 'Construction Estimate Template Excel | Contractor Bid & Job Costing Spreadsheet, Markup Margin Calculator, Proposal Generator Google Sheets',
        'subject': 'Professional Construction Estimating, Contractor Bid, Job Costing & Profit Margin Template',
        'author': 'novality store',
        'manager': 'novality store',
        'company': 'novality store',
        'category': 'Construction Management & Contractor Spreadsheets',
        'keywords': 'Construction Estimate, Contractor Bid, Job Costing, Markup Calculator, Profit Margin, Proposal Generator, Excel, Google Sheets, novality store',
        'comments': 'All formula cells are locked with the password premium. Compatible with Microsoft Excel and Google Sheets.',
        'status': 'Final Commercial Release',
    })

    # -------------------------------------------------------------------------
    # PREMIUM AESTHETIC PALETTE & CELL FORMATS
    # -------------------------------------------------------------------------
    # Colors:
    # Navy Primary: #1B365D | Slate Secondary: #2C3E50 | Accent Amber/Gold: #D97706
    # Soft Gray BG: #F8F9FA | Soft Ice Blue BG: #EBF5FB | Soft Emerald Green: #D1E7DD
    # Border Dark:  #BDC3C7 | Border Light:     #E2E8F0 | White:             #FFFFFF

    font_family = "Segoe UI"

    # 1. Main Titles & Banners (Locked = 1)
    fmt_title_banner = wb.add_format({
        'font_name': font_family, 'font_size': 15, 'bold': True,
        'font_color': '#FFFFFF', 'bg_color': '#1B365D',
        'align': 'left', 'valign': 'vcenter', 'indent': 1, 'locked': 1
    })
    fmt_section_banner = wb.add_format({
        'font_name': font_family, 'font_size': 12, 'bold': True,
        'font_color': '#FFFFFF', 'bg_color': '#2C3E50',
        'align': 'left', 'valign': 'vcenter', 'indent': 1, 'locked': 1
    })
    fmt_sub_banner = wb.add_format({
        'font_name': font_family, 'font_size': 11, 'bold': True,
        'font_color': '#1B365D', 'bg_color': '#EBF5FB',
        'align': 'left', 'valign': 'vcenter', 'indent': 1, 'border': 1, 'border_color': '#BDC3C7', 'locked': 1
    })

    # 2. Table Headers (Locked = 1)
    fmt_header = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'bold': True,
        'font_color': '#FFFFFF', 'bg_color': '#1B365D',
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#1B365D', 'locked': 1
    })
    fmt_header_left = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'bold': True,
        'font_color': '#FFFFFF', 'bg_color': '#1B365D',
        'align': 'left', 'valign': 'vcenter', 'indent': 1, 'border': 1, 'border_color': '#1B365D', 'locked': 1
    })
    fmt_header_amber = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'bold': True,
        'font_color': '#FFFFFF', 'bg_color': '#D97706',
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D97706', 'locked': 1
    })

    # 3. Editable Input Cells (Locked = 0) -> White background, light gray borders
    fmt_input_text = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#FFFFFF', 'align': 'left', 'valign': 'vcenter',
        'border': 1, 'border_color': '#CBD5E1', 'locked': 0
    })
    fmt_input_center = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter',
        'border': 1, 'border_color': '#CBD5E1', 'locked': 0
    })
    fmt_input_int = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#FFFFFF', 'align': 'right', 'valign': 'vcenter',
        'num_format': '#,##0', 'border': 1, 'border_color': '#CBD5E1', 'locked': 0
    })
    fmt_input_currency = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#FFFFFF', 'align': 'right', 'valign': 'vcenter',
        'num_format': '$#,##0.00', 'border': 1, 'border_color': '#CBD5E1', 'locked': 0
    })
    fmt_input_percent = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#FFFFFF', 'align': 'right', 'valign': 'vcenter',
        'num_format': '0.00%', 'border': 1, 'border_color': '#CBD5E1', 'locked': 0
    })
    fmt_input_date = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter',
        'num_format': 'yyyy-mm-dd', 'border': 1, 'border_color': '#CBD5E1', 'locked': 0
    })

    # 4. Locked Formula & Read-Only Cells (Locked = 1) -> Soft light backgrounds
    fmt_lock_text = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#F8F9FA', 'align': 'left', 'valign': 'vcenter',
        'border': 1, 'border_color': '#CBD5E1', 'locked': 1
    })
    fmt_lock_center = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#F8F9FA', 'align': 'center', 'valign': 'vcenter',
        'border': 1, 'border_color': '#CBD5E1', 'locked': 1
    })
    fmt_lock_currency = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#F8F9FA', 'align': 'right', 'valign': 'vcenter',
        'num_format': '$#,##0.00', 'border': 1, 'border_color': '#CBD5E1', 'locked': 1
    })
    fmt_lock_percent = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'font_color': '#1E293B',
        'bg_color': '#F8F9FA', 'align': 'right', 'valign': 'vcenter',
        'num_format': '0.00%', 'border': 1, 'border_color': '#CBD5E1', 'locked': 1
    })

    # 5. Total & Summary Rows (Locked = 1)
    fmt_total_label = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#1B365D',
        'bg_color': '#EBF5FB', 'align': 'right', 'valign': 'vcenter',
        'border': 1, 'border_color': '#BDC3C7', 'locked': 1
    })
    fmt_total_currency = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#1B365D',
        'bg_color': '#EBF5FB', 'align': 'right', 'valign': 'vcenter',
        'num_format': '$#,##0.00', 'border': 1, 'border_color': '#BDC3C7', 'locked': 1
    })
    fmt_total_percent = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#1B365D',
        'bg_color': '#EBF5FB', 'align': 'right', 'valign': 'vcenter',
        'num_format': '0.00%', 'border': 1, 'border_color': '#BDC3C7', 'locked': 1
    })

    # 6. Grand Total Accounting Style (Locked = 1, Double bottom border)
    fmt_grand_label = wb.add_format({
        'font_name': font_family, 'font_size': 11, 'bold': True, 'font_color': '#FFFFFF',
        'bg_color': '#1B365D', 'align': 'right', 'valign': 'vcenter',
        'border': 1, 'border_color': '#1B365D', 'locked': 1
    })
    fmt_grand_currency = wb.add_format({
        'font_name': font_family, 'font_size': 11, 'bold': True, 'font_color': '#FFFFFF',
        'bg_color': '#1B365D', 'align': 'right', 'valign': 'vcenter',
        'num_format': '$#,##0.00', 'border': 1, 'border_color': '#1B365D', 'locked': 1
    })

    # 7. KPI Card Styles (Locked = 1)
    fmt_kpi_label = wb.add_format({
        'font_name': font_family, 'font_size': 9, 'bold': True, 'font_color': '#475569',
        'bg_color': '#EBF5FB', 'align': 'center', 'valign': 'vcenter',
        'top': 1, 'left': 1, 'right': 1, 'border_color': '#CBD5E1', 'locked': 1
    })
    fmt_kpi_val_curr = wb.add_format({
        'font_name': font_family, 'font_size': 14, 'bold': True, 'font_color': '#1B365D',
        'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter',
        'bottom': 1, 'left': 1, 'right': 1, 'border_color': '#CBD5E1',
        'num_format': '$#,##0.00', 'locked': 1
    })
    fmt_kpi_val_pct = wb.add_format({
        'font_name': font_family, 'font_size': 14, 'bold': True, 'font_color': '#D97706',
        'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter',
        'bottom': 1, 'left': 1, 'right': 1, 'border_color': '#CBD5E1',
        'num_format': '0.0%', 'locked': 1
    })

    # 8. Proposal / Document Clean Formats
    fmt_prop_title = wb.add_format({
        'font_name': font_family, 'font_size': 16, 'bold': True, 'font_color': '#1B365D',
        'align': 'center', 'valign': 'vcenter', 'locked': 1
    })
    fmt_prop_subtitle = wb.add_format({
        'font_name': font_family, 'font_size': 11, 'font_color': '#64748B',
        'align': 'center', 'valign': 'vcenter', 'locked': 1
    })
    fmt_label_bold = wb.add_format({
        'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#1E293B',
        'align': 'left', 'valign': 'vcenter', 'locked': 1
    })
    fmt_note_text = wb.add_format({
        'font_name': font_family, 'font_size': 9, 'font_color': '#475569', 'italic': True,
        'align': 'left', 'valign': 'top', 'text_wrap': True, 'locked': 1
    })

    # Protect option dictionary used across all worksheets
    prot_options = {
        'select_locked_cells': 1,
        'select_unlocked_cells': 1,
        'format_cells': 1,
        'format_columns': 1,
        'format_rows': 1,
        'insert_rows': 1,
        'delete_rows': 1,
        'sort': 1,
        'autofilter': 1
    }

    # =========================================================================
    # TAB 1: INSTRUCTIONS
    # =========================================================================
    ws_inst = wb.add_worksheet('Instructions')
    ws_inst.set_tab_color('#1B365D')
    ws_inst.set_column('A:A', 4)
    ws_inst.set_column('B:G', 16)
    ws_inst.set_column('H:H', 24)

    # Title Block
    ws_inst.set_row(1, 35)
    ws_inst.merge_range('B2:H2', '“Construction Estimate Template Excel | Contractor Bid & Job Costing Spreadsheet, Markup Margin Calculator, Proposal Generator Google Sheets”', fmt_title_banner)
    ws_inst.set_row(2, 22)
    ws_inst.merge_range('B3:H3', '  Template Author: novality store  |  Formula Protection Password: premium  |  Excel & Google Sheets Compatible', fmt_sub_banner)

    # Overview Block
    ws_inst.merge_range('B5:H5', 'TEMPLATE OVERVIEW & KEY SELLING POINTS', fmt_section_banner)
    ws_inst.set_row(5, 45)
    intro_text = (
        "Welcome to the novality store Construction Estimate & Contractor Bid Template! This commercial-grade spreadsheet "
        "is engineered for general contractors, remodelers, builders, electricians, plumbers, roofers, painters, and specialty trades. "
        "It features automatic cost calculations, itemized trade breakdown sheets, an advanced markup & margin calculator, job costing "
        "variance tracking, and a printable client proposal generator."
    )
    fmt_wrap_lock = wb.add_format({'font_name': font_family, 'font_size': 10, 'text_wrap': True, 'valign': 'top', 'border': 1, 'border_color': '#CBD5E1', 'locked': 1})
    ws_inst.merge_range('B6:H6', intro_text, fmt_wrap_lock)

    # Color Legend Table
    ws_inst.merge_range('B8:D8', 'CELL COLOR & INPUT LEGEND', fmt_header_left)
    ws_inst.write('B9', 'White Cells with Border', fmt_input_center)
    ws_inst.merge_range('C9:D9', 'Editable User Input Cells (Data entry, quantities, prices, notes)', fmt_lock_text)
    ws_inst.write('B10', 'Soft Blue / Gray Cells', fmt_lock_center)
    ws_inst.merge_range('C10:D10', 'Locked Formula & Read-Only Cells (Protected with password: premium)', fmt_lock_text)
    ws_inst.write('B11', 'Navy / Slate / Gold Banners', fmt_header)
    ws_inst.merge_range('C11:D11', 'Section Headers, Table Banners & Accounting Grand Totals', fmt_lock_text)

    # Step-by-step instructions
    ws_inst.merge_range('B13:H13', 'STEP-BY-STEP USER WORKFLOW', fmt_section_banner)

    steps = [
        ("Step 1: Project Info & Global Rates", "Go to the 'Project Info' tab. Enter client details, contractor company info, and global rates (Sales Tax %, Discount %, Overhead %, Contingency %, and default trade Markups). These rates link automatically across the entire workbook."),
        ("Step 2: Enter Itemized Costs", "Use the 'Materials', 'Labor', 'Equipment', and 'Subcontractors' tabs to enter detailed line items for each trade. Quantities, unit costs, and hours will automatically calculate line totals and sheet subtotals."),
        ("Step 3: Build & Review Estimate", "Open the 'Estimate' tab. This master sheet compiles your itemized trade costs, applies Overhead and Contingency allowances, adds your calculated Markup, and computes the final Contract Bid Price."),
        ("Step 4: Configure Markup & Margin", "Visit the 'Markup & Margin Calculator' tab to fine-tune separate markup percentages for Materials, Labor, Equipment, and Subcontractors. Check the comparison table to see equivalent profit margins or use the Target Margin Reverse Calculator."),
        ("Step 5: Generate Client Proposal", "Switch to the 'Client Proposal' tab. Your project scope, contractor details, and pricing summary are auto-populated into a clean, print-ready, invoice-style proposal page ready for client signature."),
        ("Step 6: Track Actual Job Costs", "During construction, log your actual field costs on the 'Job Costing & Actuals' tab. Track budget variances in real time and view profitability performance on the executive 'Summary Dashboard'.")
    ]
    row_idx = 14
    for title, desc in steps:
        ws_inst.set_row(row_idx, 22)
        ws_inst.merge_range(row_idx, 1, row_idx, 7, f"  {title}", fmt_sub_banner)
        row_idx += 1
        ws_inst.set_row(row_idx, 32)
        ws_inst.merge_range(row_idx, 1, row_idx, 7, desc, fmt_wrap_lock)
        row_idx += 1

    ws_inst.protect('premium', prot_options)

    # =========================================================================
    # TAB 2: PROJECT INFO
    # =========================================================================
    ws_proj = wb.add_worksheet('Project Info')
    ws_proj.set_tab_color('#1B365D')
    ws_proj.set_column('A:A', 3)
    ws_proj.set_column('B:B', 24)
    ws_proj.set_column('C:C', 38)
    ws_proj.set_column('D:D', 4)
    ws_proj.set_column('E:E', 28)
    ws_proj.set_column('F:F', 18)

    ws_proj.set_row(1, 30)
    ws_proj.merge_range('B2:F2', 'PROJECT, CLIENT & CONTRACTOR SETUP (GLOBAL PARAMETERS)', fmt_title_banner)

    # Left Column: Project & Client Details
    ws_proj.merge_range('B4:C4', '1. PROJECT & CLIENT DETAILS', fmt_section_banner)
    project_fields = [
        ('Project Name', 'Luxury Kitchen & Living Area Remodel'),
        ('Client / Owner Name', 'Robert Sterling / Apex Properties'),
        ('Jobsite Location', '742 Evergreen Terrace, Suite 300'),
        ('Estimate Number', 'EST-2026-042'),
        ('Estimate Date', '2026-07-31'),
        ('Proposal Expiration Date', '2026-08-30'),
        ('Scope of Work Summary', 'Full kitchen demolition, custom cabinetry, quartz countertops, electrical & plumbing relocation, hardwood flooring, and ambient LED lighting.')
    ]
    for i, (label, val) in enumerate(project_fields):
        r = 4 + i
        ws_proj.write(r, 1, label, fmt_lock_text)
        if label == 'Scope of Work Summary':
            ws_proj.set_row(r, 45)
            fmt_input_wrap = wb.add_format({'font_name': font_family, 'font_size': 10, 'bg_color': '#FFFFFF', 'text_wrap': True, 'valign': 'top', 'border': 1, 'border_color': '#CBD5E1', 'locked': 0})
            ws_proj.write(r, 2, val, fmt_input_wrap)
        else:
            ws_proj.write(r, 2, val, fmt_input_text)

    # Left Column: Contractor Details
    ws_proj.merge_range('B13:C13', '2. CONTRACTOR & COMPANY DETAILS', fmt_section_banner)
    contractor_fields = [
        ('Contractor / Company Name', 'Novality Construction & Design LLC'),
        ('Primary Contact / Estimator', 'David Novality, Master Contractor'),
        ('Company Business Address', '100 Professional Way, Suite 400'),
        ('Phone Number', '(555) 382-9100'),
        ('Email Address', 'estimating@novalitystore.com'),
        ('Contractor License #', 'LIC-GEN-884210'),
        ('Tax ID / EIN #', 'XX-XXX8842')
    ]
    for i, (label, val) in enumerate(contractor_fields):
        r = 13 + i
        ws_proj.write(r, 1, label, fmt_lock_text)
        ws_proj.write(r, 2, val, fmt_input_text)

    # Right Column: Financial & Tax Configuration
    ws_proj.merge_range('E4:F4', '3. FINANCIAL, TAX & CONTINGENCY SETTINGS', fmt_section_banner)
    fin_fields = [
        ('Sales Tax Rate (%)', 0.0825, fmt_input_percent),
        ('Preferred Client Discount Rate (%)', 0.0250, fmt_input_percent),
        ('Contingency Allowance Rate (%)', 0.0500, fmt_input_percent),
        ('General Overhead Rate (%)', 0.1000, fmt_input_percent),
        ('Target Profit Margin (%)', 0.2250, fmt_input_percent),
        ('Default Material Markup (%)', 0.2000, fmt_input_percent),
        ('Default Labor Markup (%)', 0.3500, fmt_input_percent),
        ('Default Subcontractor Markup (%)', 0.1500, fmt_input_percent),
        ('Default Equipment Markup (%)', 0.2500, fmt_input_percent)
    ]
    for i, (label, val, fmt_cell) in enumerate(fin_fields):
        r = 4 + i
        ws_proj.write(r, 4, label, fmt_lock_text)
        ws_proj.write(r, 5, val, fmt_cell)

    # Summary box of configured rates
    ws_proj.merge_range('E15:F15', 'GLOBAL RATE SUMMARY NOTE', fmt_sub_banner)
    ws_proj.set_row(15, 60)
    note_rates = (
        "These global percentages feed directly into the master Estimate and Markup Calculator sheets. "
        "Overhead (E8) and Contingency (E7) allowances are applied to Total Direct Costs before Markup. "
        "Sales Tax (E5) is applied after preferred discount deductions."
    )
    ws_proj.merge_range('E16:F18', note_rates, fmt_wrap_lock)

    ws_proj.protect('premium', prot_options)

    # =========================================================================
    # TAB 4: MATERIALS (We create trade sheets first so Estimate can link to them)
    # =========================================================================
    ws_mat = wb.add_worksheet('Materials')
    ws_mat.set_tab_color('#3B82F6')
    ws_mat.set_column('A:A', 10)
    ws_mat.set_column('B:B', 32)
    ws_mat.set_column('C:C', 18)
    ws_mat.set_column('D:D', 22)
    ws_mat.set_column('E:E', 12)
    ws_mat.set_column('F:F', 10)
    ws_mat.set_column('G:I', 16)

    ws_mat.set_row(1, 30)
    ws_mat.merge_range('A2:I2', 'ITEMIZED MATERIALS COST SHEET (DIRECT MATERIAL QUANTITY & PRICE BREAKDOWN)', fmt_title_banner)
    ws_mat.merge_range('A3:I3', '  Enter material SKUs, supplier vendors, quantities, and unit prices. Line totals calculate automatically.', fmt_sub_banner)

    mat_headers = ['Item SKU', 'Material Description & Specification', 'Trade Category', 'Supplier / Vendor', 'Quantity', 'Unit', 'Unit Cost ($)', 'Tax / Freight ($)', 'Total Cost ($)']
    ws_mat.set_row(4, 25)
    for col_idx, h in enumerate(mat_headers):
        ws_mat.write(4, col_idx, h, fmt_header)

    mat_items = [
        ('MAT-101', 'Custom Shaker Shading Shaker Base Cabinets', 'Cabinetry', 'Craftsman Cabinet Co.', 12, 'Set', 700.00, 140.00),
        ('MAT-102', 'Custom Shaker Wall & Pantry Cabinets', 'Cabinetry', 'Craftsman Cabinet Co.', 8, 'Set', 350.00, 70.00),
        ('MAT-103', 'Calacatta Gold White Quartz Countertop Slab', 'Countertops', 'Apex Stone & Tile', 60, 'SqFt', 75.00, 300.00),
        ('MAT-104', 'Engineered Oak Wide Plank Hardwood Flooring', 'Flooring', 'Prime Flooring Dist.', 450, 'SqFt', 6.50, 225.00),
        ('MAT-105', 'Hardwood Underlayment & Moisture Barrier', 'Flooring', 'Prime Flooring Dist.', 450, 'SqFt', 0.65, 7.50),
        ('MAT-106', 'Recessed 6-inch LED Dimmable Fixture Bundle', 'Electrical', 'Lumens Direct Wholesale', 16, 'Unit', 45.00, 30.00),
        ('MAT-107', 'Decorative Kitchen Pendant Lighting Array', 'Electrical', 'Lumens Direct Wholesale', 3, 'Unit', 220.00, 20.00),
        ('MAT-108', 'Architectural Interior Primer (5-Gal Pail)', 'Paint', 'Sherwin-Williams Pro', 4, 'Pail', 45.00, 20.00),
        ('MAT-109', 'Premium Satin Enamel Paint (5-Gal Pail)', 'Paint', 'Sherwin-Williams Pro', 6, 'Pail', 65.00, 30.00),
        ('MAT-110', 'High-End Kitchen Sink & Commercial Faucet', 'Plumbing', 'Ferguson Plumbing Sup.', 1, 'Set', 950.00, 50.00),
        ('MAT-111', 'Ceramic Subway Tile Backsplash Bundle', 'Tile', 'Apex Stone & Tile', 80, 'SqFt', 8.50, 40.00),
        ('MAT-112', 'Drywall, Compound & Trim Lumber Package', 'Lumber / Prep', 'Home Depot Pro', 1, 'Lot', 650.00, 50.00)
    ]

    for i, item in enumerate(mat_items):
        r = 5 + i
        ws_mat.write(r, 0, item[0], fmt_input_center)
        ws_mat.write(r, 1, item[1], fmt_input_text)
        ws_mat.write(r, 2, item[2], fmt_input_center)
        ws_mat.write(r, 3, item[3], fmt_input_text)
        ws_mat.write(r, 4, item[4], fmt_input_int)
        ws_mat.write(r, 5, item[5], fmt_input_center)
        ws_mat.write(r, 6, item[6], fmt_input_currency)
        ws_mat.write(r, 7, item[7], fmt_input_currency)
        ws_mat.write_formula(r, 8, f"=E{r+1}*G{r+1}+H{r+1}", fmt_lock_currency)

    # Total Materials Row
    tot_r_mat = 5 + len(mat_items)
    ws_mat.set_row(tot_r_mat, 25)
    ws_mat.merge_range(tot_r_mat, 0, tot_r_mat, 7, 'TOTAL DIRECT MATERIAL COST ($)', fmt_total_label)
    ws_mat.write_formula(tot_r_mat, 8, f"=SUM(I6:I{tot_r_mat})", fmt_total_currency)

    ws_mat.protect('premium', prot_options)

    # =========================================================================
    # TAB 5: LABOR
    # =========================================================================
    ws_lab = wb.add_worksheet('Labor')
    ws_lab.set_tab_color('#3B82F6')
    ws_lab.set_column('A:A', 10)
    ws_lab.set_column('B:B', 32)
    ws_lab.set_column('C:D', 14)
    ws_lab.set_column('E:G', 16)
    ws_lab.set_column('H:H', 18)

    ws_lab.set_row(1, 30)
    ws_lab.merge_range('A2:H2', 'ITEMIZED LABOR COST SHEET (CRAFT HOURS, REGULAR WAGES & OVERTIME)', fmt_title_banner)
    ws_lab.merge_range('A3:H3', '  Enter craft roles, crew size, standard hours, and hourly billing rates. Overtime is tracked separately.', fmt_sub_banner)

    lab_headers = ['Task Code', 'Craft Role & Task Description', 'Crew Size', 'Reg. Hours', 'Reg. Rate ($/hr)', 'OT Hours', 'OT Rate ($/hr)', 'Total Labor Cost ($)']
    ws_lab.set_row(4, 25)
    for col_idx, h in enumerate(lab_headers):
        ws_lab.write(4, col_idx, h, fmt_header)

    lab_items = [
        ('LAB-201', 'Selective Kitchen Demolition & Haul Prep', 2, 32, 40.00, 0, 60.00),
        ('LAB-202', 'Structural Header & Framing Modification', 2, 24, 55.00, 4, 82.50),
        ('LAB-203', 'Custom Cabinetry & Island Box Installation', 2, 40, 65.00, 0, 97.50),
        ('LAB-204', 'Finish Carpentry, Crown Molding & Trim', 1, 32, 60.00, 0, 90.00),
        ('LAB-205', 'Drywall Hanging, Taping & Sanding Prep', 2, 28, 45.00, 0, 67.50),
        ('LAB-206', 'Engineered Hardwood Floor Installation', 2, 24, 50.00, 0, 75.00),
        ('LAB-207', 'Ceramic Tile Backsplash Layout & Setting', 1, 16, 55.00, 0, 82.50),
        ('LAB-208', 'Interior Spray Primer & Architectural Paint', 2, 36, 45.00, 0, 67.50),
        ('LAB-209', 'Project Supervisor & Quality Control Audit', 1, 24, 75.00, 0, 112.50),
        ('LAB-210', 'Jobsite Cleanup, Dust Control & Sanitizing', 1, 16, 35.00, 0, 52.50)
    ]

    for i, item in enumerate(lab_items):
        r = 5 + i
        ws_lab.write(r, 0, item[0], fmt_input_center)
        ws_lab.write(r, 1, item[1], fmt_input_text)
        ws_lab.write(r, 2, item[2], fmt_input_int)
        ws_lab.write(r, 3, item[3], fmt_input_int)
        ws_lab.write(r, 4, item[4], fmt_input_currency)
        ws_lab.write(r, 5, item[5], fmt_input_int)
        ws_lab.write(r, 6, item[6], fmt_input_currency)
        # Total Labor Cost = (Crew * RegHours * RegRate) + (Crew * OTHours * OTRate)
        ws_lab.write_formula(r, 7, f"=(C{r+1}*D{r+1}*E{r+1})+(C{r+1}*F{r+1}*G{r+1})", fmt_lock_currency)

    tot_r_lab = 5 + len(lab_items)
    ws_lab.set_row(tot_r_lab, 25)
    ws_lab.merge_range(tot_r_lab, 0, tot_r_lab, 6, 'TOTAL DIRECT LABOR COST ($)', fmt_total_label)
    ws_lab.write_formula(tot_r_lab, 7, f"=SUM(H6:H{tot_r_lab})", fmt_total_currency)

    ws_lab.protect('premium', prot_options)

    # =========================================================================
    # TAB 6: EQUIPMENT
    # =========================================================================
    ws_eqp = wb.add_worksheet('Equipment')
    ws_eqp.set_tab_color('#3B82F6')
    ws_eqp.set_column('A:A', 10)
    ws_eqp.set_column('B:B', 32)
    ws_eqp.set_column('C:D', 18)
    ws_eqp.set_column('E:H', 16)

    ws_eqp.set_row(1, 30)
    ws_eqp.merge_range('A2:H2', 'EQUIPMENT RENTAL & SPECIAL TOOLING COST SHEET', fmt_title_banner)
    ws_eqp.merge_range('A3:H3', '  Track equipment rentals, daily/weekly rates, delivery fees, and owned tool allowances.', fmt_sub_banner)

    eqp_headers = ['Equip. ID', 'Equipment Name & Specification', 'Ownership Type', 'Rate Period', 'Rate ($)', 'Qty / Days', 'Delivery Fee ($)', 'Total Cost ($)']
    ws_eqp.set_row(4, 25)
    for col_idx, h in enumerate(eqp_headers):
        ws_eqp.write(4, col_idx, h, fmt_header)

    eqp_items = [
        ('EQP-301', '20-Yard Roll-Off Waste Dumpster Rental', 'Rental', 'Flat Rate', 650.00, 1, 200.00),
        ('EQP-302', 'HEPA Commercial Air Scrubber & Filtration', 'Owned / Allowance', 'Daily Rate', 45.00, 10, 50.00),
        ('EQP-303', 'Commercial Wet Tile Saw & Diamond Blade', 'Owned / Allowance', 'Daily Rate', 35.00, 4, 0.00),
        ('EQP-304', 'Hardwood Floor Drum Sander & Edger Kit', 'Rental', 'Daily Rate', 85.00, 3, 50.00),
        ('EQP-305', 'Interior Aluminum Scaffolding & Work Lift', 'Rental', 'Weekly Rate', 250.00, 2, 75.00),
        ('EQP-306', 'Temporary Site Distribution Electrical Box', 'Owned / Allowance', 'Flat Rate', 120.00, 1, 0.00),
        ('EQP-307', 'High-Capacity Dehumidifier & Blower Array', 'Rental', 'Daily Rate', 40.00, 5, 25.00),
        ('EQP-308', 'Material Hoist & Cabinet Dolly Package', 'Owned / Allowance', 'Flat Rate', 150.00, 1, 0.00)
    ]

    for i, item in enumerate(eqp_items):
        r = 5 + i
        ws_eqp.write(r, 0, item[0], fmt_input_center)
        ws_eqp.write(r, 1, item[1], fmt_input_text)
        ws_eqp.write(r, 2, item[2], fmt_input_center)
        ws_eqp.write(r, 3, item[3], fmt_input_center)
        ws_eqp.write(r, 4, item[4], fmt_input_currency)
        ws_eqp.write(r, 5, item[5], fmt_input_int)
        ws_eqp.write(r, 6, item[6], fmt_input_currency)
        # Total Equipment Cost = Rate * Days + DeliveryFee
        ws_eqp.write_formula(r, 7, f"=E{r+1}*F{r+1}+G{r+1}", fmt_lock_currency)

    tot_r_eqp = 5 + len(eqp_items)
    ws_eqp.set_row(tot_r_eqp, 25)
    ws_eqp.merge_range(tot_r_eqp, 0, tot_r_eqp, 6, 'TOTAL DIRECT EQUIPMENT & RENTAL COST ($)', fmt_total_label)
    ws_eqp.write_formula(tot_r_eqp, 7, f"=SUM(H6:H{tot_r_eqp})", fmt_total_currency)

    ws_eqp.protect('premium', prot_options)

    # =========================================================================
    # TAB 7: SUBCONTRACTORS
    # =========================================================================
    ws_sub = wb.add_worksheet('Subcontractors')
    ws_sub.set_tab_color('#3B82F6')
    ws_sub.set_column('A:A', 10)
    ws_sub.set_column('B:B', 28)
    ws_sub.set_column('C:C', 32)
    ws_sub.set_column('D:G', 16)
    ws_sub.set_column('H:H', 18)

    ws_sub.set_row(1, 30)
    ws_sub.merge_range('A2:H2', 'LICENSED SUBCONTRACTOR PACKAGE & QUOTE SHEET', fmt_title_banner)
    ws_sub.merge_range('A3:H3', '  Manage subcontractor bids, verify license/insurance, and calculate markup amounts per package.', fmt_sub_banner)

    sub_headers = ['Sub. Code', 'Subcontractor Company', 'Scope of Work & Package Description', 'Quoted Cost ($)', 'Markup %', 'Markup Amt ($)', 'Billed Amount ($)', 'License Verified']
    ws_sub.set_row(4, 25)
    for col_idx, h in enumerate(sub_headers):
        ws_sub.write(4, col_idx, h, fmt_header)

    sub_items = [
        ('SUB-401', 'Spark Systems Electric LLC', '200A Service Panel Upgrade & LED Wiring Rough/Finish', 4200.00, 0.15, 'Yes - Verified'),
        ('SUB-402', 'Apex Mechanical Plumbing Co.', 'Kitchen Sink Reroute, Gas Line Extension & Appliance Hookup', 3800.00, 0.15, 'Yes - Verified'),
        ('SUB-403', 'Breeze Comfort HVAC LLC', 'Range Hood Duct Reroute & Supply Register Diffusers', 1750.00, 0.15, 'Yes - Verified')
    ]

    for i, item in enumerate(sub_items):
        r = 5 + i
        ws_sub.write(r, 0, item[0], fmt_input_center)
        ws_sub.write(r, 1, item[1], fmt_input_text)
        ws_sub.write(r, 2, item[2], fmt_input_text)
        ws_sub.write(r, 3, item[3], fmt_input_currency)
        ws_sub.write(r, 4, item[4], fmt_input_percent)
        ws_sub.write_formula(r, 5, f"=D{r+1}*E{r+1}", fmt_lock_currency)
        ws_sub.write_formula(r, 6, f"=D{r+1}+F{r+1}", fmt_lock_currency)
        ws_sub.write(r, 7, item[5], fmt_input_center)

    tot_r_sub = 5 + len(sub_items)
    ws_sub.set_row(tot_r_sub, 25)
    ws_sub.merge_range(tot_r_sub, 0, tot_r_sub, 2, 'TOTAL SUBCONTRACTOR PACKAGE COSTS ($)', fmt_total_label)
    ws_sub.write_formula(tot_r_sub, 3, f"=SUM(D6:D{tot_r_sub})", fmt_total_currency)
    ws_sub.write_formula(tot_r_sub, 4, "", fmt_total_label)
    ws_sub.write_formula(tot_r_sub, 5, f"=SUM(F6:F{tot_r_sub})", fmt_total_currency)
    ws_sub.write_formula(tot_r_sub, 6, f"=SUM(G6:G{tot_r_sub})", fmt_total_currency)
    ws_sub.write(tot_r_sub, 7, "", fmt_total_label)

    ws_sub.protect('premium', prot_options)

    # =========================================================================
    # TAB 3: ESTIMATE (Master Estimate Sheet & Contractor Bid Template)
    # =========================================================================
    ws_est = wb.add_worksheet('Estimate')
    ws_est.set_tab_color('#2C3E50')
    ws_est.set_paper(1)  # Letter size
    ws_est.fit_to_pages(1, 0)
    ws_est.set_margins(0.5, 0.5, 0.75, 0.75)
    ws_est.set_footer('&L&"Segoe UI"&8Template Author: novality store&R&"Segoe UI"&8Page &P of &N')

    ws_est.set_column('A:A', 8)
    ws_est.set_column('B:B', 34)
    ws_est.set_column('C:C', 12)
    ws_est.set_column('D:D', 10)
    ws_est.set_column('E:G', 16)

    # Title Header Block
    ws_est.set_row(1, 35)
    ws_est.merge_range('A2:G2', 'CONSTRUCTION ESTIMATE & CONTRACTOR BID TEMPLATE', fmt_title_banner)
    ws_est.merge_range('A3:G3', '  Dynamic bid sheet compiling itemized costs, overhead, contingency, and markup.', fmt_sub_banner)

    # Project Info Summary Table (Linked to Project Info tab)
    ws_est.merge_range('A5:C5', 'PROJECT & CLIENT INFORMATION', fmt_section_banner)
    ws_est.merge_range('E5:G5', 'BID ESTIMATE & CONTRACTOR DETAILS', fmt_section_banner)

    ws_est.write('A6', 'Project Name:', fmt_lock_text)
    ws_est.merge_range('B6:C6', "='Project Info'!B4", fmt_lock_text)
    ws_est.write('E6', 'Estimate Number:', fmt_lock_text)
    ws_est.merge_range('F6:G6', "='Project Info'!B7", fmt_lock_text)

    ws_est.write('A7', 'Client / Owner:', fmt_lock_text)
    ws_est.merge_range('B7:C7', "='Project Info'!B5", fmt_lock_text)
    ws_est.write('E7', 'Estimate Date:', fmt_lock_text)
    ws_est.merge_range('F7:G7', "='Project Info'!B8", fmt_lock_text)

    ws_est.write('A8', 'Job Location:', fmt_lock_text)
    ws_est.merge_range('B8:C8', "='Project Info'!B6", fmt_lock_text)
    ws_est.write('E8', 'Contractor:', fmt_lock_text)
    ws_est.merge_range('F8:G8', "='Project Info'!B13", fmt_lock_text)

    ws_est.write('A9', 'Scope Summary:', fmt_lock_text)
    ws_est.merge_range('B9:G9', "='Project Info'!B10", fmt_lock_text)

    # Table Header
    est_headers = ['Item #', 'Cost Category & Item Specification', 'Quantity', 'Unit', 'Unit Cost ($)', 'Total Cost ($)', 'Trade Source']
    ws_est.set_row(11, 25)
    for col_idx, h in enumerate(est_headers):
        ws_est.write(11, col_idx, h, fmt_header)

    # Section 1: Direct Materials (Linked to Materials total or breakdown)
    ws_est.merge_range('A13:G13', '1. DIRECT MATERIALS & FIXTURE PACKAGES', fmt_section_banner)
    mat_rows = [
        ('1.01', 'Custom Cabinetry & Millwork Package', 1, 'Lot', 11200.00, 'Materials Sheet'),
        ('1.02', 'Calacatta Gold Quartz Countertops & Backsplash', 1, 'Lot', 4800.00, 'Materials Sheet'),
        ('1.03', 'Engineered Hardwood Flooring & Moisture Barrier', 1, 'Lot', 3450.00, 'Materials Sheet'),
        ('1.04', 'Recessed & Decorative LED Lighting Bundle', 1, 'Lot', 1400.00, 'Materials Sheet'),
        ('1.05', 'Architectural Interior Primers & Enamel Paints', 1, 'Lot', 800.00, 'Materials Sheet')
    ]
    for i, (num, desc, qty, unit, cost, src) in enumerate(mat_rows):
        r = 13 + i
        ws_est.write(r, 0, num, fmt_input_center)
        ws_est.write(r, 1, desc, fmt_input_text)
        ws_est.write(r, 2, qty, fmt_input_int)
        ws_est.write(r, 3, unit, fmt_input_center)
        ws_est.write(r, 4, cost, fmt_input_currency)
        ws_est.write_formula(r, 5, f"=C{r+1}*E{r+1}", fmt_lock_currency)
        ws_est.write(r, 6, src, fmt_lock_center)
    ws_est.merge_range('A19:E19', 'SUBTOTAL DIRECT MATERIALS ($)', fmt_total_label)
    ws_est.write_formula('F19', '=SUM(F14:F18)', fmt_total_currency)
    ws_est.write('G19', "='Materials'!I18", fmt_lock_currency)

    # Section 2: Direct Labor
    ws_est.merge_range('A21:G21', '2. DIRECT LABOR & CRAFTSMANSHIP HOURS', fmt_section_banner)
    lab_rows = [
        ('2.01', 'Selective Kitchen Demolition, Haul & Structural Prep', 1, 'Lot', 3200.00, 'Labor Sheet'),
        ('2.02', 'Master Cabinetry, Millwork & Island Installation', 1, 'Lot', 6400.00, 'Labor Sheet'),
        ('2.03', 'Drywall Framing, Hanging, Taping & Sanding', 1, 'Lot', 3600.00, 'Labor Sheet'),
        ('2.04', 'Precision Interior Spray Painting & Enamel Trim Finish', 1, 'Lot', 3600.00, 'Labor Sheet')
    ]
    for i, (num, desc, qty, unit, cost, src) in enumerate(lab_rows):
        r = 21 + i
        ws_est.write(r, 0, num, fmt_input_center)
        ws_est.write(r, 1, desc, fmt_input_text)
        ws_est.write(r, 2, qty, fmt_input_int)
        ws_est.write(r, 3, unit, fmt_input_center)
        ws_est.write(r, 4, cost, fmt_input_currency)
        ws_est.write_formula(r, 5, f"=C{r+1}*E{r+1}", fmt_lock_currency)
        ws_est.write(r, 6, src, fmt_lock_center)
    ws_est.merge_range('A26:E26', 'SUBTOTAL DIRECT LABOR ($)', fmt_total_label)
    ws_est.write_formula('F26', '=SUM(F22:F25)', fmt_total_currency)
    ws_est.write('G26', "='Labor'!H16", fmt_lock_currency)

    # Section 3: Licensed Subcontractors
    ws_est.merge_range('A28:G28', '3. LICENSED SUBCONTRACTOR SPECIALTY SERVICES', fmt_section_banner)
    sub_rows = [
        ('3.01', 'Licensed Electrical Wiring & Service Panel Upgrade', 1, 'Pkg', 4200.00, 'Subcontractors'),
        ('3.02', 'Licensed Plumbing Gas/Water Reroute & Fixtures', 1, 'Pkg', 3800.00, 'Subcontractors'),
        ('3.03', 'Licensed HVAC Duct Modification & Range Hood Exhaust', 1, 'Pkg', 1750.00, 'Subcontractors')
    ]
    for i, (num, desc, qty, unit, cost, src) in enumerate(sub_rows):
        r = 28 + i
        ws_est.write(r, 0, num, fmt_input_center)
        ws_est.write(r, 1, desc, fmt_input_text)
        ws_est.write(r, 2, qty, fmt_input_int)
        ws_est.write(r, 3, unit, fmt_input_center)
        ws_est.write(r, 4, cost, fmt_input_currency)
        ws_est.write_formula(r, 5, f"=C{r+1}*E{r+1}", fmt_lock_currency)
        ws_est.write(r, 6, src, fmt_lock_center)
    ws_est.merge_range('A32:E32', 'SUBTOTAL SUBCONTRACTORS ($)', fmt_total_label)
    ws_est.write_formula('F32', '=SUM(F29:F31)', fmt_total_currency)
    ws_est.write('G32', "='Subcontractors'!D9", fmt_lock_currency)

    # Section 4: Equipment & Rentals
    ws_est.merge_range('A34:G34', '4. EQUIPMENT RENTALS & SPECIAL TOOLING ALLOWANCE', fmt_section_banner)
    eqp_rows = [
        ('4.01', '20-Yard Roll-Off Waste Dumpster Rental & Removal', 1, 'Lot', 850.00, 'Equipment Sheet'),
        ('4.02', 'HEPA Air Scrubber, Dust Containment & Dehumidifiers', 1, 'Lot', 750.00, 'Equipment Sheet'),
        ('4.03', 'Scaffolding Towers, Lifts & Specialty Flooring Sanders', 1, 'Lot', 1250.00, 'Equipment Sheet')
    ]
    for i, (num, desc, qty, unit, cost, src) in enumerate(eqp_rows):
        r = 34 + i
        ws_est.write(r, 0, num, fmt_input_center)
        ws_est.write(r, 1, desc, fmt_input_text)
        ws_est.write(r, 2, qty, fmt_input_int)
        ws_est.write(r, 3, unit, fmt_input_center)
        ws_est.write(r, 4, cost, fmt_input_currency)
        ws_est.write_formula(r, 5, f"=C{r+1}*E{r+1}", fmt_lock_currency)
        ws_est.write(r, 6, src, fmt_lock_center)
    ws_est.merge_range('A38:E38', 'SUBTOTAL EQUIPMENT & RENTALS ($)', fmt_total_label)
    ws_est.write_formula('F38', '=SUM(F35:F37)', fmt_total_currency)
    ws_est.write('G38', "='Equipment'!H14", fmt_lock_currency)

    # Section 5: Permits & Other Costs
    ws_est.merge_range('A40:G40', '5. PERMITS, FEES & SITE PROTECTION ALLOWANCE', fmt_section_banner)
    oth_rows = [
        ('5.01', 'Municipal Building, Trade & Electrical Permit Fees', 1, 'Lot', 950.00, 'Municipal'),
        ('5.02', 'Site Floor Protection, Dust Barriers & Sanitization', 1, 'Lot', 500.00, 'Jobsite Prep')
    ]
    for i, (num, desc, qty, unit, cost, src) in enumerate(oth_rows):
        r = 40 + i
        ws_est.write(r, 0, num, fmt_input_center)
        ws_est.write(r, 1, desc, fmt_input_text)
        ws_est.write(r, 2, qty, fmt_input_int)
        ws_est.write(r, 3, unit, fmt_input_center)
        ws_est.write(r, 4, cost, fmt_input_currency)
        ws_est.write_formula(r, 5, f"=C{r+1}*E{r+1}", fmt_lock_currency)
        ws_est.write(r, 6, src, fmt_lock_center)
    ws_est.merge_range('A43:E43', 'SUBTOTAL PERMITS & OTHER COSTS ($)', fmt_total_label)
    ws_est.write_formula('F43', '=SUM(F41:F42)', fmt_total_currency)
    ws_est.write('G43', '-', fmt_lock_center)

    # Contractor Bid Pricing Summary Table (Rows 45-56)
    ws_est.merge_range('A45:G45', '6. CONTRACTOR BID PRICING SUMMARY & CONTRACT BID TOTALS', fmt_header_amber)
    
    bid_summary_rows = [
        ('Total Direct Job Costs (Sum of Sections 1 to 5)', '=F19+F26+F32+F38+F43', fmt_total_currency, fmt_total_label),
        ('General Overhead Allowance (10.00% of Direct Costs)', "='Project Info'!F7*F46", fmt_lock_currency, fmt_lock_text),
        ('Contingency Allowance (5.00% of Direct Costs)', "='Project Info'!F6*F46", fmt_lock_currency, fmt_lock_text),
        ('Subtotal Total Estimated Job Cost (Before Markup)', '=F46+F47+F48', fmt_total_currency, fmt_total_label),
        ('Total Professional Markup & Contractor Gross Profit', "='Markup & Margin Calculator'!D13", fmt_lock_currency, fmt_lock_text),
        ('Estimated Selling Price (Contract Bid Subtotal)', '=F49+F50', fmt_total_currency, fmt_total_label),
        ('Preferred Client Contract Discount (-2.50%)', "=-F51*'Project Info'!F5", fmt_lock_currency, fmt_lock_text),
        ('Net Taxable Bid Contract Amount', '=F51+F52', fmt_total_currency, fmt_total_label),
        ('Estimated Sales Tax (8.25% on Net Taxable Amount)', "=F53*'Project Info'!F4", fmt_lock_currency, fmt_lock_text),
        ('GRAND TOTAL CONTRACT BID PRICE ($)', '=F53+F54', fmt_grand_currency, fmt_grand_label)
    ]

    for i, (label, form, f_val, f_lbl) in enumerate(bid_summary_rows):
        r = 45 + i
        ws_est.set_row(r, 24 if i == len(bid_summary_rows)-1 else 20)
        ws_est.merge_range(r, 0, r, 4, label, f_lbl)
        ws_est.write_formula(r, 5, form, f_val)
        ws_est.write(r, 6, '', f_lbl)

    # Terms & Acceptance Area
    ws_est.merge_range('A57:G57', '7. CONTRACTOR BID ACCEPTANCE & SIGN-OFF AREA', fmt_section_banner)
    ws_est.set_row(58, 45)
    terms_text = (
        "NOTES & ASSUMPTIONS: Bid price is valid for 30 days from the estimate date. "
        "Any deviations, structural modifications, or owner-requested scope changes will be executed only upon signed Change Order. "
        "Includes standard 1-year workmanship warranty from novality store."
    )
    ws_est.merge_range('A59:G59', terms_text, fmt_note_text)
    
    ws_est.write('A61', 'Client Signature:', fmt_lock_text)
    ws_est.merge_range('B61:C61', '', fmt_input_text)
    ws_est.write('E61', 'Date Accepted:', fmt_lock_text)
    ws_est.merge_range('F61:G61', '', fmt_input_date)

    ws_est.write('A62', 'Contractor Representative:', fmt_lock_text)
    ws_est.merge_range('B62:C62', 'David Novality, Novality Construction', fmt_lock_text)
    ws_est.write('E62', 'Date Signed:', fmt_lock_text)
    ws_est.merge_range('F62:G62', '2026-07-31', fmt_lock_center)

    ws_est.protect('premium', prot_options)

    # =========================================================================
    # TAB 8: MARKUP & MARGIN CALCULATOR
    # =========================================================================
    ws_mm = wb.add_worksheet('Markup & Margin Calculator')
    ws_mm.set_tab_color('#D97706')
    ws_mm.set_column('A:A', 26)
    ws_mm.set_column('B:G', 16)
    ws_mm.set_column('H:H', 22)

    ws_mm.set_row(1, 30)
    ws_mm.merge_range('A2:G2', 'ADVANCED MARKUP & PROFIT MARGIN CALCULATOR', fmt_title_banner)
    ws_mm.merge_range('A3:G3', '  Separate category markups, Markup vs. Margin comparison table, and Target Profit Margin reverse calculator.', fmt_sub_banner)

    # Section 1: Itemized Markup by Category
    ws_mm.merge_range('A5:G5', '1. CATEGORY-SPECIFIC MARKUP & PROFIT MARGIN CALCULATOR', fmt_header_amber)
    mm_headers = ['Cost Category', 'Base Cost ($)', 'Markup %', 'Markup Amt ($)', 'Selling Price ($)', 'Category Share', 'Profit Margin %']
    ws_mm.set_row(5, 25)
    for col_idx, h in enumerate(mm_headers):
        ws_mm.write(5, col_idx, h, fmt_header)

    mm_rows = [
        ('Direct Materials', "='Estimate'!F19", "='Project Info'!F9"),
        ('Direct Labor', "='Estimate'!F26", "='Project Info'!F10"),
        ('Equipment & Rentals', "='Estimate'!F38", "='Project Info'!F12"),
        ('Subcontractors', "='Estimate'!F32", "='Project Info'!F11"),
        ('Permits & Other Costs', "='Estimate'!F43", "0.1500"),
        ('Overhead & Contingency', "='Estimate'!F47+'Estimate'!F48", "0.2000")
    ]
    for i, (cat, f_cost, f_mkp) in enumerate(mm_rows):
        r = 6 + i
        ws_mm.write(r, 0, cat, fmt_lock_text)
        ws_mm.write_formula(r, 1, f_cost, fmt_lock_currency)
        if f_mkp.startswith('='):
            ws_mm.write_formula(r, 2, f_mkp, fmt_input_percent)
        else:
            ws_mm.write(r, 2, float(f_mkp), fmt_input_percent)
        ws_mm.write_formula(r, 3, f"=B{r+1}*C{r+1}", fmt_lock_currency)
        ws_mm.write_formula(r, 4, f"=B{r+1}+D{r+1}", fmt_lock_currency)
        ws_mm.write_formula(r, 5, f"=E{r+1}/$E$13", fmt_lock_percent)
        ws_mm.write_formula(r, 6, f"=IFERROR(D{r+1}/E{r+1}, 0)", fmt_lock_percent)

    # Total row for Section 1 (Row 12 -> index 12 -> row 13 in Excel)
    ws_mm.set_row(12, 25)
    ws_mm.write(12, 0, 'TOTAL JOB MARKUP & MARGIN', fmt_total_label)
    ws_mm.write_formula(12, 1, '=SUM(B7:B12)', fmt_total_currency)
    ws_mm.write_formula(12, 2, '=D13/B13', fmt_total_percent)
    ws_mm.write_formula(12, 3, '=SUM(D7:D12)', fmt_total_currency)
    ws_mm.write_formula(12, 4, '=SUM(E7:E12)', fmt_total_currency)
    ws_mm.write_formula(12, 5, '=SUM(F7:F12)', fmt_total_percent)
    ws_mm.write_formula(12, 6, '=IFERROR(D13/E13, 0)', fmt_total_percent)

    # Section 2: Markup vs. Profit Margin Comparison Table
    ws_mm.merge_range('A15:D15', '2. MARKUP VS. PROFIT MARGIN COMPARISON TABLE', fmt_section_banner)
    ws_mm.write(15, 0, 'Markup %', fmt_header)
    ws_mm.write(15, 1, 'Equivalent Margin %', fmt_header)
    ws_mm.write(15, 2, 'Price on $1,000 Cost ($)', fmt_header)
    ws_mm.write(15, 3, 'Gross Profit ($)', fmt_header)

    sample_markups = [0.10, 0.15, 0.20, 0.25, 0.30, 0.35, 0.40, 0.50, 0.75, 1.00]
    for i, mkp in enumerate(sample_markups):
        r = 16 + i
        ws_mm.write(r, 0, mkp, fmt_lock_percent)
        ws_mm.write_formula(r, 1, f"=A{r+1}/(1+A{r+1})", fmt_lock_percent)
        ws_mm.write_formula(r, 2, f"=1000*(1+A{r+1})", fmt_lock_currency)
        ws_mm.write_formula(r, 3, f"=C{r+1}-1000", fmt_lock_currency)

    # Section 3: Target Profit Margin Reverse Calculator
    ws_mm.merge_range('E15:G15', '3. TARGET MARGIN REVERSE CALCULATOR', fmt_section_banner)
    target_fields = [
        ('Input Total Job Cost ($):', '=B13', fmt_input_currency, 'Cost to price'),
        ('Desired Profit Margin (%):', "='Project Info'!F8", fmt_input_percent, 'Target %'),
        ('Required Selling Price ($):', '=F17/(1-F18)', fmt_total_currency, 'Formula'),
        ('Required Gross Profit ($):', '=F19-F17', fmt_lock_currency, 'Profit $'),
        ('Equivalent Markup % Needed:', '=IFERROR(F20/F17, 0)', fmt_total_percent, 'Equivalent')
    ]
    for i, (label, val, f_cell, note) in enumerate(target_fields):
        r = 16 + i
        ws_mm.write(r, 4, label, fmt_lock_text)
        if val.startswith('='):
            ws_mm.write_formula(r, 5, val, f_cell)
        else:
            ws_mm.write(r, 5, val, f_cell)
        ws_mm.write(r, 6, note, fmt_lock_center)

    ws_mm.protect('premium', prot_options)

    # =========================================================================
    # TAB 9: JOB COSTING & ACTUAL COSTS
    # =========================================================================
    ws_jc = wb.add_worksheet('Job Costing & Actuals')
    ws_jc.set_tab_color('#D97706')
    ws_jc.set_column('A:A', 10)
    ws_jc.set_column('B:B', 32)
    ws_jc.set_column('C:E', 18)
    ws_jc.set_column('F:F', 16)
    ws_jc.set_column('G:G', 20)

    ws_jc.set_row(1, 30)
    ws_jc.merge_range('A2:G2', 'ESTIMATED VS. ACTUAL JOB COSTING TRACKER & VARIANCE ANALYSIS', fmt_title_banner)
    ws_jc.merge_range('A3:G3', '  Track field expenditures against estimated budget. Automatic budget variance and status indicator.', fmt_sub_banner)

    jc_headers = ['Category Code', 'Cost Category Description', 'Estimated Budget ($)', 'Actual Cost Incurred ($)', 'Dollar Variance ($)', 'Budget Variance %', 'Cost Performance Status']
    ws_jc.set_row(4, 25)
    for col_idx, h in enumerate(jc_headers):
        ws_jc.write(4, col_idx, h, fmt_header)

    jc_items = [
        ('CAT-100', 'Direct Materials & Fixtures', "='Estimate'!F19", 20950.00),
        ('CAT-200', 'Direct Labor & Craft Hours', "='Estimate'!F26", 17200.00),
        ('CAT-300', 'Licensed Subcontractors', "='Estimate'!F32", 9750.00),
        ('CAT-400', 'Equipment Rentals & Special Tools', "='Estimate'!F38", 2650.00),
        ('CAT-500', 'Permits, Fees & Site Protection', "='Estimate'!F43", 1450.00),
        ('CAT-600', 'General Overhead & Contingency Allowance', "='Estimate'!F47+'Estimate'!F48", 7500.00)
    ]
    for i, (code, desc, f_est, act) in enumerate(jc_items):
        r = 5 + i
        ws_jc.write(r, 0, code, fmt_lock_center)
        ws_jc.write(r, 1, desc, fmt_lock_text)
        ws_jc.write_formula(r, 2, f_est, fmt_lock_currency)
        ws_jc.write(r, 3, act, fmt_input_currency)
        # Dollar Variance = Estimated - Actual (Positive means under budget)
        ws_jc.write_formula(r, 4, f"=C{r+1}-D{r+1}", fmt_lock_currency)
        ws_jc.write_formula(r, 5, f"=IFERROR(E{r+1}/C{r+1}, 0)", fmt_lock_percent)
        ws_jc.write_formula(r, 6, f'=IF(E{r+1}>=0, "UNDER BUDGET", "OVER BUDGET")', fmt_lock_center)

    tot_r_jc = 5 + len(jc_items)
    ws_jc.set_row(tot_r_jc, 25)
    ws_jc.merge_range(tot_r_jc, 0, tot_r_jc, 1, 'TOTAL CONSTRUCTION JOB COST ($)', fmt_total_label)
    ws_jc.write_formula(tot_r_jc, 2, f"=SUM(C6:C{tot_r_jc})", fmt_total_currency)
    ws_jc.write_formula(tot_r_jc, 3, f"=SUM(D6:D{tot_r_jc})", fmt_total_currency)
    ws_jc.write_formula(tot_r_jc, 4, f"=C{tot_r_jc+1}-D{tot_r_jc+1}", fmt_total_currency)
    ws_jc.write_formula(tot_r_jc, 5, f"=IFERROR(E{tot_r_jc+1}/C{tot_r_jc+1}, 0)", fmt_total_percent)
    ws_jc.write_formula(tot_r_jc, 6, f'=IF(E{tot_r_jc+1}>=0, "UNDER BUDGET", "OVER BUDGET")', fmt_total_label)

    # Executive Profitability Summary Block
    ws_jc.merge_range('B14:F14', 'EXECUTIVE JOB PROFITABILITY & MARGIN VARIANCE SUMMARY', fmt_section_banner)
    summary_fields = [
        ('Contract Bid Selling Price ($)', "='Estimate'!F51", fmt_lock_currency),
        ('Actual Total Job Cost Incurred ($)', f"=D{tot_r_jc+1}", fmt_lock_currency),
        ('Actual Job Gross Profit ($)', '=C15-C16', fmt_total_currency),
        ('Actual Realized Profit Margin (%)', '=IFERROR(C17/C15, 0)', fmt_total_percent),
        ('Estimated Profit Margin (%)', "='Markup & Margin Calculator'!G13", fmt_lock_percent),
        ('Profit Margin Variance (%)', '=C18-C19', fmt_total_percent)
    ]
    for i, (label, f_val, f_fmt) in enumerate(summary_fields):
        r = 14 + i
        ws_jc.write(r, 1, label, fmt_lock_text)
        ws_jc.write_formula(r, 2, f_val, f_fmt)

    ws_jc.protect('premium', prot_options)

    # =========================================================================
    # TAB 10: SUMMARY DASHBOARD (With Native Chart!)
    # =========================================================================
    ws_dash = wb.add_worksheet('Summary Dashboard')
    ws_dash.set_tab_color('#D97706')
    ws_dash.set_column('A:A', 3)
    ws_dash.set_column('B:C', 24)
    ws_dash.set_column('D:G', 16)
    ws_dash.set_column('H:J', 18)

    ws_dash.set_row(1, 35)
    ws_dash.merge_range('B2:J2', 'EXECUTIVE CONSTRUCTION SUMMARY DASHBOARD & KPI OVERVIEW', fmt_title_banner)

    # 4 Top KPI Cards (Rows 4 to 6)
    kpis = [
        ('B4:C4', 'B5:C6', 'TOTAL ESTIMATED JOB COST', "='Estimate'!F49", fmt_kpi_val_curr),
        ('D4:E4', 'D5:E6', 'TOTAL CONTRACT BID PRICE', "='Estimate'!F55", fmt_kpi_val_curr),
        ('F4:G4', 'F5:G6', 'TOTAL MARKUP & GROSS PROFIT', "='Markup & Margin Calculator'!D13", fmt_kpi_val_curr),
        ('H4:I4', 'H5:I6', 'ESTIMATED PROFIT MARGIN %', "='Markup & Margin Calculator'!G13", fmt_kpi_val_pct)
    ]
    for top_range, val_range, title, f_val, fmt_val in kpis:
        ws_dash.merge_range(top_range, title, fmt_kpi_label)
        ws_dash.merge_range(val_range, '', fmt_val)
        # Write formula to the top-left cell of the merged value range
        tl_cell = val_range.split(':')[0]
        ws_dash.write_formula(tl_cell, f_val, fmt_val)

    # Table 1: Estimated vs Actual Breakdown (Cols B to F, Rows 8 to 15)
    ws_dash.merge_range('B8:F8', 'COST PERFORMANCE BY CATEGORY ($)', fmt_section_banner)
    dash_headers = ['Cost Category', 'Estimated ($)', 'Actual ($)', 'Variance ($)', 'Status']
    ws_dash.set_row(8, 25)
    for col_idx, h in enumerate(dash_headers):
        ws_dash.write(8, col_idx + 1, h, fmt_header)

    dash_rows = [
        ('Direct Materials', "='Job Costing & Actuals'!C6", "='Job Costing & Actuals'!D6"),
        ('Direct Labor', "='Job Costing & Actuals'!C7", "='Job Costing & Actuals'!D7"),
        ('Subcontractors', "='Job Costing & Actuals'!C8", "='Job Costing & Actuals'!D8"),
        ('Equipment & Rentals', "='Job Costing & Actuals'!C9", "='Job Costing & Actuals'!D9"),
        ('Permits & Other', "='Job Costing & Actuals'!C10", "='Job Costing & Actuals'!D10"),
        ('Overhead & Contingency', "='Job Costing & Actuals'!C11", "='Job Costing & Actuals'!D11")
    ]
    for i, (cat, f_est, f_act) in enumerate(dash_rows):
        r = 9 + i
        ws_dash.write(r, 1, cat, fmt_lock_text)
        ws_dash.write_formula(r, 2, f_est, fmt_lock_currency)
        ws_dash.write_formula(r, 3, f_act, fmt_lock_currency)
        ws_dash.write_formula(r, 4, f"=C{r+1}-D{r+1}", fmt_lock_currency)
        ws_dash.write_formula(r, 5, f'=IF(E{r+1}>=0, "UNDER BUDGET", "OVER BUDGET")', fmt_lock_center)

    # Table 1 Totals
    ws_dash.set_row(15, 25)
    ws_dash.write(15, 1, 'TOTAL JOB COSTS', fmt_total_label)
    ws_dash.write_formula(15, 2, '=SUM(C10:C15)', fmt_total_currency)
    ws_dash.write_formula(15, 3, '=SUM(D10:D15)', fmt_total_currency)
    ws_dash.write_formula(15, 4, '=C16-D16', fmt_total_currency)
    ws_dash.write_formula(15, 5, '=IF(E16>=0, "UNDER BUDGET", "OVER BUDGET")', fmt_total_label)

    # Table 2: Contract Pricing & Tax Breakdown (Cols H to J, Rows 8 to 15)
    ws_dash.merge_range('H8:J8', 'CONTRACT BID & PROFIT BREAKDOWN', fmt_header_amber)
    fin_rows = [
        ('Subtotal Direct Costs', "='Estimate'!F46"),
        ('Overhead Allowance', "='Estimate'!F47"),
        ('Contingency Allowance', "='Estimate'!F48"),
        ('Total Professional Markup', "='Estimate'!F50"),
        ('Subtotal Bid Selling Price', "='Estimate'!F51"),
        ('Preferred Client Discount', "='Estimate'!F52"),
        ('Sales Tax (8.25%)', "='Estimate'!F54")
    ]
    for i, (label, f_val) in enumerate(fin_rows):
        r = 9 + i
        ws_dash.merge_range(r, 7, r, 8, label, fmt_lock_text)
        ws_dash.write_formula(r, 9, f_val, fmt_lock_currency)

    ws_dash.set_row(16, 25)
    ws_dash.merge_range(16, 7, 16, 8, 'GRAND TOTAL BID PRICE ($)', fmt_grand_label)
    ws_dash.write_formula(16, 9, "='Estimate'!F55", fmt_grand_currency)

    # -------------------------------------------------------------------------
    # EXCEL-SAFE COST PERFORMANCE VISUAL DASHBOARD (Cell-based, no drawings)
    # -------------------------------------------------------------------------
    # Native Excel chart was removed because Excel was repairing /xl/drawings/drawing1.xml
    # This cell-based visual dashboard is 100% Excel-compatible with zero drawing objects.

    # Visual dashboard header
    ws_dash.merge_range('B18:J18', 'COST PERFORMANCE VISUAL SNAPSHOT (EXCEL-SAFE CELL-BASED DASHBOARD)', fmt_section_banner)

    # Visual dashboard data table with REPT-based inline bar chart
    dash_viz_headers = ['Cost Category', 'Estimated ($)', 'Actual ($)', 'Variance ($)', 'Budget Perf.', 'Visual Bar (Est.)', 'Visual Bar (Act.)']
    ws_dash.set_row(18, 25)
    for col_idx, h in enumerate(dash_viz_headers):
        ws_dash.write(18, col_idx + 1, h, fmt_header)

    # Define visual bar formats (locked, proportional font, colored)
    fmt_bar_navy = wb.add_format({
        'font_name': 'Consolas', 'font_size': 10, 'font_color': '#1B365D',
        'bg_color': '#FFFFFF', 'align': 'left', 'valign': 'vcenter',
        'border': 1, 'border_color': '#CBD5E1', 'locked': 1
    })
    fmt_bar_amber = wb.add_format({
        'font_name': 'Consolas', 'font_size': 10, 'font_color': '#D97706',
        'bg_color': '#FFFFFF', 'align': 'left', 'valign': 'vcenter',
        'border': 1, 'border_color': '#CBD5E1', 'locked': 1
    })

    dash_viz_rows = [
        ('Direct Materials', "='Job Costing & Actuals'!C6", "='Job Costing & Actuals'!D6"),
        ('Direct Labor', "='Job Costing & Actuals'!C7", "='Job Costing & Actuals'!D7"),
        ('Subcontractors', "='Job Costing & Actuals'!C8", "='Job Costing & Actuals'!D8"),
        ('Equipment & Rentals', "='Job Costing & Actuals'!C9", "='Job Costing & Actuals'!D9"),
        ('Permits & Other', "='Job Costing & Actuals'!C10", "='Job Costing & Actuals'!D10"),
        ('Overhead & Contingency', "='Job Costing & Actuals'!C11", "='Job Costing & Actuals'!D11")
    ]
    for i, (cat, f_est, f_act) in enumerate(dash_viz_rows):
        r = 19 + i
        ws_dash.write(r, 1, cat, fmt_lock_text)
        ws_dash.write_formula(r, 2, f_est, fmt_lock_currency)
        ws_dash.write_formula(r, 3, f_act, fmt_lock_currency)
        ws_dash.write_formula(r, 4, f"=C{r+1}-D{r+1}", fmt_lock_currency)
        ws_dash.write_formula(r, 5, f'=IF(E{r+1}>=0, "ON BUDGET", "OVER BUDGET")', fmt_lock_center)
        # Visual bars using REPT with Unicode Full Block character (█)
        # Scale: each block represents roughly $500. Clamped to show 1-50 blocks.
        ws_dash.write_formula(r, 6, f'=REPT("█", MAX(1, MIN(50, ROUND(C{r+1}/500, 0))))', fmt_bar_navy)
        ws_dash.write_formula(r, 7, f'=REPT("█", MAX(1, MIN(50, ROUND(D{r+1}/500, 0))))', fmt_bar_amber)

    # Visual dashboard total row
    ws_dash.set_row(25, 25)
    ws_dash.write(25, 1, 'TOTAL JOB COSTS', fmt_total_label)
    ws_dash.write_formula(25, 2, '=SUM(C20:C25)', fmt_total_currency)
    ws_dash.write_formula(25, 3, '=SUM(D20:D25)', fmt_total_currency)
    ws_dash.write_formula(25, 4, '=C26-D26', fmt_total_currency)
    ws_dash.write_formula(25, 5, '=IF(E26>=0, "ON BUDGET", "OVER BUDGET")', fmt_total_label)
    ws_dash.write_formula(25, 6, f'=REPT("█", MAX(1, MIN(50, ROUND(C26/500, 0))))', fmt_bar_navy)
    ws_dash.write_formula(25, 7, f'=REPT("█", MAX(1, MIN(50, ROUND(D26/500, 0))))', fmt_bar_amber)

    # Note explaining the removal of native chart
    ws_dash.merge_range('B27:J27', 'NOTE: Native embedded chart removed to ensure 100% Excel compatibility with zero repair errors. The REPT()-based visual bars above provide a drawing-free alternative.', fmt_note_text)

    ws_dash.protect('premium', prot_options)

    # =========================================================================
    # TAB 11: CLIENT PROPOSAL (Proposal Generator - Printable Page)
    # =========================================================================
    ws_prop = wb.add_worksheet('Client Proposal')
    ws_prop.set_tab_color('#2C3E50')
    ws_prop.set_paper(1)
    ws_prop.fit_to_pages(1, 0)
    ws_prop.set_margins(0.5, 0.5, 0.75, 0.75)
    ws_prop.hide_gridlines(2)  # Clean invoice style
    ws_prop.set_footer('&L&"Segoe UI"&8Template Author: novality store&R&"Segoe UI"&8Page &P of &N')

    ws_prop.set_column('A:A', 5)
    ws_prop.set_column('B:E', 20)
    ws_prop.set_column('F:G', 16)

    # Top Header Banner
    ws_prop.set_row(1, 35)
    ws_prop.merge_range('B2:G2', 'CONSTRUCTION PROJECT PROPOSAL & BID ACCEPTANCE', fmt_prop_title)
    ws_prop.merge_range('B3:G3', 'Professional Contractor Bid & Scope Specification', fmt_prop_subtitle)

    # Contractor & Client Info Box
    ws_prop.merge_range('B5:D5', 'CONTRACTOR / COMPANY INFO', fmt_section_banner)
    ws_prop.merge_range('E5:G5', 'CLIENT / PROJECT INFO', fmt_section_banner)

    ws_prop.write('B6', 'Company Name:', fmt_label_bold)
    ws_prop.merge_range('C6:D6', "='Project Info'!B13", fmt_lock_text)
    ws_prop.write('E6', 'Client Name:', fmt_label_bold)
    ws_prop.merge_range('F6:G6', "='Project Info'!B5", fmt_lock_text)

    ws_prop.write('B7', 'Contact Person:', fmt_label_bold)
    ws_prop.merge_range('C7:D7', "='Project Info'!B14", fmt_lock_text)
    ws_prop.write('E7', 'Project Name:', fmt_label_bold)
    ws_prop.merge_range('F7:G7', "='Project Info'!B4", fmt_lock_text)

    ws_prop.write('B8', 'Phone Number:', fmt_label_bold)
    ws_prop.merge_range('C8:D8', "='Project Info'!B16", fmt_lock_text)
    ws_prop.write('E8', 'Job Location:', fmt_label_bold)
    ws_prop.merge_range('F8:G8', "='Project Info'!B6", fmt_lock_text)

    ws_prop.write('B9', 'Email Address:', fmt_label_bold)
    ws_prop.merge_range('C9:D9', "='Project Info'!B17", fmt_lock_text)
    ws_prop.write('E9', 'Estimate Number:', fmt_label_bold)
    ws_prop.merge_range('F9:G9', "='Project Info'!B7", fmt_lock_text)

    # Scope of work summary
    ws_prop.merge_range('B11:G11', 'SCOPE OF WORK SPECIFICATION', fmt_section_banner)
    ws_prop.set_row(11, 40)
    ws_prop.merge_range('B12:G12', "='Project Info'!B10", fmt_wrap_lock)

    # Itemized Proposal Summary Table
    ws_prop.merge_range('B14:G14', 'PROJECT PRICING SUMMARY BY TRADE & SERVICE', fmt_header)
    ws_prop.merge_range('B15:E15', 'Cost Category / Trade Description', fmt_header_left)
    ws_prop.merge_range('F15:G15', 'Contract Price ($)', fmt_header)

    prop_lines = [
        ('Direct Materials & Custom Architectural Fixtures', "='Estimate'!F19"),
        ('Direct Labor & Skilled Craftsmanship Hours', "='Estimate'!F26"),
        ('Licensed Specialty Subcontractor Packages', "='Estimate'!F32"),
        ('Equipment Rentals, Waste Removal & Specialty Tooling', "='Estimate'!F38"),
        ('Municipal Building & Trade Permits, Site Protection & Cleanup', "='Estimate'!F43"),
        ('Contractor Overhead, Management & Contingency Allowance', "='Estimate'!F47+'Estimate'!F48"),
        ('Professional Contractor Markup & Value Engineering', "='Estimate'!F50"),
        ('SUBTOTAL CONTRACT BID PRICE', "='Estimate'!F51"),
        ('Preferred Client Discount (-2.50%)', "='Estimate'!F52"),
        ('Net Taxable Contract Bid Amount', "='Estimate'!F53"),
        ('Applicable Sales Tax (8.25%)', "='Estimate'!F54")
    ]

    for i, (desc, form) in enumerate(prop_lines):
        r = 15 + i
        ws_prop.set_row(r, 22)
        ws_prop.merge_range(r, 1, r, 4, desc, fmt_total_label if 'SUBTOTAL' in desc or 'Net Taxable' in desc else fmt_lock_text)
        ws_prop.merge_range(r, 5, r, 6, '', fmt_total_currency if 'SUBTOTAL' in desc or 'Net Taxable' in desc else fmt_lock_currency)
        ws_prop.write_formula(f'F{r+1}', form, fmt_total_currency if 'SUBTOTAL' in desc or 'Net Taxable' in desc else fmt_lock_currency)

    # Grand Total Contract Proposal Price
    gt_r = 15 + len(prop_lines)
    ws_prop.set_row(gt_r, 28)
    ws_prop.merge_range(gt_r, 1, gt_r, 4, 'GRAND TOTAL CONTRACT BID PRICE ($)', fmt_grand_label)
    ws_prop.merge_range(gt_r, 5, gt_r, 6, '', fmt_grand_currency)
    ws_prop.write_formula(f'F{gt_r+1}', "='Estimate'!F55", fmt_grand_currency)

    # Exclusions & Terms
    ws_prop.merge_range('B28:G28', 'EXCLUSIONS, CLARIFICATIONS & PAYMENT TERMS', fmt_section_banner)
    ws_prop.set_row(28, 50)
    excl_text = (
        "EXCLUSIONS: Excludes hazardous material abatement, unforeseen subterranean repairs, and utility company connection fees. "
        "PAYMENT SCHEDULE: 20% deposit upon contract signing, 50% upon progress milestones, and 30% upon substantial completion. "
        "PROPOSAL VALIDITY: 30 days from date of issue."
    )
    ws_prop.merge_range('B29:G29', excl_text, fmt_note_text)

    # Client Acceptance Sign-off Area
    ws_prop.merge_range('B31:G31', 'CLIENT ACCEPTANCE & AUTHORIZATION SIGNATURE AREA', fmt_section_banner)
    ws_prop.write('B32', 'Client Signature:', fmt_label_bold)
    ws_prop.merge_range('C32:D32', '', fmt_input_text)
    ws_prop.write('E32', 'Date Signed:', fmt_label_bold)
    ws_prop.merge_range('F32:G32', '', fmt_input_date)

    ws_prop.write('B33', 'Printed Name:', fmt_label_bold)
    ws_prop.merge_range('C33:D33', 'Robert Sterling, Apex Properties', fmt_input_text)
    ws_prop.write('E33', 'Contractor Sign-off:', fmt_label_bold)
    ws_prop.merge_range('F33:G33', 'David Novality, Master Contractor', fmt_lock_text)

    ws_prop.protect('premium', prot_options)

    # =========================================================================
    # TAB 12: TERMS & CONDITIONS
    # =========================================================================
    ws_tc = wb.add_worksheet('Terms & Conditions')
    ws_tc.set_tab_color('#64748B')
    ws_tc.set_paper(1)
    ws_tc.fit_to_pages(1, 0)
    ws_tc.set_margins(0.5, 0.5, 0.75, 0.75)
    ws_tc.hide_gridlines(2)
    ws_tc.set_footer('&L&"Segoe UI"&8Template Author: novality store&R&"Segoe UI"&8Page &P of &N')

    ws_tc.set_column('A:A', 5)
    ws_tc.set_column('B:G', 16)

    ws_tc.set_row(1, 35)
    ws_tc.merge_range('B2:G2', 'STANDARD CONTRACTOR TERMS, CONDITIONS & WORKMANSHIP WARRANTY', fmt_prop_title)
    ws_tc.merge_range('B3:G3', 'Legal Agreement & Operational Policies for Construction Contracts', fmt_prop_subtitle)

    tc_articles = [
        ("1. SCOPE OF WORK & PROPOSAL VALIDITY", "The contractor agrees to perform the scope of work specified in the Client Proposal. The pricing stated in this document is valid for 30 calendar days from the date of issuance. Any work not explicitly listed is excluded."),
        ("2. PAYMENT SCHEDULE & INVOICING", "Payments shall be made according to the agreed milestone payment schedule: 20% mobilization deposit upon signing, milestone progress invoices during construction, and final payment upon completion of punchlist items."),
        ("3. CHANGE ORDERS & MODIFICATIONS", "Any alteration or deviation from the original specifications involving extra costs or schedule extensions will be executed only upon a written Change Order signed by both Owner and Contractor."),
        ("4. SITE ACCESS, WORKING HOURS & UTILITIES", "The owner agrees to provide reasonable access to the property during standard working hours (7:00 AM - 5:00 PM) and make water and electricity available for construction purposes without charge."),
        ("5. WORKMANSHIP WARRANTY & MATERIAL GUARANTEE", "Novality Construction & Design LLC provides a 1-year limited workmanship warranty on all installations. Standard manufacturer warranties apply to all installed equipment and material fixtures."),
        ("6. INSURANCE & LICENSING", "The contractor shall maintain general liability and workers' compensation insurance throughout the project duration and hold valid municipal trade licenses."),
        ("7. DELAYS & FORCE MAJEURE", "The contractor is not responsible for project delays caused by severe weather, material vendor backorders, strikes, or municipal permitting delays beyond contractor control.")
    ]

    row_idx = 5
    for title, text in tc_articles:
        ws_tc.set_row(row_idx, 22)
        ws_tc.merge_range(row_idx, 1, row_idx, 6, title, fmt_section_banner)
        row_idx += 1
        ws_tc.set_row(row_idx, 40)
        ws_tc.merge_range(row_idx, 1, row_idx, 6, text, fmt_wrap_lock)
        row_idx += 2

    ws_tc.protect('premium', prot_options)

    # =========================================================================
    # TAB 13: MATERIAL PRICE DB (Optional Advanced Feature)
    # =========================================================================
    ws_mpdb = wb.add_worksheet('Material Price DB')
    ws_mpdb.set_tab_color('#64748B')
    ws_mpdb.set_column('A:A', 10)
    ws_mpdb.set_column('B:B', 32)
    ws_mpdb.set_column('C:D', 18)
    ws_mpdb.set_column('E:G', 14)

    ws_mpdb.set_row(1, 30)
    ws_mpdb.merge_range('A2:G2', 'MATERIAL PRICE DATABASE (MASTER INVENTORY & UNIT PRICE REFERENCE)', fmt_title_banner)
    ws_mpdb.merge_range('A3:G3', '  Reference table of standard construction materials, default suppliers, and current market unit costs.', fmt_sub_banner)

    mp_headers = ['SKU Code', 'Material Description & Specification', 'Trade Category', 'Default Supplier', 'Unit', 'Standard Price ($)', 'Lead Time (Days)']
    ws_mpdb.set_row(4, 25)
    for col_idx, h in enumerate(mp_headers):
        ws_mpdb.write(4, col_idx, h, fmt_header)

    mp_data = [
        ('MAT-101', 'Custom Shaker Base Cabinet Box', 'Cabinetry', 'Craftsman Cabinet Co.', 'Set', 700.00, 14),
        ('MAT-102', 'Custom Shaker Wall Cabinet Box', 'Cabinetry', 'Craftsman Cabinet Co.', 'Set', 350.00, 14),
        ('MAT-103', 'Calacatta Gold Quartz Slab (3cm)', 'Countertops', 'Apex Stone & Tile', 'SqFt', 75.00, 7),
        ('MAT-104', 'Engineered Oak Hardwood Plank', 'Flooring', 'Prime Flooring Dist.', 'SqFt', 6.50, 5),
        ('MAT-105', 'Hardwood Acoustic Underlayment', 'Flooring', 'Prime Flooring Dist.', 'SqFt', 0.65, 3),
        ('MAT-106', 'Recessed 6-inch LED Dimmable Fixture', 'Electrical', 'Lumens Direct', 'Unit', 45.00, 2),
        ('MAT-107', 'Decorative Pendant Light Fixture', 'Electrical', 'Lumens Direct', 'Unit', 220.00, 5),
        ('MAT-108', 'Architectural Primer (5-Gal Pail)', 'Paint', 'Sherwin-Williams Pro', 'Pail', 45.00, 1),
        ('MAT-109', 'Premium Satin Enamel (5-Gal Pail)', 'Paint', 'Sherwin-Williams Pro', 'Pail', 65.00, 1),
        ('MAT-110', 'Commercial Stainless Sink & Faucet', 'Plumbing', 'Ferguson Plumbing', 'Set', 950.00, 7),
        ('MAT-111', 'Ceramic Subway Backsplash Tile', 'Tile', 'Apex Stone & Tile', 'SqFt', 8.50, 4),
        ('MAT-112', 'Drywall Sheet 1/2-inch x 8ft', 'Drywall', 'Home Depot Pro', 'Sheet', 14.50, 1),
        ('MAT-113', 'All-Purpose Joint Compound Box', 'Drywall', 'Home Depot Pro', 'Box', 18.00, 1),
        ('MAT-114', 'Pressure Treated 2x4 Lumber Stud', 'Framing', 'Home Depot Pro', 'Unit', 6.25, 1),
        ('MAT-115', 'Interior Mold-Resistant Silicone Caulk', 'Finish', 'Home Depot Pro', 'Tube', 8.50, 1)
    ]
    for i, row in enumerate(mp_data):
        r = 5 + i
        ws_mpdb.write(r, 0, row[0], fmt_lock_center)
        ws_mpdb.write(r, 1, row[1], fmt_lock_text)
        ws_mpdb.write(r, 2, row[2], fmt_lock_center)
        ws_mpdb.write(r, 3, row[3], fmt_lock_text)
        ws_mpdb.write(r, 4, row[4], fmt_lock_center)
        ws_mpdb.write(r, 5, row[5], fmt_input_currency)
        ws_mpdb.write(r, 6, row[6], fmt_lock_center)

    ws_mpdb.protect('premium', prot_options)

    # =========================================================================
    # TAB 14: LABOR RATE DB (Optional Advanced Feature)
    # =========================================================================
    ws_lrdb = wb.add_worksheet('Labor Rate DB')
    ws_lrdb.set_tab_color('#64748B')
    ws_lrdb.set_column('A:A', 10)
    ws_lrdb.set_column('B:B', 28)
    ws_lrdb.set_column('C:G', 16)

    ws_lrdb.set_row(1, 30)
    ws_lrdb.merge_range('A2:G2', 'LABOR RATE DATABASE (STANDARD HOURLY WAGES, BURDEN & BILLING RATES)', fmt_title_banner)
    ws_lrdb.merge_range('A3:G3', '  Master labor rate table with base wage, payroll burden, fully burdened cost, and standard billing rates.', fmt_sub_banner)

    lr_headers = ['Trade Code', 'Craft Role / Trade Description', 'Base Wage ($/hr)', 'Burden & Ins. %', 'Burdened Cost ($/hr)', 'Billing Rate ($/hr)', 'OT Billing Rate ($/hr)']
    ws_lrdb.set_row(4, 25)
    for col_idx, h in enumerate(lr_headers):
        ws_lrdb.write(4, col_idx, h, fmt_header)

    lr_data = [
        ('TRD-01', 'General Construction Laborer', 25.00, 0.25, 40.00, 60.00),
        ('TRD-02', 'Skilled Carpenter & Framing Lead', 35.00, 0.28, 55.00, 82.50),
        ('TRD-03', 'Master Millwork & Cabinet Specialist', 42.00, 0.28, 65.00, 97.50),
        ('TRD-04', 'Drywall Hanger & Finisher', 30.00, 0.25, 45.00, 67.50),
        ('TRD-05', 'Interior Painter & Trim Finisher', 28.00, 0.25, 45.00, 67.50),
        ('TRD-06', 'Licensed Electrician (Master)', 48.00, 0.30, 75.00, 112.50),
        ('TRD-07', 'Licensed Plumber & Gas Fitter', 48.00, 0.30, 75.00, 112.50),
        ('TRD-08', 'HVAC Duct & Mechanical Technician', 42.00, 0.30, 70.00, 105.00),
        ('TRD-09', 'Ceramic & Stone Tile Setter', 35.00, 0.26, 55.00, 82.50),
        ('TRD-10', 'Hardwood Flooring Installer', 32.00, 0.26, 50.00, 75.00),
        ('TRD-11', 'Project Supervisor / QC Manager', 50.00, 0.32, 75.00, 112.50),
        ('TRD-12', 'Site Safety & Sanitization Tech', 22.00, 0.25, 35.00, 52.50)
    ]
    for i, row in enumerate(lr_data):
        r = 5 + i
        ws_lrdb.write(r, 0, row[0], fmt_lock_center)
        ws_lrdb.write(r, 1, row[1], fmt_lock_text)
        ws_lrdb.write(r, 2, row[2], fmt_input_currency)
        ws_lrdb.write(r, 3, row[3], fmt_input_percent)
        # Fully Burdened Cost = Base Wage * (1 + Burden)
        ws_lrdb.write_formula(r, 4, f"=C{r+1}*(1+D{r+1})", fmt_lock_currency)
        ws_lrdb.write(r, 5, row[4], fmt_input_currency)
        ws_lrdb.write(r, 6, row[5], fmt_input_currency)

    ws_lrdb.protect('premium', prot_options)

    # =========================================================================
    # TAB 15: ESTIMATE VERSIONS (Optional Advanced Feature)
    # =========================================================================
    ws_ver = wb.add_worksheet('Estimate Versions')
    ws_ver.set_tab_color('#64748B')
    ws_ver.set_column('A:A', 32)
    ws_ver.set_column('B:D', 18)

    ws_ver.set_row(1, 30)
    ws_ver.merge_range('A2:D2', 'MULTIPLE ESTIMATE VERSIONS COMPARISON (GOOD / BETTER / BEST)', fmt_title_banner)
    ws_ver.merge_range('A3:D3', '  Side-by-side comparison of 3 client proposal packages to facilitate upselling and option analysis.', fmt_sub_banner)

    ver_headers = ['Cost Category / Metric', 'Option 1: Standard / Base', 'Option 2: Upgraded / Preferred', 'Option 3: Premium / Luxury']
    ws_ver.set_row(4, 25)
    for col_idx, h in enumerate(ver_headers):
        ws_ver.write(4, col_idx, h, fmt_header)

    ver_rows = [
        ('Direct Materials & Fixture Packages ($)', 16500.00, 21650.00, 29800.00, fmt_input_currency),
        ('Direct Labor & Craft Hours ($)', 14200.00, 16800.00, 21500.00, fmt_input_currency),
        ('Licensed Subcontractor Services ($)', 8500.00, 9750.00, 12400.00, fmt_input_currency),
        ('Equipment Rentals & Tool Allowance ($)', 2200.00, 2850.00, 3600.00, fmt_input_currency),
        ('Permits, Fees & Site Protection ($)', 1450.00, 1450.00, 1850.00, fmt_input_currency),
        ('TOTAL DIRECT JOB COSTS ($)', '=SUM(B6:B10)', '=SUM(C6:C10)', '=SUM(D6:D10)', fmt_total_currency),
        ('Overhead Allowance (10.0%)', "=B11*'Project Info'!F7", "=C11*'Project Info'!F7", "=D11*'Project Info'!F7", fmt_lock_currency),
        ('Contingency Allowance (5.0%)', "=B11*'Project Info'!F6", "=C11*'Project Info'!F6", "=D11*'Project Info'!F6", fmt_lock_currency),
        ('SUBTOTAL JOB COST (Before Markup)', '=B11+B12+B13', '=C11+C12+C13', '=D11+D12+D13', fmt_total_currency),
        ('Contractor Markup % Applied', 0.2200, 0.2348, 0.2500, fmt_input_percent),
        ('Total Markup Amount ($)', '=B14*B15', '=C14*C15', '=D14*D15', fmt_lock_currency),
        ('ESTIMATED CONTRACT SELLING PRICE ($)', '=B14+B16', '=C14+C16', '=D14+D16', fmt_grand_currency),
        ('Equivalent Profit Margin %', '=IFERROR(B16/B17, 0)', '=IFERROR(C16/C17, 0)', '=IFERROR(D16/D17, 0)', fmt_total_percent)
    ]
    for i, (label, f1, f2, f3, f_fmt) in enumerate(ver_rows):
        r = 5 + i
        ws_ver.write(r, 0, label, fmt_total_label if 'TOTAL' in label or 'SUBTOTAL' in label or 'PRICE' in label else fmt_lock_text)
        for col_idx, f_val in enumerate([f1, f2, f3]):
            if isinstance(f_val, str) and f_val.startswith('='):
                ws_ver.write_formula(r, col_idx + 1, f_val, f_fmt)
            else:
                ws_ver.write(r, col_idx + 1, f_val, f_fmt)

    ws_ver.protect('premium', prot_options)

    # =========================================================================
    # TAB 16: CHANGE ORDER TRACKER (Optional Advanced Feature)
    # =========================================================================
    ws_co = wb.add_worksheet('Change Order Tracker')
    ws_co.set_tab_color('#64748B')
    ws_co.set_column('A:A', 10)
    ws_co.set_column('B:B', 12)
    ws_co.set_column('C:C', 32)
    ws_co.set_column('D:D', 18)
    ws_co.set_column('E:G', 16)
    ws_co.set_column('H:J', 14)

    ws_co.set_row(1, 30)
    ws_co.merge_range('A2:J2', 'PROJECT CHANGE ORDER TRACKER & BID MODIFICATION LOG', fmt_title_banner)
    ws_co.merge_range('A3:J3', '  Log scope changes, calculate change order markups, and track client approval status.', fmt_sub_banner)

    co_headers = ['CO Number', 'Date Req.', 'Description of Scope Modification', 'Requested By', 'Direct Cost ($)', 'Markup %', 'CO Total Price ($)', 'Schedule Add', 'Approval Status', 'Date Approved']
    ws_co.set_row(4, 25)
    for col_idx, h in enumerate(co_headers):
        ws_co.write(4, col_idx, h, fmt_header)

    co_items = [
        ('CO-001', '2026-08-05', 'Upgrade Quartz Island Countertop to Waterfall Edge', 'Robert Sterling', 1200.00, 0.20, '2 Days', 'Approved', '2026-08-06'),
        ('CO-002', '2026-08-10', 'Add 4 Additional Dimmable Recessed LED Lights', 'Robert Sterling', 450.00, 0.25, '1 Day', 'Approved', '2026-08-11'),
        ('CO-003', '2026-08-15', 'Relocate Secondary Laundry Plumbing Drain Line', 'Robert Sterling', 850.00, 0.20, '1 Day', 'Pending', '-')
    ]
    for i, row in enumerate(co_items):
        r = 5 + i
        ws_co.write(r, 0, row[0], fmt_input_center)
        ws_co.write(r, 1, row[1], fmt_input_date)
        ws_co.write(r, 2, row[2], fmt_input_text)
        ws_co.write(r, 3, row[3], fmt_input_text)
        ws_co.write(r, 4, row[4], fmt_input_currency)
        ws_co.write(r, 5, row[5], fmt_input_percent)
        # CO Total Price = Direct Cost * (1 + Markup %)
        ws_co.write_formula(r, 6, f"=E{r+1}*(1+F{r+1})", fmt_lock_currency)
        ws_co.write(r, 7, row[6], fmt_input_center)
        ws_co.write(r, 8, row[7], fmt_input_center)
        ws_co.write(r, 9, row[8], fmt_input_center)

    tot_r_co = 5 + len(co_items)
    ws_co.set_row(tot_r_co, 25)
    ws_co.merge_range(tot_r_co, 0, tot_r_co, 3, 'TOTAL CHANGE ORDER LOG ($)', fmt_total_label)
    ws_co.write_formula(tot_r_co, 4, f"=SUM(E6:E{tot_r_co})", fmt_total_currency)
    ws_co.write(tot_r_co, 5, '', fmt_total_label)
    ws_co.write_formula(tot_r_co, 6, f"=SUM(G6:G{tot_r_co})", fmt_total_currency)
    ws_co.merge_range(tot_r_co, 7, tot_r_co, 9, '', fmt_total_label)

    # Change Order Summary Table
    ws_co.merge_range('B12:E12', 'CHANGE ORDER CONTRACT PRICE IMPACT', fmt_section_banner)
    co_sum_rows = [
        ('Original Contract Bid Price ($)', "='Estimate'!F55", fmt_lock_currency),
        ('Total Approved Change Orders ($)', f'=SUMIF(I6:I{tot_r_co}, "Approved", G6:G{tot_r_co})', fmt_total_currency),
        ('REVISED GRAND TOTAL CONTRACT PRICE ($)', '=D13+D14', fmt_grand_currency)
    ]
    for i, (lbl, f_form, f_fmt) in enumerate(co_sum_rows):
        r = 12 + i
        ws_co.merge_range(r, 1, r, 2, lbl, fmt_lock_text if i < 2 else fmt_grand_label)
        ws_co.write_formula(r, 3, f_form, f_fmt)

    ws_co.protect('premium', prot_options)

    # =========================================================================
    # TAB 17: PAYMENT SCHEDULE (Optional Advanced Feature)
    # =========================================================================
    ws_pay = wb.add_worksheet('Payment Schedule')
    ws_pay.set_tab_color('#64748B')
    ws_pay.set_column('A:A', 10)
    ws_pay.set_column('B:B', 32)
    ws_pay.set_column('C:E', 18)
    ws_pay.set_column('F:H', 16)

    ws_pay.set_row(1, 30)
    ws_pay.merge_range('A2:H2', 'MILESTONE PAYMENT SCHEDULE & PROGRESS BILLING TRACKER', fmt_title_banner)
    ws_pay.merge_range('A3:H3', '  Schedule progress payments based on contract milestones and track amounts received.', fmt_sub_banner)

    pay_headers = ['Milestone #', 'Milestone Phase Description', '% of Contract', 'Scheduled Amount ($)', 'Target Due Date', 'Payment Status', 'Date Received', 'Check / Ref #']
    ws_pay.set_row(4, 25)
    for col_idx, h in enumerate(pay_headers):
        ws_pay.write(4, col_idx, h, fmt_header)

    pay_milestones = [
        ('PAY-01', 'Initial Mobilization & Deposit (Upon Contract Signing)', 0.20, '2026-08-01', 'Paid', '2026-08-01', 'CHK-9921'),
        ('PAY-02', 'Rough-In Completion (Electrical, Plumbing & HVAC)', 0.25, '2026-08-15', 'Paid', '2026-08-16', 'CHK-9945'),
        ('PAY-03', 'Drywall, Flooring & Cabinet Installation Completion', 0.25, '2026-08-30', 'Pending', '-', '-'),
        ('PAY-04', 'Countertops, Backsplash Tile & Fixture Completion', 0.20, '2026-09-10', 'Pending', '-', '-'),
        ('PAY-05', 'Final Inspection & Punchlist Sign-Off', 0.10, '2026-09-20', 'Pending', '-', '-')
    ]
    for i, row in enumerate(pay_milestones):
        r = 5 + i
        ws_pay.write(r, 0, row[0], fmt_lock_center)
        ws_pay.write(r, 1, row[1], fmt_lock_text)
        ws_pay.write(r, 2, row[2], fmt_input_percent)
        # Scheduled Amount = Grand Total Contract Price * Milestone %
        ws_pay.write_formula(r, 3, f"='Estimate'!$F$55*C{r+1}", fmt_lock_currency)
        ws_pay.write(r, 4, row[3], fmt_input_date)
        ws_pay.write(r, 5, row[4], fmt_input_center)
        ws_pay.write(r, 6, row[5], fmt_input_center)
        ws_pay.write(r, 7, row[6], fmt_input_center)

    tot_r_pay = 5 + len(pay_milestones)
    ws_pay.set_row(tot_r_pay, 25)
    ws_pay.merge_range(tot_r_pay, 0, tot_r_pay, 1, 'TOTAL CONTRACT PAYMENT SCHEDULE', fmt_total_label)
    ws_pay.write_formula(tot_r_pay, 2, f"=SUM(C6:C{tot_r_pay})", fmt_total_percent)
    ws_pay.write_formula(tot_r_pay, 3, f"=SUM(D6:D{tot_r_pay})", fmt_total_currency)
    ws_pay.merge_range(tot_r_pay, 4, tot_r_pay, 7, '', fmt_total_label)

    ws_pay.protect('premium', prot_options)

    # =========================================================================
    # TAB 18: INVOICE SUMMARY (Optional Advanced Feature)
    # =========================================================================
    ws_inv = wb.add_worksheet('Invoice Summary')
    ws_inv.set_tab_color('#64748B')
    ws_inv.set_paper(1)
    ws_inv.fit_to_pages(1, 0)
    ws_inv.set_margins(0.5, 0.5, 0.75, 0.75)
    ws_inv.hide_gridlines(2)

    ws_inv.set_column('A:A', 5)
    ws_inv.set_column('B:E', 20)
    ws_inv.set_column('F:G', 16)

    ws_inv.set_row(1, 35)
    ws_inv.merge_range('B2:G2', 'PROGRESS INVOICE & CONTRACT BILLING SUMMARY', fmt_prop_title)
    ws_inv.merge_range('B3:G3', 'Standard AIA-Style Progress Billing Summary for Construction Contractors', fmt_prop_subtitle)

    ws_inv.merge_range('B5:D5', 'CONTRACTOR / PAYEE DETAILS', fmt_section_banner)
    ws_inv.merge_range('E5:G5', 'INVOICE & CONTRACT METRICS', fmt_section_banner)

    ws_inv.write('B6', 'Company Name:', fmt_label_bold)
    ws_inv.merge_range('C6:D6', "='Project Info'!B13", fmt_lock_text)
    ws_inv.write('E6', 'Invoice Number:', fmt_label_bold)
    ws_inv.merge_range('F6:G6', 'INV-2026-001', fmt_input_center)

    ws_inv.write('B7', 'Client / Owner:', fmt_label_bold)
    ws_inv.merge_range('C7:D7', "='Project Info'!B5", fmt_lock_text)
    ws_inv.write('E7', 'Billing Period:', fmt_label_bold)
    ws_inv.merge_range('F7:G7', '2026-08-01 to 2026-08-16', fmt_input_center)

    ws_inv.merge_range('B10:G10', 'PROGRESS BILLING FINANCIAL BREAKDOWN', fmt_header)
    inv_rows = [
        ('1. Original Contract Bid Price ($)', "='Estimate'!F55", fmt_lock_currency),
        ('2. Net Addition from Approved Change Orders ($)', "='Change Order Tracker'!D14", fmt_lock_currency),
        ('3. REVISED TOTAL CONTRACT PRICE ($)', '=E11+E12', fmt_total_currency),
        ('4. Total Value of Work Completed to Date ($)', '=E13*0.45', fmt_lock_currency),
        ('5. Retainage Withheld (5.00% of Completed Work)', '=E14*0.05', fmt_lock_currency),
        ('6. Total Earned Less Retainage ($)', '=E14-E15', fmt_total_currency),
        ('7. Less Previous Payments Received ($)', "='Payment Schedule'!D6", fmt_lock_currency),
        ('8. CURRENT PROGRESS INVOICE AMOUNT DUE ($)', '=E16-E17', fmt_grand_currency),
        ('9. Balance to Finish Including Retainage ($)', '=E13-E16', fmt_total_currency)
    ]
    for i, (lbl, f_form, f_fmt) in enumerate(inv_rows):
        r = 10 + i
        ws_inv.set_row(r, 24 if 'CURRENT PROGRESS' in lbl else 22)
        ws_inv.merge_range(r, 1, r, 3, lbl, fmt_grand_label if 'CURRENT PROGRESS' in lbl else fmt_total_label if 'REVISED' in lbl or 'Earned Less' in lbl else fmt_lock_text)
        ws_inv.merge_range(r, 4, r, 5, f_form, f_fmt)

    ws_inv.protect('premium', prot_options)

    # =========================================================================
    # TAB 19: PROJECT TIMELINE (Optional Advanced Feature)
    # =========================================================================
    ws_time = wb.add_worksheet('Project Timeline')
    ws_time.set_tab_color('#64748B')
    ws_time.set_column('A:A', 8)
    ws_time.set_column('B:B', 32)
    ws_time.set_column('C:D', 14)
    ws_time.set_column('E:G', 12)
    ws_time.set_column('H:S', 6)

    ws_time.set_row(1, 30)
    ws_time.merge_range('A2:S2', 'CONSTRUCTION PROJECT TIMELINE & GANTT SCHEDULE TRACKER (12 WEEKS)', fmt_title_banner)
    ws_time.merge_range('A3:S3', '  High-level schedule tracker by phase. Visual bar indicators highlight planned week duration.', fmt_sub_banner)

    time_headers = ['Phase #', 'Construction Phase & Task', 'Lead Trade', 'Status', 'Start Wk', 'Duration', 'End Wk']
    ws_time.set_row(4, 25)
    for col_idx, h in enumerate(time_headers):
        ws_time.write(4, col_idx, h, fmt_header)
    for wk in range(1, 13):
        ws_time.write(4, 6 + wk, f"W{wk}", fmt_header)

    fmt_bar = wb.add_format({'bg_color': '#1B365D', 'border': 1, 'border_color': '#CBD5E1', 'locked': 1})
    fmt_empty = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'border_color': '#E2E8F0', 'locked': 1})

    phases = [
        ('PH-01', 'Permitting & Material Procurement', 'Management', 'Completed', 1, 2),
        ('PH-02', 'Selective Kitchen Demolition & Site Prep', 'Demolition', 'Completed', 2, 1),
        ('PH-03', 'Structural Framing & Wall Header Reroute', 'Carpentry', 'Completed', 3, 1),
        ('PH-04', 'Electrical & Plumbing Rough-In Wiring/Pipes', 'MEP Subs', 'Completed', 3, 2),
        ('PH-05', 'Municipal Inspection (Rough MEP & Frame)', 'Municipal', 'Completed', 5, 1),
        ('PH-06', 'Drywall Hanging, Taping & Plaster Sanding', 'Drywall', 'In Progress', 5, 2),
        ('PH-07', 'Engineered Hardwood Floor Installation', 'Flooring', 'Scheduled', 7, 1),
        ('PH-08', 'Custom Cabinetry & Kitchen Island Box Install', 'Carpentry', 'Scheduled', 8, 2),
        ('PH-09', 'Quartz Countertop Slab & Tile Backsplash', 'Stone / Tile', 'Scheduled', 9, 2),
        ('PH-10', 'Interior Paint, Trim Molding & Enamel Finish', 'Painting', 'Scheduled', 10, 2),
        ('PH-11', 'MEP Fixtures, Faucet & LED Trim Finish', 'MEP Subs', 'Scheduled', 11, 1),
        ('PH-12', 'Final Cleanup, Punchlist & Owner Sign-Off', 'Management', 'Scheduled', 12, 1)
    ]

    for i, (p_num, p_name, p_lead, p_stat, p_start, p_dur) in enumerate(phases):
        r = 5 + i
        ws_time.write(r, 0, p_num, fmt_lock_center)
        ws_time.write(r, 1, p_name, fmt_lock_text)
        ws_time.write(r, 2, p_lead, fmt_lock_center)
        ws_time.write(r, 3, p_stat, fmt_input_center)
        ws_time.write(r, 4, p_start, fmt_input_int)
        ws_time.write(r, 5, p_dur, fmt_input_int)
        ws_time.write_formula(r, 6, f"=E{r+1}+F{r+1}-1", fmt_lock_center)

        # Draw Gantt timeline blocks for Weeks 1..12
        for wk in range(1, 13):
            if p_start <= wk < (p_start + p_dur):
                ws_time.write(r, 6 + wk, '', fmt_bar)
            else:
                ws_time.write(r, 6 + wk, '', fmt_empty)

    ws_time.protect('premium', prot_options)

    # -------------------------------------------------------------------------
    # SAVE AND CLOSE WORKBOOK
    # -------------------------------------------------------------------------
    wb.close()
    print(f"[SUCCESS] Successfully generated 19-tab Construction Estimate Template at '{output_path}'")
    print("All formula cells are locked with password 'premium'. Author set to 'novality store'.")


if __name__ == '__main__':
    output_filename = sys.argv[1] if len(sys.argv) > 1 else "construction_estimate_template.xlsx"
    generate_construction_estimate_template(output_filename)
