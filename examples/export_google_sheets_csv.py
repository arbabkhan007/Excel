#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Sheets CSV Exporter for Construction Estimate Template
=============================================================

This script exports all 19 sheets of the Construction Estimate Template
into clean .csv files inside the 'google_sheets_import/' directory so they
can be easily imported into Google Sheets individually if desired.

Author: novality store
"""

import os
import csv


def export_google_sheets_csvs(output_dir="google_sheets_import"):
    os.makedirs(output_dir, exist_ok=True)

    # Sheet 1: Instructions
    with open(os.path.join(output_dir, "01_Instructions.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["CONSTRUCTION ESTIMATE TEMPLATE EXCEL | CONTRACTOR BID & JOB COSTING SPREADSHEET, MARKUP MARGIN CALCULATOR, PROPOSAL GENERATOR GOOGLE SHEETS"])
        w.writerow(["Template Author: novality store", "Formula Protection Password: premium", "Excel & Google Sheets Compatible"])
        w.writerow([])
        w.writerow(["SECTION", "DESCRIPTION"])
        w.writerow(["Overview", "Commercial-grade construction estimate & contractor bid spreadsheet engineered for contractors, remodelers, builders, electricians, plumbers, roofers, and specialty trades."])
        w.writerow(["Step 1: Project Info", "Configure client details, contractor company info, and global tax/discount/contingency rates."])
        w.writerow(["Step 2: Enter Costs", "Use Materials, Labor, Equipment, and Subcontractors tabs to enter detailed itemized trade costs."])
        w.writerow(["Step 3: Build Estimate", "Open the Estimate tab to review compiled direct job costs, overhead, contingency, and contract bid total."])
        w.writerow(["Step 4: Markup & Margin", "Configure category-specific markups and compare markup vs. margin percentages."])
        w.writerow(["Step 5: Client Proposal", "Generate printable invoice-style client proposal page with signature acceptance block."])
        w.writerow(["Step 6: Track Actuals", "Track field expenditures on Job Costing & Actuals tab and view dashboard on Summary Dashboard."])

    # Sheet 2: Project Info
    with open(os.path.join(output_dir, "02_Project_Info.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["PROJECT, CLIENT & CONTRACTOR SETUP (GLOBAL PARAMETERS)"])
        w.writerow([])
        w.writerow(["Category", "Parameter Name", "Value / Configuration", "Notes"])
        w.writerow(["Project", "Project Name", "Luxury Kitchen & Living Area Remodel", "Editable"])
        w.writerow(["Project", "Client / Owner Name", "Robert Sterling / Apex Properties", "Editable"])
        w.writerow(["Project", "Jobsite Location", "742 Evergreen Terrace, Suite 300", "Editable"])
        w.writerow(["Project", "Estimate Number", "EST-2026-042", "Editable"])
        w.writerow(["Project", "Estimate Date", "2026-07-31", "Editable"])
        w.writerow(["Project", "Proposal Expiration Date", "2026-08-30", "Editable"])
        w.writerow(["Project", "Scope of Work Summary", "Full kitchen demolition, custom cabinetry, quartz countertops, electrical & plumbing relocation, hardwood flooring, and ambient LED lighting.", "Editable"])
        w.writerow([])
        w.writerow(["Contractor", "Contractor / Company Name", "Novality Construction & Design LLC", "Editable"])
        w.writerow(["Contractor", "Primary Contact / Estimator", "David Novality, Master Contractor", "Editable"])
        w.writerow(["Contractor", "Company Business Address", "100 Professional Way, Suite 400", "Editable"])
        w.writerow(["Contractor", "Phone Number", "(555) 382-9100", "Editable"])
        w.writerow(["Contractor", "Email Address", "estimating@novalitystore.com", "Editable"])
        w.writerow(["Contractor", "Contractor License #", "LIC-GEN-884210", "Editable"])
        w.writerow([])
        w.writerow(["Financial Rates", "Sales Tax Rate (%)", "8.25%", "Global Rate"])
        w.writerow(["Financial Rates", "Preferred Client Discount Rate (%)", "2.50%", "Global Rate"])
        w.writerow(["Financial Rates", "Contingency Allowance Rate (%)", "5.00%", "Global Rate"])
        w.writerow(["Financial Rates", "General Overhead Rate (%)", "10.00%", "Global Rate"])
        w.writerow(["Financial Rates", "Target Profit Margin (%)", "22.50%", "Global Rate"])
        w.writerow(["Financial Rates", "Default Material Markup (%)", "20.00%", "Global Rate"])
        w.writerow(["Financial Rates", "Default Labor Markup (%)", "35.00%", "Global Rate"])
        w.writerow(["Financial Rates", "Default Subcontractor Markup (%)", "15.00%", "Global Rate"])
        w.writerow(["Financial Rates", "Default Equipment Markup (%)", "25.00%", "Global Rate"])

    # Sheet 3: Estimate (Summary)
    with open(os.path.join(output_dir, "03_Estimate.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["CONSTRUCTION ESTIMATE & CONTRACTOR BID TEMPLATE"])
        w.writerow([])
        w.writerow(["Item #", "Cost Category & Item Specification", "Quantity", "Unit", "Unit Cost ($)", "Total Cost ($)", "Trade Source"])
        w.writerow(["1. DIRECT MATERIALS & FIXTURE PACKAGES"])
        w.writerow(["1.01", "Custom Cabinetry & Millwork Package", 1, "Lot", 11200.00, 11200.00, "Materials Sheet"])
        w.writerow(["1.02", "Calacatta Gold Quartz Countertops & Backsplash", 1, "Lot", 4800.00, 4800.00, "Materials Sheet"])
        w.writerow(["1.03", "Engineered Hardwood Flooring & Moisture Barrier", 1, "Lot", 3450.00, 3450.00, "Materials Sheet"])
        w.writerow(["1.04", "Recessed & Decorative LED Lighting Bundle", 1, "Lot", 1400.00, 1400.00, "Materials Sheet"])
        w.writerow(["1.05", "Architectural Interior Primers & Enamel Paints", 1, "Lot", 800.00, 800.00, "Materials Sheet"])
        w.writerow(["", "SUBTOTAL DIRECT MATERIALS ($)", "", "", "", 21650.00, ""])
        w.writerow([])
        w.writerow(["2. DIRECT LABOR & CRAFTSMANSHIP HOURS"])
        w.writerow(["2.01", "Selective Kitchen Demolition, Haul & Structural Prep", 1, "Lot", 3200.00, 3200.00, "Labor Sheet"])
        w.writerow(["2.02", "Master Cabinetry, Millwork & Island Installation", 1, "Lot", 6400.00, 6400.00, "Labor Sheet"])
        w.writerow(["2.03", "Drywall Framing, Hanging, Taping & Sanding", 1, "Lot", 3600.00, 3600.00, "Labor Sheet"])
        w.writerow(["2.04", "Precision Interior Spray Painting & Enamel Trim Finish", 1, "Lot", 3600.00, 3600.00, "Labor Sheet"])
        w.writerow(["", "SUBTOTAL DIRECT LABOR ($)", "", "", "", 16800.00, ""])
        w.writerow([])
        w.writerow(["3. LICENSED SUBCONTRACTOR SPECIALTY SERVICES"])
        w.writerow(["3.01", "Licensed Electrical Wiring & Service Panel Upgrade", 1, "Pkg", 4200.00, 4200.00, "Subcontractors"])
        w.writerow(["3.02", "Licensed Plumbing Gas/Water Reroute & Fixtures", 1, "Pkg", 3800.00, 3800.00, "Subcontractors"])
        w.writerow(["3.03", "Licensed HVAC Duct Modification & Range Hood Exhaust", 1, "Pkg", 1750.00, 1750.00, "Subcontractors"])
        w.writerow(["", "SUBTOTAL SUBCONTRACTORS ($)", "", "", "", 9750.00, ""])
        w.writerow([])
        w.writerow(["4. EQUIPMENT RENTALS & SPECIAL TOOLING ALLOWANCE"])
        w.writerow(["4.01", "20-Yard Roll-Off Waste Dumpster Rental & Removal", 1, "Lot", 850.00, 850.00, "Equipment Sheet"])
        w.writerow(["4.02", "HEPA Air Scrubber, Dust Containment & Dehumidifiers", 1, "Lot", 750.00, 750.00, "Equipment Sheet"])
        w.writerow(["4.03", "Scaffolding Towers, Lifts & Specialty Flooring Sanders", 1, "Lot", 1250.00, 1250.00, "Equipment Sheet"])
        w.writerow(["", "SUBTOTAL EQUIPMENT & RENTALS ($)", "", "", "", 2850.00, ""])
        w.writerow([])
        w.writerow(["5. PERMITS, FEES & SITE PROTECTION ALLOWANCE"])
        w.writerow(["5.01", "Municipal Building, Trade & Electrical Permit Fees", 1, "Lot", 950.00, 950.00, "Municipal"])
        w.writerow(["5.02", "Site Floor Protection, Dust Barriers & Sanitization", 1, "Lot", 500.00, 500.00, "Jobsite Prep"])
        w.writerow(["", "SUBTOTAL PERMITS & OTHER COSTS ($)", "", "", "", 1450.00, ""])
        w.writerow([])
        w.writerow(["6. CONTRACTOR BID PRICING SUMMARY & CONTRACT BID TOTALS"])
        w.writerow(["Total Direct Job Costs (Sum of Sections 1 to 5)", "", "", "", "", 52500.00, ""])
        w.writerow(["General Overhead Allowance (10.00% of Direct Costs)", "", "", "", "", 5250.00, ""])
        w.writerow(["Contingency Allowance (5.00% of Direct Costs)", "", "", "", "", 2625.00, ""])
        w.writerow(["Subtotal Total Estimated Job Cost (Before Markup)", "", "", "", "", 60375.00, ""])
        w.writerow(["Total Professional Markup & Contractor Gross Profit", "", "", "", "", 14177.50, ""])
        w.writerow(["Estimated Selling Price (Contract Bid Subtotal)", "", "", "", "", 74552.50, ""])
        w.writerow(["Preferred Client Contract Discount (-2.50%)", "", "", "", "", -1863.81, ""])
        w.writerow(["Net Taxable Bid Contract Amount", "", "", "", "", 72688.69, ""])
        w.writerow(["Estimated Sales Tax (8.25% on Net Taxable Amount)", "", "", "", "", 5996.82, ""])
        w.writerow(["GRAND TOTAL CONTRACT BID PRICE ($)", "", "", "", "", 78685.51, ""])

    # Sheet 4: Materials
    with open(os.path.join(output_dir, "04_Materials.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ITEMIZED MATERIALS COST SHEET"])
        w.writerow(["Item SKU", "Material Description & Specification", "Trade Category", "Supplier / Vendor", "Quantity", "Unit", "Unit Cost ($)", "Tax / Freight ($)", "Total Cost ($)"])
        w.writerow(["MAT-101", "Custom Shaker Base Cabinets", "Cabinetry", "Craftsman Cabinet Co.", 12, "Set", 700.00, 140.00, 8540.00])
        w.writerow(["MAT-102", "Custom Shaker Wall & Pantry Cabinets", "Cabinetry", "Craftsman Cabinet Co.", 8, "Set", 350.00, 70.00, 2870.00])
        w.writerow(["MAT-103", "Calacatta Gold Quartz Slab", "Countertops", "Apex Stone & Tile", 60, "SqFt", 75.00, 300.00, 4800.00])
        w.writerow(["MAT-104", "Engineered Oak Hardwood Flooring", "Flooring", "Prime Flooring Dist.", 450, "SqFt", 6.50, 225.00, 3150.00])
        w.writerow(["MAT-105", "Hardwood Underlayment", "Flooring", "Prime Flooring Dist.", 450, "SqFt", 0.65, 7.50, 300.00])
        w.writerow(["MAT-106", "Recessed LED Dimmable Fixture Bundle", "Electrical", "Lumens Direct", 16, "Unit", 45.00, 30.00, 750.00])
        w.writerow(["MAT-107", "Kitchen Pendant Lighting Array", "Electrical", "Lumens Direct", 3, "Unit", 220.00, 20.00, 680.00])
        w.writerow(["MAT-108", "Architectural Primer (5-Gal Pail)", "Paint", "Sherwin-Williams Pro", 4, "Pail", 45.00, 20.00, 200.00])
        w.writerow(["MAT-109", "Premium Satin Enamel (5-Gal Pail)", "Paint", "Sherwin-Williams Pro", 6, "Pail", 65.00, 30.00, 420.00])
        w.writerow(["MAT-110", "Kitchen Sink & Faucet", "Plumbing", "Ferguson Plumbing", 1, "Set", 950.00, 50.00, 1000.00])
        w.writerow(["MAT-111", "Ceramic Subway Backsplash Tile", "Tile", "Apex Stone & Tile", 80, "SqFt", 8.50, 40.00, 720.00])
        w.writerow(["MAT-112", "Drywall & Trim Lumber Package", "Lumber / Prep", "Home Depot Pro", 1, "Lot", 650.00, 50.00, 700.00])
        w.writerow(["", "TOTAL DIRECT MATERIAL COST ($)", "", "", "", "", "", "", 21650.00])

    # Sheet 5: Labor
    with open(os.path.join(output_dir, "05_Labor.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ITEMIZED LABOR COST SHEET"])
        w.writerow(["Task Code", "Craft Role & Task Description", "Crew Size", "Reg. Hours", "Reg. Rate ($/hr)", "OT Hours", "OT Rate ($/hr)", "Total Labor Cost ($)"])
        w.writerow(["LAB-201", "Selective Kitchen Demolition & Haul Prep", 2, 32, 40.00, 0, 60.00, 2560.00])
        w.writerow(["LAB-202", "Structural Header & Framing Modification", 2, 24, 55.00, 4, 82.50, 3300.00])
        w.writerow(["LAB-203", "Custom Cabinetry & Island Box Installation", 2, 40, 65.00, 0, 97.50, 5200.00])
        w.writerow(["LAB-204", "Finish Carpentry, Crown Molding & Trim", 1, 32, 60.00, 0, 90.00, 1920.00])
        w.writerow(["LAB-205", "Drywall Hanging, Taping & Sanding Prep", 2, 28, 45.00, 0, 67.50, 2520.00])
        w.writerow(["LAB-206", "Engineered Hardwood Floor Installation", 2, 24, 50.00, 0, 75.00, 2400.00])
        w.writerow(["LAB-207", "Ceramic Tile Backsplash Layout & Setting", 1, 16, 55.00, 0, 82.50, 880.00])
        w.writerow(["LAB-208", "Interior Spray Primer & Architectural Paint", 2, 36, 45.00, 0, 67.50, 3240.00])
        w.writerow(["LAB-209", "Project Supervisor & Quality Control Audit", 1, 24, 75.00, 0, 112.50, 1800.00])
        w.writerow(["LAB-210", "Jobsite Cleanup, Dust Control & Sanitizing", 1, 16, 35.00, 0, 52.50, 560.00])
        w.writerow(["", "TOTAL DIRECT LABOR COST ($)", "", "", "", "", "", 16800.00])

    # Sheet 6: Equipment
    with open(os.path.join(output_dir, "06_Equipment.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["EQUIPMENT RENTAL & SPECIAL TOOLING COST SHEET"])
        w.writerow(["Equip. ID", "Equipment Name & Specification", "Ownership Type", "Rate Period", "Rate ($)", "Qty / Days", "Delivery Fee ($)", "Total Cost ($)"])
        w.writerow(["EQP-301", "20-Yard Roll-Off Waste Dumpster Rental", "Rental", "Flat Rate", 650.00, 1, 200.00, 850.00])
        w.writerow(["EQP-302", "HEPA Commercial Air Scrubber & Filtration", "Owned / Allowance", "Daily Rate", 45.00, 10, 50.00, 500.00])
        w.writerow(["EQP-303", "Commercial Wet Tile Saw & Diamond Blade", "Owned / Allowance", "Daily Rate", 35.00, 4, 0.00, 140.00])
        w.writerow(["EQP-304", "Hardwood Floor Drum Sander & Edger Kit", "Rental", "Daily Rate", 85.00, 3, 50.00, 305.00])
        w.writerow(["EQP-305", "Interior Aluminum Scaffolding & Work Lift", "Rental", "Weekly Rate", 250.00, 2, 75.00, 575.00])
        w.writerow(["EQP-306", "Temporary Site Distribution Electrical Box", "Owned / Allowance", "Flat Rate", 120.00, 1, 0.00, 120.00])
        w.writerow(["EQP-307", "High-Capacity Dehumidifier & Blower Array", "Rental", "Daily Rate", 40.00, 5, 25.00, 225.00])
        w.writerow(["EQP-308", "Material Hoist & Cabinet Dolly Package", "Owned / Allowance", "Flat Rate", 150.00, 1, 0.00, 150.00])
        w.writerow(["", "TOTAL DIRECT EQUIPMENT & RENTAL COST ($)", "", "", "", "", "", 2850.00])

    # Sheet 7: Subcontractors
    with open(os.path.join(output_dir, "07_Subcontractors.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["LICENSED SUBCONTRACTOR PACKAGE & QUOTE SHEET"])
        w.writerow(["Sub. Code", "Subcontractor Company", "Scope of Work & Package Description", "Quoted Cost ($)", "Markup %", "Markup Amt ($)", "Billed Amount ($)", "License Verified"])
        w.writerow(["SUB-401", "Spark Systems Electric LLC", "200A Service Panel Upgrade & LED Wiring Rough/Finish", 4200.00, "15%", 630.00, 4830.00, "Yes - Verified"])
        w.writerow(["SUB-402", "Apex Mechanical Plumbing Co.", "Kitchen Sink Reroute, Gas Line Extension & Appliance Hookup", 3800.00, "15%", 570.00, 4370.00, "Yes - Verified"])
        w.writerow(["SUB-403", "Breeze Comfort HVAC LLC", "Range Hood Duct Reroute & Supply Register Diffusers", 1750.00, "15%", 262.50, 2012.50, "Yes - Verified"])
        w.writerow(["", "TOTAL SUBCONTRACTOR PACKAGE COSTS ($)", "", 9750.00, "", 1462.50, 11212.50, ""])

    # Sheet 8: Markup & Margin Calculator
    with open(os.path.join(output_dir, "08_Markup_Margin_Calculator.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ADVANCED MARKUP & PROFIT MARGIN CALCULATOR"])
        w.writerow([])
        w.writerow(["Cost Category", "Base Cost ($)", "Markup %", "Markup Amt ($)", "Selling Price ($)", "Category Share", "Profit Margin %"])
        w.writerow(["Direct Materials", 21650.00, "20%", 4330.00, 25980.00, "27.4%", "16.7%"])
        w.writerow(["Direct Labor", 16800.00, "35%", 5880.00, 22680.00, "24.2%", "25.9%"])
        w.writerow(["Equipment & Rentals", 2850.00, "25%", 712.50, 3562.50, "3.8%", "20.0%"])
        w.writerow(["Subcontractors", 9750.00, "15%", 1462.50, 11212.50, "12.0%", "13.0%"])
        w.writerow(["Permits & Other Costs", 1450.00, "15%", 217.50, 1667.50, "1.8%", "13.0%"])
        w.writerow(["Overhead & Contingency", 7875.00, "20%", 1575.00, 9450.00, "10.1%", "16.7%"])
        w.writerow(["TOTAL JOB MARKUP & MARGIN", 60375.00, "20.0%", 14177.50, 74552.50, "79.1%", "19.0%"])
        w.writerow([])
        w.writerow(["Markup %", "Equivalent Margin %", "Price on $1,000 Cost ($)", "Gross Profit ($)"])
        for pct in [10, 15, 20, 25, 30, 35, 40, 50, 75, 100]:
            w.writerow([f"{pct}%", f"{round(pct/(100+pct)*100,1)}%", round(1000*(1+pct/100),2), round(1000*pct/100,2)])
        w.writerow([])
        w.writerow(["Target Margin Reverse Calculator"])
        w.writerow(["Input Total Job Cost ($):", 60375.00, "Cost to price"])
        w.writerow(["Desired Profit Margin (%):", "22.50%", "Target %"])
        w.writerow(["Required Selling Price ($):", 77935.48, "Formula"])
        w.writerow(["Required Gross Profit ($):", 17560.48, "Profit $"])
        w.writerow(["Equivalent Markup % Needed:", "29.0%", "Equivalent"])

    # Sheet 9: Job Costing & Actuals
    with open(os.path.join(output_dir, "09_Job_Costing_Actuals.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ESTIMATED VS. ACTUAL JOB COSTING TRACKER & VARIANCE ANALYSIS"])
        w.writerow([])
        w.writerow(["Category Code", "Cost Category Description", "Estimated Budget ($)", "Actual Cost Incurred ($)", "Dollar Variance ($)", "Budget Variance %", "Cost Performance Status"])
        w.writerow(["CAT-100", "Direct Materials & Fixtures", 21650.00, 20950.00, 700.00, "3.2%", "UNDER BUDGET"])
        w.writerow(["CAT-200", "Direct Labor & Craft Hours", 16800.00, 17200.00, -400.00, "-2.4%", "OVER BUDGET"])
        w.writerow(["CAT-300", "Licensed Subcontractors", 9750.00, 9750.00, 0.00, "0.0%", "UNDER BUDGET"])
        w.writerow(["CAT-400", "Equipment Rentals & Special Tools", 2850.00, 2650.00, 200.00, "7.0%", "UNDER BUDGET"])
        w.writerow(["CAT-500", "Permits, Fees & Site Protection", 1450.00, 1450.00, 0.00, "0.0%", "UNDER BUDGET"])
        w.writerow(["CAT-600", "General Overhead & Contingency Allowance", 7875.00, 7500.00, 375.00, "4.8%", "UNDER BUDGET"])
        w.writerow(["TOTAL CONSTRUCTION JOB COST ($)", "", 60375.00, 59500.00, 875.00, "1.4%", "UNDER BUDGET"])
        w.writerow([])
        w.writerow(["EXECUTIVE JOB PROFITABILITY & MARGIN VARIANCE SUMMARY"])
        w.writerow(["Contract Bid Selling Price ($)", 78685.51])
        w.writerow(["Actual Total Job Cost Incurred ($)", 59500.00])
        w.writerow(["Actual Job Gross Profit ($)", 19185.51])
        w.writerow(["Actual Realized Profit Margin (%)", "24.4%"])
        w.writerow(["Estimated Profit Margin (%)", "19.0%"])
        w.writerow(["Profit Margin Variance (%)", "5.4%"])

    # Sheet 10: Summary Dashboard
    with open(os.path.join(output_dir, "10_Summary_Dashboard.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["EXECUTIVE CONSTRUCTION SUMMARY DASHBOARD & KPI OVERVIEW"])
        w.writerow([])
        w.writerow(["TOTAL ESTIMATED JOB COST", "", "TOTAL CONTRACT BID PRICE", "", "TOTAL MARKUP & GROSS PROFIT", "", "ESTIMATED PROFIT MARGIN %"])
        w.writerow(["$60,375.00", "", "$78,685.51", "", "$14,177.50", "", "19.0%"])
        w.writerow([])
        w.writerow(["COST PERFORMANCE BY CATEGORY ($)"])
        w.writerow(["Cost Category", "Estimated ($)", "Actual ($)", "Variance ($)", "Status"])
        w.writerow(["Direct Materials", 21650.00, 20950.00, 700.00, "UNDER BUDGET"])
        w.writerow(["Direct Labor", 16800.00, 17200.00, -400.00, "OVER BUDGET"])
        w.writerow(["Subcontractors", 9750.00, 9750.00, 0.00, "UNDER BUDGET"])
        w.writerow(["Equipment & Rentals", 2850.00, 2650.00, 200.00, "UNDER BUDGET"])
        w.writerow(["Permits & Other", 1450.00, 1450.00, 0.00, "UNDER BUDGET"])
        w.writerow(["Overhead & Contingency", 7875.00, 7500.00, 375.00, "UNDER BUDGET"])
        w.writerow(["TOTAL JOB COSTS", 60375.00, 59500.00, 875.00, "UNDER BUDGET"])
        w.writerow([])
        w.writerow(["CONTRACT BID & PROFIT BREAKDOWN", "", ""])
        w.writerow(["Subtotal Direct Costs", 52500.00])
        w.writerow(["Overhead Allowance", 5250.00])
        w.writerow(["Contingency Allowance", 2625.00])
        w.writerow(["Total Professional Markup", 14177.50])
        w.writerow(["Subtotal Bid Selling Price", 74552.50])
        w.writerow(["Preferred Client Discount", -1863.81])
        w.writerow(["Sales Tax (8.25%)", 5996.82])
        w.writerow(["GRAND TOTAL BID PRICE ($)", 78685.51])
        w.writerow([])
        w.writerow(["COST PERFORMANCE VISUAL SNAPSHOT (EXCEL-SAFE CELL-BASED DASHBOARD)"])
        w.writerow(["Cost Category", "Estimated ($)", "Actual ($)", "Variance ($)", "Budget Perf.", "Visual Bar (Est.)", "Visual Bar (Act.)"])
        w.writerow(["Direct Materials", 21650.00, 20950.00, 700.00, "ON BUDGET", "████████████████████████████████████████████", "█████████████████████████████████████████"])
        w.writerow(["Direct Labor", 16800.00, 17200.00, -400.00, "OVER BUDGET", "██████████████████████████████████", "█████████████████████████████████████"])
        # etc.

    # Sheet 11: Client Proposal
    with open(os.path.join(output_dir, "11_Client_Proposal.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["CONSTRUCTION PROJECT PROPOSAL & BID ACCEPTANCE"])
        w.writerow([])
        w.writerow(["CONTRACTOR / COMPANY INFO", "", "", "CLIENT / PROJECT INFO"])
        w.writerow(["Company Name:", "Novality Construction & Design LLC", "", "Client Name:", "Robert Sterling / Apex Properties"])
        w.writerow(["Contact Person:", "David Novality, Master Contractor", "", "Project Name:", "Luxury Kitchen & Living Area Remodel"])
        w.writerow(["Phone Number:", "(555) 382-9100", "", "Job Location:", "742 Evergreen Terrace, Suite 300"])
        w.writerow(["Email Address:", "estimating@novalitystore.com", "", "Estimate Number:", "EST-2026-042"])
        w.writerow([])
        w.writerow(["SCOPE OF WORK SPECIFICATION"])
        w.writerow(["Full kitchen demolition, custom cabinetry, quartz countertops, electrical & plumbing relocation, hardwood flooring, and ambient LED lighting."])
        w.writerow([])
        w.writerow(["PROJECT PRICING SUMMARY BY TRADE & SERVICE"])
        w.writerow(["Cost Category / Trade Description", "", "", "", "Contract Price ($)"])
        w.writerow(["Direct Materials & Custom Architectural Fixtures", "", "", "", 21650.00])
        w.writerow(["Direct Labor & Skilled Craftsmanship Hours", "", "", "", 16800.00])
        w.writerow(["Licensed Specialty Subcontractor Packages", "", "", "", 9750.00])
        w.writerow(["Equipment Rentals, Waste Removal & Specialty Tooling", "", "", "", 2850.00])
        w.writerow(["Municipal Building & Trade Permits, Site Protection & Cleanup", "", "", "", 1450.00])
        w.writerow(["Contractor Overhead, Management & Contingency Allowance", "", "", "", 7875.00])
        w.writerow(["Professional Contractor Markup & Value Engineering", "", "", "", 14177.50])
        w.writerow(["SUBTOTAL CONTRACT BID PRICE", "", "", "", 74552.50])
        w.writerow(["Preferred Client Discount (-2.50%)", "", "", "", -1863.81])
        w.writerow(["Net Taxable Contract Bid Amount", "", "", "", 72688.69])
        w.writerow(["Applicable Sales Tax (8.25%)", "", "", "", 5996.82])
        w.writerow(["GRAND TOTAL CONTRACT BID PRICE ($)", "", "", "", 78685.51])
        w.writerow([])
        w.writerow(["EXCLUSIONS, CLARIFICATIONS & PAYMENT TERMS"])
        w.writerow(["EXCLUSIONS: Excludes hazardous material abatement, unforeseen subterranean repairs, and utility company connection fees. PAYMENT SCHEDULE: 20% deposit upon contract signing, 50% upon progress milestones, and 30% upon substantial completion. PROPOSAL VALIDITY: 30 days from date of issue."])

    # Sheet 12: Terms & Conditions
    with open(os.path.join(output_dir, "12_Terms_Conditions.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["STANDARD CONTRACTOR TERMS, CONDITIONS & WORKMANSHIP WARRANTY"])
        w.writerow([])
        w.writerow(["1. SCOPE OF WORK & PROPOSAL VALIDITY", "The contractor agrees to perform the scope of work specified in the Client Proposal. The pricing stated in this document is valid for 30 calendar days from the date of issuance. Any work not explicitly listed is excluded."])
        w.writerow(["2. PAYMENT SCHEDULE & INVOICING", "Payments shall be made according to the agreed milestone payment schedule: 20% mobilization deposit upon signing, milestone progress invoices during construction, and final payment upon completion of punchlist items."])
        w.writerow(["3. CHANGE ORDERS & MODIFICATIONS", "Any alteration or deviation from the original specifications involving extra costs or schedule extensions will be executed only upon a written Change Order signed by both Owner and Contractor."])
        w.writerow(["4. SITE ACCESS, WORKING HOURS & UTILITIES", "The owner agrees to provide reasonable access to the property during standard working hours (7:00 AM - 5:00 PM) and make water and electricity available for construction purposes without charge."])
        w.writerow(["5. WORKMANSHIP WARRANTY & MATERIAL GUARANTEE", "Novality Construction & Design LLC provides a 1-year limited workmanship warranty on all installations. Standard manufacturer warranties apply to all installed equipment and material fixtures."])
        w.writerow(["6. INSURANCE & LICENSING", "The contractor shall maintain general liability and workers' compensation insurance throughout the project duration and hold valid municipal trade licenses."])
        w.writerow(["7. DELAYS & FORCE MAJEURE", "The contractor is not responsible for project delays caused by severe weather, material vendor backorders, strikes, or municipal permitting delays beyond contractor control."])

    # Sheet 13: Material Price DB
    with open(os.path.join(output_dir, "13_Material_Price_DB.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["MATERIAL PRICE DATABASE"])
        w.writerow(["SKU Code", "Material Description & Specification", "Trade Category", "Default Supplier", "Unit", "Standard Price ($)", "Lead Time (Days)"])
        materials_db = [
            ("MAT-101", "Custom Shaker Base Cabinet Box", "Cabinetry", "Craftsman Cabinet Co.", "Set", 700.00, 14),
            ("MAT-102", "Custom Shaker Wall Cabinet Box", "Cabinetry", "Craftsman Cabinet Co.", "Set", 350.00, 14),
            ("MAT-103", "Calacatta Gold Quartz Slab (3cm)", "Countertops", "Apex Stone & Tile", "SqFt", 75.00, 7),
            ("MAT-104", "Engineered Oak Hardwood Plank", "Flooring", "Prime Flooring Dist.", "SqFt", 6.50, 5),
            ("MAT-105", "Hardwood Acoustic Underlayment", "Flooring", "Prime Flooring Dist.", "SqFt", 0.65, 3),
            ("MAT-106", "Recessed 6-inch LED Dimmable Fixture", "Electrical", "Lumens Direct", "Unit", 45.00, 2),
            ("MAT-107", "Decorative Pendant Light Fixture", "Electrical", "Lumens Direct", "Unit", 220.00, 5),
            ("MAT-108", "Architectural Primer (5-Gal Pail)", "Paint", "Sherwin-Williams Pro", "Pail", 45.00, 1),
            ("MAT-109", "Premium Satin Enamel (5-Gal Pail)", "Paint", "Sherwin-Williams Pro", "Pail", 65.00, 1),
            ("MAT-110", "Commercial Stainless Sink & Faucet", "Plumbing", "Ferguson Plumbing", "Set", 950.00, 7),
            ("MAT-111", "Ceramic Subway Backsplash Tile", "Tile", "Apex Stone & Tile", "SqFt", 8.50, 4),
            ("MAT-112", "Drywall Sheet 1/2-inch x 8ft", "Drywall", "Home Depot Pro", "Sheet", 14.50, 1),
            ("MAT-113", "All-Purpose Joint Compound Box", "Drywall", "Home Depot Pro", "Box", 18.00, 1),
            ("MAT-114", "Pressure Treated 2x4 Lumber Stud", "Framing", "Home Depot Pro", "Unit", 6.25, 1),
            ("MAT-115", "Interior Mold-Resistant Silicone Caulk", "Finish", "Home Depot Pro", "Tube", 8.50, 1),
        ]
        for row in materials_db:
            w.writerow(row)

    # Sheet 14: Labor Rate DB
    with open(os.path.join(output_dir, "14_Labor_Rate_DB.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["LABOR RATE DATABASE"])
        w.writerow(["Trade Code", "Craft Role / Trade Description", "Base Wage ($/hr)", "Burden & Ins. %", "Burdened Cost ($/hr)", "Billing Rate ($/hr)", "OT Billing Rate ($/hr)"])
        labor_db = [
            ("TRD-01", "General Construction Laborer", 25.00, "25%", 31.25, 40.00, 60.00),
            ("TRD-02", "Skilled Carpenter & Framing Lead", 35.00, "28%", 44.80, 55.00, 82.50),
            ("TRD-03", "Master Millwork & Cabinet Specialist", 42.00, "28%", 53.76, 65.00, 97.50),
            ("TRD-04", "Drywall Hanger & Finisher", 30.00, "25%", 37.50, 45.00, 67.50),
            ("TRD-05", "Interior Painter & Trim Finisher", 28.00, "25%", 35.00, 45.00, 67.50),
            ("TRD-06", "Licensed Electrician (Master)", 48.00, "30%", 62.40, 75.00, 112.50),
            ("TRD-07", "Licensed Plumber & Gas Fitter", 48.00, "30%", 62.40, 75.00, 112.50),
            ("TRD-08", "HVAC Duct & Mechanical Technician", 42.00, "30%", 54.60, 70.00, 105.00),
            ("TRD-09", "Ceramic & Stone Tile Setter", 35.00, "26%", 44.10, 55.00, 82.50),
            ("TRD-10", "Hardwood Flooring Installer", 32.00, "26%", 40.32, 50.00, 75.00),
            ("TRD-11", "Project Supervisor / QC Manager", 50.00, "32%", 66.00, 75.00, 112.50),
            ("TRD-12", "Site Safety & Sanitization Tech", 22.00, "25%", 27.50, 35.00, 52.50),
        ]
        for row in labor_db:
            w.writerow(row)

    # Sheet 15: Estimate Versions
    with open(os.path.join(output_dir, "15_Estimate_Versions.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["MULTIPLE ESTIMATE VERSIONS COMPARISON"])
        w.writerow(["Cost Category / Metric", "Option 1: Standard / Base", "Option 2: Upgraded / Preferred", "Option 3: Premium / Luxury"])
        w.writerow(["Direct Materials & Fixture Packages ($)", 16500.00, 21650.00, 29800.00])
        w.writerow(["Direct Labor & Craft Hours ($)", 14200.00, 16800.00, 21500.00])
        w.writerow(["Licensed Subcontractor Services ($)", 8500.00, 9750.00, 12400.00])
        w.writerow(["Equipment Rentals & Tool Allowance ($)", 2200.00, 2850.00, 3600.00])
        w.writerow(["Permits, Fees & Site Protection ($)", 1450.00, 1450.00, 1850.00])
        w.writerow(["TOTAL DIRECT JOB COSTS ($)", 42850.00, 52500.00, 69150.00])
        w.writerow(["Overhead Allowance (10.0%)", 4285.00, 5250.00, 6915.00])
        w.writerow(["Contingency Allowance (5.0%)", 2142.50, 2625.00, 3457.50])
        w.writerow(["SUBTOTAL JOB COST (Before Markup)", 49277.50, 60375.00, 79522.50])
        w.writerow(["Contractor Markup % Applied", "22%", "23.48%", "25%"])
        w.writerow(["Total Markup Amount ($)", 10841.05, 14177.50, 19880.63])
        w.writerow(["ESTIMATED CONTRACT SELLING PRICE ($)", 60118.55, 74552.50, 99403.13])
        w.writerow(["Equivalent Profit Margin %", "18.0%", "19.0%", "20.0%"])

    # Sheet 16: Change Order Tracker
    with open(os.path.join(output_dir, "16_Change_Order_Tracker.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["PROJECT CHANGE ORDER TRACKER & BID MODIFICATION LOG"])
        w.writerow(["CO Number", "Date Req.", "Description of Scope Modification", "Requested By", "Direct Cost ($)", "Markup %", "CO Total Price ($)", "Schedule Add", "Approval Status", "Date Approved"])
        w.writerow(["CO-001", "2026-08-05", "Upgrade Quartz Island Countertop to Waterfall Edge", "Robert Sterling", 1200.00, "20%", 1440.00, "2 Days", "Approved", "2026-08-06"])
        w.writerow(["CO-002", "2026-08-10", "Add 4 Additional Dimmable Recessed LED Lights", "Robert Sterling", 450.00, "25%", 562.50, "1 Day", "Approved", "2026-08-11"])
        w.writerow(["CO-003", "2026-08-15", "Relocate Secondary Laundry Plumbing Drain Line", "Robert Sterling", 850.00, "20%", 1020.00, "1 Day", "Pending", "-"])
        w.writerow(["TOTAL CHANGE ORDER LOG ($)", "", "", "", 2500.00, "", 3022.50, "", "", ""])
        w.writerow([])
        w.writerow(["CHANGE ORDER CONTRACT PRICE IMPACT"])
        w.writerow(["Original Contract Bid Price ($)", 78685.51])
        w.writerow(["Total Approved Change Orders ($)", 2002.50])
        w.writerow(["REVISED GRAND TOTAL CONTRACT PRICE ($)", 80688.01])

    # Sheet 17: Payment Schedule
    with open(os.path.join(output_dir, "17_Payment_Schedule.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["MILESTONE PAYMENT SCHEDULE & PROGRESS BILLING TRACKER"])
        w.writerow(["Milestone #", "Milestone Phase Description", "% of Contract", "Scheduled Amount ($)", "Target Due Date", "Payment Status", "Date Received", "Check / Ref #"])
        w.writerow(["PAY-01", "Initial Mobilization & Deposit (Upon Contract Signing)", "20%", 15737.10, "2026-08-01", "Paid", "2026-08-01", "CHK-9921"])
        w.writerow(["PAY-02", "Rough-In Completion (Electrical, Plumbing & HVAC)", "25%", 19671.38, "2026-08-15", "Paid", "2026-08-16", "CHK-9945"])
        w.writerow(["PAY-03", "Drywall, Flooring & Cabinet Installation Completion", "25%", 19671.38, "2026-08-30", "Pending", "-", "-"])
        w.writerow(["PAY-04", "Countertops, Backsplash Tile & Fixture Completion", "20%", 15737.10, "2026-09-10", "Pending", "-", "-"])
        w.writerow(["PAY-05", "Final Inspection & Punchlist Sign-Off", "10%", 7868.55, "2026-09-20", "Pending", "-", "-"])
        w.writerow(["TOTAL CONTRACT PAYMENT SCHEDULE", "", "100%", 78685.51, "", "", "", ""])

    # Sheet 18: Invoice Summary
    with open(os.path.join(output_dir, "18_Invoice_Summary.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["PROGRESS INVOICE & CONTRACT BILLING SUMMARY"])
        w.writerow([])
        w.writerow(["CONTRACTOR / PAYEE DETAILS", "", "", "INVOICE & CONTRACT METRICS"])
        w.writerow(["Company Name:", "Novality Construction & Design LLC", "", "Invoice Number:", "INV-2026-001"])
        w.writerow(["Client / Owner:", "Robert Sterling / Apex Properties", "", "Billing Period:", "2026-08-01 to 2026-08-16"])
        w.writerow([])
        w.writerow(["PROGRESS BILLING FINANCIAL BREAKDOWN", "", ""])
        w.writerow(["1. Original Contract Bid Price ($)", "", "", 78685.51])
        w.writerow(["2. Net Addition from Approved Change Orders ($)", "", "", 2002.50])
        w.writerow(["3. REVISED TOTAL CONTRACT PRICE ($)", "", "", 80688.01])
        w.writerow(["4. Total Value of Work Completed to Date ($)", "", "", 36309.60])
        w.writerow(["5. Retainage Withheld (5.00% of Completed Work)", "", "", 1815.48])
        w.writerow(["6. Total Earned Less Retainage ($)", "", "", 34494.12])
        w.writerow(["7. Less Previous Payments Received ($)", "", "", 15737.10])
        w.writerow(["8. CURRENT PROGRESS INVOICE AMOUNT DUE ($)", "", "", 18757.02])
        w.writerow(["9. Balance to Finish Including Retainage ($)", "", "", 46193.89])

    # Sheet 19: Project Timeline
    with open(os.path.join(output_dir, "19_Project_Timeline.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["CONSTRUCTION PROJECT TIMELINE & GANTT SCHEDULE TRACKER (12 WEEKS)"])
        w.writerow(["Phase #", "Construction Phase & Task", "Lead Trade", "Status", "Start Wk", "Duration", "End Wk", "W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8", "W9", "W10", "W11", "W12"])
        phases_data = [
            ("PH-01", "Permitting & Material Procurement", "Management", "Completed", 1, 2, 2, "█", "█", "", "", "", "", "", "", "", "", "", "", ""),
            ("PH-02", "Selective Kitchen Demolition & Site Prep", "Demolition", "Completed", 2, 1, 2, "", "█", "", "", "", "", "", "", "", "", "", "", ""),
            ("PH-03", "Structural Framing & Wall Header Reroute", "Carpentry", "Completed", 3, 1, 3, "", "", "█", "", "", "", "", "", "", "", "", "", ""),
            ("PH-04", "Electrical & Plumbing Rough-In Wiring/Pipes", "MEP Subs", "Completed", 3, 2, 4, "", "", "█", "█", "", "", "", "", "", "", "", "", ""),
            ("PH-05", "Municipal Inspection (Rough MEP & Frame)", "Municipal", "Completed", 5, 1, 5, "", "", "", "", "█", "", "", "", "", "", "", "", ""),
            ("PH-06", "Drywall Hanging, Taping & Plaster Sanding", "Drywall", "In Progress", 5, 2, 6, "", "", "", "", "█", "█", "", "", "", "", "", "", ""),
            ("PH-07", "Engineered Hardwood Floor Installation", "Flooring", "Scheduled", 7, 1, 7, "", "", "", "", "", "", "█", "", "", "", "", "", ""),
            ("PH-08", "Custom Cabinetry & Kitchen Island Box Install", "Carpentry", "Scheduled", 8, 2, 9, "", "", "", "", "", "", "", "█", "█", "", "", "", ""),
            ("PH-09", "Quartz Countertop Slab & Tile Backsplash", "Stone / Tile", "Scheduled", 9, 2, 10, "", "", "", "", "", "", "", "", "█", "█", "", "", ""),
            ("PH-10", "Interior Paint, Trim Molding & Enamel Finish", "Painting", "Scheduled", 10, 2, 11, "", "", "", "", "", "", "", "", "", "█", "█", "", ""),
            ("PH-11", "MEP Fixtures, Faucet & LED Trim Finish", "MEP Subs", "Scheduled", 11, 1, 11, "", "", "", "", "", "", "", "", "", "", "█", "", ""),
            ("PH-12", "Final Cleanup, Punchlist & Owner Sign-Off", "Management", "Scheduled", 12, 1, 12, "", "", "", "", "", "", "", "", "", "", "", "█"),
        ]
        for row in phases_data:
            w.writerow(row)

    print(f"[SUCCESS] Exported all 19 Google Sheets compatible CSV files into '{output_dir}/'")


if __name__ == "__main__":
    export_google_sheets_csvs()

# Also create a convenience script runner
if __name__ == "__main__":
    export_google_sheets_csvs()