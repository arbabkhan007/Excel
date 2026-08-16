#!/usr/bin/env python3
"""
Comprehensive School Management System - Excel Generator
=========================================================
Creates a fully-structured multi-sheet Excel workbook covering:
  - Student Management (Master, Parent, Medical)
  - Staff / HR Management (Master, Attendance, Payroll)
  - Academic Management (Class, Subjects, Timetable, Syllabus)
  - Attendance Management (Daily, Monthly, Dashboard)
  - Examination & Grade Management (Setup, Marks, Grades, Report Card, Analytics)
  - Fee Management (Structure, Collection, Defaulters, Dashboard)
  - Library Management (Inventory, Issue/Return, Dashboard)
  - Transport Management (Vehicle, Route, Student Mapping)
  - Inventory & Asset Management (Register, Stationery)
  - Communication & Notices (Notice Board, Parent Log)
  - Discipline & Behaviour (Incidents, Awards)
  - Hostel Management (Room Allocation, Mess)
  - Health & Infirmary (Records, Medicine)
  - Events & Calendar (Academic Calendar, Event Management)
  - Financial Management (Income, Expense, Budget, P&L)
  - Admission Management (Enquiry, Admission Tracker)
  - Transfer Certificate Register
  - Main Dashboard (Summary)
  - Config (Dropdown lists, Grade Tables, System Settings)

Author  : Arena Agent
Version : 1.0
"""

import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, NamedStyle, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from copy import copy

# ── Colour Palette ──────────────────────────────────────────────────────────
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
GREY        = "F2F2F2"
HEADER_GREY = "D9E2F3"

# ── Reusable Styles ─────────────────────────────────────────────────────────
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

header_font = Font(name='Calibri', bold=True, color=WHITE, size=11)
header_fill = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type='solid')
sub_header_font = Font(name='Calibri', bold=True, color=DARK_BLUE, size=11)
sub_header_fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
wrap_align = Alignment(wrap_text=True, vertical='top')

title_font = Font(name='Calibri', bold=True, color=WHITE, size=16)
title_fill = PatternFill(start_color=MID_BLUE, end_color=MID_BLUE, fill_type='solid')

green_fill = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type='solid')
red_fill = PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')
yellow_cond_fill = PatternFill(start_color=YELLOW_FILL, end_color=YELLOW_FILL, fill_type='solid')

kpi_font = Font(name='Calibri', bold=True, size=28, color=DARK_BLUE)
kpi_label_font = Font(name='Calibri', bold=True, size=11, color=MID_BLUE)
section_font = Font(name='Calibri', bold=True, size=13, color=DARK_BLUE)
section_fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')


# ═══════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════

def style_header_row(ws, row, max_col, font=None, fill=None):
    """Apply header styling to an entire row."""
    f = font or header_font
    fi = fill or header_fill
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = f
        cell.fill = fi
        cell.alignment = center_align
        cell.border = thin_border


def auto_width(ws, max_col, min_width=10, max_width=35):
    """Set reasonable column widths based on header text length."""
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        max_len = min_width
        for row in ws.iter_rows(min_col=col, max_col=col, values_only=False):
            for cell in row:
                if cell.value:
                    max_len = max(max_len, min(len(str(cell.value)) + 4, max_width))
        ws.column_dimensions[letter].width = max_len


def add_title_row(ws, title, max_col):
    """Merge and style a title across the top of a sheet."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    cell = ws.cell(row=1, column=1, value=title)
    cell.font = title_font
    cell.fill = title_fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 36


def add_data_validation(ws, cell_range, formula_list):
    """Add a dropdown data-validation to a range."""
    dv = DataValidation(type="list", formula1=f'"{formula_list}"', allow_blank=True)
    dv.error = "Please select a valid option"
    dv.errorTitle = "Invalid Entry"
    ws.add_data_validation(dv)
    dv.add(cell_range)


def apply_borders(ws, start_row, end_row, max_col):
    """Apply thin borders to a block of cells."""
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).border = thin_border


def write_headers(ws, headers, row=2):
    """Write a list of headers starting at the given row (default 2, after title)."""
    for col_idx, h in enumerate(headers, 1):
        ws.cell(row=row, column=col_idx, value=h)
    style_header_row(ws, row, len(headers))
    return len(headers)


def freeze_and_filter(ws, freeze_row, max_col):
    """Freeze panes and add auto-filter."""
    ws.freeze_panes = ws.cell(row=freeze_row + 1, column=1)
    ws.auto_filter.ref = f"A{freeze_row}:{get_column_letter(max_col)}{freeze_row + 500}"


# ═══════════════════════════════════════════════════════════════════════════
# MAIN BUILDER
# ═══════════════════════════════════════════════════════════════════════════

def build_workbook():
    wb = Workbook()

    # ─── We will remove the default sheet at the end ─────────────────────
    default_sheet = wb.active

    # =====================================================================
    # 0. CONFIG / HIDDEN SHEETS
    # =====================================================================
    ws = wb.create_sheet("Config_Dropdowns")
    _build_config_dropdowns(ws)

    # =====================================================================
    # 1. STUDENT MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Student Master Data")
    _build_student_master(ws)

    ws = wb.create_sheet("Parent Information")
    _build_parent_info(ws)

    ws = wb.create_sheet("Student Medical")
    _build_student_medical(ws)

    # =====================================================================
    # 2. STAFF / TEACHER MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Staff Master Data")
    _build_staff_master(ws)

    ws = wb.create_sheet("Staff Attendance")
    _build_staff_attendance(ws)

    ws = wb.create_sheet("Staff Payroll")
    _build_staff_payroll(ws)

    # =====================================================================
    # 3. ACADEMIC MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Class Section Setup")
    _build_class_setup(ws)

    ws = wb.create_sheet("Subject Master")
    _build_subject_master(ws)

    ws = wb.create_sheet("Timetable")
    _build_timetable(ws)

    ws = wb.create_sheet("Syllabus Tracker")
    _build_syllabus_tracker(ws)

    # =====================================================================
    # 4. ATTENDANCE MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Daily Attendance")
    _build_daily_attendance(ws)

    ws = wb.create_sheet("Monthly Attendance Summary")
    _build_monthly_attendance(ws)

    ws = wb.create_sheet("Attendance Dashboard")
    _build_attendance_dashboard(ws)

    # =====================================================================
    # 5. EXAMINATION & GRADE MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Exam Setup")
    _build_exam_setup(ws)

    ws = wb.create_sheet("Marks Entry")
    _build_marks_entry(ws)

    ws = wb.create_sheet("Grade Configuration")
    _build_grade_config(ws)

    ws = wb.create_sheet("Report Card")
    _build_report_card(ws)

    ws = wb.create_sheet("Exam Analytics")
    _build_exam_analytics(ws)

    # =====================================================================
    # 6. FEE MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Fee Structure")
    _build_fee_structure(ws)

    ws = wb.create_sheet("Fee Collection")
    _build_fee_collection(ws)

    ws = wb.create_sheet("Fee Defaulters")
    _build_fee_defaulters(ws)

    ws = wb.create_sheet("Fee Dashboard")
    _build_fee_dashboard(ws)

    # =====================================================================
    # 7. LIBRARY MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Book Inventory")
    _build_book_inventory(ws)

    ws = wb.create_sheet("Book Issue Return")
    _build_book_issue_return(ws)

    ws = wb.create_sheet("Library Dashboard")
    _build_library_dashboard(ws)

    # =====================================================================
    # 8. TRANSPORT MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Vehicle Master")
    _build_vehicle_master(ws)

    ws = wb.create_sheet("Route Management")
    _build_route_management(ws)

    ws = wb.create_sheet("Student Transport")
    _build_student_transport(ws)

    # =====================================================================
    # 9. INVENTORY & ASSET MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Asset Register")
    _build_asset_register(ws)

    ws = wb.create_sheet("Stationery Stock")
    _build_stationery_stock(ws)

    # =====================================================================
    # 10. COMMUNICATION & NOTICES
    # =====================================================================
    ws = wb.create_sheet("Notice Board")
    _build_notice_board(ws)

    ws = wb.create_sheet("Parent Communication")
    _build_parent_communication(ws)

    # =====================================================================
    # 11. DISCIPLINE & BEHAVIOUR
    # =====================================================================
    ws = wb.create_sheet("Incident Log")
    _build_incident_log(ws)

    ws = wb.create_sheet("Student Awards")
    _build_student_awards(ws)

    # =====================================================================
    # 12. HOSTEL MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Hostel Room Allocation")
    _build_hostel_rooms(ws)

    ws = wb.create_sheet("Mess Food Management")
    _build_mess_management(ws)

    # =====================================================================
    # 13. HEALTH & INFIRMARY
    # =====================================================================
    ws = wb.create_sheet("Health Records")
    _build_health_records(ws)

    ws = wb.create_sheet("Medical Inventory")
    _build_medical_inventory(ws)

    # =====================================================================
    # 14. EVENTS & CALENDAR
    # =====================================================================
    ws = wb.create_sheet("Academic Calendar")
    _build_academic_calendar(ws)

    ws = wb.create_sheet("Event Management")
    _build_event_management(ws)

    # =====================================================================
    # 15. FINANCIAL MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Income Tracker")
    _build_income_tracker(ws)

    ws = wb.create_sheet("Expense Tracker")
    _build_expense_tracker(ws)

    ws = wb.create_sheet("Budget Planning")
    _build_budget_planning(ws)

    ws = wb.create_sheet("Profit Loss Summary")
    _build_profit_loss(ws)

    # =====================================================================
    # 16. ADMISSION MANAGEMENT
    # =====================================================================
    ws = wb.create_sheet("Enquiry Tracker")
    _build_enquiry_tracker(ws)

    ws = wb.create_sheet("Admission Tracker")
    _build_admission_tracker(ws)

    # =====================================================================
    # 17. TRANSFER CERTIFICATE
    # =====================================================================
    ws = wb.create_sheet("TC Register")
    _build_tc_register(ws)

    # =====================================================================
    # MAIN DASHBOARD (created last so it can be moved to front)
    # =====================================================================
    ws = wb.create_sheet("DASHBOARD", 1)  # insert at position 1
    _build_dashboard(ws, wb)

    # ── Remove default sheet ──────────────────────────────────────────────
    wb.remove(default_sheet)

    # ── Hide config sheet ─────────────────────────────────────────────────
    wb["Config_Dropdowns"].sheet_state = 'hidden'

    return wb


# ═══════════════════════════════════════════════════════════════════════════
# INDIVIDUAL SHEET BUILDERS
# ═══════════════════════════════════════════════════════════════════════════

# ─── CONFIG ──────────────────────────────────────────────────────────────
def _build_config_dropdowns(ws):
    """Hidden sheet holding dropdown lists for data validation."""
    add_title_row(ws, "⚙️ CONFIG – Dropdown Lists & System Settings", 4)

    lists = {
        "Gender": ["Male", "Female", "Other"],
        "BloodGroup": ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"],
        "Category": ["General", "OBC", "SC", "ST", "EWS"],
        "StudentStatus": ["Active", "Inactive", "Transferred", "Graduated"],
        "CommPref": ["SMS", "Email", "WhatsApp"],
        "AttendanceStatus": ["Present", "Absent", "Late", "Half-day", "Excused"],
        "LeaveType": ["Casual", "Sick", "Earned", "Maternity", "Unpaid"],
        "StaffStatus": ["Active", "On Leave", "Resigned", "Retired"],
        "EmploymentType": ["Full-time", "Part-time", "Contract"],
        "Designation": ["Teacher", "Admin", "Support Staff", "Principal", "Vice-Principal", "Coordinator"],
        "SubjectType": ["Core", "Elective", "Co-curricular"],
        "ExamType": ["Internal", "External", "Practical", "Oral"],
        "FeeType": ["Tuition", "Transport", "Lab", "Library", "Sports", "Admission", "Exam", "Hostel", "Misc"],
        "PaymentMode": ["Cash", "Cheque", "Online", "Bank Transfer", "UPI"],
        "FeeStatus": ["Paid", "Partially Paid", "Unpaid", "Overdue"],
        "BookCategory": ["Textbook", "Reference", "Fiction", "Non-fiction", "Journal", "Magazine"],
        "BookCondition": ["Good", "Fair", "Damaged", "Lost"],
        "BookStatus": ["Issued", "Returned", "Overdue", "Lost"],
        "VehicleType": ["Bus", "Van", "Auto"],
        "VehicleStatus": ["Active", "Under Maintenance", "Retired"],
        "AssetCategory": ["Furniture", "Electronics", "Sports", "Lab Equipment", "Stationery", "Other"],
        "AssetCondition": ["New", "Good", "Fair", "Needs Repair", "Disposed"],
        "IncidentType": ["Misconduct", "Bullying", "Vandalism", "Truancy", "Academic Dishonesty"],
        "ActionTaken": ["Warning", "Detention", "Suspension", "Parent Meeting", "Expulsion"],
        "AchievementCategory": ["Academic", "Sports", "Cultural", "Community Service"],
        "AchievementLevel": ["School", "District", "State", "National", "International"],
        "RoomType": ["Single", "Double", "Dormitory"],
        "MealPref": ["Veg", "Non-Veg", "Vegan", "Jain"],
        "EventType": ["Holiday", "Exam", "PTM", "Sports Day", "Annual Day", "Workshop", "Other"],
        "EventStatus": ["Planning", "Ongoing", "Completed", "Cancelled"],
        "IncomeSource": ["Fees", "Donations", "Grants", "Events", "Rental", "Transport Fees", "Other"],
        "ExpenseCategory": ["Salary", "Utilities", "Maintenance", "Supplies", "Events", "Transport", "Misc"],
        "EnquirySource": ["Walk-in", "Online", "Referral", "Ad", "Social Media"],
        "EnquiryStatus": ["New", "Follow-up", "Converted", "Closed"],
        "AdmissionStatus": ["Approved", "Waitlisted", "Rejected", "Pending"],
        "TCReason": ["Transfer", "Withdrawal", "Graduated", "Expelled"],
        "NoticeAudience": ["All", "Specific Class", "Staff", "Parents", "Students"],
        "NoticeMedium": ["SMS", "Email", "WhatsApp", "Notice Board"],
        "NoticeStatus": ["Sent", "Pending", "Scheduled"],
        "CommType": ["Complaint", "Feedback", "Meeting", "Call", "General"],
        "SyllabusStatus": ["Not Started", "In Progress", "Completed"],
        "TransportPref": ["One-way", "Both ways"],
        "Classes": ["Nursery","LKG","UKG","1","2","3","4","5","6","7","8","9","10","11","12"],
        "Sections": ["A","B","C","D","E"],
        "Days": ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"],
    }

    row = 3
    for key, values in lists.items():
        ws.cell(row=row, column=1, value=key).font = Font(bold=True, color=DARK_BLUE)
        for i, v in enumerate(values):
            ws.cell(row=row, column=i + 2, value=v)
        row += 1

    # System Settings section
    row += 2
    ws.cell(row=row, column=1, value="SYSTEM SETTINGS").font = Font(bold=True, size=14, color=DARK_BLUE)
    row += 1
    settings = [
        ("School Name", "Your School Name Here"),
        ("School Address", "123 Education Street, City, State - 000000"),
        ("School Phone", "+91-XXXXXXXXXX"),
        ("School Email", "info@yourschool.com"),
        ("Academic Year", "2025-2026"),
        ("Principal Name", "Dr. [Name]"),
        ("Affiliation Number", "XXXXXXXXX"),
        ("Max Periods Per Day", 8),
        ("Library Fine Per Day", 1),
        ("Max Library Days", 14),
        ("Minimum Attendance %", 75),
    ]
    for key, val in settings:
        ws.cell(row=row, column=1, value=key).font = Font(bold=True)
        ws.cell(row=row, column=2, value=val)
        row += 1

    auto_width(ws, 15)


# ─── 1. STUDENT MANAGEMENT ──────────────────────────────────────────────
def _build_student_master(ws):
    headers = [
        "Student ID", "Admission Number", "Admission Date", "First Name",
        "Middle Name", "Last Name", "Date of Birth", "Age (Auto)",
        "Gender", "Blood Group", "Nationality", "Religion", "Caste",
        "Category", "Aadhaar Number", "Photo Link", "Current Class",
        "Section", "Roll Number", "Previous School", "Student Status"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📋 STUDENT MASTER DATA", max_col)

    # Auto-generated Student ID formula (rows 3-502)
    for r in range(3, 503):
        ws.cell(row=r, column=1).value = f'=IF(D{r}="",""&TEXT(ROW()-2,"STU0000"),TEXT(ROW()-2,"STU0000"))'

    # Age auto-calculation
    for r in range(3, 503):
        ws.cell(row=r, column=8).value = f'=IF(G{r}="","",DATEDIF(G{r},TODAY(),"Y"))'

    # Data validations
    add_data_validation(ws, f"I3:I502", "Male,Female,Other")
    add_data_validation(ws, f"J3:J502", "A+,A-,B+,B-,AB+,AB-,O+,O-")
    add_data_validation(ws, f"N3:N502", "General,OBC,SC,ST,EWS")
    add_data_validation(ws, f"Q3:Q502", "Nursery,LKG,UKG,1,2,3,4,5,6,7,8,9,10,11,12")
    add_data_validation(ws, f"R3:R502", "A,B,C,D,E")
    add_data_validation(ws, f"U3:U502", "Active,Inactive,Transferred,Graduated")

    # Conditional formatting for status
    ws.conditional_formatting.add("U3:U502",
        CellIsRule(operator='equal', formula=['"Active"'], fill=green_fill))
    ws.conditional_formatting.add("U3:U502",
        CellIsRule(operator='equal', formula=['"Transferred"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_parent_info(ws):
    headers = [
        "Student ID", "Father's Name", "Father's Occupation", "Father's Phone",
        "Father's Email", "Father's Education", "Mother's Name", "Mother's Occupation",
        "Mother's Phone", "Mother's Email", "Mother's Education", "Guardian Name",
        "Emergency Contact", "Current Address", "Permanent Address",
        "Communication Preference"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "👨‍👩‍👧 PARENT / GUARDIAN INFORMATION", max_col)

    add_data_validation(ws, "P3:P502", "SMS,Email,WhatsApp")
    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_student_medical(ws):
    headers = [
        "Student ID", "Student Name", "Class", "Section",
        "Known Allergies", "Medical Conditions", "Medications",
        "Doctor Name", "Doctor Contact", "Insurance Provider",
        "Insurance Policy No.", "Vaccination Records", "Disability (if any)",
        "Disability Type", "Last Medical Check-up"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🏥 STUDENT MEDICAL INFORMATION", max_col)
    add_data_validation(ws, "N3:N502", "Yes,No")
    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 2. STAFF MANAGEMENT ────────────────────────────────────────────────
def _build_staff_master(ws):
    headers = [
        "Employee ID", "Full Name", "Date of Birth", "Gender",
        "Designation", "Department", "Subject Specialization",
        "Qualification", "Date of Joining", "Date of Leaving",
        "Experience (Years)", "Contact Number", "Email",
        "Address", "Bank Account No.", "Bank Name", "IFSC Code",
        "Salary Grade", "Pay Scale", "Employment Type", "Status"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "👨‍🏫 STAFF MASTER DATA", max_col)

    # Auto Employee ID
    for r in range(3, 203):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"EMP000"))'

    add_data_validation(ws, "D3:D202", "Male,Female,Other")
    add_data_validation(ws, "E3:E202", "Teacher,Admin,Support Staff,Principal,Vice-Principal,Coordinator")
    add_data_validation(ws, "T3:T202", "Full-time,Part-time,Contract")
    add_data_validation(ws, "U3:U202", "Active,On Leave,Resigned,Retired")

    ws.conditional_formatting.add("U3:U202",
        CellIsRule(operator='equal', formula=['"Active"'], fill=green_fill))
    ws.conditional_formatting.add("U3:U202",
        CellIsRule(operator='equal', formula=['"Resigned"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_staff_attendance(ws):
    headers = [
        "Date", "Employee ID", "Employee Name", "Status",
        "Leave Type", "Check-in Time", "Check-out Time", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📅 STAFF ATTENDANCE", max_col)

    add_data_validation(ws, "D3:D2000", "Present,Absent,Half-day,Leave")
    add_data_validation(ws, "E3:E2000", "Casual,Sick,Earned,Maternity,Unpaid")

    # Monthly Summary section starting at row 2010
    r = 2010
    ws.cell(row=r, column=1, value="MONTHLY ATTENDANCE SUMMARY").font = section_font
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1).fill = section_fill
    r += 1
    summary_headers = ["Month", "Employee ID", "Employee Name", "Working Days",
                       "Present Days", "Leave Days", "Absent Days", "Leave Balance"]
    for ci, h in enumerate(summary_headers, 1):
        ws.cell(row=r, column=ci, value=h)
    style_header_row(ws, r, len(summary_headers))

    # Formulas for summary
    for row in range(r + 1, r + 201):
        ws.cell(row=row, column=4).value = 26  # default working days
        ws.cell(row=row, column=8).value = f"=12-D{row}+E{row}"  # leave balance calc

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_staff_payroll(ws):
    headers = [
        "Employee ID", "Employee Name", "Month", "Basic Salary",
        "HRA", "DA", "Transport Allowance", "Medical Allowance",
        "Other Allowances", "Gross Salary", "Tax (TDS)", "PF Deduction",
        "Insurance Deduction", "Loan Recovery", "Other Deductions",
        "Total Deductions", "Net Salary", "Payment Date", "Payment Mode"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "💰 STAFF PAYROLL", max_col)

    for r in range(3, 2003):
        # Gross Salary = SUM of Basic + all allowances
        ws.cell(row=r, column=10).value = f"=IF(D{r}=\"\",\"\",SUM(D{r}:I{r}))"
        # Total Deductions = SUM of all deductions
        ws.cell(row=r, column=16).value = f"=IF(D{r}=\"\",\"\",SUM(K{r}:O{r}))"
        # Net Salary = Gross - Deductions
        ws.cell(row=r, column=17).value = f"=IF(D{r}=\"\",\"\",J{r}-P{r})"

    add_data_validation(ws, "S3:S2002", "Cash,Cheque,Online,Bank Transfer,UPI")

    # Protect sheet (sensitive)
    ws.protection.sheet = True
    ws.protection.password = "payroll2026"

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 3. ACADEMIC MANAGEMENT ─────────────────────────────────────────────
def _build_class_setup(ws):
    headers = [
        "Class ID", "Class Name", "Section", "Class Teacher (Staff ID)",
        "Class Teacher Name", "Room Number", "Max Capacity",
        "Current Strength", "Academic Year"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🏫 CLASS & SECTION SETUP", max_col)

    # Pre-populate class/section combos
    classes = ["Nursery","LKG","UKG"] + [str(i) for i in range(1,13)]
    sections = ["A","B","C"]
    row = 3
    cid = 1
    for cls in classes:
        for sec in sections:
            ws.cell(row=row, column=1, value=f"CLS{cid:03d}")
            ws.cell(row=row, column=2, value=cls)
            ws.cell(row=row, column=3, value=sec)
            ws.cell(row=row, column=7, value=40)  # max capacity
            ws.cell(row=row, column=9, value="2025-2026")
            row += 1
            cid += 1

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_subject_master(ws):
    headers = [
        "Subject Code", "Subject Name", "Subject Type", "Assigned Class",
        "Assigned Section", "Assigned Teacher (Staff ID)", "Teacher Name",
        "Periods Per Week", "Credits / Weightage"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📚 SUBJECT MASTER", max_col)

    add_data_validation(ws, "C3:C502", "Core,Elective,Co-curricular")
    add_data_validation(ws, "D3:D502", "Nursery,LKG,UKG,1,2,3,4,5,6,7,8,9,10,11,12")

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_timetable(ws):
    headers = [
        "Day", "Period", "Time Slot", "Class", "Section",
        "Subject", "Teacher", "Room Number"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🕐 TIMETABLE", max_col)

    add_data_validation(ws, "A3:A2000", "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday")
    add_data_validation(ws, "B3:B2000", "1,2,3,4,5,6,7,8")
    add_data_validation(ws, "D3:D2000", "Nursery,LKG,UKG,1,2,3,4,5,6,7,8,9,10,11,12")

    # Conflict detection column
    r = 2
    ws.cell(row=r, column=9, value="Conflict Check")
    ws.cell(row=r, column=9).font = header_font
    ws.cell(row=r, column=9).fill = header_fill
    ws.cell(row=r, column=9).alignment = center_align

    for row in range(3, 2001):
        # Flag if same teacher is double-booked (simplified check)
        ws.cell(row=row, column=9).value = (
            f'=IF(COUNTIFS(G{row}:G2000,G{row},A{row}:A2000,A{row},'
            f'B{row}:B2000,B{row})>1,"⚠ CONFLICT","✓ OK")'
        )

    ws.conditional_formatting.add("I3:I2000",
        CellIsRule(operator='containsText', formula=['"CONFLICT"'],
                   fill=PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_syllabus_tracker(ws):
    headers = [
        "Class", "Section", "Subject", "Chapter/Unit No.", "Topic Name",
        "Planned Date", "Actual Date", "Status", "% Completed (Auto)",
        "Total Chapters", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📖 SYLLABUS TRACKER", max_col)

    add_data_validation(ws, "H3:H5000", "Not Started,In Progress,Completed")

    for r in range(3, 5001):
        ws.cell(row=r, column=9).value = f'=IF(H{r}="Completed",100,IF(H{r}="In Progress",50,IF(H{r}="Not Started",0,"")))'

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 4. ATTENDANCE MANAGEMENT ───────────────────────────────────────────
def _build_daily_attendance(ws):
    headers = [
        "Date", "Class", "Section", "Roll No.", "Student ID",
        "Student Name", "Attendance Status", "Marked By", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📝 DAILY STUDENT ATTENDANCE", max_col)

    add_data_validation(ws, "G3:G50000", "Present,Absent,Late,Half-day,Excused")

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_monthly_attendance(ws):
    headers = [
        "Student ID", "Student Name", "Class", "Section", "Month",
        "Total Working Days", "Days Present", "Days Absent", "Days Late",
        "Attendance %", "Status Alert"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📊 MONTHLY ATTENDANCE SUMMARY", max_col)

    for r in range(3, 5003):
        # Attendance %
        ws.cell(row=r, column=10).value = f'=IF(F{r}=0,"",ROUND(G{r}/F{r}*100,1))'
        # Status Alert
        ws.cell(row=r, column=11).value = (
            f'=IF(J{r}="","",IF(J{r}<75,"⚠ BELOW MINIMUM",'
            f'IF(J{r}<85,"⚡ NEEDS IMPROVEMENT","✓ GOOD")))'
        )

    # Conditional formatting for attendance %
    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='lessThan', formula=['75'],
                   fill=PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')))
    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='between', formula=['75', '85'],
                   fill=PatternFill(start_color=YELLOW_FILL, end_color=YELLOW_FILL, fill_type='solid')))
    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='greaterThan', formula=['85'],
                   fill=PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type='solid')))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_attendance_dashboard(ws):
    add_title_row(ws, "📈 ATTENDANCE DASHBOARD", 12)

    # KPI cards
    kpis = [
        ("B3", "📊 Today's Student Attendance %", "C4", '=IFERROR(ROUND(COUNTIF(\'Daily Attendance\'!G3:G50000,"Present")/COUNTA(\'Daily Attendance\'!G3:G50000)*100,1),0)'),
        ("E3", "👥 Today's Staff Attendance %", "F4", '=IFERROR(ROUND(COUNTIF(\'Staff Attendance\'!D3:D2000,"Present")/COUNTA(\'Staff Attendance\'!D3:D2000)*100,1),0)'),
        ("H3", "📉 Students Below 75%", "I4", '=COUNTIF(\'Monthly Attendance Summary\'!J3:J5002,"<75")'),
    ]

    for label_cell, label, value_cell, formula in kpis:
        ws[label_cell] = label
        ws[label_cell].font = kpi_label_font
        ws[value_cell] = formula
        ws[value_cell].font = kpi_font

    # Class-wise attendance table
    r = 7
    ws.cell(row=r, column=2, value="CLASS-WISE ATTENDANCE").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    r += 1
    cls_headers = ["Class", "Total Students", "Present Today", "Absent Today", "Attendance %"]
    for ci, h in enumerate(cls_headers):
        ws.cell(row=r, column=2 + ci, value=h)
    style_header_row(ws, r, 5)

    # Pre-fill rows for classes
    classes = ["Nursery","LKG","UKG"] + [str(i) for i in range(1,13)]
    for i, cls in enumerate(classes):
        row = r + 1 + i
        ws.cell(row=row, column=2, value=cls)

    freeze_and_filter(ws, 2, 12)
    auto_width(ws, 12)


# ─── 5. EXAMINATION & GRADES ────────────────────────────────────────────
def _build_exam_setup(ws):
    headers = [
        "Exam ID", "Exam Name", "Exam Type", "Class", "Section",
        "Subject", "Max Marks", "Passing Marks", "Weightage (%)",
        "Exam Date", "Duration (Min)", "Academic Year"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📝 EXAM SETUP", max_col)

    add_data_validation(ws, "C3:C500", "Internal,External,Practical,Oral")
    add_data_validation(ws, "D3:D500", "Nursery,LKG,UKG,1,2,3,4,5,6,7,8,9,10,11,12")

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_marks_entry(ws):
    headers = [
        "Exam Name", "Class", "Section", "Student ID", "Student Name",
        "Subject", "Max Marks", "Marks Obtained", "Practical Marks",
        "Internal Assessment", "Total Marks", "Percentage",
        "Grade", "Pass/Fail", "Class Rank"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📊 MARKS ENTRY", max_col)

    for r in range(3, 50003):
        # Total Marks
        ws.cell(row=r, column=11).value = f'=IF(H{r}="","",H{r}+I{r}+J{r})'
        # Percentage
        ws.cell(row=r, column=12).value = f'=IF(K{r}="","",ROUND(K{r}/(G{r}*3)*100,1))'
        # Grade (using lookup from Grade Config sheet)
        ws.cell(row=r, column=13).value = (
            f'=IF(L{r}="","",INDEX(\'Grade Configuration\'!$A$4:$A$12,'
            f'MATCH(TRUE,INDEX((L{r}>=\'Grade Configuration\'!$C$4:$C$12)*'
            f'(L{r}<=\'Grade Configuration\'!$D$4:$D$12),0),0)))'
        )
        # Pass/Fail
        ws.cell(row=r, column=14).value = f'=IF(H{r}="","",IF(H{r}>=0.33*G{r},"PASS","FAIL"))'
        # Rank
        ws.cell(row=r, column=15).value = (
            f'=IF(K{r}="","",RANK(K{r},$K$3:$K$50002,0))'
        )

    ws.conditional_formatting.add("N3:N50002",
        CellIsRule(operator='equal', formula=['"PASS"'], fill=green_fill))
    ws.conditional_formatting.add("N3:N50002",
        CellIsRule(operator='equal', formula=['"FAIL"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_grade_config(ws):
    headers = ["Grade", "Grade Description", "Min %", "Max %", "GPA", "Remarks"]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🎯 GRADE CONFIGURATION", max_col)

    grades = [
        ("A+", "Outstanding", 91, 100, 10, "Outstanding"),
        ("A",  "Excellent",   81, 90,  9,  "Excellent"),
        ("B+", "Very Good",   71, 80,  8,  "Very Good"),
        ("B",  "Good",        61, 70,  7,  "Good"),
        ("C+", "Above Average",51, 60,  6,  "Above Average"),
        ("C",  "Average",     41, 50,  5,  "Average"),
        ("D",  "Below Average",33,40,  4,  "Needs Improvement"),
        ("E",  "Fail",        0,  32,  0,  "Fail"),
        ("F",  "Fail",        0,  32,  0,  "Fail"),
    ]
    for i, g in enumerate(grades):
        row = 4 + i
        for ci, val in enumerate(g):
            ws.cell(row=row, column=ci + 1, value=val)
            ws.cell(row=row, column=ci + 1).border = thin_border
            ws.cell(row=row, column=ci + 1).alignment = center_align

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_report_card(ws):
    headers = [
        "Student ID", "Student Name", "Class", "Section", "Roll No.",
        "Academic Year", "Exam 1 Name", "Exam 1 Total", "Exam 1 %",
        "Exam 2 Name", "Exam 2 Total", "Exam 2 %",
        "Exam 3 Name", "Exam 3 Total", "Exam 3 %",
        "Cumulative Avg %", "CGPA", "Class Rank", "Section Rank",
        "Attendance %", "Teacher Remarks", "Principal Signature"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🎓 REPORT CARD", max_col)

    for r in range(3, 5003):
        ws.cell(row=r, column=16).value = f'=IF(I{r}="","",ROUND(AVERAGE(I{r},K{r},M{r}),1))'
        ws.cell(row=r, column=17).value = f'=IF(P{r}="","",ROUND(P{r}/10,1))'
        ws.cell(row=r, column=18).value = f'=IF(P{r}="","",RANK(P{r},$P$3:$P$5002,0))'

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_exam_analytics(ws):
    add_title_row(ws, "📈 EXAM ANALYTICS DASHBOARD", 12)

    # Class/Subject averages table
    r = 3
    ws.cell(row=r, column=2, value="SUBJECT-WISE ANALYSIS").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    r += 1
    analytics_headers = ["Subject", "Class Avg", "Highest", "Lowest",
                         "Pass %", "Total Students", "Pass Count", "Fail Count"]
    for ci, h in enumerate(analytics_headers):
        ws.cell(row=r, column=2 + ci, value=h)
    style_header_row(ws, r, len(analytics_headers))

    for row in range(r + 1, r + 51):
        ws.cell(row=row, column=3).value = f'=IF(B{row}="","",AVERAGEIF(\'Marks Entry\'!F:F,B{row},\'Marks Entry\'!L:L))'
        ws.cell(row=row, column=4).value = f'=IF(B{row}="","",MAXIFS(\'Marks Entry\'!L:L,\'Marks Entry\'!F:F,B{row}))'
        ws.cell(row=row, column=5).value = f'=IF(B{row}="","",MINIFS(\'Marks Entry\'!L:L,\'Marks Entry\'!F:F,B{row}))'
        ws.cell(row=row, column=6).value = f'=IF(B{row}="","",ROUND(COUNTIFS(\'Marks Entry\'!F:F,B{row},\'Marks Entry\'!N:N,"PASS")/COUNTIF(\'Marks Entry\'!F:F,B{row})*100,1))'
        ws.cell(row=row, column=7).value = f'=IF(B{row}="","",COUNTIF(\'Marks Entry\'!F:F,B{row}))'
        ws.cell(row=row, column=8).value = f'=IF(B{row}="","",COUNTIFS(\'Marks Entry\'!F:F,B{row},\'Marks Entry\'!N:N,"PASS"))'
        ws.cell(row=row, column=9).value = f'=IF(B{row}="","",G{row}-H{row})'

    # Topper list
    r2 = r + 55
    ws.cell(row=r2, column=2, value="TOPPER LIST").font = section_font
    ws.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=5)
    r2 += 1
    topper_headers = ["Rank", "Student ID", "Student Name", "Class", "Percentage"]
    for ci, h in enumerate(topper_headers):
        ws.cell(row=r2, column=2 + ci, value=h)
    style_header_row(ws, r2, len(topper_headers))

    auto_width(ws, 12)


# ─── 6. FEE MANAGEMENT ──────────────────────────────────────────────────
def _build_fee_structure(ws):
    headers = [
        "Fee ID", "Class", "Fee Type", "Annual Amount", "Monthly Amount",
        "Quarterly Amount", "Due Date", "Late Fee (% per month)",
        "Sibling Discount (%)", "Merit Scholarship (%)", "Staff Ward Discount (%)"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "💵 FEE STRUCTURE", max_col)

    add_data_validation(ws, "C3:C500", "Tuition,Transport,Lab,Library,Sports,Admission,Exam,Hostel,Misc")
    add_data_validation(ws, "B3:B500", "Nursery,LKG,UKG,1,2,3,4,5,6,7,8,9,10,11,12")

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_fee_collection(ws):
    headers = [
        "Receipt No.", "Student ID", "Student Name", "Class", "Section",
        "Fee Type", "Amount Due", "Amount Paid", "Balance",
        "Payment Date", "Payment Mode", "Transaction Ref No.",
        "Collected By", "Status"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "💰 FEE COLLECTION", max_col)

    # Auto Receipt No
    for r in range(3, 50003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"RCT00000"))'
        # Balance
        ws.cell(row=r, column=9).value = f'=IF(G{r}="","",G{r}-H{r})'
        # Status auto-detect
        ws.cell(row=r, column=14).value = (
            f'=IF(G{r}="","",IF(I{r}<=0,"Paid",IF(H{r}>0,"Partially Paid","Unpaid")))'
        )

    add_data_validation(ws, "K3:K50002", "Cash,Cheque,Online,Bank Transfer,UPI")

    ws.conditional_formatting.add("N3:N50002",
        CellIsRule(operator='equal', formula=['"Paid"'], fill=green_fill))
    ws.conditional_formatting.add("N3:N50002",
        CellIsRule(operator='equal', formula=['"Partially Paid"'], fill=yellow_cond_fill))
    ws.conditional_formatting.add("N3:N50002",
        CellIsRule(operator='equal', formula=['"Unpaid"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_fee_defaulters(ws):
    headers = [
        "Student ID", "Student Name", "Class", "Section",
        "Fee Type", "Total Due", "Total Paid", "Outstanding Amount",
        "First Overdue Date", "Days Overdue", "Severity"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "⚠️ FEE DEFAULTERS REPORT", max_col)

    for r in range(3, 5003):
        ws.cell(row=r, column=8).value = f'=IF(F{r}="","",F{r}-G{r})'
        ws.cell(row=r, column=10).value = f'=IF(I{r}="","",TODAY()-I{r})'
        ws.cell(row=r, column=11).value = (
            f'=IF(J{r}="","",IF(J{r}>90,"🔴 CRITICAL",IF(J{r}>60,"🟠 HIGH",IF(J{r}>30,"🟡 MEDIUM","🟢 LOW"))))'
        )

    # Conditional formatting on severity
    ws.conditional_formatting.add("K3:K5002",
        CellIsRule(operator='containsText', formula=['"CRITICAL"'],
                   fill=PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')))
    ws.conditional_formatting.add("K3:K5002",
        CellIsRule(operator='containsText', formula=['"HIGH"'],
                   fill=PatternFill(start_color="FFD966", end_color="FFD966", fill_type='solid')))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_fee_dashboard(ws):
    add_title_row(ws, "💰 FEE DASHBOARD", 12)

    r = 3
    kpis = [
        ("B3", "Total Fees Expected (Annual)"),
        ("D3", "Total Collected"),
        ("F3", "Outstanding Amount"),
        ("H3", "Collection %"),
    ]
    for cell, label in kpis:
        ws[cell] = label
        ws[cell].font = kpi_label_font

    ws["B4"] = '=SUM(\'Fee Collection\'!G3:G50002)'
    ws["B4"].font = kpi_font
    ws["D4"] = '=SUM(\'Fee Collection\'!H3:H50002)'
    ws["D4"].font = kpi_font
    ws["F4"] = '=SUM(\'Fee Collection\'!I3:I50002)'
    ws["F4"].font = kpi_font
    ws["H4"] = '=IFERROR(ROUND(D4/B4*100,1)&"%","N/A")'
    ws["H4"].font = kpi_font

    # Class-wise collection summary
    r = 7
    ws.cell(row=r, column=2, value="CLASS-WISE COLLECTION SUMMARY").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    r += 1
    for ci, h in enumerate(["Class", "Expected", "Collected", "Pending", "Collection %"]):
        ws.cell(row=r, column=2 + ci, value=h)
    style_header_row(ws, r, 5)

    auto_width(ws, 12)


# ─── 7. LIBRARY MANAGEMENT ──────────────────────────────────────────────
def _build_book_inventory(ws):
    headers = [
        "Book ID", "Accession No.", "Title", "Author", "Publisher",
        "ISBN", "Category", "Subject", "Class Level", "Edition",
        "Year of Publication", "Rack / Shelf", "Total Copies",
        "Available Copies", "Condition", "Cost per Copy"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📚 BOOK INVENTORY", max_col)

    add_data_validation(ws, "G3:G5002", "Textbook,Reference,Fiction,Non-fiction,Journal,Magazine")
    add_data_validation(ws, "O3:O5002", "Good,Fair,Damaged,Lost")

    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Good"'], fill=green_fill))
    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Damaged"'], fill=red_fill))
    ws.conditional_formatting.add("O3:O5002",
        CellIsRule(operator='equal', formula=['"Lost"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_book_issue_return(ws):
    headers = [
        "Transaction ID", "Book ID", "Book Title", "Issued To (ID)",
        "Borrower Name", "Borrower Type", "Issue Date", "Due Date",
        "Return Date", "Status", "Overdue Days", "Fine Amount", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📖 BOOK ISSUE & RETURN", max_col)

    for r in range(3, 20003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"LIB00000"))'
        # Due Date = Issue Date + 14
        ws.cell(row=r, column=8).value = f'=IF(G{r}="","",G{r}+14)'
        # Overdue Days
        ws.cell(row=r, column=11).value = (
            f'=IF(OR(J{r}="Returned",J{r}=""),0,MAX(0,TODAY()-H{r}))'
        )
        # Fine = overdue days * 1 (from config)
        ws.cell(row=r, column=12).value = f'=IF(K{r}=0,0,K{r}*1)'

    add_data_validation(ws, "F3:F20002", "Student,Staff")
    add_data_validation(ws, "J3:J20002", "Issued,Returned,Overdue,Lost")

    ws.conditional_formatting.add("J3:J20002",
        CellIsRule(operator='equal', formula=['"Overdue"'], fill=red_fill))
    ws.conditional_formatting.add("J3:J20002",
        CellIsRule(operator='equal', formula=['"Returned"'], fill=green_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_library_dashboard(ws):
    add_title_row(ws, "📚 LIBRARY DASHBOARD", 10)

    ws["B3"] = "Total Books"
    ws["B3"].font = kpi_label_font
    ws["B4"] = '=SUM(\'Book Inventory\'!M3:M5002)'
    ws["B4"].font = kpi_font

    ws["D3"] = "Available Books"
    ws["D3"].font = kpi_label_font
    ws["D4"] = '=SUM(\'Book Inventory\'!N3:N5002)'
    ws["D4"].font = kpi_font

    ws["F3"] = "Books Issued"
    ws["F3"].font = kpi_label_font
    ws["F4"] = '=COUNTIF(\'Book Issue Return\'!J3:J20002,"Issued")'
    ws["F4"].font = kpi_font

    ws["H3"] = "Overdue Books"
    ws["H3"].font = kpi_label_font
    ws["H4"] = '=COUNTIF(\'Book Issue Return\'!J3:J20002,"Overdue")'
    ws["H4"].font = kpi_font

    # Overdue list section
    r = 7
    ws.cell(row=r, column=2, value="OVERDUE BOOKS LIST").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    r += 1
    for ci, h in enumerate(["Book ID", "Book Title", "Borrower", "Issue Date", "Due Date", "Fine"]):
        ws.cell(row=r, column=2 + ci, value=h)
    style_header_row(ws, r, 6)

    auto_width(ws, 10)


# ─── 8. TRANSPORT MANAGEMENT ────────────────────────────────────────────
def _build_vehicle_master(ws):
    headers = [
        "Vehicle No.", "Registration No.", "Vehicle Type", "Capacity",
        "Driver Name", "Driver Contact", "Driver License No.",
        "Helper / Conductor", "Insurance Expiry", "Fitness Cert. Expiry",
        "GPS Tracker ID", "Route Assigned", "Status"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🚌 VEHICLE MASTER", max_col)

    add_data_validation(ws, "C3:C200", "Bus,Van,Auto")
    add_data_validation(ws, "M3:M200", "Active,Under Maintenance,Retired")

    # Alert for insurance expiry
    for r in range(3, 203):
        ws.cell(row=r, column=14, value=f'=IF(I{r}="","",IF(I{r}-TODAY()<30,"⚠ EXPIRING SOON","✓ OK"))')

    ws.conditional_formatting.add("N3:N200",
        CellIsRule(operator='containsText', formula=['"EXPIRING"'],
                   fill=PatternFill(start_color=YELLOW_FILL, end_color=YELLOW_FILL, fill_type='solid')))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_route_management(ws):
    headers = [
        "Route ID", "Route Name", "Stop 1", "Stop 2", "Stop 3",
        "Stop 4", "Stop 5", "Pickup Time", "Drop Time",
        "Distance (KM)", "Assigned Vehicle", "Assigned Driver",
        "No. of Students", "Monthly Fee"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🗺️ ROUTE MANAGEMENT", max_col)

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_student_transport(ws):
    headers = [
        "Student ID", "Student Name", "Class", "Section",
        "Route Assigned", "Stop Name", "Pickup/Drop",
        "Transport Fee", "Payment Status"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🚐 STUDENT TRANSPORT MAPPING", max_col)

    add_data_validation(ws, "G3:G5002", "One-way,Both ways")
    add_data_validation(ws, "I3:I5002", "Paid,Unpaid,Partially Paid")

    ws.conditional_formatting.add("I3:I5002",
        CellIsRule(operator='equal', formula=['"Paid"'], fill=green_fill))
    ws.conditional_formatting.add("I3:I5002",
        CellIsRule(operator='equal', formula=['"Unpaid"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 9. INVENTORY & ASSET MANAGEMENT ────────────────────────────────────
def _build_asset_register(ws):
    headers = [
        "Asset ID", "Asset Name / Description", "Category",
        "Purchase Date", "Vendor", "Cost", "Invoice No.",
        "Location / Room", "Assigned To", "Condition",
        "Warranty Expiry", "Depreciation Value (Auto)", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🏷️ ASSET REGISTER", max_col)

    # Auto Asset ID
    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"AST0000"))'
        # Depreciation (straight-line, 10 years)
        ws.cell(row=r, column=12).value = (
            f'=IF(D{r}="","",ROUND(MAX(0,F{r}-(F{r}/10)*(DATEDIF(D{r},TODAY(),"Y"))),2))'
        )

    add_data_validation(ws, "C3:C5002", "Furniture,Electronics,Sports,Lab Equipment,Stationery,Other")
    add_data_validation(ws, "J3:J5002", "New,Good,Fair,Needs Repair,Disposed")

    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='equal', formula=['"New"'], fill=green_fill))
    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='equal', formula=['"Disposed"'], fill=red_fill))
    ws.conditional_formatting.add("J3:J5002",
        CellIsRule(operator='equal', formula=['"Needs Repair"'], fill=yellow_cond_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_stationery_stock(ws):
    headers = [
        "Item Code", "Item Name", "Category", "Opening Stock",
        "Received Qty", "Received Date", "Issued Qty", "Issued To",
        "Closing Stock (Auto)", "Reorder Level", "Stock Alert"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📦 STATIONERY / CONSUMABLES STOCK", max_col)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"ITM000"))'
        # Closing Stock = Opening + Received - Issued
        ws.cell(row=r, column=9).value = f'=IF(B{r}="","",D{r}+E{r}-G{r})'
        # Alert
        ws.cell(row=r, column=11).value = (
            f'=IF(I{r}="","",IF(I{r}<=J{r},"🔴 REORDER NOW","✓ OK"))'
        )

    ws.conditional_formatting.add("K3:K2002",
        CellIsRule(operator='containsText', formula=['"REORDER"'],
                   fill=PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')))
    ws.conditional_formatting.add("K3:K2002",
        CellIsRule(operator='containsText', formula=['"OK"'],
                   fill=green_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 10. COMMUNICATION & NOTICES ────────────────────────────────────────
def _build_notice_board(ws):
    headers = [
        "Notice ID", "Date", "Title", "Description",
        "Target Audience", "Issued By", "Medium", "Status"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📢 NOTICE BOARD LOG", max_col)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"NTC0000"))'

    add_data_validation(ws, "E3:E2002", "All,Specific Class,Staff,Parents,Students")
    add_data_validation(ws, "G3:G2002", "SMS,Email,WhatsApp,Notice Board")
    add_data_validation(ws, "H3:H2002", "Sent,Pending,Scheduled")

    ws.conditional_formatting.add("H3:H2002",
        CellIsRule(operator='equal', formula=['"Sent"'], fill=green_fill))
    ws.conditional_formatting.add("H3:H2002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=yellow_cond_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_parent_communication(ws):
    headers = [
        "Log ID", "Date", "Student ID", "Parent Name",
        "Type", "Subject", "Description", "Action Taken",
        "Resolved", "Follow-up Date", "Follow-up Status"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📞 PARENT COMMUNICATION LOG", max_col)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"COM0000"))'

    add_data_validation(ws, "E3:E5002", "Complaint,Feedback,Meeting,Call,General")
    add_data_validation(ws, "I3:I5002", "Yes,No")

    ws.conditional_formatting.add("I3:I5002",
        CellIsRule(operator='equal', formula=['"Yes"'], fill=green_fill))
    ws.conditional_formatting.add("I3:I5002",
        CellIsRule(operator='equal', formula=['"No"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 11. DISCIPLINE & BEHAVIOUR ─────────────────────────────────────────
def _build_incident_log(ws):
    headers = [
        "Incident ID", "Date", "Time", "Student ID", "Student Name",
        "Class", "Section", "Incident Type", "Description",
        "Reported By", "Witnesses", "Action Taken",
        "Parent Notified", "Follow-up Status"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "⚠️ INCIDENT LOG – DISCIPLINE & BEHAVIOUR", max_col)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(D{r}="","",TEXT(ROW()-2,"INC0000"))'

    add_data_validation(ws, "H3:H5002", "Misconduct,Bullying,Vandalism,Truancy,Academic Dishonesty")
    add_data_validation(ws, "L3:L5002", "Warning,Detention,Suspension,Parent Meeting,Expulsion")
    add_data_validation(ws, "M3:M5002", "Yes,No")
    add_data_validation(ws, "N3:N5002", "Resolved,Pending,Under Review")

    ws.conditional_formatting.add("N3:N5002",
        CellIsRule(operator='equal', formula=['"Resolved"'], fill=green_fill))
    ws.conditional_formatting.add("N3:N5002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_student_awards(ws):
    headers = [
        "Award ID", "Student ID", "Student Name", "Class",
        "Date", "Category", "Achievement Description",
        "Level", "Award / Certificate Given", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🏆 STUDENT AWARDS & ACHIEVEMENTS", max_col)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"AWD0000"))'

    add_data_validation(ws, "F3:F5002", "Academic,Sports,Cultural,Community Service")
    add_data_validation(ws, "H3:H5002", "School,District,State,National,International")

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 12. HOSTEL MANAGEMENT ──────────────────────────────────────────────
def _build_hostel_rooms(ws):
    headers = [
        "Room No.", "Block", "Floor", "Room Type", "Capacity",
        "Occupied", "Vacant (Auto)", "Student IDs Assigned",
        "Warden Name", "Amenities"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🏠 HOSTEL ROOM ALLOCATION", max_col)

    for r in range(3, 503):
        ws.cell(row=r, column=7).value = f'=IF(E{r}="","",E{r}-F{r})'

    add_data_validation(ws, "D3:D502", "Single,Double,Dormitory")

    ws.conditional_formatting.add("G3:G502",
        CellIsRule(operator='equal', formula=['0'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_mess_management(ws):
    headers = [
        "Date", "Day", "Breakfast", "Lunch", "Snacks", "Dinner",
        "Veg Count", "Non-Veg Count", "Total Meals Served",
        "Monthly Mess Fee", "Special Instructions"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🍽️ MESS / FOOD MANAGEMENT", max_col)

    for r in range(3, 2003):
        ws.cell(row=r, column=9).value = f'=IF(H{r}="","",G{r}+H{r})'

    add_data_validation(ws, "B3:B2002", "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday")

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 13. HEALTH & INFIRMARY ─────────────────────────────────────────────
def _build_health_records(ws):
    headers = [
        "Visit ID", "Date of Visit", "Student/Staff ID", "Name",
        "Type (Student/Staff)", "Complaint / Symptoms", "Diagnosis",
        "Treatment Given", "Referred to Hospital", "Doctor Name",
        "Follow-up Required", "Follow-up Date", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🏥 HEALTH RECORDS", max_col)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"HLT00000"))'

    add_data_validation(ws, "E3:E10002", "Student,Staff")
    add_data_validation(ws, "I3:I10002", "Yes,No")
    add_data_validation(ws, "K3:K10002", "Yes,No")

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_medical_inventory(ws):
    headers = [
        "Medicine ID", "Medicine Name", "Category", "Stock Available",
        "Expiry Date", "Supplier", "Usage Log", "Reorder Level",
        "Stock Alert"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "💊 MEDICAL INVENTORY", max_col)

    for r in range(3, 1003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"MED000"))'
        ws.cell(row=r, column=9).value = (
            f'=IF(D{r}="","",IF(D{r}<=H{r},"🔴 REORDER",IF(F{r}<TODAY(),"⚠ EXPIRED","✓ OK")))'
        )

    ws.conditional_formatting.add("I3:I1002",
        CellIsRule(operator='containsText', formula=['"REORDER"'],
                   fill=PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')))
    ws.conditional_formatting.add("I3:I1002",
        CellIsRule(operator='containsText', formula=['"EXPIRED"'],
                   fill=PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 14. EVENTS & CALENDAR ──────────────────────────────────────────────
def _build_academic_calendar(ws):
    headers = [
        "Date", "Day", "Event / Activity", "Type",
        "Applicable To", "Description", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📅 ACADEMIC CALENDAR", max_col)

    for r in range(3, 2003):
        ws.cell(row=r, column=2).value = f'=IF(A{r}="","",TEXT(A{r},"dddd"))'

    add_data_validation(ws, "D3:D2002", "Holiday,Exam,PTM,Sports Day,Annual Day,Workshop,Other")
    add_data_validation(ws, "E3:E2002", "All,Specific Classes,Staff,Parents,Students")

    # Highlight holidays
    ws.conditional_formatting.add("D3:D2002",
        CellIsRule(operator='equal', formula=['"Holiday"'], fill=green_fill))
    ws.conditional_formatting.add("D3:D2002",
        CellIsRule(operator='equal', formula=['"Exam"'], fill=yellow_cond_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_event_management(ws):
    headers = [
        "Event ID", "Event Name", "Date", "Venue",
        "Organizer / Coordinator", "Budget Allocated", "Budget Spent",
        "Variance (Auto)", "Participants", "Status", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🎉 EVENT MANAGEMENT", max_col)

    for r in range(3, 1003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"EVT0000"))'
        ws.cell(row=r, column=8).value = f'=IF(F{r}="","",F{r}-G{r})'

    add_data_validation(ws, "J3:J1002", "Planning,Ongoing,Completed,Cancelled")

    ws.conditional_formatting.add("J3:J1002",
        CellIsRule(operator='equal', formula=['"Completed"'], fill=green_fill))
    ws.conditional_formatting.add("J3:J1002",
        CellIsRule(operator='equal', formula=['"Cancelled"'], fill=red_fill))
    ws.conditional_formatting.add("J3:J1002",
        CellIsRule(operator='equal', formula=['"Planning"'], fill=yellow_cond_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 15. FINANCIAL MANAGEMENT ───────────────────────────────────────────
def _build_income_tracker(ws):
    headers = [
        "Entry ID", "Date", "Source", "Description",
        "Amount", "Reference Number", "Received By", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "💵 INCOME TRACKER", max_col)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"INC00000"))'

    add_data_validation(ws, "C3:C10002", "Fees,Donations,Grants,Events,Rental,Transport Fees,Other")

    # Running total row at bottom
    r = 10005
    ws.cell(row=r, column=3, value="TOTAL INCOME:").font = Font(bold=True, size=12, color=DARK_BLUE)
    ws.cell(row=r, column=5).value = '=SUM(E3:E10002)'
    ws.cell(row=r, column=5).font = Font(bold=True, size=14, color=DARK_BLUE)

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_expense_tracker(ws):
    headers = [
        "Entry ID", "Date", "Category", "Description",
        "Vendor", "Amount", "Payment Mode", "Approved By",
        "Receipt / Invoice No.", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "💸 EXPENSE TRACKER", max_col)

    for r in range(3, 10003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"EXP00000"))'

    add_data_validation(ws, "C3:C10002", "Salary,Utilities,Maintenance,Supplies,Events,Transport,Misc")
    add_data_validation(ws, "G3:G10002", "Cash,Cheque,Online,Bank Transfer,UPI")

    r = 10005
    ws.cell(row=r, column=3, value="TOTAL EXPENSES:").font = Font(bold=True, size=12, color=RED)
    ws.cell(row=r, column=6).value = '=SUM(F3:F10002)'
    ws.cell(row=r, column=6).font = Font(bold=True, size=14, color=RED)

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_budget_planning(ws):
    headers = [
        "Department / Category", "Annual Budget Allocated",
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        "YTD Spend", "Remaining Budget", "Variance (Auto)"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📊 BUDGET PLANNING", max_col)

    for r in range(3, 103):
        ws.cell(row=r, column=15).value = f'=IF(B{r}="","",SUM(C{r}:N{r}))'
        ws.cell(row=r, column=16).value = f'=IF(B{r}="","",B{r}-O{r})'
        ws.cell(row=r, column=17).value = f'=IF(B{r}="","",IF(P{r}<0,"⚠ OVER BUDGET","✓ Within Budget"))'

    # Totals row
    r = 104
    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
    for c in range(2, 18):
        ws.cell(row=r, column=c).value = f'=SUM({get_column_letter(c)}3:{get_column_letter(c)}103)'
        ws.cell(row=r, column=c).font = Font(bold=True)

    ws.conditional_formatting.add("Q3:Q103",
        CellIsRule(operator='containsText', formula=['"OVER"'],
                   fill=PatternFill(start_color=LIGHT_RED, end_color=LIGHT_RED, fill_type='solid')))
    ws.conditional_formatting.add("Q3:Q103",
        CellIsRule(operator='containsText', formula=['"Within"'],
                   fill=green_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_profit_loss(ws):
    add_title_row(ws, "📊 PROFIT & LOSS SUMMARY", 8)

    r = 3
    ws.cell(row=r, column=2, value="METRIC").font = section_font
    ws.cell(row=r, column=4, value="AMOUNT").font = section_font
    r += 1
    p_headers = ["Category", "Amount", "Notes"]
    ws.cell(row=r, column=2, value="Category")
    ws.cell(row=r, column=3, value="Amount")
    ws.cell(row=r, column=4, value="Notes")
    style_header_row(ws, r, 3)

    metrics = [
        ("Total Income (YTD)", "='Income Tracker'!E10005", "From Income Tracker"),
        ("Total Expenses (YTD)", "='Expense Tracker'!F10005", "From Expense Tracker"),
        ("Surplus / Deficit", "=B7-B8", "Positive = Surplus, Negative = Deficit"),
        ("", "", ""),
        ("Monthly Income (Avg)", "=IFERROR(B7/12,0)", "12-month average"),
        ("Monthly Expenses (Avg)", "=IFERROR(B8/12,0)", "12-month average"),
        ("Monthly Surplus / Deficit", "=B11-B12", ""),
    ]

    for i, (label, formula, note) in enumerate(metrics):
        row = r + 1 + i
        ws.cell(row=row, column=2, value=label).font = Font(bold=True) if label else Font()
        ws.cell(row=row, column=3).value = formula
        ws.cell(row=row, column=4, value=note)
        for c in range(2, 5):
            ws.cell(row=row, column=c).border = thin_border

    # Highlight surplus/deficit
    ws.conditional_formatting.add("C9",
        CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill))
    ws.conditional_formatting.add("C9",
        CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))

    auto_width(ws, 8)


# ─── 16. ADMISSION MANAGEMENT ───────────────────────────────────────────
def _build_enquiry_tracker(ws):
    headers = [
        "Enquiry ID", "Date", "Student Name", "Parent Name",
        "Contact Number", "Email", "Class Applied For",
        "Source", "Status", "Follow-up Date", "Follow-up Notes"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "🔍 ENQUIRY TRACKER", max_col)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(C{r}="","",TEXT(ROW()-2,"ENQ0000"))'

    add_data_validation(ws, "G3:G5002", "Nursery,LKG,UKG,1,2,3,4,5,6,7,8,9,10,11,12")
    add_data_validation(ws, "H3:H5002", "Walk-in,Online,Referral,Ad,Social Media")
    add_data_validation(ws, "I3:I5002", "New,Follow-up,Converted,Closed")

    ws.conditional_formatting.add("I3:I5002",
        CellIsRule(operator='equal', formula=['"New"'], fill=yellow_cond_fill))
    ws.conditional_formatting.add("I3:I5002",
        CellIsRule(operator='equal', formula=['"Converted"'], fill=green_fill))
    ws.conditional_formatting.add("I3:I5002",
        CellIsRule(operator='equal', formula=['"Closed"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


def _build_admission_tracker(ws):
    headers = [
        "Application No.", "Student Name", "Class Applied For",
        "Date of Application", "Birth Certificate", "Previous TC",
        "Photos", "Transfer Certificate", "Marks Sheet",
        "Entrance Test Marks", "Interview Status", "Interview Date",
        "Admission Status", "Admission Date", "Class Allotted",
        "Section Allotted", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📝 ADMISSION TRACKER", max_col)

    for r in range(3, 5003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"APP0000"))'

    # Document checklist (Yes/No dropdowns)
    for col in ["E", "F", "G", "H", "I"]:
        add_data_validation(ws, f"{col}3:{col}5002", "Yes,No")

    add_data_validation(ws, "K3:K5002", "Scheduled,Completed,Cancelled,No Show")
    add_data_validation(ws, "M3:M5002", "Approved,Waitlisted,Rejected,Pending")
    add_data_validation(ws, "O3:O5002", "Nursery,LKG,UKG,1,2,3,4,5,6,7,8,9,10,11,12")
    add_data_validation(ws, "P3:P5002", "A,B,C,D,E")

    ws.conditional_formatting.add("M3:M5002",
        CellIsRule(operator='equal', formula=['"Approved"'], fill=green_fill))
    ws.conditional_formatting.add("M3:M5002",
        CellIsRule(operator='equal', formula=['"Rejected"'], fill=red_fill))
    ws.conditional_formatting.add("M3:M5002",
        CellIsRule(operator='equal', formula=['"Pending"'], fill=yellow_cond_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── 17. TRANSFER CERTIFICATE ───────────────────────────────────────────
def _build_tc_register(ws):
    headers = [
        "TC Number", "Student ID", "Student Name", "Class at Leaving",
        "Section", "Date of Admission", "Date of Leaving", "Reason",
        "Conduct & Character", "Fee Clearance", "TC Issued Date",
        "Received By (Parent/Guardian)", "Remarks"
    ]
    max_col = write_headers(ws, headers)
    add_title_row(ws, "📄 TRANSFER CERTIFICATE REGISTER", max_col)

    for r in range(3, 2003):
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",TEXT(ROW()-2,"TC-0000"))'

    add_data_validation(ws, "H3:H2002", "Transfer,Withdrawal,Graduated,Expelled")
    add_data_validation(ws, "J3:J2002", "Yes,No")
    add_data_validation(ws, "I3:I2002", "Excellent,Good,Satisfactory,Needs Improvement")

    ws.conditional_formatting.add("J3:J2002",
        CellIsRule(operator='equal', formula=['"Yes"'], fill=green_fill))
    ws.conditional_formatting.add("J3:J2002",
        CellIsRule(operator='equal', formula=['"No"'], fill=red_fill))

    freeze_and_filter(ws, 2, max_col)
    auto_width(ws, max_col)


# ─── MAIN DASHBOARD ─────────────────────────────────────────────────────
def _build_dashboard(ws, wb):
    """Main summary dashboard with KPIs and quick links."""
    max_col = 14

    # Title
    ws.merge_cells("A1:N1")
    ws["A1"] = "🏫  COMPREHENSIVE SCHOOL MANAGEMENT SYSTEM – DASHBOARD"
    ws["A1"].font = Font(name='Calibri', bold=True, color=WHITE, size=18)
    ws["A1"].fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type='solid')
    ws["A1"].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 45

    ws.merge_cells("A2:N2")
    ws["A2"] = "School Name: Your School Name Here    |    Academic Year: 2025-2026    |    Principal: Dr. [Name]"
    ws["A2"].font = Font(name='Calibri', italic=True, color=MID_BLUE, size=11)
    ws["A2"].alignment = Alignment(horizontal='center')

    # ── KPI Cards ───────────────────────────────────────────────────────
    kpi_data = [
        ("B4", "Total Students", "B5", "=COUNTA('Student Master Data'!D3:D502)"),
        ("D4", "Total Staff", "D5", "=COUNTA('Staff Master Data'!B3:B202)"),
        ("F4", "Today's Attendance %", "F5",
         "=IFERROR(ROUND(COUNTIF('Daily Attendance'!G3:G50000,\"Present\")/COUNTA('Daily Attendance'!G3:G50000)*100,1)&\"%\",\"N/A\")"),
        ("H4", "Fee Collected %", "H5",
         "=IFERROR(ROUND(SUM('Fee Collection'!H3:H50002)/SUM('Fee Collection'!G3:G50002)*100,1)&\"%\",\"N/A\")"),
        ("J4", "Active Students", "J5", "=COUNTIF('Student Master Data'!U3:U502,\"Active\")"),
        ("L4", "Books Available", "L5", "=SUM('Book Inventory'!N3:N5002)"),
    ]

    for label_cell, label, value_cell, formula in kpi_data:
        ws[label_cell] = label
        ws[label_cell].font = Font(name='Calibri', bold=True, size=10, color=MID_BLUE)
        ws[label_cell].alignment = Alignment(horizontal='center')
        ws[label_cell].fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')
        ws[value_cell] = formula
        ws[value_cell].font = Font(name='Calibri', bold=True, size=22, color=DARK_BLUE)
        ws[value_cell].alignment = Alignment(horizontal='center')

    # ── Module Quick Links ──────────────────────────────────────────────
    r = 8
    ws.cell(row=r, column=2, value="📋 MODULE QUICK LINKS").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1

    modules = [
        ("📁 Student Module", "Student Master Data", "Parent Information", "Student Medical"),
        ("📁 Academic Module", "Class Section Setup", "Subject Master", "Timetable"),
        ("📁 Attendance Module", "Daily Attendance", "Monthly Attendance Summary", "Attendance Dashboard"),
        ("📁 Exam Module", "Exam Setup", "Marks Entry", "Grade Configuration"),
        ("📁 Fee Module", "Fee Structure", "Fee Collection", "Fee Defaulters"),
        ("📁 HR Module", "Staff Master Data", "Staff Attendance", "Staff Payroll"),
        ("📁 Library Module", "Book Inventory", "Book Issue Return", "Library Dashboard"),
        ("📁 Transport Module", "Vehicle Master", "Route Management", "Student Transport"),
        ("📁 Finance Module", "Income Tracker", "Expense Tracker", "Budget Planning"),
        ("📁 Admission Module", "Enquiry Tracker", "Admission Tracker", "TC Register"),
    ]

    for mod in modules:
        ws.cell(row=r, column=2, value=mod[0]).font = Font(bold=True, color=DARK_BLUE)
        ws.cell(row=r, column=2).fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type='solid')
        for ci, name in enumerate(mod[1:], 3):
            ws.cell(row=r, column=ci, value=f"→ {name}")
            ws.cell(row=r, column=ci).font = Font(color=MID_BLUE, underline='single')
        r += 1

    # ── Financial Summary ───────────────────────────────────────────────
    r += 1
    ws.cell(row=r, column=2, value="💰 FINANCIAL SUMMARY").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1

    fin_headers = ["Metric", "Amount"]
    for ci, h in enumerate(fin_headers):
        ws.cell(row=r, column=2 + ci, value=h)
    style_header_row(ws, r, 2)

    fin_rows = [
        ("Total Income (YTD)", "='Income Tracker'!E10005"),
        ("Total Expenses (YTD)", "='Expense Tracker'!F10005"),
        ("Net Surplus / Deficit", "=B{0}-B{1}".format(r+2, r+3)),
        ("Fee Collection (Collected)", "=SUM('Fee Collection'!H3:H50002)"),
        ("Fee Outstanding", "=SUM('Fee Collection'!I3:I50002)"),
    ]
    for i, (label, formula) in enumerate(fin_rows):
        row = r + 1 + i
        ws.cell(row=row, column=2, value=label).font = Font(bold=True)
        ws.cell(row=row, column=3).value = formula
        ws.cell(row=row, column=3).number_format = '#,##0'
        for c in range(2, 4):
            ws.cell(row=row, column=c).border = thin_border

    # ── Upcoming Events ─────────────────────────────────────────────────
    r = r + len(fin_rows) + 3
    ws.cell(row=r, column=2, value="📅 UPCOMING EVENTS").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1
    for ci, h in enumerate(["Date", "Event", "Type", "Applicable To"]):
        ws.cell(row=r, column=2 + ci, value=h)
    style_header_row(ws, r, 4)

    # ── Recent Notices ──────────────────────────────────────────────────
    r += 12
    ws.cell(row=r, column=2, value="📢 RECENT NOTICES").font = section_font
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2).fill = section_fill
    r += 1
    for ci, h in enumerate(["Date", "Title", "Audience", "Status"]):
        ws.cell(row=r, column=2 + ci, value=h)
    style_header_row(ws, r, 4)

    # ── Print setup ─────────────────────────────────────────────────────
    ws.sheet_properties.pageSetUpPr = None

    # Set column widths
    for col in range(1, 15):
        ws.column_dimensions[get_column_letter(col)].width = 18

    # Freeze
    ws.freeze_panes = "A3"


# ═══════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("🏫 Building Comprehensive School Management System...")
    wb = build_workbook()

    output_path = "School_Management_System.xlsx"
    wb.save(output_path)
    print(f"✅ Successfully created: {output_path}")
    print(f"   Total sheets: {len(wb.sheetnames)}")
    print(f"   Sheets: {', '.join(wb.sheetnames)}")
