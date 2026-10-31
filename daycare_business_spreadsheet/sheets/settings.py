"""`Settings & Lists` - every number and every list the workbook runs on.

Nothing in this product is hard coded to one country, one licence or
one price list.  The age bands and the staff-to-child ratios that go
with them, the tax switches, the fee basis, the late collection
charge, the overtime threshold, the budget and all forty-odd
drop-downs are set here, once, and every other tab follows.

The ratio table in particular is a starting point and not a legal
opinion.  How many two-year-olds one adult may supervise is a local
licensing decision; the workbook's job is to apply whatever you tell
it, consistently, on every room and every day.
"""

from __future__ import annotations

import datetime as _dt
from typing import List, Tuple

from ..config import (AGE_BANDS, ALLERGY_SEVERITY, ATTENDANCE_STATUSES,
                      BILLING_BASES, CHARGE_SOURCES, COMM_CATEGORIES,
                      COMM_DIRECTIONS, COMM_STATUSES, CONTACT_METHODS,
                      DIET_TYPES, DISCLAIMER, DOCUMENT_HOLDERS,
                      DOCUMENT_STATUSES, DOCUMENT_TYPES,
                      EATEN_AMOUNTS, EMPLOYMENT_STATUSES,
                      ENROLLMENT_STATUSES, EXPENSE_CATEGORIES,
                      FEE_TYPES, FREQUENCIES, GENDERS,
                      ID_PREFIXES, INCIDENT_CATEGORIES,
                      INCIDENT_LOCATIONS, INCIDENT_SEVERITY,
                      INQUIRY_SOURCES, LEAVE_TYPES, LEAVING_REASONS,
                      LEDGER_TYPES, MEAL_SITTINGS,
                      MEDICAL_RECORD_TYPES, MOODS, PAY_PERIODS,
                      PAY_TYPES, PAYMENT_METHODS, PRIVACY_NOTE,
                      PROGRAM_TYPES, RELATIONSHIPS, ROOM_STATUSES,
                      SCHEDULE_TYPES, SHIFT_TYPES, STAFF_ROLES,
                      SUPPLY_CATEGORIES, TOILETING_OPTIONS,
                      WAITLIST_PRIORITY, WAITLIST_STATUSES, Config,
                      Palette)
from ..layout import Registry, col_letter, dynamic_name, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Settings & Lists"
LAST_COL = 11

#: name -> values.  Everything the workbook offers you in a drop-down.
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("GenderList", list(GENDERS)),
    ("ProgramList", list(PROGRAM_TYPES)),
    ("EnrollmentStatusList", list(ENROLLMENT_STATUSES)),
    ("ScheduleTypeList", list(SCHEDULE_TYPES)),
    ("AttendanceStatusList", list(ATTENDANCE_STATUSES)),
    ("BillingBasisList", list(BILLING_BASES)),
    ("FeeTypeList", list(FEE_TYPES)),
    ("ChargeSourceList", list(CHARGE_SOURCES)),
    ("LedgerTypeList", list(LEDGER_TYPES)),
    ("PaymentMethodList", list(PAYMENT_METHODS)),
    ("RelationshipList", list(RELATIONSHIPS)),
    ("ContactMethodList", list(CONTACT_METHODS)),
    ("DirectionList", list(COMM_DIRECTIONS)),
    ("MoodList", list(MOODS)),
    ("MealSittingList", list(MEAL_SITTINGS)),
    ("EatenAmountList", list(EATEN_AMOUNTS)),
    ("DietList", list(DIET_TYPES)),
    ("ToiletingList", list(TOILETING_OPTIONS)),
    ("AllergySeverityList", list(ALLERGY_SEVERITY)),
    ("MedicalRecordList", list(MEDICAL_RECORD_TYPES)),
    ("IncidentCategoryList", list(INCIDENT_CATEGORIES)),
    ("IncidentLocationList", list(INCIDENT_LOCATIONS)),
    ("IncidentSeverityList", list(INCIDENT_SEVERITY)),
    ("RoleList", list(STAFF_ROLES)),
    ("EmploymentStatusList", list(EMPLOYMENT_STATUSES)),
    ("PayTypeList", list(PAY_TYPES)),
    ("ShiftTypeList", list(SHIFT_TYPES)),
    ("LeaveTypeList", list(LEAVE_TYPES)),
    ("PayPeriodList", list(PAY_PERIODS)),
    ("ExpenseCategoryList", list(EXPENSE_CATEGORIES)),
    ("FrequencyList", list(FREQUENCIES)),
    ("SupplyCategoryList", list(SUPPLY_CATEGORIES)),
    ("DocumentTypeList", list(DOCUMENT_TYPES)),
    ("DocumentHolderList", list(DOCUMENT_HOLDERS)),
    ("DocumentStatusList", list(DOCUMENT_STATUSES)),
    ("CommCategoryList", list(COMM_CATEGORIES)),
    ("CommStatusList", list(COMM_STATUSES)),
    ("InquirySourceList", list(INQUIRY_SOURCES)),
    ("WaitlistStatusList", list(WAITLIST_STATUSES)),
    ("WaitlistPriorityList", list(WAITLIST_PRIORITY)),
    ("LeavingReasonList", list(LEAVING_REASONS)),
    ("RoomStatusList", list(ROOM_STATUSES)),
]


def _month_keys(today: _dt.date, back: int = 15, ahead: int = 9
                ) -> List[str]:
    """`yyyy-mm` keys around today, oldest first."""
    keys: List[str] = []
    year, month = today.year, today.month
    start = month - back
    while start <= 0:
        start += 12
        year -= 1
    for step in range(back + ahead + 1):
        m = start + step
        y = year + (m - 1) // 12
        m = (m - 1) % 12 + 1
        keys.append(f"{y:04d}-{m:02d}")
    return keys


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    sheet_ref = quote_sheet(NAME)

    s.banner("Settings & Lists",
             "Your setting, your fees, your ratios, your lists. Set "
             "these once and every tab in the workbook follows.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 23)
    s.set_col(3, 3, 20)
    s.set_col(4, LAST_COL, 17)

    row = 4
    row = s.note(
        row,
        "White cells are yours to change; everything else is locked "
        "and calculated. Nothing here is specific to one country - set "
        "your own age bands and ratios, your own tax switches, your "
        "own fees and currency, and the rest of the workbook follows.",
        LAST_COL, height=32)
    row += 1

    # ---------------------------------------------------- the setting ---
    row = s.section(row, "Your setting", LAST_COL)
    row = s.field(row, "Setting name", cfg.centre.name, name="CentreName",
                  value_end=5, help_col=7,
                  help_text="Printed at the top of every report and "
                            "invoice.")
    row = s.field(row, "Owner / director", cfg.centre.owner,
                  name="OwnerName", value_end=5, help_col=7)
    row = s.field(row, "Address", cfg.centre.address, value_end=5,
                  help_col=7)
    row = s.field(row, "Phone", cfg.centre.phone, value_end=5,
                  help_col=7)
    row = s.field(row, "Email", cfg.centre.email, value_end=5,
                  help_col=7)
    row = s.field(row, "Website", cfg.centre.website, value_end=5,
                  help_col=7)
    row = s.field(row, "Licence number", cfg.centre.licence,
                  value_end=5, help_col=7,
                  help_text="Your childcare registration or licence "
                            "reference, for your own records.")
    row = s.field(row, "Tax / business number", cfg.centre.tax_id,
                  value_end=5, help_col=7)
    row = s.field(row, "Opening hours", cfg.centre.opening_hours,
                  value_end=5, help_col=7)
    row = s.field(row, "Strapline", cfg.centre.tagline, value_end=5,
                  help_col=7)
    row = s.field(row, "Currency symbol", cfg.currency_symbol,
                  name="CurrencySign",
                  help_text="Shown on screen. The cell formats follow "
                            "the symbol you build the file with.")
    row = s.field(row, "Currency code", cfg.currency_code)
    row = s.field(row, "Tracking started",
                  cfg.today - _dt.timedelta(days=90), style="date",
                  name="TrackingStart",
                  help_text="The day you began logging in this file. "
                            "Supply cover and averages measure from "
                            "here.")
    row = s.field(row, "Today", None, style="date", name="TodayDate",
                  mode="calc", formula="=TODAY()",
                  help_text="Live. Everything that says 'overdue', "
                            "'expiring' or 'this month' reads this "
                            "cell.")
    row = s.field(row, "This month", None, style="text",
                  name="ThisMonthKey", mode="calc",
                  formula='=TEXT(TODAY(),"yyyy-mm")',
                  help_text="The month key every monthly total groups "
                            "on.")
    row = s.field(row, "Weeks tracked", None, style="number1",
                  name="TrackingWeeks", mode="calc",
                  formula="=MAX(4,ROUND((TodayDate-TrackingStart)/7,0))",
                  help_text="Used for the 'used a week' and weeks of "
                            "cover figures on the supply tab.")
    row += 1

    # ------------------------------------------- opening and capacity ---
    row = s.section(row, "Opening hours, capacity and the year",
                    LAST_COL)
    row = s.field(row, "Opens at", cfg.open_hour, style="number1",
                  name="OpenHour", validate=number_check(0, 24),
                  help_text="As a 24 hour number - 7 is 7am, 18 is 6pm.")
    row = s.field(row, "Closes at", cfg.close_hour, style="number1",
                  name="CloseHour", validate=number_check(0, 24))
    row = s.field(row, "Hours open a day", None, style="number1",
                  name="OpenHoursDay", mode="calc",
                  formula="=MAX(1,CloseHour-OpenHour)",
                  help_text="The ratio figures divide rota hours by "
                            "this to work out how many adults were "
                            "actually in a room.")
    row = s.field(row, "Days open a week", cfg.open_days, style="int",
                  name="OpenDays",
                  validate=number_check(1, 7, decimal=False))
    row = s.field(row, "Weeks open a year", cfg.weeks_open, style="int",
                  name="WeeksOpen",
                  validate=number_check(1, 53, decimal=False),
                  help_text="Trading weeks, after closures. What a "
                            "child's place is worth over a year uses "
                            "this.")
    row = s.field(row, "Weeks paid a year", 52, style="int",
                  name="WeeksPaid",
                  validate=number_check(1, 53, decimal=False),
                  help_text="Used to turn a salary into an hourly "
                            "rate. Usually 52 even when you close for "
                            "two of them.")
    row = s.field(row, "Weeks in a term", cfg.term_weeks, style="int",
                  name="TermWeeks",
                  validate=number_check(1, 30, decimal=False),
                  help_text="Only used when a child is on a term "
                            "rate.")
    row = s.field(row, "Licensed capacity", cfg.licensed_capacity,
                  style="int", name="LicensedCapacity",
                  validate=number_check(1, 1000, decimal=False),
                  help_text="The number on your licence, for the whole "
                            "setting. Room by room capacity is on the "
                            "Classroom Management tab.")
    row = s.field(row, "Warn when a room is this full",
                  cfg.ratio_warning_percent, style="percent",
                  name="RatioWarning", validate=number_check(0, 1),
                  help_text="Rooms at or above this go amber, so you "
                            "open the waitlist before the last place "
                            "goes.")
    row += 1

    # ------------------------------------------------------ age bands ---
    row = s.section(row, "Age bands and staff-to-child ratios", LAST_COL)
    row = s.note(
        row,
        "These are the most country-specific numbers in the workbook, "
        "which is exactly why they are a table you can edit rather "
        "than a formula you cannot. Set them to whatever your own "
        "licence says. Every room, every ratio warning and every "
        "'moves up on' date follows this table.",
        LAST_COL, height=38)
    row += 1
    band_head = row
    s.table_head(band_head, 1, [
        "Age band", "From (months)", "To (months)",
        "Children per adult", "Max group size"],
        widths=[22, 15, 15, 18, 16])
    first_band = band_head + 1
    band_rows = len(AGE_BANDS) + 2
    for index in range(band_rows):
        line = first_band + index
        if index < len(AGE_BANDS):
            label, m_from, m_to, ratio, group = AGE_BANDS[index]
            s.write(line, 1, label, fmts.input("text"))
            s.write(line, 2, m_from, fmts.input("int"))
            s.write(line, 3, m_to, fmts.input("int"))
            s.write(line, 4, ratio, fmts.input("int"))
            s.write(line, 5, group, fmts.input("int"))
        else:
            for col in range(1, 6):
                s.write(line, col, None,
                        fmts.input("text" if col == 1 else "int"))
    last_band = first_band + band_rows
    for key, column in (("AgeBandList", 1), ("AgeBandFrom", 2),
                        ("AgeBandTo", 3), ("AgeBandRatio", 4),
                        ("AgeBandMax", 5)):
        book.define_name(
            key,
            f"={sheet_ref}!${col_letter(column)}${first_band + 1}:"
            f"${col_letter(column)}${last_band}")
    row = last_band + 1

    # ---------------------------------------------------------- fees ----
    row = s.section(row, "Fees and billing", LAST_COL)
    row = s.note(
        row,
        "Each child's own basis and rate live on the Child Enrollment "
        "tab - these are the standard rates the waitlist and the "
        "quick calculations use.",
        LAST_COL, height=26)
    row += 1
    row = s.field(row, "Full-time week", cfg.full_time_week,
                  style="money", name="FullTimeWeek",
                  validate=number_check(0, 100000))
    row = s.field(row, "Part-time day", cfg.part_time_day,
                  style="money", name="PartTimeDay",
                  validate=number_check(0, 100000))
    row = s.field(row, "Hourly rate", cfg.hourly_rate_fee,
                  style="money", name="HourlyFee",
                  validate=number_check(0, 10000))
    row = s.field(row, "Registration fee", cfg.registration_fee,
                  style="money", name="RegistrationFee",
                  validate=number_check(0, 100000))
    row = s.field(row, "Enrollment deposit", cfg.enrollment_deposit,
                  style="money", name="EnrollmentDeposit",
                  validate=number_check(0, 100000),
                  help_text="A deposit is a payment with no invoice "
                            "behind it yet, so it can never be "
                            "counted as income twice.")
    row = s.field(row, "Meal fee a day", cfg.meal_fee_day,
                  style="money", name="MealFeeDay",
                  validate=number_check(0, 10000))
    row = s.field(row, "Transport fee a week", cfg.transport_fee_week,
                  style="money", name="TransportFeeWeek",
                  validate=number_check(0, 10000))
    row = s.field(row, "Sibling discount", cfg.sibling_discount,
                  style="percent", name="SiblingDiscount",
                  validate=number_check(0, 1),
                  help_text="Applied automatically to the second and "
                            "later child of a family.")
    row = s.field(row, "Staff discount", cfg.staff_discount,
                  style="percent", name="StaffDiscount",
                  validate=number_check(0, 1))
    row = s.field(row, "Payment terms, days", cfg.payment_terms_days,
                  style="int", name="PaymentTermsDays",
                  validate=number_check(0, 365, decimal=False),
                  help_text="How long after the invoice date it falls "
                            "due.")
    row = s.field(row, "Invoice on day of month", cfg.invoice_day,
                  style="int", name="InvoiceDay",
                  validate=number_check(1, 28, decimal=False))
    row = s.field(row, "Late payment fee", cfg.late_payment_fee,
                  style="money", name="LatePaymentFee",
                  validate=number_check(0, 10000))
    row += 1

    # ----------------------------------------------- late collection ----
    row = s.section(row, "Late collection charges", LAST_COL)
    row = s.note(
        row,
        "Measured against each child's own contracted pickup time on "
        "the enrollment tab, not against closing time. Free for the "
        "grace period, then the first block, then per minute - which "
        "is how most parent handbooks word it.",
        LAST_COL, height=32)
    row += 1
    row = s.field(row, "Grace, minutes", cfg.late_pickup_grace,
                  style="int", name="LatePickupGrace",
                  validate=number_check(0, 120, decimal=False))
    row = s.field(row, "First block, minutes",
                  cfg.late_pickup_block, style="int",
                  name="LatePickupBlock",
                  validate=number_check(0, 120, decimal=False))
    row = s.field(row, "First block charge", cfg.late_pickup_fee,
                  style="money", name="LatePickupFee",
                  validate=number_check(0, 10000))
    row = s.field(row, "Then per minute", cfg.late_pickup_per,
                  style="money", name="LatePickupPerMin",
                  validate=number_check(0, 1000))
    row += 1

    # ----------------------------------------------------- tax & fees ---
    row = s.section(row, "Tax and payment processing", LAST_COL)
    row = s.note(
        row,
        "Childcare is tax exempt in a great many places, so the "
        "default here is zero and nothing is assumed. Set the rate and "
        "say what it applies to - tuition, extras, both or neither.",
        LAST_COL, height=28)
    row += 1
    row = s.field(row, "Tax rate", cfg.tax_rate, style="percent",
                  name="TaxRate", validate=number_check(0, 1),
                  help_text="20% is 0.2 or 20% - both work.")
    row = s.field(row, "Tax on tuition?", cfg.tax_on_tuition,
                  name="TaxOnTuition", validate=list_from("YesNoList"))
    row = s.field(row, "Tax on extras?", cfg.tax_on_extras,
                  name="TaxOnExtras", validate=list_from("YesNoList"),
                  help_text="Meals, trips, late fees and everything "
                            "else that is not tuition.")
    row = s.field(row, "Tax registered?", cfg.tax_registered,
                  validate=list_from("YesNoList"),
                  help_text="For your own reference - it changes "
                            "nothing in the formulas.")
    row = s.field(row, "Card fee, percent", cfg.card_fee_percent,
                  style="percent", name="CardFeePercent",
                  validate=number_check(0, 1))
    row = s.field(row, "Card fee, fixed", cfg.card_fee_fixed,
                  style="money", name="CardFeeFixed",
                  validate=number_check(0, 100))
    row += 1

    # ------------------------------------------------------ staff pay ---
    row = s.section(row, "Staff pay and hours", LAST_COL)
    row = s.field(row, "Default hourly rate", cfg.default_hourly,
                  style="money", name="DefaultHourly",
                  validate=number_check(0, 1000))
    row = s.field(row, "Overtime after, hours a week",
                  cfg.overtime_after_hours, style="number1",
                  name="OvertimeAfter", validate=number_check(0, 100),
                  help_text="Hours above this in a single week are "
                            "paid at the multiplier below.")
    row = s.field(row, "Overtime multiplier",
                  cfg.overtime_multiplier, style="number1",
                  name="OvertimeMultiplier",
                  validate=number_check(1, 5))
    row = s.field(row, "Employer costs on wages", cfg.payroll_burden,
                  style="percent", name="PayrollBurden",
                  validate=number_check(0, 1),
                  help_text="Employer tax, insurance and cover as a "
                            "percentage of gross pay. Contractors and "
                            "volunteers carry none.")
    row = s.field(row, "Unpaid break, minutes", cfg.break_minutes,
                  style="int", name="BreakMinutes",
                  validate=number_check(0, 240, decimal=False))
    row = s.field(row, "Payroll sits in expenses?", "No",
                  name="PayrollInExpenses",
                  validate=list_from("YesNoList"),
                  help_text="Leave this No and wages come from the "
                            "Payroll Tracker only. Set it Yes and the "
                            "wage rows on the Expense Tracker are "
                            "counted instead. Never both.")
    row += 1

    # -------------------------------------------------------- targets ---
    row = s.section(row, "Targets the dashboard measures you against",
                    LAST_COL)
    row = s.field(row, "Revenue target, month",
                  cfg.revenue_target_month, style="money",
                  name="RevenueTargetMonth",
                  validate=number_check(0, 10000000))
    row = s.field(row, "Revenue target, year", cfg.revenue_target_year,
                  style="money", name="RevenueTargetYear",
                  validate=number_check(0, 100000000))
    row = s.field(row, "Profit target, month", cfg.profit_target_month,
                  style="money", name="ProfitTargetMonth",
                  validate=number_check(0, 10000000))
    row = s.field(row, "Children enrolled, target",
                  cfg.enrollment_goal, style="int",
                  name="EnrollmentTarget",
                  validate=number_check(0, 1000, decimal=False))
    row = s.field(row, "Occupancy target", cfg.target_occupancy,
                  style="percent", name="TargetOccupancy",
                  validate=number_check(0, 1))
    row = s.field(row, "Attendance target", cfg.target_attendance,
                  style="percent", name="TargetAttendance",
                  validate=number_check(0, 1))
    row = s.field(row, "Collection target", cfg.target_collection,
                  style="percent", name="TargetCollection",
                  validate=number_check(0, 1),
                  help_text="The share of what you bill that actually "
                            "arrives. Below 95% and something is "
                            "wrong with the chasing, not the "
                            "families.")
    row += 1

    # ------------------------------------------------------ reminders ---
    row = s.section(row, "Reminders and chase windows", LAST_COL)
    row = s.field(row, "Warn about expiry, days",
                  cfg.document_warning_days, style="int",
                  name="DocWarningDays",
                  validate=number_check(1, 365, decimal=False),
                  help_text="Amber this many days before a document "
                            "or certificate runs out.")
    row = s.field(row, "Urgent expiry, days", cfg.document_urgent_days,
                  style="int", name="DocUrgentDays",
                  validate=number_check(1, 365, decimal=False))
    row = s.field(row, "Birthday window, days",
                  cfg.birthday_window_days, style="int",
                  name="BirthdayWindow",
                  validate=number_check(1, 365, decimal=False))
    row = s.field(row, "Chase an absence after, days",
                  cfg.absence_chase_days, style="int",
                  name="AbsenceChaseDays",
                  validate=number_check(1, 30, decimal=False),
                  help_text="Unexplained absences inside this window "
                            "raise a prompt on the attendance tab.")
    row = s.field(row, "Chase an invoice after, days",
                  cfg.overdue_chase_days, style="int",
                  name="OverdueChaseDays",
                  validate=number_check(1, 180, decimal=False))
    row += 1

    # ------------------------------------------------------- supplies ---
    row = s.section(row, "Supplies and reordering", LAST_COL)
    row = s.field(row, "Order up to, weeks of cover",
                  cfg.reorder_weeks, style="number1",
                  name="ReorderWeeks", validate=number_check(0.5, 52))
    row = s.field(row, "Warn below, weeks of cover",
                  cfg.stock_warning_weeks, style="number1",
                  name="StockWarningWeeks",
                  validate=number_check(0, 52))
    row += 1

    # --------------------------------------------------- the budget ----
    row = s.section(row, "Monthly budget by category", LAST_COL)
    row = s.note(
        row,
        "What you expect to spend in a normal month. The Expense "
        "Tracker reads this table, compares it with what you actually "
        "spent and shows the variance in the month it happened. Leave "
        "a category at zero and it simply is not measured.",
        LAST_COL, height=32)
    row += 1
    budget_head = row
    s.table_head(budget_head, 1, ["Category", "Budget a month"],
                 widths=[28, 16])
    first_budget = budget_head + 1
    budget_rows = len(EXPENSE_CATEGORIES)
    #: a sensible opening budget so the variance column has something
    #: to say on day one - every figure editable
    seed = {
        "Rent": 3200, "Utilities": 620, "Food & groceries": 1850,
        "Nappies & hygiene": 420, "Cleaning supplies": 240,
        "Toys & learning materials": 300,
        "Educational supplies": 220, "Arts & crafts": 150,
        "Insurance": 310, "Licensing & registration": 90,
        "Repairs & maintenance": 260,
        "Marketing & advertising": 180,
        "Software & subscriptions": 95,
        "Telephone & internet": 110, "Office & admin": 95,
        "Training & qualifications": 150, "Bank & card fees": 85,
        "Professional fees": 140, "Waste & recycling": 70,
    }
    for index, category in enumerate(EXPENSE_CATEGORIES):
        line = first_budget + index
        s.write(line, 1, category, fmts.input("text"))
        s.write(line, 2, float(seed.get(category, 0)),
                fmts.input("money"))
    last_budget = first_budget + budget_rows
    book.define_name(
        "BudgetCategory",
        f"={sheet_ref}!$B${first_budget + 1}:$B${last_budget}")
    book.define_name(
        "BudgetAmount",
        f"={sheet_ref}!$C${first_budget + 1}:$C${last_budget}")
    total_line = last_budget
    s.write(total_line, 1, "Total budget a month",
            fmts.chrome("total_label"))
    s.write(total_line, 2, "=SUM(BudgetAmount)",
            fmts.chrome("total_value",
                        num_format=cfg.money_format))
    book.define_name(
        "BudgetMonthlyTotal",
        f"={sheet_ref}!$C${total_line + 1}")
    row = total_line + 2

    # ----------------------------------------------- workbook totals ----
    row = s.section(row, "Workbook totals", LAST_COL)
    row = s.note(
        row,
        "Calculated once here so several tabs can share them. Nothing "
        "to change.",
        LAST_COL, height=24)
    row += 1
    row = s.field(row, "Children holding a place", None, style="int",
                  name="ChildrenEnrolled", mode="calc",
                  formula='=COUNTIFS([[Child Enrollment.holds_place]],'
                          '"Yes")')
    row = s.field(row, "Children attending", None, style="int",
                  name="ChildrenActive", mode="calc",
                  formula='=COUNTIFS([[Child Enrollment.status]],'
                          '"Active")')
    row = s.field(row, "Staff on the books", None, style="int",
                  name="StaffCount", mode="calc",
                  formula='=COUNTIFS([[Staff Database.status]],'
                          '"Active")+'
                          'COUNTIFS([[Staff Database.status]],'
                          '"Probation")')
    row = s.field(row, "Places across all rooms", None, style="int",
                  name="RoomCapacity", mode="calc",
                  formula="=SUM([[Classroom Management.capacity]])")
    row = s.field(row, "Revenue, all time", None, style="money",
                  name="RevenueAll", mode="calc",
                  formula="=ROUND(SUM([[Invoices & Receipts.billed]]),"
                          "2)")
    row = s.field(row, "Collected, all time", None, style="money",
                  name="CollectedAll", mode="calc",
                  formula="=ROUND(SUM([[Invoices & Receipts.collected]])"
                          ",2)")
    row += 1

    # ------------------------------------------------------ prefixes ----
    row = s.section(row, "ID prefixes", LAST_COL)
    row = s.note(
        row,
        "Every record numbers itself from its row. Change a prefix "
        "here and the IDs on that tab change with it - but do it "
        "before you have typed three hundred attendance rows, not "
        "after.",
        LAST_COL, height=28)
    row += 1
    half = (len(ID_PREFIXES) + 1) // 2
    top = row
    for index, (label, default, name, _digits) in enumerate(ID_PREFIXES):
        column = 0 if index < half else 4
        line = top + (index if index < half else index - half)
        s.write(line, column + 1, label, fmts.chrome("label_bold"))
        s.write(line, column + 2, default, fmts.input("text"))
        book.define_name(
            name, f"={sheet_ref}!${col_letter(column + 2)}${line + 1}")
    row = top + half + 1

    # -------------------------------------------------------- lists -----
    row = s.section(row, "Drop-down lists", LAST_COL)
    row = s.note(
        row,
        "Add new entries directly under the last one in a column - "
        "every list has room to grow and blank rows are ignored. Do "
        "not delete a whole column: a drop-down pointing at a missing "
        "list stops working.",
        LAST_COL, height=30)
    row += 1

    per_row = 4
    spare = 6
    top = row
    column = 0
    tallest = 0
    for index, (name, values) in enumerate(LISTS):
        if index and index % per_row == 0:
            top += tallest + 2
            column = 0
            tallest = 0
        s.write(top, column, name.replace("List", ""),
                fmts.chrome("header"))
        s.set_col(column, column, 26)
        for offset, value in enumerate(values):
            s.write(top + 1 + offset, column, value, fmts.input("text"))
        for offset in range(len(values), len(values) + spare):
            s.write(top + 1 + offset, column, None, fmts.input("text"))
        first = top + 2
        last = top + 1 + len(values) + spare
        book.define_name(
            name,
            f"={sheet_ref}!${col_letter(column)}${first}:"
            f"${col_letter(column)}${last}")
        tallest = max(tallest, len(values) + spare)
        column += 3
    row = top + tallest + 3

    # ---------------------------------------------------- month keys ----
    row = s.section(row, "Month keys", LAST_COL)
    row = s.note(
        row,
        "The months the reports and the budget pick from. They are "
        "plain text in yyyy-mm form so that every monthly total in "
        "the workbook groups on exactly the same thing.",
        LAST_COL, height=28)
    row += 1
    keys = _month_keys(cfg.today)
    s.write(row, 0, "Month", fmts.chrome("header"))
    s.set_col(0, 0, 14)
    for offset, key in enumerate(keys):
        s.write(row + 1 + offset, 0, key, fmts.input("text"))
    book.define_name(
        "MonthKeyList",
        f"={sheet_ref}!$A${row + 2}:$A${row + 1 + len(keys)}")
    book.define_name(
        "MonthKeys",
        dynamic_name(NAME, "A", row + 2, row + 1 + len(keys)))
    row = row + len(keys) + 2

    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=30)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=74)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
