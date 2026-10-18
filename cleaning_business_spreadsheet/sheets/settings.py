"""
The `Settings` tab - business profile, pricing, pay rates, targets, alerts.

Everything the rest of the workbook needs to know about your cleaning
business is entered once here and published as a defined name (BizName,
HourlyRate, MinimumCharge, PayRate, TaxRate, PaymentTerms ...) so the
formulas stay readable and there is exactly one place to change a price.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

SHEET = "Settings"
LAST_COL = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, SHEET, tab_color=Palette.accent)
    b = cfg.business

    s.banner("Business Settings",
             "Fill in the white cells once. Shaded cells are locked "
             "formulas. Every price, every invoice and every alert in "
             "the workbook reads these values.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 34, fmts.chrome("label"))
    s.set_col(2, 2, 14, fmts.chrome("label"))
    s.set_col(3, 3, 24)
    s.set_col(4, 4, 18)
    s.set_col(5, 8, 17, fmts.chrome("note"))

    r = 4  # 0 based row index -> Excel row 5

    # -------------------------------------------------------- 1. the business
    r = s.section(r, "1. Your business", LAST_COL)
    r = s.field(r, "Business name", b.name, name="BizName", value_end=4,
                help_text="Shown on every quote, invoice and printable "
                          "page.")
    r = s.field(r, "Owner / manager", b.owner, name="BizOwner", value_end=4)
    r = s.field(r, "Email", b.email, name="BizEmail", value_end=4)
    r = s.field(r, "Phone", b.phone, name="BizPhone", value_end=4)
    r = s.field(r, "Address", b.address, name="BizAddress", value_end=4)
    r = s.field(r, "Website", b.website, name="BizWebsite", value_end=4)
    r = s.field(r, "Area you cover", b.area, name="BizArea", value_end=4,
                help_text="Used on quotes and to sanity check how far a "
                          "job is.")
    r = s.field(r, "Tagline", b.tagline, name="BizTagline", value_end=4)
    r = s.field(r, "Business / tax number", b.reg_number, name="BizRegNo",
                value_end=4)
    r = s.field(r, "Insurance policy", b.insurance, name="BizInsurance",
                value_end=4,
                help_text="Clients ask for this more often than you think "
                          "- keep it here so it is easy to quote.")

    # ---------------------------------------------------------- 2. the diary
    r += 1
    r = s.section(r, "2. Your working week", LAST_COL)
    r = s.field(r, "Year", cfg.year, style="int", name="PlanYear",
                validate=number_check(2000, 2100, decimal=False),
                help_text="The year the calendar, the planner and the "
                          "reports open on.")
    r = s.field(r, "Week starts on", cfg.week_start, name="WeekStart",
                validate=list_from("WeekStartList"),
                help_text="Used by the calendar grid, the staff rota and "
                          "the week numbers.")
    r = s.field(r, "Billable hours in a working day", cfg.productive_hours,
                style="number1", name="ProductiveHours",
                validate=number_check(1, 16),
                help_text="Hours one cleaner can actually charge for "
                          "after travel, setup and breaks. The capacity "
                          "and utilisation figures use this.")
    r = s.field(r, "Cleaners on a standard job", cfg.cleaners_per_job,
                style="int", name="CleanersPerJob",
                validate=number_check(1, 20, decimal=False),
                help_text="Two cleaners for three hours is six paid "
                          "hours - the costing tab works that out for "
                          "you.")
    r = s.field(r, "Travel allowance between jobs (minutes)",
                cfg.travel_minutes, style="int", name="TravelMinutes",
                validate=number_check(0, 240, decimal=False),
                help_text="Added to each job when the scheduler checks "
                          "whether two bookings clash.")
    r = s.field(r, "Area measured in", cfg.area_unit, name="AreaUnit",
                validate=list_from("AreaUnitList"),
                help_text="Square feet or square metres - the price per "
                          "unit below should match.")
    r = s.field(r, "Distance measured in", cfg.distance_unit,
                name="DistanceUnit", validate=list_from("DistanceUnitList"))
    r = s.field(r, "Today", None, style="date", mode="calc",
                name="TodayDate", formula="=TODAY()",
                help_text="Every 'due', 'overdue' and 'coming up' flag in "
                          "the workbook counts from this cell, so they all "
                          "move on by themselves overnight.")
    r = s.field(r, "This month starts", None, style="date", mode="calc",
                name="MonthStart",
                formula="=DATE(YEAR(TodayDate),MONTH(TodayDate),1)")
    r = s.field(r, "This month ends", None, style="date", mode="calc",
                name="MonthEnd",
                formula="=DATE(YEAR(TodayDate),MONTH(TodayDate)+1,1)-1")
    r = s.field(r, "This week starts", None, style="date", mode="calc",
                name="WeekStartDate",
                formula="=TodayDate-WEEKDAY(TodayDate,"
                        "IF(WeekStart=\"Sunday\",1,2))+1")

    # --------------------------------------------------------- 3. your prices
    r += 1
    r = s.section(r, "3. What you charge", LAST_COL)
    r = s.note(r, "These are the numbers the quote calculator multiplies. "
                  "Set them once from your own costs and local market - "
                  "the figures shipped with the workbook are a realistic "
                  "starting point, not advice on what to charge.",
               LAST_COL, first_col=1, height=30)
    r = s.field(r, "Standard hourly rate", cfg.hourly_rate, style="money",
                name="HourlyRate", validate=number_check(0, 10000),
                help_text="What one cleaner's hour costs the client.")
    r = s.field(r, "Minimum charge for any job", cfg.minimum_charge,
                style="money", name="MinimumCharge",
                validate=number_check(0, 10000),
                help_text="No quote leaves below this, however small the "
                          "job. It covers getting there and back.")
    r = s.field(r, f"Price per {cfg.area_unit}", cfg.rate_per_area,
                style="money", name="RatePerArea",
                validate=number_check(0, 100),
                help_text="Used by the area based pricing method - "
                          "common for offices and end of tenancy work.")
    r = s.field(r, "Price per bedroom", cfg.rate_per_bedroom, style="money",
                name="RatePerBedroom", validate=number_check(0, 10000))
    r = s.field(r, "Price per bathroom", cfg.rate_per_bathroom,
                style="money", name="RatePerBathroom",
                validate=number_check(0, 10000),
                help_text="Bathrooms and kitchens take the longest - "
                          "price them higher than a bedroom.")
    r = s.field(r, "Price per other room", cfg.rate_per_room, style="money",
                name="RatePerRoom", validate=number_check(0, 10000))
    r = s.field(r, "Call out / travel fee", cfg.callout_fee, style="money",
                name="CalloutFee", validate=number_check(0, 10000))
    r = s.field(r, "Deep clean uplift", cfg.deep_clean_uplift,
                style="percent0", name="DeepCleanUplift",
                validate=number_check(0, 5),
                help_text="A deep clean takes far longer than a weekly "
                          "visit. 45% means a $200 clean quotes at $290.")
    r = s.field(r, "Move in / move out uplift", cfg.move_clean_uplift,
                style="percent0", name="MoveCleanUplift",
                validate=number_check(0, 5))
    r = s.field(r, "Pet surcharge", cfg.pet_surcharge, style="money",
                name="PetSurcharge", validate=number_check(0, 10000),
                help_text="Hair and paw prints are real extra minutes. "
                          "Set it to 0 if you would rather absorb it.")
    r = s.field(r, "Deposit to take up front", cfg.deposit_percent,
                style="percent0", name="DepositPercent",
                validate=number_check(0, 1),
                help_text="Shown on quotes and contracts for larger jobs.")
    r = s.field(r, "Price rounding", "Nearest 5", name="PriceRounding",
                validate=list_from("RoundingList"),
                help_text="Quotes look more professional rounded. "
                          "'Nearest 5' turns 187.40 into 185.")

    # ----------------------------------------------------- 4. what it costs
    r += 1
    r = s.section(r, "4. What the work costs you", LAST_COL)
    r = s.field(r, "Cleaner pay rate an hour", cfg.pay_rate, style="money",
                name="PayRate", validate=number_check(0, 10000),
                help_text="Gross pay before your on-costs.")
    r = s.field(r, "Overtime multiplier", cfg.overtime_multiplier,
                style="number1", name="OvertimeRate",
                validate=number_check(1, 5),
                help_text="1.5 = time and a half on overtime hours.")
    r = s.field(r, "Employer on-costs", cfg.employer_oncost,
                style="percent0", name="EmployerOncost",
                validate=number_check(0, 2),
                help_text="Payroll tax, insurance, holiday pay, sick pay "
                          "- the part of an employee most owners forget "
                          "to price in.")
    r = s.field(r, "True cost of one cleaner hour", None, style="money",
                mode="calc", name="CostPerHour",
                formula="=PayRate*(1+EmployerOncost)",
                help_text="Pay plus on-costs. Every job costing and "
                          "profit figure in the workbook uses this.")
    r = s.field(r, "Supplies used per hour", cfg.supplies_per_hour,
                style="money", name="SuppliesPerHour",
                validate=number_check(0, 1000),
                help_text="Chemicals, cloths, bags - the consumables a "
                          "cleaning hour burns through.")
    r = s.field(r, f"Mileage rate per {cfg.distance_unit[:-1]}",
                cfg.travel_rate, style="money", name="MileageRate",
                validate=number_check(0, 100),
                help_text="What you reimburse (or claim) for travel.")
    r = s.field(r, "Target gross margin", cfg.target_margin,
                style="percent0", name="TargetMargin",
                validate=number_check(0, 1),
                help_text="What a healthy job should make after wages "
                          "and supplies. The quote calculator shows you "
                          "the price that hits it.")
    r = s.field(r, "Flag a job below this margin", cfg.margin_floor,
                style="percent0", name="MarginFloor",
                validate=number_check(0, 1),
                help_text="Anything thinner than this turns red on the "
                          "costing tab and the dashboard.")

    # --------------------------------------------------------- 5. invoicing
    r += 1
    r = s.section(r, "5. Invoicing and payment", LAST_COL)
    r = s.field(r, "Sales tax / VAT rate", cfg.tax_rate, style="percent0",
                name="TaxRate", validate=number_check(0, 1),
                help_text="Leave at 0 if you are not registered. Every "
                          "invoice total adds this to the net amount.")
    r = s.field(r, "What you call the tax", cfg.tax_label, name="TaxLabel",
                help_text="VAT, GST, Sales tax - whatever the invoice "
                          "should say.")
    r = s.field(r, "Payment terms (days)", cfg.payment_terms_days,
                style="int", name="PaymentTerms",
                validate=number_check(0, 180, decimal=False),
                help_text="Due date = invoice date plus this. The overdue "
                          "flags count from the due date.")
    r = s.field(r, "Late fee", cfg.late_fee_percent, style="percent0",
                name="LateFeePercent", validate=number_check(0, 1),
                help_text="Only charge it if your contract says you can.")
    r = s.field(r, "Grace days before 'overdue'", cfg.overdue_grace_days,
                style="int", name="OverdueGrace",
                validate=number_check(0, 60, decimal=False))
    r = s.field(r, "A quote stays valid for (days)", cfg.quote_valid_days,
                style="int", name="QuoteValidDays",
                validate=number_check(1, 365, decimal=False))
    r = s.field(r, "Bank / payment details", "Bright & Tidy Ltd - "
                                             "00-00-00 12345678",
                name="BizPayDetails", value_end=4,
                help_text="Printed at the bottom of every invoice.")

    # ------------------------------------------------------- 6. the targets
    r += 1
    r = s.section(r, "6. Targets the dashboard measures you against",
                  LAST_COL)
    r = s.field(r, "Revenue target a month", cfg.revenue_target_month,
                style="money0", name="MonthlyRevenueTarget",
                validate=number_check(0, 100000000))
    r = s.field(r, "Revenue target a year", cfg.revenue_target_year,
                style="money0", name="AnnualRevenueTarget",
                validate=number_check(0, 1000000000))
    r = s.field(r, "Expense budget a month", cfg.expense_budget_month,
                style="money0", name="MonthlyExpenseBudget",
                validate=number_check(0, 100000000))
    r = s.field(r, "Utilisation target", cfg.utilisation_target,
                style="percent0", name="UtilisationTarget",
                validate=number_check(0, 1),
                help_text="Share of the hours you pay for that you "
                          "actually bill. 75% is good for a small team.")
    r = s.field(r, "Client retention target", cfg.retention_target,
                style="percent0", name="RetentionTarget",
                validate=number_check(0, 1))
    r = s.field(r, "Quote win rate target", cfg.quote_win_target,
                style="percent0", name="QuoteWinTarget",
                validate=number_check(0, 1),
                help_text="Winning every quote usually means you are too "
                          "cheap. Half is a healthy place to be.")
    r = s.field(r, "Satisfaction target (out of 5)",
                cfg.satisfaction_target, style="number1",
                name="SatisfactionTarget", validate=number_check(0, 5))

    # ------------------------------------------------------- 7. the alerts
    r += 1
    r = s.section(r, "7. When the workbook should nudge you", LAST_COL)
    r = s.field(r, "Remind me about a job this many days early",
                cfg.reminder_days, style="int", name="ReminderDays",
                validate=number_check(0, 60, decimal=False),
                help_text="Drives the 'coming up' list on the dashboard "
                          "and the confirmation chase list.")
    r = s.field(r, "Chase an unanswered quote after (days)",
                cfg.quote_followup_days, style="int",
                name="QuoteFollowupDays",
                validate=number_check(0, 120, decimal=False),
                help_text="The single highest value number on this tab. "
                          "Most lost quotes were never followed up.")
    r = s.field(r, "A client is 'going quiet' after (days)",
                cfg.inactive_client_days, style="int",
                name="InactiveClientDays",
                validate=number_check(0, 999, decimal=False),
                help_text="Used by the retention tab to spot clients "
                          "drifting away before they are gone.")
    r = s.field(r, "Days of stock that counts as low", cfg.low_stock_days,
                style="int", name="LowStockDays",
                validate=number_check(0, 365, decimal=False))
    r = s.field(r, "Service equipment every (days)", cfg.maintenance_days,
                style="int", name="MaintenanceDays",
                validate=number_check(0, 3650, decimal=False))
    r = s.field(r, "Warn before a certificate expires (days)",
                cfg.certificate_warning_days, style="int",
                name="CertWarningDays",
                validate=number_check(0, 365, decimal=False),
                help_text="Insurance, background checks, training - the "
                          "Safety tab turns amber this far ahead.")
    r = s.field(r, "Warn before a contract renews (days)",
                cfg.renewal_warning_days, style="int",
                name="RenewalWarningDays",
                validate=number_check(0, 365, decimal=False))

    # ------------------------------------------------------- 8. ID prefixes
    r += 1
    r = s.section(r, "8. ID prefixes", LAST_COL)
    r = s.note(r, "The letters in front of the automatic numbers. Change "
                  "them before you start entering data, not after.",
               LAST_COL, first_col=1, height=24)
    for label, value, name in [
            ("Client IDs", "CL-", "ClientPrefix"),
            ("Property IDs", "PR-", "PropertyPrefix"),
            ("Lead IDs", "LD-", "LeadPrefix"),
            ("Service IDs", "SV-", "ServicePrefix"),
            ("Quote IDs", "Q-", "QuotePrefix"),
            ("Job IDs", "JOB-", "JobPrefix"),
            ("Invoice IDs", "INV-", "InvoicePrefix"),
            ("Payment IDs", "PAY-", "PaymentPrefix"),
            ("Employee IDs", "EMP-", "EmployeePrefix"),
            ("Supply IDs", "SUP-", "SupplyPrefix"),
            ("Equipment IDs", "EQ-", "EquipmentPrefix"),
            ("Expense IDs", "EXP-", "ExpensePrefix"),
            ("Complaint IDs", "CMP-", "ComplaintPrefix")]:
        r = s.field(r, label, value, name=name)

    # ---------------------------------------------------------- 9. about it
    r += 1
    r = s.section(r, "9. About this workbook", LAST_COL)
    r = s.field(r, "Product", cfg.product_name, mode="calc", value_end=5)
    r = s.field(r, "Version", cfg.version, mode="calc")
    r = s.field(r, "Created by", cfg.author, mode="calc")
    r = s.field(r, "Sheet protection password", cfg.password, mode="calc",
                help_text="Review > Unprotect Sheet if you need to change "
                          "a formula.")
    r = s.field(r, "Currency", f"{cfg.currency_symbol} "
                               f"({cfg.currency_code})", mode="calc")
    r += 1
    s.note(r,
           "Changed a price here? Everything updates straight away - the "
           "quote calculator, every open quote, the job costing and the "
           "dashboard all read these cells rather than keeping their own "
           "copy. Invoices you have already sent keep the amount you "
           "typed on them.", LAST_COL, first_col=1, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
