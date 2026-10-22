"""`Settings` - every number and every list the workbook runs on.

Nothing in this workbook is hard coded to one country, one pay model
or one price list. The tax switches, the commission basis, the
rounding rule, the loyalty scheme and all twenty-odd drop-downs are
set here, once, and every other tab follows.
"""

from __future__ import annotations

import datetime as _dt
from typing import List, Tuple

from ..config import (APPOINTMENT_STATUSES, BOOKING_SOURCES,
                      CAMPAIGN_STATUSES, CANCEL_REASONS, CLIENT_STATUSES,
                      COMMISSION_METHODS, COMMISSION_TIERS,
                      CONTACT_METHODS, EXPENSE_CATEGORIES, FREQUENCIES,
                      ID_PREFIXES, LOYALTY_REWARDS, MARKETING_CHANNELS,
                      PAY_PERIODS, PAYMENT_KINDS, PAYMENT_METHODS,
                      PRODUCT_CATEGORIES, PURCHASE_STATUSES,
                      DISCLAIMER, ROUNDING_RULES, SALE_KINDS,
                      SERVICE_CATEGORIES,
                      SERVICE_STATUSES, STAFF_STATUSES, STAFF_TYPES,
                      STYLIST_LEVELS, Config, Palette)
from ..layout import Registry, col_letter, dynamic_name, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Settings"
LAST_COL = 11

#: name -> values. Everything the workbook offers you in a drop-down.
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("ServiceCategoryList", list(SERVICE_CATEGORIES)),
    ("ServiceStatusList", list(SERVICE_STATUSES)),
    ("StylistLevelList", list(STYLIST_LEVELS)),
    ("AppointmentStatusList", list(APPOINTMENT_STATUSES)),
    ("SourceList", list(BOOKING_SOURCES)),
    ("CancelReasonList", list(CANCEL_REASONS)),
    ("ContactList", list(CONTACT_METHODS)),
    ("ClientStatusList", list(CLIENT_STATUSES)),
    ("SaleKindList", list(SALE_KINDS)),
    ("PaymentKindList", list(PAYMENT_KINDS)),
    ("PaymentMethodList", list(PAYMENT_METHODS)),
    ("StaffTypeList", list(STAFF_TYPES)),
    ("StaffStatusList", list(STAFF_STATUSES)),
    ("CommissionMethodList", list(COMMISSION_METHODS)),
    ("PayPeriodList", list(PAY_PERIODS)),
    ("ExpenseCategoryList", list(EXPENSE_CATEGORIES)),
    ("FrequencyList", list(FREQUENCIES)),
    ("ProductCategoryList", list(PRODUCT_CATEGORIES)),
    ("PurchaseStatusList", list(PURCHASE_STATUSES)),
    ("MarketingChannelList", list(MARKETING_CHANNELS)),
    ("CampaignStatusList", list(CAMPAIGN_STATUSES)),
    ("LoyaltyRewardList", list(LOYALTY_REWARDS)),
    ("RoundingRuleList", list(ROUNDING_RULES)),
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
             "Your salon, your prices, your pay model, your lists. Set "
             "these once and every tab in the workbook follows.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 22)
    s.set_col(3, 3, 20)
    s.set_col(4, LAST_COL, 17)

    row = 4
    row = s.note(
        row,
        "Blue cells are yours to change. Everything else is locked and "
        "calculated. Nothing here is specific to one country: set your "
        "own tax rate and switches, your own commission basis, your own "
        "currency, and the workbook does the rest.",
        LAST_COL, height=32)
    row += 1

    # ------------------------------------------------------- the salon ---
    row = s.section(row, "The salon", LAST_COL)
    row = s.field(row, "Salon name", cfg.salon.name, name="SalonName",
                  value_end=5, help_col=7,
                  help_text="Printed at the top of every report.")
    row = s.field(row, "Owner", cfg.salon.owner, name="OwnerName",
                  value_end=5, help_col=7)
    row = s.field(row, "Address", cfg.salon.address, value_end=5,
                  help_col=7)
    row = s.field(row, "Phone", cfg.salon.phone, value_end=5, help_col=7)
    row = s.field(row, "Email", cfg.salon.email, value_end=5, help_col=7)
    row = s.field(row, "Website", cfg.salon.website, value_end=5,
                  help_col=7)
    row = s.field(row, "Booking link", cfg.salon.booking_link,
                  value_end=5, help_col=7,
                  help_text="Where clients book. This workbook does "
                            "not connect to it.")
    row = s.field(row, "Social", cfg.salon.instagram, value_end=5,
                  help_col=7, help_text="Handle used on the marketing "
                                        "tab.")
    row = s.field(row, "Strapline", cfg.salon.tagline, value_end=5,
                  help_col=7)
    row = s.field(row, "Opening hours", cfg.salon.opening_hours,
                  value_end=5, help_col=7)
    row = s.field(row, "Tax / business number", cfg.salon.tax_id,
                  value_end=5, help_col=7)
    row = s.field(row, "Currency symbol", cfg.currency_symbol,
                  name="CurrencySign",
                  help_text="Used on screen. The cell formats follow the "
                            "symbol you build the file with.")
    row = s.field(row, "Currency code", cfg.currency_code)
    row = s.field(row, "Tracking started", cfg.today - _dt.timedelta(days=70),
                  style="date", name="TrackingStart",
                  help_text="The day you began logging in this file. "
                            "Stock cover and averages measure from here.")
    row = s.field(row, "Today", None, style="date", name="TodayDate",
                  mode="calc", formula="=TODAY()",
                  help_text="Live. Everything that says 'overdue', "
                            "'lapsing' or 'this month' reads this cell.")
    row = s.field(row, "This month", None, style="text",
                  name="ThisMonthKey", mode="calc",
                  formula='=TEXT(TODAY(),"yyyy-mm")',
                  help_text="The month key the dashboard totals on.")
    row = s.field(row, "Weeks tracked", None, style="number1",
                  name="TrackingWeeks", mode="calc",
                  formula="=MAX(4,ROUND((TodayDate-TrackingStart)/7,0))",
                  help_text="Used for 'used a week' and the stock cover "
                            "figures.")
    row += 1

    # ------------------------------------------------- opening & capacity --
    row = s.section(row, "Opening hours and capacity", LAST_COL)
    row = s.field(row, "Opens at", cfg.open_hour, style="number1",
                  name="OpenHour", validate=number_check(0, 24),
                  help_text="As a 24 hour number - 9 is 9am, 19 is 7pm.")
    row = s.field(row, "Closes at", cfg.close_hour, style="number1",
                  name="CloseHour", validate=number_check(0, 24))
    row = s.field(row, "Days open a week", cfg.open_days, style="int",
                  name="OpenDays", validate=number_check(1, 7,
                                                         decimal=False))
    row = s.field(row, "Chairs / stations", cfg.chairs, style="int",
                  name="Chairs", validate=number_check(1, 60,
                                                       decimal=False))
    row = s.field(row, "Sellable chair hours a week", None,
                  style="number1", name="ChairHoursWeek", mode="calc",
                  formula="=MAX(1,(CloseHour-OpenHour)*OpenDays*Chairs)",
                  help_text="Hours x days x chairs. The utilisation "
                            "figure on the dashboard divides by this.")
    row += 1

    # --------------------------------------------------------- money & tax
    row = s.section(row, "Money, tax and deposits", LAST_COL)
    row = s.note(
        row,
        "Sales tax works differently everywhere, so nothing is assumed. "
        "Set the rate and say what it applies to - retail only, services "
        "only, both, or neither. Set both switches to No if tax is not "
        "something you charge.",
        LAST_COL, height=30)
    row += 1
    row = s.field(row, "Sales tax rate", cfg.tax_rate, style="percent",
                  name="TaxRate", validate=number_check(0, 1),
                  help_text="8.25% is 0.0825 or 8.25% - both work.")
    row = s.field(row, "Tax on services?", cfg.tax_on_services,
                  name="TaxOnServices", validate=list_from("YesNoList"))
    row = s.field(row, "Tax on retail?", cfg.tax_on_retail,
                  name="TaxOnRetail", validate=list_from("YesNoList"))
    row = s.field(row, "Tax registered?", cfg.tax_registered,
                  validate=list_from("YesNoList"),
                  help_text="For your own reference - it changes "
                            "nothing in the formulas.")
    row = s.field(row, "Card fee, percent", cfg.card_fee_percent,
                  style="percent", name="CardFeePercent",
                  validate=number_check(0, 1),
                  help_text="What the card machine takes on each "
                            "payment.")
    row = s.field(row, "Card fee, fixed", cfg.card_fee_fixed,
                  style="money", name="CardFeeFixed",
                  validate=number_check(0, 100))
    row = s.field(row, "Deposit, percent of price", cfg.deposit_percent,
                  style="percent", name="DepositPercent",
                  validate=number_check(0, 1))
    row = s.field(row, "Ask for a deposit above", cfg.deposit_threshold,
                  style="money", name="DepositThreshold",
                  validate=number_check(0, 100000),
                  help_text="Bookings worth more than this suggest a "
                            "deposit on the appointment tab.")
    row = s.field(row, "Cancellation notice, hours",
                  cfg.cancellation_hours, style="int",
                  name="CancellationHours",
                  validate=number_check(0, 336, decimal=False))
    row = s.field(row, "No-show fee", cfg.no_show_fee, style="money",
                  name="NoShowFee", validate=number_check(0, 10000))
    row = s.field(row, "Price rounding", cfg.rounding_rule,
                  name="RoundingRule",
                  validate=list_from("RoundingRuleList"),
                  help_text="How calculated prices are rounded. Every "
                            "rule rounds up, never down.")
    row = s.field(row, "Target chair rate an hour", cfg.target_chair_rate,
                  style="money", name="TargetChairRate",
                  validate=number_check(0, 10000),
                  help_text="What an hour of chair time has to earn "
                            "before commission. The suggested prices "
                            "on the service menu are built from it.")
    row = s.field(row, "Review prices after, days", cfg.price_review_days,
                  style="int", name="PriceReviewDays",
                  validate=number_check(0, 3650, decimal=False))
    row += 1

    # ----------------------------------------------------------- staff pay
    row = s.section(row, "How the team is paid", LAST_COL)
    row = s.note(
        row,
        "These are defaults, not rules. Each person carries their own "
        "method and their own rate on the Staff tab; these fill the gaps "
        "when a rate is left blank.",
        LAST_COL, height=28)
    row += 1
    row = s.field(row, "Default service commission",
                  cfg.service_commission, style="percent",
                  name="ServiceCommission", validate=number_check(0, 1))
    row = s.field(row, "Default retail commission", cfg.retail_commission,
                  style="percent", name="RetailCommission",
                  validate=number_check(0, 1))
    row = s.field(row, "Apprentice commission", cfg.apprentice_commission,
                  style="percent", name="ApprenticeCommission",
                  validate=number_check(0, 1))
    row = s.field(row, "Default hourly rate", cfg.hourly_rate,
                  style="money", name="DefaultHourly",
                  validate=number_check(0, 1000))
    row = s.field(row, "Employer costs on wages", cfg.payroll_burden,
                  style="percent", name="PayrollBurden",
                  validate=number_check(0, 1),
                  help_text="Employer tax, insurance and holiday "
                            "accrual as a percentage of gross pay.")
    row = s.field(row, "Commission is paid on", cfg.commission_on,
                  value_end=5, name="CommissionOn",
                  validate={"validate": "list",
                            "source": ["Net of discounts",
                                       "Gross before discounts"]},
                  help_col=7,
                  help_text="One switch, applied to everybody.")
    row = s.field(row, "Chair rent a week", cfg.chair_rent_week,
                  style="money", name="ChairRentWeek",
                  validate=number_check(0, 10000))
    row += 1

    row = s.subsection(row, "Tiered commission ladder", LAST_COL)
    row = s.note(
        row,
        "Used by the 'tiered on monthly service revenue' method. The "
        "rate that applies is the highest band the stylist's month "
        "reaches. Edit the bands to match your own ladder - keep the "
        "'from' column in ascending order.",
        LAST_COL, height=30)
    row += 1
    s.table_head(row, 1, ["Monthly service revenue from", "Rate"])
    first_tier = row + 1
    tier_rows = len(COMMISSION_TIERS) + 3
    for index in range(tier_rows):
        if index < len(COMMISSION_TIERS):
            floor, rate = COMMISSION_TIERS[index]
            s.write(first_tier + index, 1, floor, fmts.input("money"))
            s.write(first_tier + index, 2, rate, fmts.input("percent"))
        else:
            s.write(first_tier + index, 1, None, fmts.input("money"))
            s.write(first_tier + index, 2, None, fmts.input("percent"))
    book.define_name(
        "TierFrom",
        f"={sheet_ref}!$B${first_tier + 1}:$B${first_tier + tier_rows}")
    book.define_name(
        "TierRate",
        f"={sheet_ref}!$C${first_tier + 1}:$C${first_tier + tier_rows}")
    row = first_tier + tier_rows + 1

    # ------------------------------------------------------------- targets
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
    row = s.field(row, "Average ticket target", cfg.target_average_ticket,
                  style="money", name="TargetTicket",
                  validate=number_check(0, 100000))
    row = s.field(row, "Retail share target", cfg.target_retail_percent,
                  style="percent", name="TargetRetailShare",
                  validate=number_check(0, 1),
                  help_text="Retail as a share of everything you sell. "
                            "Ten to fifteen percent is a good salon.")
    row = s.field(row, "Rebooking target", cfg.target_rebook_rate,
                  style="percent", name="TargetRebook",
                  validate=number_check(0, 1))
    row = s.field(row, "Retention target", cfg.target_retention,
                  style="percent", name="TargetRetention",
                  validate=number_check(0, 1))
    row = s.field(row, "Utilisation target", cfg.target_utilisation,
                  style="percent", name="TargetUtilisation",
                  validate=number_check(0, 1))
    row = s.field(row, "VIP spend, 12 months", cfg.vip_spend_year,
                  style="money", name="VipSpend",
                  validate=number_check(0, 1000000),
                  help_text="Spend in a year that marks a client as "
                            "VIP on the CRM.")
    row += 1

    # --------------------------------------------------------------- stock
    row = s.section(row, "Stock and ordering", LAST_COL)
    row = s.field(row, "Retail markup on cost", cfg.retail_markup,
                  style="number1", name="RetailMarkup",
                  validate=number_check(1, 20),
                  help_text="2.0 means a $10 product retails at $20 "
                            "before rounding.")
    row = s.field(row, "Reorder cover, weeks",
                  round(cfg.default_reorder_days / 7, 1),
                  style="number1", name="ReorderWeeks",
                  validate=number_check(0.5, 52),
                  help_text="How many weeks of stock an order should "
                            "bring in.")
    row = s.field(row, "Warn when stock falls below, weeks",
                  cfg.stock_warning_weeks, style="number1",
                  name="StockWarningWeeks", validate=number_check(0, 52))
    row = s.field(row, "Backbar cost as a share of a service",
                  cfg.product_cost_percent, style="percent",
                  name="ProductCostPercent", validate=number_check(0, 1),
                  help_text="Only used when a service has no product "
                            "cost of its own.")
    row += 1

    # -------------------------------------------------------- client care
    row = s.section(row, "Client care, loyalty and referrals", LAST_COL)
    row = s.field(row, "Usual gap between visits, weeks",
                  cfg.rebook_weeks, style="int", name="RebookWeeks",
                  validate=number_check(1, 104, decimal=False))
    row = s.field(row, "Lapsing after, days", cfg.lapsing_days,
                  style="int", name="LapsingDays",
                  validate=number_check(1, 1000, decimal=False))
    row = s.field(row, "Inactive after, days", cfg.inactive_days,
                  style="int", name="InactiveDays",
                  validate=number_check(1, 2000, decimal=False))
    row = s.field(row, "Patch test valid for, days", cfg.patch_test_days,
                  style="int", name="PatchTestDays",
                  validate=number_check(1, 2000, decimal=False))
    row = s.field(row, "Loyalty points per unit spent",
                  cfg.loyalty_points_per, style="number1",
                  name="PointsPerUnit", validate=number_check(0, 100))
    row = s.field(row, "Reward at, points", cfg.loyalty_reward_at,
                  style="int", name="RewardAt",
                  validate=number_check(1, 100000, decimal=False))
    row = s.field(row, "Reward is worth", cfg.loyalty_reward_value,
                  style="money", name="RewardValue",
                  validate=number_check(0, 10000),
                  help_text="One point per dollar with a $20 reward at "
                            "200 points is a 10% scheme. Worth knowing "
                            "before you promise it.")
    row = s.field(row, "Referral reward", cfg.referral_reward,
                  style="money", name="ReferralReward",
                  validate=number_check(0, 10000))
    row += 1

    # ------------------------------------------------------------ prefixes
    row = s.section(row, "ID prefixes", LAST_COL)
    row = s.note(
        row,
        "Every record numbers itself from its row. Change a prefix here "
        "and the IDs on that tab change with it - but do it before you "
        "have typed a hundred appointments, not after.",
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
            name,
            f"={sheet_ref}!${col_letter(column + 2)}${line + 1}")
    row = top + half + 1

    # -------------------------------------------------- calculated totals
    row = s.section(row, "Workbook totals", LAST_COL)
    row = s.note(
        row,
        "Calculated once here so that several tabs can share them. "
        "Nothing to change.",
        LAST_COL, height=24)
    row += 1
    row = s.field(row, "Service revenue, all time", None, style="money",
                  name="ServiceRevenueAll", mode="calc",
                  formula='=ROUND(SUMIFS([[Sales.net]],[[Sales.kind]],'
                          '"Service"),2)')
    row = s.field(row, "Retail revenue, all time", None, style="money",
                  name="RetailRevenueAll", mode="calc",
                  formula='=ROUND(SUMIFS([[Sales.net]],[[Sales.kind]],'
                          '"Retail"),2)')
    row = s.field(row, "Net sales, all time", None, style="money",
                  name="NetSalesAll", mode="calc",
                  formula="=ROUND(SUM([[Sales.net]]),2)")
    row = s.field(row, "Clients on file", None, style="int",
                  name="ClientCount", mode="calc",
                  formula='=COUNTIF([[Clients.name]],"<>")')
    row += 1

    # ---------------------------------------------------------- the lists
    row = s.section(row, "Drop-down lists", LAST_COL)
    row = s.note(
        row,
        "Add new entries directly under the last one in a column - every "
        "list has room to grow and blank rows are ignored. Do not delete "
        "a whole column: a drop-down pointing at a missing list stops "
        "working.",
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

    # ------------------------------------------------------- month keys --
    row = s.section(row, "Month keys", LAST_COL)
    row = s.note(
        row,
        "The months the commission, payroll and budget tabs pick from. "
        "They are plain text in yyyy-mm form so that every monthly total "
        "in the workbook groups on exactly the same thing.",
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

    row = s.note(row, DISCLAIMER, LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
