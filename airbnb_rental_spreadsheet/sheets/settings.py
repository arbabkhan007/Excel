"""`Settings & Lists` - every number and every list the workbook runs on.

Nothing here is hard coded to one country, one platform or one pricing
model.  The commission each platform keeps, the card processing fee on
top, the extra guest charge, the lodging tax and who remits it, the
cost of a turnover, the targets, the alert windows and all forty-odd
drop-downs are set here, once, and every other tab follows.

The platform fee table matters most.  Airbnb runs a split fee in some
markets and a host-only fee in others; Vrbo has two completely
different models; Booking.com commission is negotiated per property.
A spreadsheet that hard codes "3%" is wrong for most of the people who
buy it, so this one ships with sensible defaults in cells you can type
over.
"""

from __future__ import annotations

import datetime as _dt
from typing import List, Tuple

from ..config import (AVAILABILITY, BOOKING_STATUSES, CLEANING_STATUSES,
                      CONTRACT_STATUSES, COUNTRIES, DISCLAIMER,
                      EXPENSE_CATEGORIES, FREQUENCY_PER_YEAR,
                      GUEST_TYPES, ID_PREFIX_FIELDS,
                      MAINTENANCE_CATEGORIES, MAINTENANCE_STATUSES,
                      MONTH_NAMES, OWNERSHIP_TYPES, PAY_BASIS,
                      PAYMENT_METHODS, PAYMENT_STATUSES, PAYOUT_STATUSES,
                      PRIORITY_LEVELS, PRIVACY_NOTE, PROPERTY_STATUSES,
                      PROPERTY_TYPES, RECURRING_FREQUENCIES,
                      REVIEW_STATUSES, SEASONS, STAFF_ROLES, STAY_BANDS,
                      STOCK_STATUSES, SUPPLY_CATEGORIES, TAX_CATEGORIES,
                      TRIP_PURPOSES, YES_NO, YES_NO_NA, Config, Palette)
from ..layout import Registry, col_letter, dynamic_name, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Settings & Lists"
LAST_COL = 11

#: name -> values.  Everything the workbook offers you in a drop-down.
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", list(YES_NO)),
    ("YesNoNaList", list(YES_NO_NA)),
    ("BookingStatusList", list(BOOKING_STATUSES)),
    ("PaymentStatusList", list(PAYMENT_STATUSES)),
    ("PropertyTypeList", list(PROPERTY_TYPES)),
    ("PropertyStatusList", list(PROPERTY_STATUSES)),
    ("OwnershipList", list(OWNERSHIP_TYPES)),
    ("CleaningStatusList", list(CLEANING_STATUSES)),
    ("MaintenanceCategoryList", list(MAINTENANCE_CATEGORIES)),
    ("MaintenanceStatusList", list(MAINTENANCE_STATUSES)),
    ("PriorityList", list(PRIORITY_LEVELS)),
    ("ExpenseCategoryList", list(EXPENSE_CATEGORIES)),
    ("TaxCategoryList", list(TAX_CATEGORIES)),
    ("PaymentMethodList", list(PAYMENT_METHODS)),
    ("FrequencyList", list(RECURRING_FREQUENCIES)),
    ("SupplyCategoryList", list(SUPPLY_CATEGORIES)),
    ("StockStatusList", list(STOCK_STATUSES)),
    ("StaffRoleList", list(STAFF_ROLES)),
    ("PayBasisList", list(PAY_BASIS)),
    ("ContractStatusList", list(CONTRACT_STATUSES)),
    ("AvailabilityList", list(AVAILABILITY)),
    ("ReviewStatusList", list(REVIEW_STATUSES)),
    ("GuestTypeList", list(GUEST_TYPES)),
    ("TripPurposeList", list(TRIP_PURPOSES)),
    ("PayoutStatusList", list(PAYOUT_STATUSES)),
    ("CountryList", list(COUNTRIES)),
    ("MonthNameList", list(MONTH_NAMES)),
    ("SeasonList", [s[0] for s in SEASONS]),
    ("StayBandList", [b[0] for b in STAY_BANDS]),
    ("RatingList", ["5", "4.5", "4", "3.5", "3", "2", "1"]),
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
             "Your properties, your platforms, your fees, your lists. "
             "Set these once and every tab in the workbook follows.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 24)
    s.set_col(3, 3, 20)
    s.set_col(4, LAST_COL, 17)

    row = 4
    row = s.note(
        row,
        "White cells are yours to change; everything else is locked and "
        "calculated. Nothing here assumes a country, a platform or a "
        "pricing model - set your own commission rates, your own tax "
        "position, your own currency and your own targets, and the "
        "rest of the workbook follows.",
        LAST_COL, height=34)
    row += 1

    # ------------------------------------------------------ the business --
    row = s.section(row, "Your business", LAST_COL)
    row = s.field(row, "Business name", cfg.host.name, name="HostName",
                  value_end=5, help_col=7,
                  help_text="Printed at the top of every report and "
                            "every printable checklist.")
    row = s.field(row, "Owner / host", cfg.host.owner, name="OwnerName",
                  value_end=5, help_col=7)
    row = s.field(row, "Address", cfg.host.address, value_end=5,
                  help_col=7)
    row = s.field(row, "Phone", cfg.host.phone, value_end=5, help_col=7)
    row = s.field(row, "Email", cfg.host.email, value_end=5, help_col=7)
    row = s.field(row, "Website", cfg.host.website, value_end=5,
                  help_col=7)
    row = s.field(row, "Let registration / permit", cfg.host.licence,
                  value_end=5, help_col=7,
                  help_text="Your short-term let registration number, "
                            "if your city requires one. Per-property "
                            "licences live on the Property Database.")
    row = s.field(row, "Tax / business number", cfg.host.tax_id,
                  value_end=5, help_col=7)
    row = s.field(row, "Strapline", cfg.host.tagline, value_end=5,
                  help_col=7)
    row = s.field(row, "Currency symbol", cfg.currency_symbol,
                  name="CurrencySign",
                  help_text="Shown on screen. The cell formats follow "
                            "the symbol you build the file with.")
    row = s.field(row, "Currency code", cfg.currency_code)
    row = s.field(row, "Tracking started",
                  cfg.today - _dt.timedelta(days=120), style="date",
                  name="TrackingStart",
                  help_text="The day you began logging in this file.")
    row = s.field(row, "Today", None, style="date", name="TodayDate",
                  mode="calc", formula="=TODAY()",
                  help_text="Live. Everything that says 'overdue', "
                            "'arriving' or 'this month' reads this "
                            "cell.")
    row = s.field(row, "This month", None, style="text",
                  name="ThisMonthKey", mode="calc",
                  formula='=TEXT(TODAY(),"yyyy-mm")',
                  help_text="The month key every monthly total groups "
                            "on.")
    row = s.field(row, "Properties on the books", None, style="int",
                  name="PropertyCount", mode="calc",
                  formula='=COUNTIF([[Property Database.name]],"?*")')
    row = s.field(row, "Properties active", None, style="int",
                  name="ActiveProperties", mode="calc",
                  formula='=COUNTIFS([[Property Database.status]],'
                          '"Active")')
    row += 1

    # ------------------------------------------------- the platform fees --
    row = s.section(row, "Platform fees - the most important table here",
                    LAST_COL)
    row = s.note(
        row,
        "Every booking prices its commission and card fee from this "
        "table. The defaults below are the common cases, not the law: "
        "Airbnb runs a split fee in most markets and a host-only fee "
        "in others, Vrbo has two different models, and Booking.com "
        "commission is negotiated per property. Put your own numbers "
        "in - the Platform Revenue & Fees tab will tell you when they "
        "are wrong, because it compares what you assumed with what the "
        "payout actually showed.",
        LAST_COL, height=50)
    row += 1
    plat_head = row
    s.table_head(plat_head, 1, [
        "Platform", "Commission", "Card fee", "Is a platform?",
        "Notes"],
        widths=[22, 14, 12, 15, 46])
    first_plat = plat_head + 1
    table = cfg.platform_table
    plat_rows = len(table) + 2
    for index in range(plat_rows):
        line = first_plat + index
        if index < len(table):
            label, commission, payment, counts, note = table[index]
            s.write(line, 1, label, fmts.input("text"))
            s.write(line, 2, commission, fmts.input("percent"))
            s.write(line, 3, payment, fmts.input("percent"))
            s.write(line, 4, counts, fmts.input("flag"))
            s.write(line, 5, note, fmts.chrome("note"))
        else:
            s.write(line, 1, None, fmts.input("text"))
            s.write(line, 2, None, fmts.input("percent"))
            s.write(line, 3, None, fmts.input("percent"))
            s.write(line, 4, None, fmts.input("flag"))
            s.write(line, 5, None, fmts.chrome("note"))
    last_plat = first_plat + plat_rows
    for key, column in (("PlatformList", 1), ("PlatformCommission", 2),
                        ("PlatformPayment", 3), ("PlatformIsOTA", 4)):
        book.define_name(
            key,
            f"={sheet_ref}!${col_letter(column)}${first_plat + 1}:"
            f"${col_letter(column)}${last_plat}")
    row = last_plat + 1

    # --------------------------------------------------- pricing defaults --
    row = s.section(row, "Pricing defaults", LAST_COL)
    row = s.note(
        row,
        "Each property has its own nightly rate and cleaning fee on the "
        "Property Database. These are the portfolio-wide rules that "
        "sit behind them, and the charges a booking works out for "
        "itself when you leave the cell blank.",
        LAST_COL, height=32)
    row += 1
    row = s.field(row, "Standard nightly rate", cfg.default_nightly_rate,
                  style="money", name="DefaultNightly",
                  validate=number_check(0, 100000),
                  help_text="Only a starting point for a new property. "
                            "The booking reads the property's rate, "
                            "not this one.")
    row = s.field(row, "Standard cleaning fee", cfg.default_cleaning_fee,
                  style="money", name="DefaultCleaningFee",
                  validate=number_check(0, 100000))
    row = s.field(row, "Standard cleaning cost",
                  cfg.default_cleaning_cost, style="money",
                  name="DefaultCleaningCost",
                  validate=number_check(0, 100000),
                  help_text="What you pay a cleaner, as opposed to "
                            "what you charge the guest.")
    row = s.field(row, "Guests included in the rate",
                  cfg.extra_guest_after, style="int",
                  name="ExtraGuestAfter",
                  validate=number_check(0, 40, decimal=False),
                  help_text="Guests above this number are charged the "
                            "extra guest fee, per night.")
    row = s.field(row, "Extra guest fee a night", cfg.extra_guest_fee,
                  style="money", name="ExtraGuestFee",
                  validate=number_check(0, 10000))
    row = s.field(row, "Pet fee a stay", cfg.pet_fee, style="money",
                  name="PetFee", validate=number_check(0, 10000))
    row = s.field(row, "Security deposit",
                  cfg.default_security_deposit, style="money",
                  name="DefaultDeposit", validate=number_check(0, 100000))
    row = s.field(row, "Minimum nights", cfg.default_min_nights,
                  style="int", name="DefaultMinNights",
                  validate=number_check(1, 365, decimal=False))
    row = s.field(row, "Management fee", cfg.management_fee_percent,
                  style="percent", name="ManagementFee",
                  validate=number_check(0, 1),
                  help_text="For co-hosts: the share of revenue you "
                            "keep, or that an agency keeps from you. "
                            "Zero if it does not apply.")
    row += 1

    # ------------------------------------------------------------- tax ----
    row = s.section(row, "Lodging tax", LAST_COL)
    row = s.note(
        row,
        "Occupancy, lodging, tourist or sales tax on the accommodation "
        "total. Most platforms collect and remit it for you, which is "
        "why the default says so and the rate is zero. If you collect "
        "it yourself, put your rate in and switch the second cell to "
        "No - every booking will then show the tax it owes and keep it "
        "out of your net revenue.",
        LAST_COL, height=42)
    row += 1
    row = s.field(row, "Tax name", cfg.tax_label, name="TaxLabel")
    row = s.field(row, "Tax rate", cfg.tax_rate, style="percent",
                  name="TaxRate", validate=number_check(0, 1))
    row = s.field(row, "Platform collects and remits it",
                  cfg.tax_collected_by_platform, style="flag",
                  name="TaxCollectedByPlatform",
                  validate=list_from("YesNoList"))
    row += 1

    # ---------------------------------------------------- turnover costs --
    row = s.section(row, "What a turnover really costs", LAST_COL)
    row = s.note(
        row,
        "A changeover is never just the cleaner's fee. These two "
        "numbers are added to every cleaning job unless you type "
        "something different on the job itself, and they are what "
        "makes the true cost per booking honest.",
        LAST_COL, height=32)
    row += 1
    row = s.field(row, "Laundry a turnover", cfg.laundry_cost_turnover,
                  style="money", name="LaundryPerTurn",
                  validate=number_check(0, 10000))
    row = s.field(row, "Consumables a turnover",
                  cfg.supplies_cost_turnover, style="money",
                  name="SuppliesPerTurn", validate=number_check(0, 10000),
                  help_text="Loo roll, coffee, soap, bin bags. Compare "
                            "it with the figure the Supplies tab works "
                            "out from your actual stock usage.")
    row = s.field(row, "Welcome basket", cfg.welcome_cost_turnover,
                  style="money", name="WelcomePerTurn",
                  validate=number_check(0, 10000))
    row = s.field(row, "Default cleaner rate", cfg.cleaner_default_rate,
                  style="money", name="CleanerRate",
                  validate=number_check(0, 10000))
    row = s.field(row, "A tight turnaround is under",
                  cfg.turnover_warning_hours, style="number1",
                  name="TightTurnHours", validate=number_check(0, 48),
                  help_text="Hours between check-out and the next "
                            "check-in on a same-day turn. Below this "
                            "the cleaning job goes red.")
    row += 1

    # -------------------------------------------------- double counting --
    row = s.section(row, "Counting costs once", LAST_COL)
    row = s.note(
        row,
        "Cleaning has its own tab and so does maintenance, because "
        "both need a workflow rather than a line in a ledger. If you "
        "would rather type them as ordinary expenses instead, switch "
        "the relevant cell to Yes and the expense rows will start "
        "counting towards profit - but then stop using the other tab "
        "for money, or you will count the same cost twice.",
        LAST_COL, height=46)
    row += 1
    row = s.field(row, "Cleaning expenses count in profit", "No",
                  style="flag", name="CleaningInExpenses",
                  validate=list_from("YesNoList"),
                  help_text="No = the Cleaning Schedule provides the "
                            "cleaning cost, and cleaning rows on the "
                            "expense tab are ignored.")
    row = s.field(row, "Maintenance expenses count in profit", "No",
                  style="flag", name="MaintenanceInExpenses",
                  validate=list_from("YesNoList"),
                  help_text="No = the Maintenance Log provides the "
                            "repair costs.")
    row += 1

    # --------------------------------------------------------- targets --
    row = s.section(row, "Targets", LAST_COL)
    row = s.field(row, "Occupancy target", cfg.occupancy_target,
                  style="percent", name="TargetOccupancy",
                  validate=number_check(0, 1),
                  help_text="Every occupancy figure in the workbook is "
                            "green at or above this and amber below.")
    row = s.field(row, "ADR target", cfg.adr_target, style="money",
                  name="TargetADR", validate=number_check(0, 100000))
    row = s.field(row, "RevPAR target", cfg.revpar_target, style="money",
                  name="TargetRevPAR", validate=number_check(0, 100000))
    row = s.field(row, "Margin target", cfg.margin_target,
                  style="percent", name="TargetMargin",
                  validate=number_check(0, 1))
    row = s.field(row, "Rating target", cfg.rating_target,
                  style="number", name="TargetRating",
                  validate=number_check(0, 5))
    row = s.field(row, "Revenue target a month",
                  cfg.revenue_target_month, style="money0",
                  name="TargetRevenueMonth",
                  validate=number_check(0, 100000000))
    row = s.field(row, "Revenue target a year", cfg.revenue_target_year,
                  style="money0", name="TargetRevenueYear",
                  validate=number_check(0, 100000000))
    row = s.field(row, "Profit target a month", cfg.profit_target_month,
                  style="money0", name="TargetProfitMonth",
                  validate=number_check(0, 100000000))
    row += 1

    # ---------------------------------------------------------- alerts --
    row = s.section(row, "When to start nagging", LAST_COL)
    row = s.field(row, "'Arriving soon' window (days)",
                  cfg.upcoming_window_days, style="int",
                  name="UpcomingDays",
                  validate=number_check(0, 90, decimal=False),
                  help_text="A booking inside this window counts as "
                            "arriving soon, and its pre-arrival "
                            "messages start being chased.")
    row = s.field(row, "Chase a review after (days)",
                  cfg.review_chase_days, style="int",
                  name="ReviewChaseDays",
                  validate=number_check(0, 90, decimal=False))
    row = s.field(row, "Chase a payout after (days)",
                  cfg.payout_chase_days, style="int",
                  name="PayoutChaseDays",
                  validate=number_check(0, 180, decimal=False))
    row = s.field(row, "Warn about maintenance (days)",
                  cfg.maintenance_warning_days, style="int",
                  name="MaintenanceWarnDays",
                  validate=number_check(0, 180, decimal=False))
    row = s.field(row, "Reorder when cover drops below (weeks)",
                  cfg.restock_warning_weeks, style="number1",
                  name="RestockWeeks", validate=number_check(0, 52))
    row = s.field(row, "A long stay is (nights)", cfg.long_stay_nights,
                  style="int", name="LongStayNights",
                  validate=number_check(1, 365, decimal=False))
    row += 1

    # ---------------------------------------------- recurring frequency --
    row = s.section(row, "How often a recurring cost repeats", LAST_COL)
    row = s.note(
        row,
        "This is what turns any expense into a monthly equivalent, so "
        "a yearly insurance premium and a weekly laundry bill can sit "
        "in the same breakdown. 'One off' and 'Per booking' are zero "
        "because they are not a standing monthly cost.",
        LAST_COL, height=32)
    row += 1
    freq_head = row
    s.table_head(freq_head, 1, ["How often", "Times a year"],
                 widths=[22, 16])
    first_freq = freq_head + 1
    for index, label in enumerate(RECURRING_FREQUENCIES):
        line = first_freq + index
        s.write(line, 1, label, fmts.input("text"))
        s.write(line, 2, FREQUENCY_PER_YEAR.get(label, 0.0),
                fmts.input("number1"))
    last_freq = first_freq + len(RECURRING_FREQUENCIES) - 1
    book.define_name(
        "FrequencyPerYear",
        f"={sheet_ref}!$C${first_freq + 1}:$C${last_freq + 1}")
    row = last_freq + 2

    # ------------------------------------------------------ ID prefixes --
    row = s.section(row, "ID prefixes", LAST_COL)
    row = s.note(
        row,
        "The letters in front of every automatic ID. Change them here "
        "and the IDs on that tab change with them - but do it before "
        "you have typed three hundred bookings, not after.",
        LAST_COL, height=28)
    row += 1
    half = (len(ID_PREFIX_FIELDS) + 1) // 2
    top = row
    for index, (label, default, name, _digits) in enumerate(
            ID_PREFIX_FIELDS):
        column = 0 if index < half else 4
        line = top + (index if index < half else index - half)
        s.write(line, column + 1, label, fmts.chrome("label_bold"))
        s.write(line, column + 2, default, fmts.input("text"))
        book.define_name(
            name, f"={sheet_ref}!${col_letter(column + 2)}${line + 1}")
    row = top + half + 1

    # ----------------------------------------------------------- lists --
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

    # ------------------------------------------------------ month keys --
    row = s.section(row, "Month keys", LAST_COL)
    row = s.note(
        row,
        "The months the reports pick from. They are plain text in "
        "yyyy-mm form so that every monthly total in the workbook "
        "groups on exactly the same thing.",
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

    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=42)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=96)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
