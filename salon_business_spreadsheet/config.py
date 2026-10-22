"""
Configuration for the Salon Business Management Spreadsheet.

Everything a salon owner may want to re-brand lives here, so the
workbook can be regenerated for a different salon, currency, country or
commission model without touching any of the sheet code.

Author of the generated workbooks: Novality Store.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import date


# ---------------------------------------------------------------------------
# Product / branding constants
# ---------------------------------------------------------------------------

AUTHOR = "Novality Store"
PRODUCT_NAME = "Salon Business Management Spreadsheet"
PRODUCT_SUBTITLE = (
    "Appointments, Clients, Services, Sales, Staff Commission, "
    "Inventory, Expenses and Profit in one workbook"
)
PRODUCT_VERSION = "1.0.0"

#: Title suggested for the shop listing.
PRODUCT_LISTING_TITLE = (
    "Salon Business Management Spreadsheet, Appointment and Client "
    "Tracker, Staff Commission Calculator, Salon Income Expense and "
    "Profit Template"
)

#: Password that protects every sheet so formula cells cannot be edited
#: by accident.  Input cells stay unlocked, formula cells are locked.
FORMULA_PASSWORD = "premium"

#: The workbook is a business organiser, not an accountant.  This line
#: appears on Start Here, on the dashboard and in the help tab.
DISCLAIMER = (
    "This workbook is a business organising tool. It is not accounting, "
    "payroll, legal or tax advice, and it does not file anything for "
    "you. Tax, commission and payroll settings are yours to set and "
    "check against your own country's rules, and nothing here replaces "
    "an accountant or a booking system that sends confirmations."
)

#: Service categories offered in the drop downs - broad enough for hair,
#: nails, lashes, skin and barbering.
SERVICE_CATEGORIES = [
    "Hair - cutting", "Hair - colour", "Hair - treatment",
    "Hair - styling", "Barbering", "Nails - manicure",
    "Nails - pedicure", "Nails - extensions", "Lashes", "Brows",
    "Waxing", "Facials & skin", "Massage & body", "Make-up",
    "Tanning", "Package", "Other service",
]

#: Where a booking came from.
BOOKING_SOURCES = [
    "Repeat client", "Walk-in", "Instagram", "Facebook", "Google",
    "TikTok", "Referral", "Booking site", "Phone", "Website",
    "Gift certificate", "Promotion", "Other",
]

#: An appointment moves through these, in this order.
APPOINTMENT_STATUSES = [
    "Enquiry", "Booked", "Confirmed", "Arrived", "Completed",
    "Cancelled", "No show", "Rescheduled",
]

#: Statuses that mean the appointment reached a final answer - the
#: completion and cancellation rates divide by these only.
FINAL_STATUSES = ["Completed", "Cancelled", "No show"]

#: How a client pays.
PAYMENT_METHODS = [
    "Cash", "Card", "Bank transfer", "Mobile wallet", "Gift certificate",
    "Package / prepaid", "Online deposit", "Other",
]

#: What a payment is for - keeps deposits out of the revenue figures
#: until the appointment they belong to is completed.
PAYMENT_KINDS = [
    "Deposit", "Balance", "Full payment", "Retail only", "Refund",
    "Gift certificate sale", "Package sale", "Tip",
]

#: How staff are paid.  Every one of these is supported by the
#: commission tab - the owner picks per person.
COMMISSION_METHODS = [
    "None", "Percent of service", "Percent of service and retail",
    "Percent after product cost", "Tiered on monthly service revenue",
    "Flat amount per service", "Hourly only", "Rent a chair",
]

#: Where a team member is up to.
STAFF_STATUSES = ["Active", "On leave", "Trial", "Inactive", "Left"]

#: Employment types.
STAFF_TYPES = ["Employed", "Self-employed", "Chair renter", "Apprentice",
               "Owner", "Contractor"]

#: Expense categories, in the order a profit and loss statement wants
#: them.
EXPENSE_CATEGORIES = [
    "Rent", "Utilities", "Product & stock", "Backbar supplies",
    "Salaries & wages", "Commission", "Payroll taxes", "Insurance",
    "Equipment", "Repairs & maintenance", "Marketing & advertising",
    "Software & subscriptions", "Card processing fees", "Laundry",
    "Cleaning", "Training & education", "Licences & memberships",
    "Accounting & legal", "Telephone & internet", "Travel & parking",
    "Bank charges", "Loan repayment", "Other",
]

#: Product categories for the stock room.
PRODUCT_CATEGORIES = [
    "Retail - hair", "Retail - skin", "Retail - nails",
    "Retail - accessories", "Backbar - colour", "Backbar - developer",
    "Backbar - shampoo & conditioner", "Backbar - treatment",
    "Backbar - nails", "Backbar - lashes", "Backbar - waxing",
    "Disposables", "Tools & equipment", "Other stock",
]

#: Marketing channels a campaign can run on.
MARKETING_CHANNELS = [
    "Instagram", "Facebook", "TikTok", "Google Ads", "Google Business",
    "Email", "SMS", "Flyers & print", "Local press", "Referral scheme",
    "Loyalty scheme", "Influencer", "Event", "Window & signage", "Other",
]

#: Loyalty reward types.
LOYALTY_REWARDS = [
    "Free service", "Percent off", "Amount off", "Free add-on",
    "Free retail product", "Birthday treat", "Referral credit", "None",
]

#: Client status, worked out from the last visit but overridable.
CLIENT_STATUSES = ["New", "Active", "Regular", "VIP", "Lapsing",
                   "Inactive", "Archived"]

#: What a till line can be.  Keeping services and retail apart is what
#: makes the retail share, the commission split and the stock figures
#: work.
SALE_KINDS = ["Service", "Retail", "Package", "Gift certificate",
              "Course", "Other"]

#: Why an appointment fell over - useful once you have fifty of them.
CANCEL_REASONS = ["Client cancelled", "Client rescheduled",
                  "Client unwell", "No show", "Salon cancelled",
                  "Stylist unavailable", "Weather", "Double booked",
                  "Price", "Other"]

#: How often a recurring cost comes round.
FREQUENCIES = ["Weekly", "Fortnightly", "Four weekly", "Monthly",
               "Quarterly", "Twice a year", "Yearly", "One off"]

#: Pay run frequency on the payroll tab.
PAY_PERIODS = ["Weekly", "Fortnightly", "Monthly", "Four weekly",
               "One off"]

#: Purchase order progress.
PURCHASE_STATUSES = ["Draft", "Ordered", "Part received", "Received",
                     "Cancelled", "Returned"]

#: Campaign progress.
CAMPAIGN_STATUSES = ["Planned", "Running", "Finished", "Paused",
                     "Cancelled"]

#: How a client prefers to be contacted.
CONTACT_METHODS = ["Text", "Phone call", "Email", "WhatsApp",
                   "Instagram DM", "Do not contact"]

#: Stylist levels - who is allowed to take which service, and the
#: price band it sits in.
STYLIST_LEVELS = ["Any stylist", "Apprentice", "Stylist", "Senior stylist",
                  "Master stylist", "Owner", "Specialist"]

#: Service, campaign and product lifecycle.
SERVICE_STATUSES = ["Active", "New", "Seasonal", "On hold", "Discontinued"]

#: How a price is rounded when a service price is calculated.
ROUNDING_RULES = ["None", "Nearest 0.99", "Nearest 0.95", "Nearest 0.50",
                  "Nearest 1", "Nearest 5", "Up to whole 1",
                  "Up to whole 5"]

#: Tiered commission bands: monthly service revenue from, rate.  Shown
#: on the Settings tab and editable there.
COMMISSION_TIERS = [
    (0.0, 0.35),
    (2000.0, 0.40),
    (3500.0, 0.45),
    (5000.0, 0.50),
]

#: ID prefixes.  Each one is an input field on the Settings tab with a
#: defined name, so a salon can renumber its own records without
#: touching a formula: (label, default, defined name, digits).
ID_PREFIXES = [
    ("Client IDs", "CLI-", "ClientPrefix", 4),
    ("Service IDs", "SRV-", "ServicePrefix", 3),
    ("Stylist IDs", "STF-", "StaffPrefix", 3),
    ("Appointment IDs", "APT-", "AppointmentPrefix", 4),
    ("Sale IDs", "TXN-", "SalePrefix", 4),
    ("Payment IDs", "PMT-", "PaymentPrefix", 4),
    ("Commission IDs", "COM-", "CommissionPrefix", 4),
    ("Pay run IDs", "RUN-", "PayrollPrefix", 4),
    ("Expense IDs", "EXP-", "ExpensePrefix", 4),
    ("Product IDs", "PRD-", "ProductPrefix", 3),
    ("Purchase order IDs", "PO-", "PurchasePrefix", 4),
    ("Campaign IDs", "CMP-", "CampaignPrefix", 3),
    ("Loyalty IDs", "LOY-", "LoyaltyPrefix", 4),
]

#: defined name -> (prefix, digits), for the demo data generator.
PREFIX_BY_NAME = {name: (prefix, digits)
                  for _label, prefix, name, digits in ID_PREFIXES}


def rid(kind: str, number: int) -> str:
    """Build the ID the workbook itself will generate on row `number`.

    ``rid("Client", 7)`` -> ``"CLI-0007"``.  The demo data has to agree
    with the ID formulas exactly, because the IDs are what tie the
    tabs together.
    """
    prefix, digits = PREFIX_BY_NAME[f"{kind}Prefix"]
    return f"{prefix}{number:0{digits}d}"


#: The twelve month names used by every monthly roll-up.
MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]


# ---------------------------------------------------------------------------
# Colours - a salon palette: plum, rose, soft gold, warm grey
# ---------------------------------------------------------------------------

class Palette:
    ink = "#2E2430"            # near black plum - titles
    deep = "#4C2A42"           # deep plum - staff & money tabs
    primary = "#6D3B5E"        # plum headers
    primary_mid = "#9A5C86"    # section bands
    primary_soft = "#F3E6EF"   # light band / alternating rows
    accent = "#C08A5E"         # warm gold accent
    accent_soft = "#F7ECE1"
    input_bg = "#FFFFFF"       # editable cell background
    input_border = "#E2D3DD"
    calc_bg = "#FAF5F8"        # locked formula cell background
    calc_border = "#EADFE7"
    white = "#FFFFFF"
    grey_text = "#6B5C66"
    line = "#DCCBD6"
    good_bg = "#DCEBD8"
    good_text = "#2F5A2B"
    warn_bg = "#FDEDD2"
    warn_text = "#8A5A00"
    bad_bg = "#F8DAD9"
    bad_text = "#9C2A24"
    info_bg = "#E3E4F2"
    info_text = "#3A4578"
    note_bg = "#FBF7FA"


# ---------------------------------------------------------------------------
# Row capacity of every register
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Start Here": 0,
    "Dashboard": 0,
    "Appointments": 1200,
    "Clients": 500,
    "Services": 200,
    "Sales": 1200,
    "Payments": 1200,
    "Staff": 40,
    "Commissions": 240,
    "Payroll": 360,
    "Expenses": 600,
    "Profit & Loss": 0,
    "Inventory": 300,
    "Purchases": 300,
    "Marketing": 100,
    "Loyalty": 500,
    "Goals & Budget": 0,
    "Analytics": 0,
    "Reports": 0,
    "Settings": 0,
}


@dataclass
class SalonDefaults:
    """Sample values written into the business profile on Settings."""

    name: str = "Gloss & Grace Salon"
    owner: str = "Nadia Ellison"
    brand: str = "Gloss & Grace"
    email: str = "hello@glossandgrace.co"
    phone: str = "(512) 555 0184"
    address: str = "118 Juniper Street, Austin, TX 78702"
    website: str = "glossandgrace.co"
    instagram: str = "@glossandgracesalon"
    booking_link: str = "glossandgrace.co/book"
    tagline: str = "Colour, cuts and care since 2019"
    tax_id: str = "TX-SAL-2019-8820"
    opening_hours: str = "Tue-Fri 9:00-19:00, Sat 9:00-17:00"


# ---------------------------------------------------------------------------
# The configuration object every sheet module reads
# ---------------------------------------------------------------------------

@dataclass
class Config:
    """Everything the generator needs to know to write one workbook."""

    author: str = AUTHOR
    product_name: str = PRODUCT_NAME
    product_subtitle: str = PRODUCT_SUBTITLE
    listing_title: str = PRODUCT_LISTING_TITLE
    version: str = PRODUCT_VERSION
    password: str = FORMULA_PASSWORD
    protect_sheets: bool = True
    include_sample_data: bool = True

    # Formatting ------------------------------------------------------------
    currency_symbol: str = "$"
    currency_code: str = "USD"
    date_format: str = "dd-mmm-yyyy"
    time_format: str = "h:mm AM/PM"
    decimals: int = 2
    week_start: str = "Monday"

    # Salon profile pre-filled on the Settings tab --------------------------
    salon: SalonDefaults = field(default_factory=SalonDefaults)
    plan_year: int = 0                   # 0 = the year of `today`

    # Opening hours used by the planner and the utilisation figures ---------
    open_hour: float = 9.0
    close_hour: float = 19.0
    open_days: int = 5                   # trading days a week
    chairs: int = 4                      # stations available

    # Money -----------------------------------------------------------------
    tax_rate: float = 0.0825             # sales tax on retail, editable
    tax_on_services: str = "No"          # many places do not tax services
    tax_on_retail: str = "Yes"
    tax_registered: str = "Yes"
    card_fee_percent: float = 0.0275     # what the card machine takes
    card_fee_fixed: float = 0.25
    deposit_percent: float = 0.25        # asked for on big bookings
    deposit_threshold: float = 80.00     # bookings above this want one
    cancellation_hours: int = 24
    no_show_fee: float = 25.00
    rounding_rule: str = "Nearest 0.50"

    # Staff pay -------------------------------------------------------------
    default_commission_method: str = "Percent of service"
    service_commission: float = 0.40
    retail_commission: float = 0.10
    apprentice_commission: float = 0.15
    hourly_rate: float = 16.00           # assistant / front desk hourly
    payroll_burden: float = 0.12         # employer tax, insurance, holiday
    tip_handling: str = "Paid to staff in full"
    commission_on: str = "Net of discounts"
    chair_rent_week: float = 180.00

    # Targets the dashboard measures against --------------------------------
    monthly_revenue_target: float = 28000.0
    annual_revenue_target: float = 336000.0
    monthly_profit_target: float = 6000.0
    target_average_ticket: float = 78.00
    target_retail_percent: float = 0.12   # retail as a share of revenue
    target_rebook_rate: float = 0.60
    target_retention: float = 0.70
    target_utilisation: float = 0.75
    # the example salon is smaller than the defaults assume, so the demo
    # file carries targets it can actually be measured against
    sample_monthly_revenue_target: float = 18000.0
    sample_annual_revenue_target: float = 216000.0
    sample_monthly_profit_target: float = 2400.0

    # Stock -----------------------------------------------------------------
    retail_markup: float = 2.00          # retail price = cost x this
    default_reorder_days: int = 14
    stock_warning_weeks: float = 2.0
    product_cost_percent: float = 0.08   # backbar cost as a share of a
    #                                      service, when nothing is logged
    target_chair_rate: float = 70.00     # what an hour of chair time has
    #                                      to earn before commission
    price_review_days: int = 365         # nag when a price is older
    patch_test_days: int = 180           # colour patch test validity

    # Client care -----------------------------------------------------------
    vip_spend_year: float = 1200.00      # spend in 12m that makes a VIP
    rebook_weeks: int = 6                # the usual gap between visits
    lapsing_days: int = 90               # after this a client is lapsing
    inactive_days: int = 180
    loyalty_points_per: float = 1.0      # points per unit of spend
    loyalty_reward_at: int = 200         # points needed for a reward
    loyalty_reward_value: float = 20.00
    referral_reward: float = 10.00
    review_period_days: int = 30

    # Behaviour -------------------------------------------------------------
    top_rows: int = 5                    # size of the "top / best" tables
    months_shown: int = 12
    rows: dict = field(default_factory=lambda: dict(DEFAULT_ROWS))
    today: date = field(default_factory=date.today)

    # ---------------------------------------------------------------- helpers
    def row_count(self, sheet: str) -> int:
        return int(self.rows.get(sheet, 100))

    @property
    def year(self) -> int:
        return self.plan_year or self.today.year

    @property
    def revenue_target_month(self) -> float:
        """Monthly revenue target, scaled to whoever is reading it."""
        return (self.sample_monthly_revenue_target
                if self.include_sample_data else self.monthly_revenue_target)

    @property
    def revenue_target_year(self) -> float:
        return (self.sample_annual_revenue_target
                if self.include_sample_data else self.annual_revenue_target)

    @property
    def profit_target_month(self) -> float:
        return (self.sample_monthly_profit_target
                if self.include_sample_data else self.monthly_profit_target)

    @property
    def labor_cost_per_hour(self) -> float:
        """An hour of paid time once employer on-costs are added."""
        return round(self.hourly_rate * (1 + self.payroll_burden), 2)

    @property
    def chair_hours_week(self) -> float:
        """Sellable hours in a week - the denominator of utilisation."""
        return round((self.close_hour - self.open_hour) * self.open_days
                     * self.chairs, 1)

    def scale_rows(self, main_rows: int) -> None:
        """Scale every register in place so Appointments holds N rows."""
        base = DEFAULT_ROWS["Appointments"]
        factor = max(0.02, float(main_rows) / float(base))
        self.rows = {k: (max(5, int(round(v * factor))) if v else v)
                     for k, v in self.rows.items()}

    def scaled(self, factor: float) -> "Config":
        """Return a copy with every register scaled."""
        rows = {k: max(5, int(round(v * factor))) if v else v
                for k, v in self.rows.items()}
        return replace(self, rows=rows)

    # Number format strings derived from the chosen currency ----------------
    @property
    def money_format(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0.00;[Red]-{s}#,##0.00'

    @property
    def money_format_0(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0;[Red]-{s}#,##0'

    @property
    def money_format_4(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0.0000;[Red]-{s}#,##0.0000'
