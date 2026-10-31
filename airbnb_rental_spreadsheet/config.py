"""
Every default, list, rate and colour the workbook is built from.

Nothing in the generated spreadsheet is hard coded to one country, one
platform or one pricing model.  Platform commission rates, payment
processing fees, tax rates, cleaning costs and every drop down list are
settings the host edits on the *Settings & Lists* tab, because an Airbnb
host in Lisbon, a Vrbo host in Colorado and a Booking.com partner in
Manchester all pay different fees on different terms.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional, Tuple

PRODUCT_NAME = "Airbnb & Short-Term Rental Management Spreadsheet"
PRODUCT_SHORT = "Short-Term Rental Manager"
LISTING_TITLE = (
    "Airbnb Host Spreadsheet, Short-Term Rental Management, Booking "
    "Calendar, Cleaning, Guest Tracker, Expenses and Profit Dashboard"
)
PRODUCT_SUBTITLE = (
    "Bookings, cleaning turnovers, platform fees, expenses, "
    "occupancy and profit - across every property and every "
    "platform"
)
AUTHOR = "Novality Store"
VERSION = "1.0.0"
PASSWORD = "premium"


# ---------------------------------------------------------------------------
# Booking platforms
#
# `commission` is what the platform keeps out of the host's money, and
# `payment_fee` is the card / payout processing charge on top.  Both ship
# as editable cells: Airbnb alone runs a split fee in most markets and a
# host-only fee in others, Vrbo has two completely different models, and
# Booking.com commission is negotiated per property.
# ---------------------------------------------------------------------------

#: (platform, host commission %, payment processing %, counts as a platform
#:  booking, note shown on Settings)
PLATFORMS: List[Tuple[str, float, float, str, str]] = [
    ("Airbnb", 0.030, 0.000, "Yes",
     "host-only service fee is usually 14-16%, split fee about 3%"),
    ("Vrbo", 0.050, 0.030, "Yes",
     "pay-per-booking is around 5% plus 3% card processing"),
    ("Booking.com", 0.150, 0.000, "Yes",
     "commission is negotiated, commonly 15-18%"),
    ("Expedia", 0.150, 0.000, "Yes", "similar terms to Booking.com"),
    ("Direct booking", 0.000, 0.029, "No",
     "no commission, just whatever your card processor charges"),
    ("Returning guest", 0.000, 0.029, "No", "direct repeat booking"),
    ("Corporate / relocation", 0.000, 0.020, "No",
     "invoiced stays, often monthly"),
    ("Other", 0.100, 0.000, "Yes", "anything else - set your own rate"),
]

PLATFORM_NAMES = [p[0] for p in PLATFORMS]


# ---------------------------------------------------------------------------
# Booking states
#
# A block is just a reservation that earns nothing: it takes the dates off
# the calendar, it shows up in the availability grid, and it is counted as
# a blocked night rather than a booked night.  That keeps one register
# behind the calendar instead of two.
# ---------------------------------------------------------------------------

BOOKING_STATUSES = [
    "Inquiry",
    "Confirmed",
    "Checked in",
    "Completed",
    "Cancelled",
    "Owner block",
    "Maintenance block",
]

#: statuses that hold the dates on the calendar
HOLDING_STATUSES = ["Confirmed", "Checked in", "Completed",
                    "Owner block", "Maintenance block"]

#: statuses that earn money
EARNING_STATUSES = ["Confirmed", "Checked in", "Completed"]

#: statuses that take the night off the market without earning
BLOCK_STATUSES = ["Owner block", "Maintenance block"]

PAYMENT_STATUSES = [
    "Not due yet",
    "Awaiting payout",
    "Part paid",
    "Paid in full",
    "Refunded",
    "Written off",
]

PROPERTY_TYPES = [
    "Entire apartment", "Entire house", "Villa", "Cabin", "Studio",
    "Condo", "Townhouse", "Cottage", "Loft", "Guest suite",
    "Private room", "Serviced apartment", "Tiny home", "Chalet",
]

PROPERTY_STATUSES = [
    "Active", "Inactive", "Under maintenance", "Seasonal closure",
    "Being set up", "Sold / ended",
]

OWNERSHIP_TYPES = [
    "Owned outright", "Owned with mortgage", "Rental arbitrage",
    "Co-host (managed for owner)", "Family owned", "Head lease",
]

CLEANING_STATUSES = [
    "Not scheduled", "Scheduled", "In progress", "Complete",
    "Complete - issues", "Skipped", "Cancelled",
]

CLEANING_TASKS: List[Tuple[str, str]] = [
    ("task_bedroom", "Bedrooms"),
    ("task_bathroom", "Bathrooms"),
    ("task_kitchen", "Kitchen"),
    ("task_living", "Living area"),
    ("task_floors", "Floors"),
    ("task_laundry", "Laundry & linen"),
    ("task_rubbish", "Rubbish out"),
    ("task_restock", "Restock supplies"),
    ("task_inspection", "Final inspection"),
]

#: the printable room-by-room checklist on the Turnover Checklist tab
TURNOVER_CHECKLIST: List[Tuple[str, str, List[str]]] = [
    ("task_bedroom", "Bedrooms", [
        "Strip and change all bed sheets",
        "Replace pillowcases and check pillows",
        "Check mattress and protector for stains",
        "Make beds to listing-photo standard",
        "Dust surfaces, headboards and skirting",
        "Vacuum floors and under the beds",
        "Check bedside lights, lamps and sockets",
        "Empty and wipe out wardrobes and drawers",
    ]),
    ("task_bathroom", "Bathrooms", [
        "Clean and disinfect the toilet",
        "Clean shower, bath, screen and grout",
        "Clean sink, taps and mirrors",
        "Replace all towels and bath mats",
        "Refill shampoo, conditioner and soap",
        "Replace toilet roll (and leave a spare)",
        "Check drains, plumbing and extractor fan",
        "Empty the bin and replace the liner",
    ]),
    ("task_kitchen", "Kitchen", [
        "Clear, clean and disinfect all countertops",
        "Wash, dry and put away every dish",
        "Clean hob, oven, microwave and kettle",
        "Empty the fridge and freezer, wipe shelves",
        "Refill coffee, tea, sugar and condiments",
        "Restock dishwasher tablets and washing-up liquid",
        "Check cutlery, glasses and crockery counts",
        "Empty bins and replace liners",
    ]),
    ("task_living", "Living area", [
        "Dust all surfaces, shelves and skirting",
        "Vacuum or mop all floors and rugs",
        "Plump cushions and straighten furniture",
        "Check TV, remote and batteries",
        "Test the Wi-Fi and leave the code visible",
        "Tidy books, games and welcome folder",
    ]),
    ("task_inspection", "Final inspection", [
        "Check every window and door locks properly",
        "Set heating or cooling to the welcome setting",
        "Check the access code or key box works",
        "Walk through and photograph each room",
        "Log any damage or missing items",
        "Confirm the property is guest-ready",
    ]),
]

MAINTENANCE_CATEGORIES = [
    "Plumbing", "Electrical", "Heating / cooling", "Appliance",
    "Furniture", "Internet / TV", "Locks & access", "Pest control",
    "Garden / outdoor", "Pool / hot tub", "Damp & mould",
    "Paint & decorating", "Flooring", "Glass & windows",
    "Guest damage", "Deep clean", "Safety equipment", "Other",
]

MAINTENANCE_STATUSES = [
    "Reported", "Quoted", "Scheduled", "In progress", "Complete",
    "Waiting on parts", "Waiting on owner", "Cancelled",
]

PRIORITY_LEVELS = ["Urgent", "High", "Normal", "Low", "When convenient"]

EXPENSE_CATEGORIES = [
    "Cleaning", "Laundry", "Utilities - electric", "Utilities - gas",
    "Utilities - water", "Internet & TV", "Maintenance", "Repairs",
    "Furniture", "Appliances", "Guest amenities", "Welcome basket",
    "Insurance", "Property taxes", "Mortgage interest", "Rent",
    "Service charge", "Supplies", "Management fees", "Co-host fees",
    "Marketing & photography", "Software & subscriptions",
    "Licences & permits", "Accounting & legal", "Travel & mileage",
    "Bank & card fees", "Guest refunds", "Miscellaneous",
]

#: categories that are capital rather than running costs - excluded from the
#: operating profit figure but still totalled for the tax summary
CAPITAL_CATEGORIES = ["Furniture", "Appliances"]

#: categories the tax summary treats as a financing or ownership cost
OWNERSHIP_CATEGORIES = ["Mortgage interest", "Rent", "Property taxes",
                        "Service charge", "Insurance"]

TAX_CATEGORIES = [
    "Operating expense", "Capital / asset", "Financing cost",
    "Owner drawing", "Not deductible", "Review with accountant",
]

PAYMENT_METHODS = [
    "Bank transfer", "Debit card", "Credit card", "Direct debit",
    "Cash", "PayPal", "Platform deducted", "Owner paid", "Other",
]

RECURRING_FREQUENCIES = [
    "One off", "Weekly", "Fortnightly", "Monthly", "Quarterly",
    "Twice a year", "Yearly", "Per booking",
]

#: how many times a year each frequency happens - used to turn any cost
#: into a monthly equivalent
FREQUENCY_PER_YEAR: Dict[str, float] = {
    "One off": 0.0,
    "Weekly": 52.0,
    "Fortnightly": 26.0,
    "Monthly": 12.0,
    "Quarterly": 4.0,
    "Twice a year": 2.0,
    "Yearly": 1.0,
    "Per booking": 0.0,
}

SUPPLY_CATEGORIES = [
    "Towels", "Bed linen", "Pillows & duvets", "Toilet paper",
    "Shampoo & conditioner", "Soap & hand wash", "Cleaning products",
    "Coffee & tea", "Kitchen supplies", "Bin bags", "Laundry detergent",
    "Welcome treats", "Batteries & bulbs", "First aid", "Paper & stationery",
    "Pool & hot tub", "Garden", "Other",
]

STOCK_STATUSES = ["In stock", "Low", "Order now", "Out of stock",
                  "Overstocked", "Not tracked"]

STAFF_ROLES = [
    "Cleaner", "Head cleaner", "Laundry service", "Handyman",
    "Plumber", "Electrician", "Gardener", "Pest control",
    "Pool technician", "Co-host", "Property manager", "Photographer",
    "Linen hire", "Locksmith", "Window cleaner", "Accountant", "Other",
]

PAY_BASIS = ["Per turnover", "Per hour", "Per job", "Monthly retainer",
             "Per property per month", "% of revenue", "Invoice each time"]

CONTRACT_STATUSES = ["Active", "Trial", "Occasional", "Paused",
                     "Ended", "Do not use"]

AVAILABILITY = ["Any day", "Weekdays only", "Weekends only",
                "Mornings", "Afternoons", "By arrangement", "Unavailable"]

REVIEW_STATUSES = ["Not requested", "Requested", "Received",
                   "No review left", "Removed", "Disputed"]

RATING_CATEGORIES = [
    ("rate_overall", "Overall"),
    ("rate_clean", "Cleanliness"),
    ("rate_accuracy", "Accuracy"),
    ("rate_checkin", "Check-in"),
    ("rate_comm", "Communication"),
    ("rate_location", "Location"),
    ("rate_value", "Value"),
]

GUEST_TYPES = ["Couple", "Family", "Solo", "Friends", "Business",
               "Group", "Contractor", "Long stay", "Unknown"]

TRIP_PURPOSES = ["Holiday", "Business", "Family visit", "Event",
                 "Relocation", "Medical", "Contract work", "Other"]

PAYOUT_STATUSES = ["Expected", "Received", "Part received",
                   "Late", "Disputed", "Written off"]

YES_NO = ["Yes", "No"]
YES_NO_NA = ["Yes", "No", "N/A"]

COUNTRIES = [
    "United States", "United Kingdom", "Canada", "Ireland", "Australia",
    "New Zealand", "Germany", "France", "Spain", "Portugal", "Italy",
    "Netherlands", "Belgium", "Sweden", "Norway", "Denmark", "Poland",
    "Mexico", "Brazil", "Argentina", "Japan", "South Korea", "Singapore",
    "India", "United Arab Emirates", "South Africa", "Other",
]

MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]

WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
                 "Saturday", "Sunday"]

#: how the length of a stay is bucketed on the performance tab
STAY_BANDS: List[Tuple[str, int, int]] = [
    ("1 night", 1, 1),
    ("2 nights", 2, 2),
    ("3-4 nights", 3, 4),
    ("5-6 nights", 5, 6),
    ("1 week", 7, 9),
    ("10+ nights", 10, 27),
    ("Monthly stay", 28, 400),
]

#: seasons, so a host can see peak against shoulder without a pivot table
SEASONS: List[Tuple[str, int, int]] = [
    ("Winter", 12, 2),
    ("Spring", 3, 5),
    ("Summer", 6, 8),
    ("Autumn", 9, 11),
]


# ---------------------------------------------------------------------------
# Colours - a coastal / host palette: deep navy, sea green, warm sand
# ---------------------------------------------------------------------------

class Palette:
    ink = "#17262F"            # near black navy - titles
    deep = "#113A4C"           # deep navy - money tabs
    primary = "#1C6E78"        # sea green headers
    primary_mid = "#3F97A0"    # section bands
    primary_soft = "#E6F3F4"   # light band / alternating rows
    accent = "#E2703A"         # terracotta accent
    accent_soft = "#FCEFE7"
    input_bg = "#FFFFFF"
    input_border = "#D2E6E8"
    calc_bg = "#F4FAFA"
    calc_border = "#E2F0F1"
    white = "#FFFFFF"
    grey_text = "#5A6B70"
    line = "#C4DEE1"
    good_bg = "#D8EDDD"
    good_text = "#23583A"
    warn_bg = "#FCEBD4"
    warn_text = "#8A5300"
    bad_bg = "#F9DAD6"
    bad_text = "#9B2D20"
    info_bg = "#E1E8F2"
    info_text = "#2F4576"
    note_bg = "#F6FBFB"


# ---------------------------------------------------------------------------
# Row capacity of every register
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Dashboard": 0,
    "Getting Started": 0,
    "Settings & Lists": 0,
    "Property Database": 30,
    "Booking Tracker": 900,
    "Booking Calendar": 0,
    "Guest Database": 400,
    "Check-in & Check-out": 900,
    "Cleaning Schedule": 900,
    "Turnover Checklist": 0,
    "Platform Revenue & Fees": 900,
    "Expense Tracker": 700,
    "Supplies Inventory": 300,
    "Maintenance Log": 300,
    "Reviews Tracker": 600,
    "Staff & Vendors": 60,
    "Tax Summary": 0,
    "Occupancy & Performance": 400,
    "Monthly Reports": 0,
    "Annual Reports": 0,
}

#: the register the `--rows` switch scales everything else against
SCALE_ANCHOR = "Booking Tracker"

#: how many properties the calendar grid and the comparison tables cover
MAX_PROPERTIES = 12

#: days shown on the multi-property availability strip
CALENDAR_STRIP_DAYS = 14

#: rows on the printable turnover checklist's extra-notes block
TURNOVER_NOTE_LINES = 4


# ---------------------------------------------------------------------------
# Automatic ID prefixes
# ---------------------------------------------------------------------------

ID_PREFIXES: Dict[str, Tuple[str, int]] = {
    "property": ("PRP-", 2),
    "booking": ("BKG-", 4),
    "guest": ("GST-", 4),
    "arrival": ("ARR-", 4),
    "clean": ("CLN-", 4),
    "payout": ("PAY-", 4),
    "expense": ("EXP-", 4),
    "supply": ("SUP-", 3),
    "maintenance": ("MNT-", 3),
    "review": ("REV-", 4),
    "vendor": ("VND-", 3),
    "period": ("PER-", 3),
}


@dataclass
class HostDefaults:
    """Sample values written into the business profile on Settings."""

    name: str = "Harbour & Pine Stays"
    owner: str = "Nadia Whitlock"
    brand: str = "Harbour & Pine"
    email: str = "stay@harbourandpine.example"
    phone: str = "(503) 555 0188"
    address: str = "18 Fenwick Street, Portland, OR 97209"
    website: str = "harbourandpine.example"
    tax_id: str = "EIN 87-4410923"
    licence: str = "STR-PDX-2021-00714"
    tagline: str = "Four coastal and city stays, run properly"


# ---------------------------------------------------------------------------
# The configuration object every sheet module reads
# ---------------------------------------------------------------------------

@dataclass
class Config:
    """Everything the build needs to know, in one place."""

    # --------------------------------------------------------- identity ---
    author: str = AUTHOR
    version: str = VERSION
    password: Optional[str] = PASSWORD
    protect_sheets: bool = True
    include_sample_data: bool = True

    host: HostDefaults = field(default_factory=HostDefaults)

    # ------------------------------------------------------------ money ---
    currency_symbol: str = "$"
    currency_code: str = "USD"
    date_format: str = "dd mmm yyyy"
    time_format: str = "hh:mm"

    #: sales / occupancy / lodging tax on the accommodation total
    tax_rate: float = 0.0
    tax_label: str = "Lodging tax"
    #: does the platform collect and remit the lodging tax for you
    tax_collected_by_platform: str = "Yes"

    # ------------------------------------------------- portfolio shape ---
    property_count: int = 4
    default_nightly_rate: float = 185.0
    default_cleaning_fee: float = 95.0
    default_cleaning_cost: float = 65.0
    default_security_deposit: float = 300.0
    default_min_nights: int = 2
    default_check_in: str = "15:00"
    default_check_out: str = "11:00"
    default_max_guests: int = 4
    extra_guest_fee: float = 25.0
    extra_guest_after: int = 2
    pet_fee: float = 40.0

    #: nights a year the host expects a property to be lettable at all
    lettable_nights_year: int = 365
    #: nights the owner expects to block for their own use
    owner_nights_year: int = 14

    # ------------------------------------------------------- platforms ---
    default_platform: str = "Airbnb"
    airbnb_commission: float = 0.030
    vrbo_commission: float = 0.050
    booking_commission: float = 0.150
    direct_card_fee: float = 0.029
    payout_days: int = 1

    # -------------------------------------------------------- operating ---
    laundry_cost_turnover: float = 18.0
    supplies_cost_turnover: float = 12.0
    welcome_cost_turnover: float = 8.0
    cleaner_default_rate: float = 65.0
    management_fee_percent: float = 0.0

    # ---------------------------------------------------------- targets ---
    occupancy_target: float = 0.72
    adr_target: float = 195.0
    revpar_target: float = 140.0
    margin_target: float = 0.38
    rating_target: float = 4.80
    revenue_target_month: float = 14000.0
    revenue_target_year: float = 168000.0
    profit_target_month: float = 5200.0

    # ----------------------------------------------------------- alerts ---
    turnover_warning_hours: int = 6
    review_chase_days: int = 3
    payout_chase_days: int = 7
    maintenance_warning_days: int = 7
    restock_warning_weeks: float = 2.0
    upcoming_window_days: int = 7
    long_stay_nights: int = 28
    lead_time_warning_days: int = 2

    # ------------------------------------------------------------ build ---
    year: int = 0
    today_override: Optional[date] = None
    rows_override: Optional[int] = None
    output: str = ""

    # ----------------------------------------------------------- lookups ---
    def __post_init__(self) -> None:
        if not self.year:
            self.year = self.today.year

    # ------------------------------------------------------------- dates ---
    @property
    def today(self) -> date:
        return self.today_override or date.today()

    # ------------------------------------------------------------ formats ---
    @property
    def money_format(self) -> str:
        sign = self.currency_symbol
        return f'{sign}#,##0.00;[Red]-{sign}#,##0.00'

    @property
    def money_format_0(self) -> str:
        sign = self.currency_symbol
        return f'{sign}#,##0;[Red]-{sign}#,##0'

    @property
    def money_format_4(self) -> str:
        sign = self.currency_symbol
        return f'{sign}#,##0.0000'

    # ------------------------------------------------------------- labels ---
    @property
    def product_name(self) -> str:
        return PRODUCT_NAME

    @property
    def listing_title(self) -> str:
        return LISTING_TITLE

    @property
    def tax_name(self) -> str:
        return self.tax_label

    @property
    def nights_per_year(self) -> int:
        return self.lettable_nights_year

    @property
    def platform_table(self) -> List[Tuple[str, float, float, str, str]]:
        """Platform fee table with the CLI overrides applied."""
        table: List[Tuple[str, float, float, str, str]] = []
        overrides = {
            "Airbnb": self.airbnb_commission,
            "Vrbo": self.vrbo_commission,
            "Booking.com": self.booking_commission,
        }
        for name, commission, payment, counts, note in PLATFORMS:
            commission = overrides.get(name, commission)
            if name in ("Direct booking", "Returning guest"):
                payment = self.direct_card_fee
            table.append((name, commission, payment, counts, note))
        return table

    # -------------------------------------------------------------- rows ---
    def row_count(self, sheet: str) -> int:
        base = DEFAULT_ROWS.get(sheet, 0)
        if not base or not self.rows_override:
            return base
        anchor = DEFAULT_ROWS[SCALE_ANCHOR]
        scaled = int(round(base * self.rows_override / anchor))
        return max(12, scaled)

    def scale_rows(self, count: int) -> int:
        """Scale a hand written row count the same way a register scales."""
        if not self.rows_override:
            return count
        anchor = DEFAULT_ROWS[SCALE_ANCHOR]
        return max(6, int(round(count * self.rows_override / anchor)))

    @property
    def scaled(self) -> bool:
        return bool(self.rows_override)

    # ---------------------------------------------------------------- ids ---
    def rid(self, kind: str, row: int) -> str:
        """The ID a given record row will show, as plain text."""
        prefix, digits = ID_PREFIXES[kind]
        return f"{prefix}{row:0{digits}d}"

    def prefix(self, kind: str) -> str:
        return ID_PREFIXES[kind][0]

    def digits(self, kind: str) -> int:
        return ID_PREFIXES[kind][1]


# ---------------------------------------------------------------------------
# Drop down list names, so a typo is a build error and not a broken workbook
# ---------------------------------------------------------------------------

LIST_NAMES: List[str] = [
    "YesNoList", "YesNoNaList", "BookingStatusList",
    "PaymentStatusList", "PropertyTypeList", "PropertyStatusList",
    "OwnershipList", "CleaningStatusList", "MaintenanceCategoryList",
    "MaintenanceStatusList", "PriorityList", "ExpenseCategoryList",
    "TaxCategoryList", "PaymentMethodList", "FrequencyList",
    "SupplyCategoryList", "StockStatusList", "StaffRoleList",
    "PayBasisList", "ContractStatusList", "AvailabilityList",
    "ReviewStatusList", "GuestTypeList", "TripPurposeList",
    "PayoutStatusList", "CountryList", "MonthNameList", "SeasonList",
    "StayBandList", "RatingList", "PlatformList", "PlatformCommission",
    "PlatformPayment", "PlatformIsOTA", "FrequencyPerYear",
    "MonthKeyList", "MonthKeys", "PropertyList", "PropertyNameList",
    "BookingList", "GuestList", "GuestNameList", "VendorList",
    "VendorNameList", "SupplyList", "SupplyNameList", "CleanTaskList",
    "MaintenanceList",
]


# ---------------------------------------------------------------------------
# The ID prefix table as it appears on Settings
# ---------------------------------------------------------------------------

#: (label on Settings, default prefix, defined name, digits)
ID_PREFIX_FIELDS: List[Tuple[str, str, str, int]] = [
    ("Property", "PRP-", "PropertyPrefix", 2),
    ("Booking", "BKG-", "BookingPrefix", 4),
    ("Guest", "GST-", "GuestPrefix", 4),
    ("Arrival", "ARR-", "ArrivalPrefix", 4),
    ("Cleaning job", "CLN-", "CleanPrefix", 4),
    ("Payout", "PAY-", "PayoutPrefix", 4),
    ("Expense", "EXP-", "ExpensePrefix", 4),
    ("Supply line", "SUP-", "SupplyPrefix", 3),
    ("Maintenance job", "MNT-", "MaintenancePrefix", 3),
    ("Review", "REV-", "ReviewPrefix", 4),
    ("Vendor", "VND-", "VendorPrefix", 3),
    ("Performance row", "PER-", "PeriodPrefix", 3),
]


PRIVACY_NOTE = (
    "A word about guest data. This workbook asks for a name, a way to "
    "contact someone and what they asked for, because that is what it "
    "takes to run the business. It deliberately has nowhere to put a "
    "passport number, a card number or a scan of an ID document, and "
    "you should not add one. Keep the file somewhere only you can "
    "reach, only mail the guests whose consent column says Yes, and "
    "delete the rows you no longer need."
)

DISCLAIMER = (
    "What this workbook is, and is not. It is a management spreadsheet "
    "for running short-term lets: bookings, turnovers, costs and "
    "performance. It is not connected to Airbnb, Vrbo, Booking.com or "
    "any other platform, it is not affiliated with or endorsed by any "
    "of them, and it cannot read or write their calendars - every "
    "figure in it is one you typed or pasted. It does not take "
    "payments and stores no card details. It is not accounting, legal "
    "or tax software, and nothing in it is accounting, legal or tax "
    "advice: the Tax Summary tab groups your own figures under your "
    "own headings so your accountant has less work to do, but what is "
    "deductible, how rental income is classified, whether lodging tax "
    "applies and who must remit it all depend on where the property "
    "is, where you live and how you hold it. It is also not a "
    "licensing or compliance system - short-term let registration, "
    "safety certificates, insurance and local caps on letting nights "
    "are your responsibility. Check your own rules, and check them "
    "with someone qualified."
)
