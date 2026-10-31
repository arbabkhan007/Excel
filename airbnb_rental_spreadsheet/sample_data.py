"""A worked example: four properties, four months, every tab populated.

The numbers are deliberately not flattering.  One property is a
rent-to-rent that barely washes its face, one is a one-bedroom that
takes short stays and gets eaten alive by cleaning costs, and there is
a boiler failure in February that blocks a week.  A sample portfolio
where everything works teaches nobody anything.

Every row here is checked twice before it ships: the keys must all be
*input* columns on their register, and every value that feeds a
drop-down must exist in that drop-down's list.
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Callable, Dict, List

from .config import Config

Rows = List[Dict[str, Any]]


def _day(cfg: Config, offset: int) -> _dt.date:
    """A date `offset` days from the build date."""
    return cfg.today + _dt.timedelta(days=offset)


def _time(hour: int, minute: int = 0) -> _dt.time:
    return _dt.time(hour, minute)


# ---------------------------------------------------------------------------
# 1. Property Database
# ---------------------------------------------------------------------------

def properties(cfg: Config) -> Rows:
    return [
        {
            "name": "Harbour View Loft", "nickname": "Harbour",
            "address": "14 Quayside Walk, Apt 3B", "city": "Portland",
            "country": "United States", "type": "Loft",
            "bedrooms": 1, "bathrooms": 1, "beds": 2, "max_guests": 3,
            "status": "Active", "listed_from": _day(cfg, -690),
            "owner": "Nadia Whitlock", "manager": "Nadia Whitlock",
            "ownership": "Owned outright",
            "main_platform": "Airbnb",
            "listing_url": "https://example.com/listing/harbour-loft",
            "licence": "STR-2023-04417",
            "check_in_time": _time(15), "check_out_time": _time(11),
            "min_nights": 2, "wifi_name": "HarbourLoft-5G",
            "access": "Lockbox by the lift, code 4417",
            "nightly_rate": 165, "cleaning_fee": 85,
            "cleaning_cost": 60, "deposit": 250,
            "monthly_rent": 0, "monthly_mortgage": 0,
            "purchase_price": 289000,
            "notes": "The reliable one. Short stays, high turnover, "
                     "cleaning is the whole game here.",
        },
        {
            "name": "Pine Ridge Cabin", "nickname": "Cabin",
            "address": "2 Pine Ridge Road", "city": "Mount Hood",
            "country": "United States", "type": "Cabin",
            "bedrooms": 3, "bathrooms": 2, "beds": 4, "max_guests": 8,
            "status": "Active", "listed_from": _day(cfg, -520),
            "owner": "Nadia Whitlock", "manager": "Nadia Whitlock",
            "ownership": "Owned with mortgage",
            "main_platform": "Vrbo",
            "listing_url": "https://example.com/listing/pine-ridge",
            "licence": "STR-2023-09112",
            "check_in_time": _time(16), "check_out_time": _time(10),
            "min_nights": 3, "wifi_name": "PineRidge",
            "access": "Key safe on the porch post, code 2291",
            "nightly_rate": 395, "cleaning_fee": 180,
            "cleaning_cost": 135, "deposit": 600,
            "monthly_rent": 0, "monthly_mortgage": 1840,
            "purchase_price": 612000,
            "notes": "Long stays, families, strong summer. The "
                     "mortgage makes winter look thin.",
        },
        {
            "name": "Old Mill Townhouse", "nickname": "Mill",
            "address": "71 Old Mill Lane", "city": "Portland",
            "country": "United States", "type": "Townhouse",
            "bedrooms": 2, "bathrooms": 2, "beds": 3, "max_guests": 5,
            "status": "Active", "listed_from": _day(cfg, -400),
            "owner": "Marcus Oyelaran", "manager": "Nadia Whitlock",
            "ownership": "Rental arbitrage",
            "main_platform": "Booking.com",
            "listing_url": "https://example.com/listing/old-mill",
            "licence": "STR-2024-01188",
            "check_in_time": _time(15), "check_out_time": _time(11),
            "min_nights": 2, "wifi_name": "OldMill-Guest",
            "access": "Smart lock, code sent the morning of arrival",
            "nightly_rate": 210, "cleaning_fee": 110,
            "cleaning_cost": 80, "deposit": 300,
            "monthly_rent": 1950, "monthly_mortgage": 0,
            "purchase_price": 0,
            "notes": "Rent-to-rent. The rent goes out whether anyone "
                     "stays or not, which is the whole risk.",
        },
        {
            "name": "Creekside Bungalow", "nickname": "Creek",
            "address": "9 Creekside Avenue", "city": "Hood River",
            "country": "United States", "type": "Cottage",
            "bedrooms": 2, "bathrooms": 1, "beds": 3, "max_guests": 4,
            "status": "Active", "listed_from": _day(cfg, -210),
            "owner": "Nadia Whitlock", "manager": "Priya Raman",
            "ownership": "Owned with mortgage",
            "main_platform": "Airbnb",
            "listing_url": "https://example.com/listing/creekside",
            "licence": "STR-2025-02204",
            "check_in_time": _time(15), "check_out_time": _time(11),
            "min_nights": 2, "wifi_name": "Creekside",
            "access": "Lockbox on the gate, code 8830",
            "nightly_rate": 185, "cleaning_fee": 95,
            "cleaning_cost": 70, "deposit": 300,
            "monthly_rent": 0, "monthly_mortgage": 980,
            "purchase_price": 338000,
            "notes": "Newest. Still building a review count, which is "
                     "why the rate is where it is.",
        },
        {
            "name": "Fir Street Apartment", "nickname": "Fir St",
            "address": "140 Fir Street, Unit 7", "city": "Portland",
            "country": "United States", "type": "Entire apartment",
            "bedrooms": 1, "bathrooms": 1, "beds": 1, "max_guests": 2,
            "status": "Being set up", "listed_from": _day(cfg, 21),
            "owner": "Nadia Whitlock", "manager": "Nadia Whitlock",
            "ownership": "Owned with mortgage",
            "main_platform": "Airbnb",
            "listing_url": "",
            "licence": "Applied for",
            "check_in_time": _time(15), "check_out_time": _time(11),
            "min_nights": 2, "wifi_name": "FirSt-Guest",
            "access": "To be fitted",
            "nightly_rate": 145, "cleaning_fee": 75,
            "cleaning_cost": 55, "deposit": 200,
            "monthly_rent": 0, "monthly_mortgage": 720,
            "purchase_price": 241000,
            "notes": "Opens next month. Here to show what a property "
                     "looks like before it earns anything.",
        },
    ]


# ---------------------------------------------------------------------------
# 2. Staff & Vendors
# ---------------------------------------------------------------------------

def vendors(cfg: Config) -> Rows:
    return [
        {"name": "Priya Raman", "role": "Head cleaner",
         "company": "Raman Turnovers", "phone": "+1 503 555 0142",
         "email": "priya@example.com", "properties": "All Portland",
         "availability": "Any day",
         "pay_basis": "Per turnover", "rate": 65,
         "contract": "Active", "started": _day(cfg, -640),
         "insured": "Yes", "insurance_expiry": _day(cfg, 142),
         "keyholder": "Yes",
         "notes": "Does the bulk of the work and is worth every "
                  "penny. Also co-manages Creekside."},
        {"name": "Tomas Beric", "role": "Cleaner",
         "company": "", "phone": "+1 503 555 0188",
         "email": "tomas@example.com", "properties": "Cabin, Creek",
         "availability": "Weekends only",
         "pay_basis": "Per turnover", "rate": 120,
         "contract": "Active", "started": _day(cfg, -300),
         "insured": "No", "insurance_expiry": "",
         "keyholder": "Yes",
         "notes": "Covers the cabin, which is a four-hour turnover "
                  "and priced accordingly. Chase him for insurance."},
        {"name": "Glenbrook Laundry", "role": "Laundry service",
         "company": "Glenbrook Laundry Co.",
         "phone": "+1 503 555 0310", "email": "orders@example.com",
         "properties": "All", "availability": "Weekdays only",
         "pay_basis": "Invoice each time", "rate": 3.40,
         "contract": "Active", "started": _day(cfg, -610),
         "insured": "Yes", "insurance_expiry": _day(cfg, 310),
         "keyholder": "No",
         "notes": "Collects Tuesday and Friday. Cheaper than "
                  "in-house once you count the machine."},
        {"name": "Dale Kirkbride", "role": "Handyman",
         "company": "Kirkbride Property Services",
         "phone": "+1 503 555 0266", "email": "dale@example.com",
         "properties": "All", "availability": "By arrangement",
         "pay_basis": "Per hour", "rate": 72,
         "contract": "Active", "started": _day(cfg, -580),
         "insured": "Yes", "insurance_expiry": _day(cfg, 24),
         "keyholder": "Yes",
         "notes": "Insurance is about to lapse - the Action column "
                  "has already noticed."},
        {"name": "Hollis Plumbing", "role": "Plumber",
         "company": "Hollis & Sons", "phone": "+1 503 555 0401",
         "email": "jobs@example.com", "properties": "All",
         "availability": "By arrangement", "pay_basis": "Per job",
         "rate": 240, "contract": "Active",
         "started": _day(cfg, -420), "insured": "Yes",
         "insurance_expiry": _day(cfg, 201), "keyholder": "No",
         "notes": "Emergency rate after six. Did the cabin boiler."},
        {"name": "Ines Fuentes", "role": "Co-host",
         "company": "", "phone": "+1 503 555 0529",
         "email": "ines@example.com", "properties": "Harbour, Mill",
         "availability": "Weekdays only", "pay_basis": "% of revenue",
         "rate": 0, "contract": "Active",
         "started": _day(cfg, -120), "insured": "No",
         "insurance_expiry": "", "keyholder": "Yes",
         "notes": "Takes 12% of net on the two Portland units for "
                  "guest messaging. Logged as a co-host fee."},
        {"name": "Evergreen Gardens", "role": "Gardener",
         "company": "Evergreen Gardens LLC",
         "phone": "+1 503 555 0633", "email": "hello@example.com",
         "properties": "Cabin, Creek", "availability": "Weekdays only",
         "pay_basis": "Monthly retainer", "rate": 180,
         "contract": "Active", "started": _day(cfg, -500),
         "insured": "Yes", "insurance_expiry": _day(cfg, 95),
         "keyholder": "No", "notes": "Fortnightly in summer."},
        {"name": "Pacific Pest Control", "role": "Pest control",
         "company": "Pacific Pest Control",
         "phone": "+1 503 555 0744", "email": "book@example.com",
         "properties": "All", "availability": "By arrangement",
         "pay_basis": "Per job", "rate": 150,
         "contract": "Occasional", "started": _day(cfg, -360),
         "insured": "Yes", "insurance_expiry": _day(cfg, 180),
         "keyholder": "No",
         "notes": "Annual preventative visit. Not used since."},
    ]


# ---------------------------------------------------------------------------
# 3. Guest Database
# ---------------------------------------------------------------------------

_GUESTS = [
    ("Rebecca Thorne", "rebecca.t@example.com", "+1 206 555 0118",
     "United States", "Couple", "Holiday", "Airbnb", "Yes", "Received",
     "Ground floor if possible", "Late check-in, arriving 10pm",
     "Loved it. Said the bed was the best part.",
     "Excellent guest. Would have back tomorrow."),
    ("Dai Nguyen", "dai.nguyen@example.com", "+1 503 555 0227",
     "United States", "Business", "Business", "Direct booking",
     "Yes", "Received", "Desk and fast wifi",
     "Needs a receipt addressed to the company",
     "Quiet, tidy, books the same week every quarter.",
     "Repeat corporate. Give him the direct rate."),
    ("Marion Esseltine", "marion.e@example.com", "+44 7700 900412",
     "United Kingdom", "Family", "Holiday", "Vrbo", "No", "Received",
     "Travel cot and high chair",
     "Flying in late, needs flexible check-in",
     "Said the cabin was perfect for the children.",
     "Family of five. Books the cabin in August."),
    ("Grant Oyelaran", "grant.o@example.com", "+1 971 555 0340",
     "United States", "Friends", "Holiday", "Booking.com", "No",
     "Not requested", "", "", "", "One night, left a towel behind."),
    ("Suzanna Pell", "suzanna.pell@example.com", "+1 415 555 0455",
     "United States", "Couple", "Holiday", "Airbnb", "Yes", "Received",
     "Prefers a firm mattress", "Early check-in if free",
     "Marked us down on cleanliness. Fair - that was the "
     "Tuesday Tomas could not cover.",
     "The review that cost us. Worth winning back."),
    ("Hector Bellweather", "hector.b@example.com", "+1 503 555 0562",
     "United States", "Couple", "Holiday", "Airbnb", "Yes", "Received",
     "Dog friendly", "Bringing one small dog",
     "Great stay, will return with the dog.",
     "Pet fee paid without being asked. Good sign."),
    ("Amira Sadiq", "amira.sadiq@example.com", "+1 646 555 0677",
     "United States", "Long stay", "Relocation", "Corporate / "
     "relocation", "No", "Not requested", "Long stay, needs laundry",
     "Six weeks while her flat is found",
     "", "Longest booking of the year. Worth a discount."),
    ("Niall Fergusson", "niall.f@example.com", "+353 87 555 0781",
     "Ireland", "Couple", "Holiday", "Expedia", "No", "Received", "",
     "", "Nothing but five stars.",
     "Booked through Expedia - the expensive channel."),
    ("Petra Lindqvist", "petra.l@example.com", "+46 70 555 0894",
     "Sweden", "Couple", "Event", "Airbnb", "Yes", "Received",
     "Quiet room, late check-out",
     "Celebrating an anniversary - a bottle would be welcome",
     "Said the welcome basket made the trip.",
     "Welcome basket costs eleven dollars and bought a "
     "five-star review."),
    ("Owen Castellanos", "owen.c@example.com", "+1 503 555 0902",
     "United States", "Group", "Event", "Vrbo", "No",
     "Received", "Eight people, two cars",
     "Wedding party - needs the full cabin",
     "Great space, slightly tired bathroom.",
     "Flagged the bathroom. Booked Dale to look at it."),
    ("Fenella Adebayo", "fenella.a@example.com", "+1 212 555 1014",
     "United States", "Couple", "Holiday", "Direct booking", "Yes",
     "Not requested", "", "Returning - third stay", "",
     "Our best direct guest. Zero commission three times over."),
    ("Jurgen Haas", "jurgen.haas@example.com", "+49 151 555 1127",
     "Germany", "Couple", "Holiday", "Booking.com", "No",
     "Not requested", "", "Arriving by train, no car", "",
     "Cancelled late. Booking.com terms meant no fee."),
]


def guests(cfg: Config) -> Rows:
    keys = ("name", "email", "phone", "country", "guest_type",
            "purpose", "first_platform", "marketing_ok",
            "review_status", "preferences", "requests", "feedback",
            "internal")
    return [dict(zip(keys, row)) for row in _GUESTS]


# ---------------------------------------------------------------------------
# 4. Booking Tracker
# ---------------------------------------------------------------------------

#: (property, guest, platform, status, booked_offset, in_offset,
#:  nights, adults, children, pets, rate, payment, requests)
_BOOKINGS = [
    # ---- finished stays, three months back ----------------------------
    ("PRP-01", "GST-01", "Airbnb", "Completed", -112, -96, 3, 2, 0, 0,
     165, "Paid in full", "Late check-in, arriving 10pm"),
    ("PRP-02", "GST-03", "Vrbo", "Completed", -120, -93, 7, 2, 3, 0,
     410, "Paid in full", "Travel cot and high chair please"),
    ("PRP-03", "GST-04", "Booking.com", "Completed", -95, -90, 1, 2, 0,
     0, 210, "Paid in full", ""),
    ("PRP-01", "GST-02", "Direct booking", "Completed", -100, -88, 2,
     1, 0, 0, 155, "Paid in full", "Receipt to the company please"),
    ("PRP-04", "GST-06", "Airbnb", "Completed", -99, -85, 4, 2, 0, 1,
     185, "Paid in full", "One small dog"),
    ("PRP-01", "GST-05", "Airbnb", "Completed", -90, -82, 2, 2, 0, 0,
     165, "Paid in full", "Early check-in if free"),
    ("PRP-02", "GST-10", "Vrbo", "Completed", -102, -78, 3, 8, 0, 0,
     450, "Paid in full", "Wedding party, eight people"),
    ("PRP-03", "GST-08", "Expedia", "Completed", -84, -74, 4, 2, 0, 0,
     215, "Paid in full", ""),
    ("PRP-01", "GST-09", "Airbnb", "Completed", -80, -70, 3, 2, 0, 0,
     175, "Paid in full", "Anniversary - quiet room"),
    ("PRP-04", "GST-01", "Airbnb", "Completed", -75, -66, 2, 2, 0, 0,
     185, "Paid in full", "Second stay"),

    # ---- two months back ----------------------------------------------
    ("PRP-02", "GST-03", "Vrbo", "Completed", -78, -62, 5, 2, 3, 0,
     395, "Paid in full", "Returning family"),
    ("PRP-03", "GST-11", "Direct booking", "Completed", -70, -58, 3, 2,
     0, 0, 195, "Paid in full", "Third stay - direct"),
    ("PRP-01", "GST-04", "Booking.com", "Completed", -62, -55, 1, 2, 0,
     0, 170, "Paid in full", ""),
    ("PRP-04", "GST-05", "Airbnb", "Completed", -66, -52, 3, 2, 0, 0,
     185, "Paid in full", "Firm mattress if possible"),
    ("PRP-01", "GST-06", "Airbnb", "Completed", -58, -48, 2, 2, 0, 1,
     165, "Paid in full", "Dog again"),
    ("PRP-02", "GST-10", "Vrbo", "Cancelled", -64, -46, 4, 6, 0, 0,
     395, "Refunded", "Cancelled eleven days out"),
    ("PRP-03", "GST-12", "Booking.com", "Cancelled", -55, -44, 2, 2, 0,
     0, 210, "Refunded", "Cancelled, no fee under their terms"),
    ("PRP-01", "GST-02", "Direct booking", "Completed", -50, -41, 2, 1,
     0, 0, 155, "Paid in full", "Quarterly work trip"),
    ("PRP-04", "GST-08", "Expedia", "Completed", -52, -38, 3, 2, 0, 0,
     190, "Paid in full", ""),
    ("PRP-02", "GST-07", "Corporate / relocation", "Completed", -60,
     -36, 14, 1, 0, 0, 340, "Paid in full",
     "Relocation - six weeks split over two bookings"),

    # ---- last month, including a maintenance block ---------------------
    ("PRP-02", "", "Direct booking", "Maintenance block", -34, -32, 6,
     0, 0, 0, 0, "Not due yet", "Boiler failed - closed the cabin"),
    ("PRP-01", "GST-09", "Airbnb", "Completed", -40, -29, 3, 2, 0, 0,
     175, "Paid in full", ""),
    ("PRP-03", "GST-01", "Booking.com", "Completed", -36, -26, 2, 2, 0,
     0, 210, "Paid in full", "Third property for this guest"),
    ("PRP-04", "GST-11", "Direct booking", "Completed", -38, -24, 4, 2,
     0, 0, 185, "Paid in full", ""),
    ("PRP-01", "GST-05", "Airbnb", "Completed", -30, -21, 1, 2, 0, 0,
     165, "Paid in full", "One night - the expensive kind"),
    ("PRP-02", "GST-07", "Corporate / relocation", "Completed", -45,
     -20, 14, 1, 0, 0, 340, "Paid in full", "Second half"),
    ("PRP-03", "GST-06", "Airbnb", "Completed", -28, -18, 3, 2, 0, 1,
     215, "Paid in full", "Dog"),
    ("PRP-01", "GST-08", "Expedia", "Completed", -26, -14, 2, 2, 0, 0,
     170, "Paid in full", ""),
    ("PRP-04", "GST-09", "Airbnb", "Completed", -24, -11, 3, 2, 0, 0,
     190, "Paid in full", ""),

    # ---- this month, in progress ---------------------------------------
    ("PRP-01", "GST-11", "Direct booking", "Completed", -20, -8, 2, 2,
     0, 0, 160, "Paid in full", "Direct again - no commission"),
    ("PRP-03", "GST-04", "Booking.com", "Completed", -18, -6, 2, 2, 0,
     0, 210, "Awaiting payout", ""),
    ("PRP-02", "GST-03", "Vrbo", "Completed", -30, -5, 4, 2, 3, 0,
     420, "Awaiting payout", "Half term"),
    ("PRP-04", "GST-05", "Airbnb", "Checked in", -16, -2, 5, 2, 0, 0,
     195, "Awaiting payout", "Currently in the property"),
    ("PRP-01", "GST-06", "Airbnb", "Checked in", -14, -1, 3, 2, 0, 1,
     175, "Awaiting payout", "Dog, as usual"),

    # ---- today and the days around it ----------------------------------
    ("PRP-03", "GST-09", "Airbnb", "Confirmed", -10, 0, 3, 2, 0, 0,
     220, "Awaiting payout", "Arriving today"),
    ("PRP-02", "GST-10", "Vrbo", "Confirmed", -12, 1, 4, 7, 0, 0, 430,
     "Part paid", "Arriving tomorrow - seven guests"),
    ("PRP-01", "GST-01", "Airbnb", "Confirmed", -8, 2, 2, 2, 0, 0, 180,
     "Part paid", ""),
    ("PRP-04", "GST-08", "Expedia", "Confirmed", -6, 4, 3, 2, 0, 0,
     195, "Not due yet", ""),
    ("PRP-03", "GST-11", "Direct booking", "Confirmed", -5, 6, 4, 2, 0,
     0, 200, "Part paid", "Fourth stay"),
    ("PRP-01", "GST-02", "Direct booking", "Confirmed", -4, 9, 2, 1, 0,
     0, 155, "Part paid", "Quarterly work trip"),
    ("PRP-02", "GST-03", "Vrbo", "Confirmed", -22, 12, 7, 2, 3, 0, 440,
     "Part paid", "Booked well ahead"),
    ("PRP-04", "GST-06", "Airbnb", "Confirmed", -3, 15, 3, 2, 0, 1,
     195, "Not due yet", "Dog"),
    ("PRP-01", "", "Airbnb", "Owner block", -2, 18, 4, 0, 0, 0, 0,
     "Not due yet", "Keeping it for family"),
    ("PRP-03", "GST-12", "Booking.com", "Inquiry", 0, 22, 3, 2, 0, 0,
     0, "Not due yet", "Asked about the train station"),
    ("PRP-02", "GST-10", "Vrbo", "Confirmed", -1, 26, 5, 8, 0, 0, 460,
     "Part paid", "Full house again"),
    ("PRP-04", "GST-05", "Airbnb", "Inquiry", 0, 31, 2, 2, 0, 0, 0,
     "Not due yet", "Asked whether the creek is swimmable"),
]


def bookings(cfg: Config) -> Rows:
    rows: Rows = []
    for (prop, guest, platform, status, booked, start, nights, adults,
         children, pets, rate, payment, requests) in _BOOKINGS:
        row: Dict[str, Any] = {
            "property": prop,
            "platform": platform,
            "status": status,
            "booked_date": _day(cfg, booked),
            "check_in": _day(cfg, start),
            "check_out": _day(cfg, start + nights),
            "adults": adults,
            "payment_status": payment,
        }
        if guest:
            row["guest"] = guest
        if children:
            row["children"] = children
        if pets:
            row["pets"] = pets
        if rate:
            row["rate_in"] = rate
        if requests:
            row["requests"] = requests
        if platform in ("Airbnb", "Vrbo", "Booking.com", "Expedia"):
            row["reference"] = f"HM{abs(start):04d}{adults}{nights}"
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 5. Check-in & Check-out
# ---------------------------------------------------------------------------

def arrivals(cfg: Config) -> Rows:
    """One row per recent and upcoming stay, in varying states."""
    rows: Rows = []
    finished = [
        (22, "Smart lock code", True, "Guest left a towel behind."),
        (23, "Lockbox", True, ""),
        (24, "Lockbox", True, "Shower pressure was mentioned again."),
        (25, "Lockbox", True, "One night. Barely worth the turnover."),
        (26, "Key safe", True, "Long stay - laundry done mid-way."),
        (27, "Smart lock code", True, ""),
        (28, "Lockbox", True, "Late departure, cleaner had to wait."),
        (29, "Lockbox", True, ""),
        (30, "Lockbox", True, ""),
        (31, "Smart lock code", True, ""),
        (32, "Key safe", True, ""),
    ]
    for index, method, complete, issues in finished:
        row: Dict[str, Any] = {
            "booking": f"BKG-{index:04d}",
            "arrival_time": _time(16),
            "departure_time": _time(10, 30),
            "late_checkout": "No", "early_checkin": "No",
            "key_method": method,
            "issues": issues,
        }
        for key in ("sent_directions", "sent_access", "sent_wifi",
                    "sent_rules", "sent_welcome"):
            row[key] = "Yes"
        if complete:
            for key in ("sent_checkout", "checked_out", "inspected",
                        "photographed"):
                row[key] = "Yes"
        rows.append(row)

    # in-house now: pre-arrival done, nothing after departure yet
    for index, method in ((33, "Lockbox"), (34, "Smart lock code")):
        rows.append({
            "booking": f"BKG-{index:04d}",
            "arrival_time": _time(15, 30),
            "departure_time": _time(11),
            "late_checkout": "No", "early_checkin": "Yes",
            "key_method": method,
            "sent_directions": "Yes", "sent_access": "Yes",
            "sent_wifi": "Yes", "sent_rules": "Yes",
            "sent_welcome": "Yes",
            "notes": "Guest is in the property now.",
        })

    # arriving today: everything sent
    rows.append({
        "booking": "BKG-0035", "arrival_time": _time(15),
        "departure_time": _time(11), "late_checkout": "No",
        "early_checkin": "No", "key_method": "Lockbox",
        "sent_directions": "Yes", "sent_access": "Yes",
        "sent_wifi": "Yes", "sent_rules": "Yes", "sent_welcome": "Yes",
        "notes": "Arriving today and fully briefed.",
    })
    # arriving tomorrow: access code not sent - the one to chase
    rows.append({
        "booking": "BKG-0036", "arrival_time": _time(16),
        "departure_time": _time(10), "late_checkout": "No",
        "early_checkin": "No", "key_method": "Key safe",
        "sent_directions": "Yes", "sent_access": "No",
        "sent_wifi": "Yes", "sent_rules": "Yes", "sent_welcome": "No",
        "notes": "Seven guests arriving tomorrow and the access code "
                 "has not gone out. This is what the Action column is "
                 "for.",
    })
    # arriving in a few days: barely started
    for index in (37, 38, 39):
        rows.append({
            "booking": f"BKG-{index:04d}",
            "arrival_time": _time(15),
            "departure_time": _time(11),
            "late_checkout": "No", "early_checkin": "No",
            "key_method": "Lockbox",
            "sent_directions": "Yes", "sent_access": "No",
            "sent_wifi": "No", "sent_rules": "No",
            "sent_welcome": "No",
        })
    return rows


# ---------------------------------------------------------------------------
# 6. Cleaning Schedule
# ---------------------------------------------------------------------------

def cleaning(cfg: Config) -> Rows:
    """A turnover for every completed stay, plus the ones coming up."""
    every_task = ("task_bedroom", "task_bathroom", "task_kitchen",
                  "task_living", "task_floors", "task_laundry",
                  "task_rubbish", "task_restock", "task_inspection")
    # booking number -> (cleaner, hours, status, skipped tasks, note)
    jobs = [
        (1, "VND-01", 2.0, "Complete", (), ""),
        (2, "VND-02", 4.0, "Complete", (), "Cabin after a week - "
                                            "four hours is normal."),
        (3, "VND-01", 2.5, "Complete", (), ""),
        (4, "VND-01", 2.0, "Complete", (), ""),
        (5, "VND-01", 2.5, "Complete", (), ""),
        (6, "VND-01", 2.0, "Complete", ("task_restock",),
         "Ran out of coffee pods - restock not done."),
        (7, "VND-02", 4.5, "Complete", (), "Eight guests. Enough "
                                            "said."),
        (8, "VND-01", 2.5, "Complete", (), ""),
        (9, "VND-01", 2.0, "Complete", (), ""),
        (10, "VND-01", 2.5, "Complete", (), ""),
        (11, "VND-02", 4.0, "Complete", (), ""),
        (12, "VND-01", 2.5, "Complete", (), ""),
        (13, "VND-01", 2.0, "Complete", (), ""),
        (14, "VND-01", 2.5, "Complete", ("task_floors",),
         "Vacuum broke mid-clean. This is the turnover that cost us "
         "a cleanliness score."),
        (15, "VND-01", 2.0, "Complete", (), ""),
        (18, "VND-01", 2.0, "Complete", (), ""),
        (19, "VND-01", 2.5, "Complete", (), ""),
        (20, "VND-02", 3.5, "Complete", (), "Mid-stay clean on the "
                                             "relocation booking."),
        (22, "VND-01", 2.0, "Complete", (), ""),
        (23, "VND-01", 2.5, "Complete", (), ""),
        (24, "VND-01", 2.5, "Complete", (), ""),
        (25, "VND-01", 2.0, "Complete", (), "One-night stay. Same "
                                             "cost as a week."),
        (26, "VND-02", 3.5, "Complete", (), ""),
        (27, "VND-01", 2.5, "Complete", (), ""),
        (28, "VND-01", 2.0, "Complete", (), ""),
        (29, "VND-01", 2.5, "Complete", (), ""),
        (30, "VND-01", 2.0, "Complete", (), ""),
        (31, "VND-01", 2.5, "Complete", (), ""),
        (32, "VND-02", 4.0, "Complete", (), ""),
    ]
    rows: Rows = []
    for number, cleaner, hours, status, skipped, note in jobs:
        start_hour = 11
        row: Dict[str, Any] = {
            "booking": f"BKG-{number:04d}",
            "cleaner": cleaner,
            "start_time": _time(start_hour),
            "end_time": _time(start_hour + int(hours),
                              30 if hours % 1 else 0),
            "status": status,
            "laundry_cost": 28 if cleaner == "VND-02" else 16,
            "supplies_cost": 14 if cleaner == "VND-02" else 9,
            "paid": "Yes",
        }
        for key in every_task:
            row[key] = "No" if key in skipped else "Yes"
        if note:
            row["notes"] = note
        rows.append(row)

    # the in-flight ones
    rows.append({
        "booking": "BKG-0033", "cleaner": "VND-01",
        "status": "Scheduled", "laundry_cost": 16, "supplies_cost": 9,
        "paid": "No",
        "notes": "Guest still in. Clean is booked for check-out day.",
    })
    rows.append({
        "booking": "BKG-0034", "cleaner": "VND-01",
        "start_time": _time(11), "end_time": _time(13),
        "status": "In progress", "laundry_cost": 16,
        "supplies_cost": 9, "paid": "No",
        "task_bedroom": "Yes", "task_bathroom": "Yes",
        "task_kitchen": "Yes", "task_living": "No",
        "task_floors": "No", "task_laundry": "Yes",
        "task_rubbish": "No", "task_restock": "No",
        "task_inspection": "No",
        "notes": "Same-day turnover happening right now - a guest "
                 "arrives this afternoon.",
    })
    rows.append({
        "booking": "BKG-0035", "cleaner": "VND-01",
        "status": "Scheduled", "laundry_cost": 16, "supplies_cost": 9,
        "paid": "No", "notes": "Booked ahead for check-out day.",
    })
    rows.append({
        "booking": "BKG-0036", "cleaner": "VND-02",
        "status": "Scheduled", "laundry_cost": 28, "supplies_cost": 14,
        "paid": "No",
        "notes": "Cabin, eight guests. Tomas, four hours.",
    })
    return rows


# ---------------------------------------------------------------------------
# 7. Platform Revenue & Fees
# ---------------------------------------------------------------------------

def payouts(cfg: Config) -> Rows:
    """Reconciled payouts, including two where the fee was not as
    assumed - which is the entire point of the tab."""
    settled = [
        (1, 0.0, 0.0, 0.0), (2, 0.0, 0.0, 0.0), (3, 0.0, 0.0, 0.0),
        (4, 0.0, 0.0, 0.0), (5, 0.0, 0.0, 0.0), (6, 0.0, 0.0, 0.0),
        (7, 0.0, 0.0, 0.0), (8, -18.40, 0.0, 0.0),
        (9, 0.0, 0.0, 0.0), (10, 0.0, 0.0, 0.0),
        (11, 0.0, 0.0, 0.0), (12, 0.0, 0.0, 0.0),
        (13, 0.0, 0.0, 0.0), (14, 0.0, 0.0, -35.00),
        (15, 0.0, 0.0, 0.0), (18, 0.0, 0.0, 0.0),
        (19, -22.10, 0.0, 0.0), (20, 0.0, 0.0, 0.0),
        (22, 0.0, 0.0, 0.0), (23, 0.0, 0.0, 0.0),
        (24, 0.0, 0.0, 0.0), (25, 0.0, 0.0, 0.0),
        (26, 0.0, 0.0, 0.0), (27, 0.0, 0.0, 0.0),
        (28, 0.0, 0.0, 0.0), (29, 0.0, 0.0, 0.0),
        (30, 0.0, 0.0, 0.0),
    ]
    rows: Rows = []
    for number, variance, refund, adjustment in settled:
        rows.append({
            "booking": f"BKG-{number:04d}",
            "status": "Received",
            "refund": refund or "",
            "adjustment": adjustment or "",
            "reference": f"PO-{number:05d}",
            "notes": ("Expedia kept more than the Settings rate - "
                      "worth checking the contract."
                      if variance else ""),
        })
    # awaiting
    for number in (31, 32):
        rows.append({
            "booking": f"BKG-{number:04d}",
            "status": "Expected",
            "reference": f"PO-{number:05d}",
            "notes": "Due in the next payout run.",
        })
    # overdue - the one the dashboard chases
    rows.append({
        "booking": "BKG-0016",
        "status": "Late",
        "reference": "PO-00016",
        "notes": "Cancelled stay. Platform still owes the "
                 "non-refundable portion and it is now well past "
                 "their usual terms.",
    })
    return rows


# ---------------------------------------------------------------------------
# 8. Expense Tracker
# ---------------------------------------------------------------------------

def expenses(cfg: Config) -> Rows:
    base: Rows = []
    recurring = [
        ("PRP-03", "Rent", "Fulcrum Lettings",
         "Monthly rent to the landlord", 1950.00, "Monthly",
         "Operating expense", "Bank transfer"),
        ("PRP-02", "Mortgage interest", "Cascade Mutual",
         "Interest portion of the cabin mortgage", 1180.00, "Monthly",
         "Financing cost", "Direct debit"),
        ("PRP-04", "Mortgage interest", "Cascade Mutual",
         "Interest portion of the bungalow mortgage", 604.00,
         "Monthly", "Financing cost", "Direct debit"),
        ("PRP-01", "Utilities - electric", "Northwest Power",
         "Electricity", 86.40, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-02", "Utilities - electric", "Northwest Power",
         "Electricity - the cabin is not cheap to heat", 214.80,
         "Monthly", "Operating expense", "Direct debit"),
        ("PRP-03", "Utilities - electric", "Northwest Power",
         "Electricity", 104.20, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-04", "Utilities - electric", "Northwest Power",
         "Electricity", 92.10, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-01", "Internet & TV", "Stream Fibre",
         "Guest wifi", 55.00, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-02", "Internet & TV", "Stream Fibre",
         "Guest wifi", 62.00, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-03", "Internet & TV", "Stream Fibre",
         "Guest wifi", 55.00, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-04", "Internet & TV", "Stream Fibre",
         "Guest wifi", 55.00, "Monthly", "Operating expense",
         "Direct debit"),
        ("", "Insurance", "Harbourline Insurance",
         "Short-term let policy, all properties", 268.00, "Monthly",
         "Operating expense", "Direct debit"),
        ("", "Software & subscriptions", "PriceLabs",
         "Dynamic pricing tool", 39.00, "Monthly",
         "Operating expense", "Credit card"),
        ("", "Software & subscriptions", "Hospitable",
         "Guest messaging", 48.00, "Monthly", "Operating expense",
         "Credit card"),
        ("PRP-02", "Service charge", "Pine Ridge Association",
         "Road and gate maintenance levy", 95.00, "Monthly",
         "Operating expense", "Bank transfer"),
    ]
    for offset_months in range(4):
        day = _day(cfg, -offset_months * 30 - 4)
        for (prop, category, vendor, description, amount, frequency,
             tax_category, method) in recurring:
            row: Dict[str, Any] = {
                "date": day, "category": category, "vendor": vendor,
                "description": description, "amount": amount,
                "method": method, "receipt": "Yes",
                "frequency": frequency, "tax_category": tax_category,
            }
            if prop:
                row["property"] = prop
            base.append(row)

    one_offs = [
        (-104, "PRP-02", "Furniture", "Rowan & Field",
         "Replacement sofa bed for the cabin", 1240.00,
         "Capital / asset", "Credit card",
         "Capital, not a repair. It will not touch this year's "
         "profit and your accountant will be pleased."),
        (-92, "", "Marketing & photography", "Lumen Studio",
         "Re-shoot of the Portland listings", 780.00,
         "Operating expense", "Bank transfer",
         "Paid for itself in six weeks of better conversion."),
        (-76, "PRP-01", "Appliances", "Harbour Appliance Co.",
         "New washer-dryer", 689.00, "Capital / asset", "Credit card", ""),
        (-64, "PRP-04", "Guest amenities", "Cedar & Salt",
         "Welcome baskets, twelve", 132.00, "Operating expense",
         "Credit card", "Eleven dollars a basket bought a five-star "
                 "review. Best money spent all quarter."),
        (-58, "", "Accounting & legal", "Mercer & Vance",
         "Quarterly bookkeeping", 420.00, "Operating expense",
         "Bank transfer", ""),
        (-46, "PRP-02", "Repairs", "Hollis & Sons",
         "Emergency boiler repair - cabin closed six nights",
         1180.00, "Operating expense", "Bank transfer",
         "The week that cost the cabin its February."),
        (-44, "PRP-02", "Guest refunds", "Marion Esseltine",
         "Partial refund, boiler outage", 395.00,
         "Operating expense", "Bank transfer", ""),
        (-38, "", "Licences & permits", "City of Portland",
         "Short-term let registration renewal", 310.00,
         "Operating expense", "Credit card", ""),
        (-31, "PRP-03", "Co-host fees", "Ines Fuentes",
         "12% of net on the Portland units", 288.00,
         "Operating expense", "Bank transfer", ""),
        (-24, "PRP-01", "Supplies", "Restock Wholesale",
         "Linen, towels, consumables", 246.00, "Operating expense",
         "Credit card", ""),
        (-18, "", "Travel & mileage", "",
         "Driving between properties, 340 miles", 221.00,
         "Operating expense", "Cash", ""),
        (-12, "PRP-04", "Repairs", "Kirkbride Property Services",
         "Re-seal the bathroom, re-grout tiles", 340.00,
         "Operating expense", "Bank transfer",
         "Raised by a guest in a review. Fixed within the week."),
        (-9, "", "Bank & card fees", "Northline Bank",
         "Account and card fees", 42.00, "Operating expense",
         "Direct debit", ""),
        (-6, "PRP-05", "Furniture", "Rowan & Field",
         "Furnishing the new Fir Street flat", 2840.00,
         "Capital / asset", "Credit card",
         "Setting up a property that has not earned a night yet."),
        (-3, "PRP-05", "Appliances", "Harbour Appliance Co.",
         "Fridge, washer and oven for Fir Street", 1690.00,
         "Capital / asset", "Credit card", ""),
    ]
    for (offset, prop, category, vendor, description, amount,
         tax_category, method, note) in one_offs:
        row = {
            "date": _day(cfg, offset), "category": category,
            "vendor": vendor, "description": description,
            "amount": amount, "method": method, "receipt": "Yes",
            "frequency": "One off", "tax_category": tax_category,
        }
        if prop:
            row["property"] = prop
        if note:
            row["notes"] = note
        base.append(row)
    return base


# ---------------------------------------------------------------------------
# 9. Supplies Inventory
# ---------------------------------------------------------------------------

def supplies(cfg: Config) -> Rows:
    items = [
        ("Toilet rolls", "Soap & hand wash", "", "Pack of 9",
         "Restock Wholesale", "RW-TR09", 14, 6, 3, 8.40, 12,
         "Across the whole portfolio."),
        ("Kitchen roll", "Kitchen supplies", "", "Pack of 4",
         "Restock Wholesale", "RW-KR04", 9, 4, 1.5, 6.20, 10, ""),
        ("Hand soap refill", "Soap & hand wash", "", "500ml",
         "Restock Wholesale", "RW-HS50", 7, 6, 0.5, 3.10, 12,
         "At the minimum. The Action column has noticed."),
        ("Shampoo refill", "Soap & hand wash", "", "500ml",
         "Restock Wholesale", "RW-SH50", 11, 6, 0.6, 3.80, 12, ""),
        ("Dishwasher tablets", "Kitchen supplies", "", "Box of 60",
         "Restock Wholesale", "RW-DT60", 3, 3, 4, 11.50, 6,
         "Exactly at minimum with four turnovers a month."),
        ("Coffee pods", "Welcome treats", "", "Box of 50",
         "Cedar & Salt", "CS-CP50", 0, 2, 6, 18.90, 6,
         "Out of stock. This is why one turnover shows restock "
         "not done."),
        ("Tea bags", "Welcome treats", "", "Box of 80",
         "Cedar & Salt", "CS-TB80", 4, 2, 4, 6.40, 4, ""),
        ("Bin bags", "Cleaning products", "", "Roll of 50",
         "Restock Wholesale", "RW-BB50", 6, 3, 3, 7.20, 6, ""),
        ("Multi-surface cleaner", "Cleaning products", "", "750ml",
         "Restock Wholesale", "RW-MS75", 8, 4, 1, 4.30, 8, ""),
        ("Bleach", "Cleaning products", "", "1 litre",
         "Restock Wholesale", "RW-BL10", 5, 3, 0.5, 2.90, 6, ""),
        ("King duvet covers", "Bed linen", "PRP-02", "Each",
         "Glenbrook Laundry", "GL-DK", 6, 4, 0, 42.00, 4,
         "Cabin only. Counted, not consumed."),
        ("Double duvet covers", "Bed linen", "", "Each",
         "Glenbrook Laundry", "GL-DD", 14, 8, 0, 34.00, 6, ""),
        ("Pillowcases", "Bed linen", "", "Each", "Glenbrook Laundry",
         "GL-PC", 38, 24, 0, 7.50, 12, ""),
        ("Bath towels", "Towels", "", "Each", "Glenbrook Laundry",
         "GL-BT", 31, 24, 0, 11.00, 12,
         "Guests take them. Budget for shrinkage."),
        ("Hand towels", "Towels", "", "Each", "Glenbrook Laundry",
         "GL-HT", 22, 16, 0, 5.50, 10, ""),
        ("Tea towels", "Kitchen supplies", "", "Each",
         "Glenbrook Laundry", "GL-TT", 12, 8, 0, 3.20, 8, ""),
        ("Mattress protectors", "Bed linen", "", "Each",
         "Glenbrook Laundry", "GL-MP", 7, 6, 0, 24.00, 4, ""),
        ("Light bulbs", "Batteries & bulbs", "", "Each",
         "Kirkbride Property Services", "KP-LB", 9, 6, 0.3, 3.60, 10,
         ""),
        ("Spare smart lock batteries", "Batteries & bulbs", "",
         "Pack of 4", "Kirkbride Property Services", "KP-BT", 2, 2,
         0.1, 9.80, 4, "One pack per property is the rule."),
        ("Welcome basket - local goods", "Welcome treats", "",
         "Each", "Cedar & Salt", "CS-WB", 5, 4, 1, 11.00, 10,
         "The eleven dollars that keeps buying five-star reviews."),
    ]
    rows: Rows = []
    for (name, category, prop, unit, supplier, sku, on_hand, min_level,
         per_turn, unit_cost, reorder, note) in items:
        row: Dict[str, Any] = {
            "name": name, "category": category, "unit": unit,
            "supplier": supplier, "sku": sku, "on_hand": on_hand,
            "min_level": min_level, "last_count": _day(cfg, -9),
            "last_restock": _day(cfg, -24),
            "per_turn": per_turn, "unit_cost": unit_cost,
            "reorder_qty": reorder,
        }
        if prop:
            row["property"] = prop
        if note:
            row["notes"] = note
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 10. Maintenance Log
# ---------------------------------------------------------------------------

def maintenance(cfg: Config) -> Rows:
    jobs = [
        (-108, "PRP-01", "Plumbing", "Dripping kitchen tap", "Low",
         "Cleaner", "", "VND-04", -104, -103, "Complete", "No",
         80, 94, "KP-2291", "Yes", 365,
         "Small job, slight overrun. Normal."),
        (-96, "PRP-03", "Appliance", "Dishwasher not draining",
         "Normal", "Guest", "BKG-0003", "VND-04", -94, -93,
         "Complete", "No", 150, 150, "KP-2314", "Yes", 180, ""),
        (-72, "PRP-02", "Heating / cooling", "Thermostat replaced", "Normal",
         "Owner inspection", "", "VND-04", -70, -69, "Complete",
         "No", 220, 265, "KP-2402", "Yes", 730, ""),
        (-47, "PRP-02", "Heating / cooling",
         "Boiler failure - no hot water or heating", "Urgent",
         "Guest", "BKG-0011", "VND-05", -46, -41, "Complete", "Yes",
         900, 1180, "HS-88410", "Yes", 365,
         "Six nights blocked, one partial refund and a long "
         "weekend of phone calls. The single most expensive thing "
         "that happened this quarter."),
        (-44, "PRP-01", "Electrical", "Hallway light flickering",
         "Low", "Cleaner", "", "VND-04", -40, -39, "Complete", "No",
         60, 72, "KP-2477", "Yes", 180, ""),
        (-33, "PRP-04", "Furniture",
         "Bed slat broken in the second bedroom", "Normal", "Guest",
         "BKG-0024", "VND-04", -30, -29, "Complete", "No", 90, 118,
         "KP-2511", "Yes", 90, ""),
        (-26, "PRP-04", "Plumbing",
         "Re-seal bath, re-grout tiles - raised in a review",
         "Normal", "Review", "BKG-0010", "VND-04", -14, -12,
         "Complete", "No", 300, 340, "KP-2566", "Yes", 365,
         "Guest mentioned a tired bathroom. Fixed inside a "
         "fortnight, which is the right answer to a review."),
        (-16, "PRP-03", "Locks & access",
         "Smart lock battery warning", "Normal", "Owner inspection",
         "", "VND-04", -6, "", "Scheduled", "No", 40, "", "", "No",
         "", "Batteries on order. Not urgent until it is."),
        (-11, "PRP-02", "Garden / outdoor",
         "Decking boards lifting at the north end", "Normal",
         "Cleaner", "", "VND-04", 4, "", "Scheduled", "No", 480, "",
         "", "No", "",
         "Quoted. Booked for the gap between bookings."),
        (-5, "PRP-01", "Appliance",
         "Washer-dryer making a grinding noise", "High", "Cleaner",
         "", "VND-04", 2, "", "Reported", "No", 150, "", "", "No",
         "", "If it fails, every Harbour turnover needs the "
             "laundry service instead."),
        (-2, "PRP-03", "Plumbing",
         "Shower pressure poor - mentioned twice now", "High",
         "Guest", "BKG-0031", "VND-05", 1, "", "Scheduled", "Yes",
         260, "", "", "No", "",
         "Blocking the calendar on the day of the visit. Two "
         "guests have now mentioned it, so it is a rating "
         "problem rather than a plumbing one."),
        (-1, "PRP-05", "Other", "Snagging list for the new flat",
         "Low", "Owner inspection", "", "VND-04", 7, "", "Reported",
         "No", 400, "", "", "No", "",
         "Pre-launch. Nothing earning here yet."),
    ]
    keys = ("report_date", "property", "category", "description",
            "priority", "reported_by", "booking", "vendor",
            "scheduled", "completed", "status", "blocks_calendar",
            "est_cost", "act_cost", "invoice_ref", "paid",
            "warranty_until", "notes")
    rows: Rows = []
    for job in jobs:
        row: Dict[str, Any] = {}
        for key, value in zip(keys, job):
            if value == "":
                continue
            if key in ("report_date", "scheduled", "completed"):
                value = _day(cfg, value)
            elif key == "warranty_until":
                value = _day(cfg, value)
            row[key] = value
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 11. Reviews Tracker
# ---------------------------------------------------------------------------

def reviews(cfg: Config) -> Rows:
    #  booking, overall, clean, accuracy, checkin, comm, location,
    #  value, responded, public, action
    entries = [
        (1, 5, 5, 5, 5, 5, 5, 5, "Yes",
         "Spotless, easy check-in, great spot by the water.", ""),
        (2, 5, 5, 5, 5, 5, 4, 5, "Yes",
         "Perfect for the children. We will be back.", ""),
        (3, 4, 4, 4, 5, 5, 4, 3, "Yes",
         "Fine for one night, but pricey once the cleaning fee "
         "goes on.",
         "Looked at the one-night pricing - raised the minimum "
         "stay to two."),
        (4, 5, 5, 5, 5, 5, 5, 4, "No",
         "Everything I need for a work trip.", ""),
        (5, 5, 5, 5, 5, 5, 5, 5, "Yes",
         "The dog was made as welcome as we were.", ""),
        (6, 3, 2, 4, 4, 5, 4, 3, "Yes",
         "The floors had not been vacuumed and there was no "
         "coffee. Everything else was lovely, which made it more "
         "annoying.",
         "Replaced the vacuum, added a coffee minimum to the "
         "inventory, and rang Suzanna to apologise."),
        (7, 4, 4, 5, 4, 5, 5, 4, "Yes",
         "Great space for a wedding party. Bathroom is tired.",
         "Booked the re-seal and re-grout the same week."),
        (8, 5, 5, 5, 5, 5, 5, 5, "Yes", "Nothing but five stars.",
         ""),
        (9, 5, 5, 5, 5, 5, 5, 5, "Yes",
         "The welcome basket made our anniversary.", ""),
        (11, 5, 5, 5, 5, 5, 4, 5, "Yes",
         "Second stay and just as good.", ""),
        (12, 5, 5, 5, 5, 5, 5, 5, "No",
         "Booked direct, which was easy. Lovely house.", ""),
        (14, 4, 3, 5, 5, 5, 5, 4, "Yes",
         "Comfortable but not quite as clean as the photos "
         "suggest.",
         "Same week as the broken vacuum. Two reviews, one "
         "cause."),
        (18, 5, 5, 5, 5, 5, 5, 5, "Yes", "Faultless.", ""),
        (19, 4, 5, 4, 4, 4, 5, 4, "No",
         "Good stay. Check-in instructions arrived late.",
         "Moved the access message to two days before arrival."),
        (22, 5, 5, 5, 5, 5, 5, 5, "Yes", "Home from home.", ""),
        (23, 4, 4, 4, 5, 5, 5, 4, "Yes", "Very good value.", ""),
        (24, 5, 5, 5, 5, 5, 5, 4, "Yes",
         "Third stay. Says it all.", ""),
        (26, 5, 5, 5, 5, 5, 4, 5, "Yes",
         "Six weeks and it felt like a home.", ""),
        (27, 5, 5, 5, 5, 5, 5, 5, "Yes", "The dog loved it too.",
         ""),
        (28, 4, 4, 4, 4, 5, 5, 4, "No", "Comfortable and quiet.",
         ""),
        (29, 5, 5, 5, 5, 5, 5, 5, "Yes", "Beautiful spot.", ""),
        (30, 5, 5, 5, 5, 5, 5, 5, "Yes",
         "Booking direct was simple and the house was perfect.",
         ""),
    ]
    rows: Rows = []
    for (number, overall, clean, accuracy, checkin, comm, location,
         value, responded, public, action) in entries:
        row: Dict[str, Any] = {
            "booking": f"BKG-{number:04d}",
            "requested": _day(cfg, -70),
            "status": "Received",
            "review_date": _day(cfg, -68),
            "rate_overall": overall, "rate_clean": clean,
            "rate_accuracy": accuracy, "rate_checkin": checkin,
            "rate_comm": comm, "rate_location": location,
            "rate_value": value,
            "public_review": public,
            "responded": responded,
        }
        if responded == "Yes":
            row["response_date"] = _day(cfg, -66)
            row["response"] = ("Thank you - we are glad you enjoyed "
                               "it." if overall >= 4 else
                               "You are right and I am sorry. Here "
                               "is what we have changed.")
        if action:
            row["action_taken"] = action
        rows.append(row)

    # asked for but not yet left - the chase case
    rows.append({
        "booking": "BKG-0025", "requested": _day(cfg, -19),
        "status": "Requested", "responded": "No",
        "public_review": "",
        "action_taken": "One-night stay, no review. The window is "
                        "closing.",
    })
    rows.append({
        "booking": "BKG-0032", "requested": _day(cfg, -1),
        "status": "Requested", "responded": "No",
        "public_review": "",
    })
    rows.append({
        "booking": "BKG-0013", "status": "No review left",
        "responded": "No", "public_review": "",
        "action_taken": "Guest never reviews. Marked so it stops "
                        "being chased.",
    })
    return rows


# ---------------------------------------------------------------------------
# 12. Occupancy & Performance
# ---------------------------------------------------------------------------

def performance(cfg: Config) -> Rows:
    rows: Rows = []
    notes = {
        (1, 3): "Boiler failure blocked six nights and cost a "
                "partial refund. The month the cabin lost money.",
        (0, 0): "Strongest month so far. Direct bookings doing the "
                "heavy lifting.",
        (2, 2): "Two one-night stays. Cleaning ate the margin.",
    }
    for back in range(4):
        first = cfg.today.replace(day=1)
        month = first.month - back
        year = first.year
        while month <= 0:
            month += 12
            year -= 1
        period = _dt.date(year, month, 1)
        for index in range(4):
            row: Dict[str, Any] = {
                "property": f"PRP-{index + 1:02d}",
                "period": period,
            }
            note = notes.get((back, index + 1))
            if note:
                row["notes"] = note
            rows.append(row)
    return rows


# ---------------------------------------------------------------------------

_BUILDERS: Dict[str, Callable[[Config], Rows]] = {
    "properties": properties,
    "bookings": bookings,
    "guests": guests,
    "arrivals": arrivals,
    "cleaning": cleaning,
    "payouts": payouts,
    "expenses": expenses,
    "supplies": supplies,
    "maintenance": maintenance,
    "reviews": reviews,
    "vendors": vendors,
    "performance": performance,
}


def get(name: str, cfg: Config) -> Rows:
    """Sample rows for one register, or `[]` when there are none."""
    builder = _BUILDERS.get(name)
    if builder is None:
        return []
    return builder(cfg)
