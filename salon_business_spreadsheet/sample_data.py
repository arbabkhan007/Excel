"""Demo data: ten weeks at Gloss & Grace Salon.

A four-chair salon in Austin: an owner who still works the floor, a
commission-only colourist, a stylist, a self-employed nail and lash
tech, an apprentice, a front desk and one chair renter who bills her
own clients. Every tab opens with numbers that hang together - the
diary feeds the till, the till feeds the commissions, the commissions
feed the payroll, and the payroll lands in the profit and loss.

Everything is dated relative to the build date, so the workbook never
looks stale, and `--blank` strips the lot out for your own data.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from .config import COMMISSION_TIERS, Config, rid

SEED = 20261022


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def t(hour: int, minute: int = 0) -> _dt.time:
    return _dt.time(hour, minute)


def money(value: float) -> float:
    return round(value + 1e-9, 2)


def month_key(day: _dt.date) -> str:
    return f"{day.year:04d}-{day.month:02d}"


@dataclass
class Sample:
    services: List[Dict[str, Any]] = field(default_factory=list)
    clients: List[Dict[str, Any]] = field(default_factory=list)
    appointments: List[Dict[str, Any]] = field(default_factory=list)
    sales: List[Dict[str, Any]] = field(default_factory=list)
    payments: List[Dict[str, Any]] = field(default_factory=list)
    staff: List[Dict[str, Any]] = field(default_factory=list)
    commissions: List[Dict[str, Any]] = field(default_factory=list)
    payroll: List[Dict[str, Any]] = field(default_factory=list)
    expenses: List[Dict[str, Any]] = field(default_factory=list)
    inventory: List[Dict[str, Any]] = field(default_factory=list)
    purchases: List[Dict[str, Any]] = field(default_factory=list)
    marketing: List[Dict[str, Any]] = field(default_factory=list)
    loyalty: List[Dict[str, Any]] = field(default_factory=list)
    suppliers: List[Dict[str, Any]] = field(default_factory=list)


_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg) if cfg.include_sample_data else Sample()
    return _CACHE[key]


# ---------------------------------------------------------------------------
# The team.  Between them they use every commission method the
# workbook supports, which is the point: nothing is assumed.
# ---------------------------------------------------------------------------

STAFF: List[Dict[str, Any]] = [
    dict(name="Nadia Ellison", role="Owner & master stylist",
         level="Owner", type="Owner", status="Active",
         phone="512-555-0118", email="nadia@glossandgrace.example",
         hours_week=32, days="Tue, Wed, Thu, Fri, Sat",
         method="None", rate_service=None, rate_retail=None,
         flat_rate=None, hourly=None, rent_week=None,
         target_month=5800,
         notes="Takes the profit rather than a commission - her column "
               "is the salon's."),
    dict(name="Marcus Webb", role="Senior stylist & colourist",
         level="Senior stylist", type="Employed", status="Active",
         phone="512-555-0146", email="marcus@glossandgrace.example",
         hours_week=38, days="Tue, Wed, Thu, Fri, Sat",
         method="Tiered on monthly service revenue", rate_service=None,
         rate_retail=0.10, flat_rate=None, hourly=None, rent_week=None,
         target_month=5200,
         notes="Commission only, on the tiered ladder. A strong month "
               "pays him 50%."),
    dict(name="Priya Raman", role="Stylist", level="Stylist",
         type="Employed", status="Active", phone="512-555-0172",
         email="priya@glossandgrace.example", hours_week=34,
         days="Tue, Wed, Thu, Fri, Sat",
         method="Percent of service and retail", rate_service=0.38,
         rate_retail=0.10, flat_rate=None, hourly=None, rent_week=None,
         target_month=4000,
         notes="Best retail seller in the salon - the 10% is earned."),
    dict(name="Elena Ruiz", role="Nail & lash technician",
         level="Specialist", type="Self-employed", status="Active",
         phone="512-555-0193", email="elena@glossandgrace.example",
         hours_week=30, days="Wed, Thu, Fri, Sat",
         method="Percent after product cost", rate_service=0.45,
         rate_retail=0.10, flat_rate=None, hourly=None, rent_week=None,
         target_month=3000,
         notes="Self-employed, so no employer costs. Her gels and lash "
               "trays come off the top before the 45%."),
    dict(name="Jade Cooper", role="Apprentice stylist",
         level="Apprentice", type="Apprentice", status="Active",
         phone="512-555-0121", email="jade@glossandgrace.example",
         hours_week=30, days="Tue, Wed, Thu, Fri, Sat",
         method="Flat amount per service", rate_service=None,
         rate_retail=0.05, flat_rate=8.00, hourly=13.00, rent_week=None,
         target_month=1400,
         notes="Hourly plus a flat $8 a service while she trains."),
    dict(name="Tom Fenwick", role="Front desk & salon coordinator",
         level="Any stylist", type="Employed", status="Active",
         phone="512-555-0160", email="tom@glossandgrace.example",
         hours_week=24, days="Thu, Fri, Sat",
         method="Hourly only", rate_service=None, rate_retail=0.05,
         flat_rate=None, hourly=15.00, rent_week=None,
         target_month=None,
         notes="Books, rebooks and chases the gaps. Paid by the hour."),
    dict(name="Ruby Shah", role="Stylist, part time",
         level="Stylist", type="Part time", status="Active",
         phone="512-555-0205", email="ruby@glossandgrace.example",
         hours_week=16, days="Fri, Sat",
         method="Percent of service", rate_service=0.35,
         rate_retail=0.08, flat_rate=None, hourly=None, rent_week=None,
         target_month=2200,
         notes="Two long days a week on a straight 35% of what she "
               "takes."),
    dict(name="Sofia Marchetti", role="Barber (chair renter)",
         level="Specialist", type="Chair renter", status="Active",
         phone="512-555-0188", email="sofia@glossandgrace.example",
         hours_week=28, days="Wed, Thu, Fri, Sat",
         method="Rent a chair", rate_service=None, rate_retail=None,
         flat_rate=None, hourly=None, rent_week=180.00,
         target_month=None,
         notes="Rents the fourth chair for $180 a week and bills her "
               "own clients - her bookings are in the diary at zero "
               "so the room is planned properly."),
]

(OWNER, COLOURIST, STYLIST, NAILS, APPRENTICE, DESK,
 PART_TIME, RENTER) = range(8)


# ---------------------------------------------------------------------------
# The menu
# ---------------------------------------------------------------------------
# name, category, minutes, turnaround, price, product cost, level,
# commissionable, add-on, who usually does it
SERVICES: List[Tuple] = [
    ("Cut & finish", "Hair - cutting", 45, 10, 48, 1.20,
     "Stylist", "Yes", "No", [OWNER, COLOURIST, STYLIST]),
    ("Restyle & consultation", "Hair - cutting", 60, 15, 65, 1.50,
     "Senior stylist", "Yes", "No", [OWNER, COLOURIST]),
    ("Dry cut", "Hair - cutting", 30, 5, 35, 0.60,
     "Stylist", "Yes", "No", [STYLIST, COLOURIST]),
    ("Fringe trim", "Hair - cutting", 10, 5, 8, 0.20,
     "Any stylist", "No", "Yes", [APPRENTICE, STYLIST]),
    ("Children's cut", "Hair - cutting", 25, 5, 24, 0.40,
     "Any stylist", "Yes", "No", [APPRENTICE, STYLIST]),
    ("Blow dry", "Hair - styling", 30, 10, 32, 0.90,
     "Any stylist", "Yes", "No", [APPRENTICE, STYLIST, OWNER]),
    ("Blow dry, long or thick", "Hair - styling", 45, 10, 40, 1.30,
     "Any stylist", "Yes", "No", [APPRENTICE, STYLIST]),
    ("Hair up & occasion styling", "Hair - styling", 60, 15, 68, 2.40,
     "Senior stylist", "Yes", "No", [OWNER, STYLIST]),
    ("Root tint", "Hair - colour", 60, 15, 58, 8.50,
     "Stylist", "Yes", "No", [COLOURIST, OWNER, STYLIST]),
    ("Full head colour", "Hair - colour", 90, 15, 88, 14.00,
     "Stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Half head foils", "Hair - colour", 105, 15, 102, 16.50,
     "Senior stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Full head foils", "Hair - colour", 135, 20, 138, 24.00,
     "Senior stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Balayage", "Hair - colour", 150, 20, 168, 28.00,
     "Senior stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Toner & gloss", "Hair - colour", 30, 10, 34, 6.00,
     "Stylist", "Yes", "Yes", [COLOURIST, STYLIST]),
    ("Colour correction", "Hair - colour", 240, 30, 295, 46.00,
     "Master stylist", "Yes", "No", [COLOURIST]),
    ("Bond building treatment", "Hair - treatment", 25, 5, 35, 9.00,
     "Any stylist", "Yes", "Yes", [COLOURIST, STYLIST, OWNER]),
    ("Keratin smoothing", "Hair - treatment", 120, 20, 185, 38.00,
     "Senior stylist", "Yes", "No", [COLOURIST]),
    ("Scalp treatment", "Hair - treatment", 30, 5, 42, 7.00,
     "Any stylist", "Yes", "Yes", [STYLIST, APPRENTICE]),
    ("Skin fade", "Barbering", 35, 5, 30, 0.60,
     "Specialist", "Yes", "No", [RENTER]),
    ("Beard trim & hot towel", "Barbering", 25, 5, 20, 0.80,
     "Specialist", "Yes", "Yes", [RENTER]),
    ("Gel manicure", "Nails - manicure", 45, 10, 40, 4.20,
     "Specialist", "Yes", "No", [NAILS]),
    ("Luxury pedicure", "Nails - pedicure", 60, 10, 52, 5.60,
     "Specialist", "Yes", "No", [NAILS]),
    ("Acrylic full set", "Nails - extensions", 90, 15, 68, 9.50,
     "Specialist", "Yes", "No", [NAILS]),
    ("Acrylic infills", "Nails - extensions", 60, 10, 46, 6.00,
     "Specialist", "Yes", "No", [NAILS]),
    ("Classic lash set", "Lashes", 90, 15, 88, 12.50,
     "Specialist", "Yes", "No", [NAILS]),
    ("Lash infill", "Lashes", 60, 10, 50, 7.00,
     "Specialist", "Yes", "No", [NAILS]),
    ("Lash lift & tint", "Lashes", 45, 10, 58, 8.00,
     "Specialist", "Yes", "No", [NAILS]),
    ("Brow shape & tint", "Brows", 30, 5, 28, 3.00,
     "Any stylist", "Yes", "Yes", [NAILS, APPRENTICE]),
    ("Leg wax", "Waxing", 30, 5, 32, 4.00,
     "Any stylist", "Yes", "No", [NAILS]),
    ("Facial wax", "Waxing", 15, 5, 16, 1.80,
     "Any stylist", "Yes", "Yes", [NAILS, APPRENTICE]),
    ("Express facial", "Facials & skin", 30, 10, 45, 6.50,
     "Specialist", "Yes", "No", [NAILS]),
    ("Signature facial", "Facials & skin", 60, 10, 78, 11.00,
     "Specialist", "Yes", "No", [NAILS]),
    ("Colour & cut package", "Package", 150, 20, 138, 16.00,
     "Senior stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Bridal trial", "Package", 90, 15, 95, 8.00,
     "Master stylist", "Yes", "No", [OWNER]),
]

SERVICE_NOTES = {
    "Colour correction": "Always consult and patch test first. Priced "
                         "after the consultation, never over the phone.",
    "Balayage": "Our signature service and the best earner on the menu.",
    "Fringe trim": "Free to colour clients - kept on the menu so the "
                   "time is still planned.",
    "Keratin smoothing": "Needs the extractor on and the room to "
                         "itself for two hours.",
}

#: service index -> how often it is booked (relative weight)
SERVICE_WEIGHT = [
    18, 5, 6, 7, 5, 14, 7, 3,      # cutting and styling
    11, 8, 7, 6, 6, 5, 1,          # colour
    6, 1, 3,                       # treatments
    9, 5,                          # barbering (chair renter)
    9, 5, 4, 5,                    # nails
    5, 5, 4, 7,                    # lashes and brows
    4, 4,                          # waxing
    3, 3,                          # facials
    3, 1,                          # packages
]


# ---------------------------------------------------------------------------
# The stock room
# ---------------------------------------------------------------------------
# name, brand, category, size, cost, retail, opening, retail?, supplier
PRODUCTS: List[Tuple] = [
    ("Repair shampoo", "Nova Hair", "Retail - hair", "300ml", 9.40,
     22.00, 14, "Yes", "Lumiere Wholesale"),
    ("Repair conditioner", "Nova Hair", "Retail - hair", "300ml", 9.80,
     23.00, 12, "Yes", "Lumiere Wholesale"),
    ("Colour-safe shampoo", "Nova Hair", "Retail - hair", "300ml", 10.20,
     24.00, 11, "Yes", "Lumiere Wholesale"),
    ("Blonde toning shampoo", "Nova Hair", "Retail - hair", "250ml",
     11.00, 26.00, 9, "Yes", "Lumiere Wholesale"),
    ("Leave-in bond serum", "Nova Hair", "Retail - hair", "100ml", 14.50,
     34.00, 8, "Yes", "Lumiere Wholesale"),
    ("Heat protect spray", "Studio Six", "Retail - hair", "200ml", 8.20,
     19.00, 13, "Yes", "Lumiere Wholesale"),
    ("Texture spray", "Studio Six", "Retail - hair", "200ml", 9.00,
     21.00, 10, "Yes", "Lumiere Wholesale"),
    ("Dry shampoo", "Studio Six", "Retail - hair", "200ml", 7.40, 17.00,
     12, "Yes", "Lumiere Wholesale"),
    ("Styling cream", "Studio Six", "Retail - hair", "100ml", 8.60,
     20.00, 7, "Yes", "Lumiere Wholesale"),
    ("Hair oil", "Nova Hair", "Retail - hair", "50ml", 12.00, 28.00, 9,
     "Yes", "Lumiere Wholesale"),
    ("Cuticle oil pen", "Pollen Nails", "Retail - nails", "4ml", 4.20,
     11.00, 16, "Yes", "Pollen Nail Supply"),
    ("Hand cream", "Pollen Nails", "Retail - skin", "75ml", 5.80, 14.00,
     14, "Yes", "Pollen Nail Supply"),
    ("Lash cleanser", "Lume Lash", "Retail - skin", "50ml", 6.40, 16.00,
     11, "Yes", "Lume Lash Co"),
    ("Brow pencil", "Lume Lash", "Retail - accessories", "1g", 5.20,
     13.00, 9, "Yes", "Lume Lash Co"),
    ("Silk scrunchie set", "Gloss & Grace", "Retail - accessories",
     "3 pack", 4.00, 12.00, 18, "Yes", "Austin Salon Supplies"),
    ("Detangle brush", "Studio Six", "Retail - accessories", "each",
     6.00, 16.00, 8, "Yes", "Austin Salon Supplies"),
    ("Permanent colour tube", "Nova Colour", "Backbar - colour", "60ml",
     6.80, None, 48, "No", "Lumiere Wholesale"),
    ("High lift colour tube", "Nova Colour", "Backbar - colour", "60ml",
     7.60, None, 22, "No", "Lumiere Wholesale"),
    ("Developer 20 vol", "Nova Colour", "Backbar - developer", "1L",
     7.20, None, 9, "No", "Lumiere Wholesale"),
    ("Developer 30 vol", "Nova Colour", "Backbar - developer", "1L",
     7.20, None, 6, "No", "Lumiere Wholesale"),
    ("Bleach powder", "Nova Colour", "Backbar - colour", "500g", 19.50,
     None, 7, "No", "Lumiere Wholesale"),
    ("Backbar shampoo", "Nova Hair", "Backbar - shampoo & conditioner",
     "5L", 28.00, None, 4, "No", "Lumiere Wholesale"),
    ("Backbar conditioner", "Nova Hair",
     "Backbar - shampoo & conditioner", "5L", 30.00, None, 3, "No",
     "Lumiere Wholesale"),
    ("Bond builder salon kit", "Nova Hair", "Backbar - treatment", "1L",
     46.00, None, 2, "No", "Lumiere Wholesale"),
    ("Gel polish", "Pollen Nails", "Backbar - nails", "15ml", 7.90,
     None, 26, "No", "Pollen Nail Supply"),
    ("Acrylic powder", "Pollen Nails", "Backbar - nails", "250g", 16.00,
     None, 4, "No", "Pollen Nail Supply"),
    ("Lash tray, classic", "Lume Lash", "Backbar - lashes", "16 rows",
     9.50, None, 12, "No", "Lume Lash Co"),
    ("Lash adhesive", "Lume Lash", "Backbar - lashes", "5ml", 18.00,
     None, 3, "No", "Lume Lash Co"),
    ("Strip wax", "Wax Works", "Backbar - waxing", "800ml", 11.00, None,
     5, "No", "Austin Salon Supplies"),
    ("Foils, pre-cut", "Studio Six", "Disposables", "500 sheets", 14.00,
     None, 6, "No", "Austin Salon Supplies"),
    ("Gloves, nitrile", "Studio Six", "Disposables", "box of 100", 9.00,
     None, 11, "No", "Austin Salon Supplies"),
    ("Towels, black", "Studio Six", "Disposables", "12 pack", 22.00,
     None, 4, "No", "Austin Salon Supplies"),
]

SUPPLIERS: List[Dict[str, Any]] = [
    dict(supplier="Lumiere Wholesale", contact="Dana Pryce",
         phone="512-555-0410", email="orders@lumiere.example",
         account_number="GG-4471", minimum_order=150.00, lead_days=3,
         pays_by="Card on order",
         notes="Colour and retail haircare. Free delivery over $150."),
    dict(supplier="Pollen Nail Supply", contact="Rhea Dawson",
         phone="512-555-0433", email="hello@pollennails.example",
         account_number="PN-2290", minimum_order=80.00, lead_days=2,
         pays_by="Card on order", notes="Gel, acrylic and nail retail."),
    dict(supplier="Lume Lash Co", contact="Bea Tanaka",
         phone="512-555-0456", email="trade@lumelash.example",
         account_number="LL-0917", minimum_order=60.00, lead_days=4,
         pays_by="30 days",
         notes="Lash trays, adhesive, aftercare. Adhesive is "
               "date-stamped - check it in."),
    dict(supplier="Austin Salon Supplies", contact="Will Brady",
         phone="512-555-0478", email="sales@austinsalon.example",
         account_number="ASS-1182", minimum_order=50.00, lead_days=2,
         pays_by="30 days",
         notes="Towels, foils, gloves, sundries. Cheapest on "
               "disposables."),
    dict(supplier="Wax Works", contact="Nina Alvarez",
         phone="512-555-0491", email="orders@waxworks.example",
         account_number="WW-0330", minimum_order=40.00, lead_days=5,
         pays_by="Card on order", notes="Waxing consumables only."),
    dict(supplier="Barker Print", contact="Joe Barker",
         phone="512-555-0502", email="studio@barkerprint.example",
         account_number="BP-7781", minimum_order=0.00, lead_days=7,
         pays_by="Card on order",
         notes="Loyalty cards, price lists, window vinyl."),
]


FIRST_NAMES = [
    "Amara", "Beth", "Carla", "Dana", "Elise", "Farrah", "Gemma",
    "Hana", "Imani", "Jess", "Kiara", "Lena", "Maya", "Nadine",
    "Orla", "Paige", "Quinn", "Rosa", "Sienna", "Tessa", "Ursula",
    "Vera", "Willa", "Xenia", "Yara", "Zoe", "Alice", "Bridget",
    "Chiara", "Delphine", "Erin", "Freya", "Greta", "Heidi", "Isla",
    "Juno", "Kerry", "Lucia", "Marnie", "Nova", "Opal", "Petra",
    "Raina", "Sana", "Thea", "Una", "Vita", "Wren", "Xia", "Yuki",
    "Adam", "Ben", "Caleb", "Dev", "Eli", "Felix", "Gus", "Hugo",
    "Isaac", "Jonah", "Kofi", "Liam", "Mateo", "Noah", "Omar",
    "Pablo", "Reuben", "Sam", "Theo", "Vince",
]

LAST_NAMES = [
    "Alvarez", "Brennan", "Castillo", "Doyle", "Ellison", "Fischer",
    "Garcia", "Hughes", "Iverson", "Jackson", "Keller", "Lombardi",
    "Mehta", "Nolan", "Okafor", "Pereira", "Quigley", "Rossi",
    "Sandoval", "Thorne", "Upton", "Vance", "Whitfield", "Yates",
    "Zimmer", "Barlow", "Chen", "Delgado", "Esposito", "Farrell",
    "Gallagher", "Hartley", "Ingram", "Jenkins", "Kowalski", "Lind",
    "Moreau", "Nakamura", "Osei", "Patel",
]

PREFERENCE_NOTES = [
    "Half head foils, 7/1 and 8/3, toner at the basin. Coffee, oat "
    "milk, no sugar.",
    "Hates a round brush finish - rough dry and texture spray only.",
    "Colour formula on file: 6/0 roots, 30 minutes, no heat.",
    "Books the 9am slot, always early, never stays for a treatment.",
    "Sensitive scalp - cool water, barrier cream at the hairline.",
    "Growing out a fringe. Do not offer to trim it.",
    "Gel on hands, acrylic on toes. Prefers Elena, will wait for her.",
    "Lashes: classic, 9-11mm, C curl. Allergic to the strong adhesive.",
    "Wedding in the spring - trial booked, keeping the length.",
    "Quiet appointment please. No small talk, headphones in.",
]

TAGS = ["Colour client", "Blow dry regular", "Bridal", "Student",
        "Lash client", "Nails", "Barbering", "Treatment lover", ""]


def _client_pool(rng: random.Random, count: int) -> List[str]:
    names: List[str] = []
    used = set()
    while len(names) < count:
        name = f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)}"
        if name in used:
            continue
        used.add(name)
        names.append(name)
    return names


# ---------------------------------------------------------------------------
# The generator
# ---------------------------------------------------------------------------

OPEN_WEEKDAYS = (1, 2, 3, 4, 5)        # Tuesday to Saturday
PER_WEEKDAY = {1: 11, 2: 13, 3: 15, 4: 17, 5: 19}
HISTORY_DAYS = 60
FUTURE_DAYS = 16


def _build(cfg: Config) -> Sample:          # noqa: C901 - demo data
    rng = random.Random(SEED)
    today = cfg.today
    sample = Sample()

    # ------------------------------------------------------------- staff --
    for index, person in enumerate(STAFF):
        row = dict(person)
        row["start_date"] = d(today, -(900 - index * 120))
        sample.staff.append(row)

    staff_id = {index: rid("Staff", index + 1) for index in range(len(STAFF))}

    # ---------------------------------------------------------- services --
    for index, (name, category, minutes, buffer_, price, cost, level,
                commissionable, addon, _who) in enumerate(SERVICES):
        sample.services.append(dict(
            name=name, category=category, duration=minutes,
            buffer=buffer_, price=price, product_cost=cost,
            addon=addon, commissionable=commissionable,
            online="No" if name in ("Colour correction", "Bridal trial")
                   else "Yes",
            level=level, status="Active",
            reviewed=d(today, -rng.randint(40, 330)),
            description="", notes=SERVICE_NOTES.get(name, ""),
        ))
    service_id = {index: rid("Service", index + 1)
                  for index in range(len(SERVICES))}

    # --------------------------------------------------------- inventory --
    product_rows: List[Dict[str, Any]] = []
    for index, (name, brand, category, size, cost, retail, opening,
                is_retail, supplier) in enumerate(PRODUCTS):
        used = 0
        if category.startswith("Backbar") or category == "Disposables":
            used = rng.randint(3, 22)
        product_rows.append(dict(
            name=name, brand=brand, category=category, size=size,
            sku=f"{brand[:2].upper()}-{1000 + index * 7}",
            supplier=supplier, unit_cost=cost, retail_price=retail,
            opening=opening, used=used,
            waste=rng.choice([0, 0, 0, 1]),
            adjust=rng.choice([0, 0, 0, 0, -1, 1]),
            reorder_level=None, reorder_qty=None,
            counted_on=d(today, -rng.randint(3, 48)),
            expires=(d(today, rng.randint(20, 400))
                     if "Developer" in name or "adhesive" in name
                     or "Bleach" in name else None),
            retail_item=is_retail, notes="",
        ))
    sample.inventory = product_rows
    product_id = {index: rid("Product", index + 1)
                  for index in range(len(PRODUCTS))}
    retail_indexes = [i for i, p in enumerate(PRODUCTS) if p[7] == "Yes"]

    sample.suppliers = [dict(row) for row in SUPPLIERS]

    # ----------------------------------------------------------- clients --
    names = _client_pool(rng, 132)
    sources = (["Repeat client"] * 4 + ["Instagram"] * 7 +
               ["Referral"] * 5 + ["Walk-in"] * 4 + ["Google"] * 4 +
               ["Booking site"] * 3 + ["Facebook"] * 2 + ["Website"] * 2 +
               ["TikTok"] + ["Gift certificate"] + ["Promotion"])
    for index, name in enumerate(names):
        joined = d(today, -rng.randint(20, 1400))
        colour_client = rng.random() < 0.45
        sample.clients.append(dict(
            name=name,
            phone=f"512-555-{rng.randint(1000, 9989):04d}",
            email=(f"{name.split()[0].lower()}."
                   f"{name.split()[1].lower()}@example.com"),
            contact=rng.choice(["Text", "Text", "Email", "Phone call",
                                "WhatsApp"]),
            joined=joined,
            source=rng.choice(sources),
            referred_by="",
            staff=staff_id[rng.choice(
                [OWNER, COLOURIST, COLOURIST, STYLIST, STYLIST, NAILS,
                 NAILS, APPRENTICE])],
            birthday=_dt.date(1965 + rng.randint(0, 40),
                              rng.randint(1, 12), rng.randint(1, 28)),
            preferences=(rng.choice(PREFERENCE_NOTES)
                         if rng.random() < 0.45 else ""),
            allergies=("Patch tested - no reaction"
                       if colour_client and rng.random() < 0.3 else ""),
            patch_test=(d(today, -rng.randint(10, 260))
                        if colour_client else None),
            marketing_ok="Yes" if rng.random() < 0.82 else "No",
            photos_ok="Yes" if rng.random() < 0.6 else "No",
            tags=rng.choice(TAGS),
            set_status="Active",
            notes="",
        ))
    client_id = {index: rid("Client", index + 1)
                 for index in range(len(names))}

    # a few referrals, so the loyalty tab has something to pay out on
    for index in range(len(names)):
        if rng.random() < 0.14:
            sender = rng.randrange(len(names))
            if sender != index:
                sample.clients[index]["referred_by"] = client_id[sender]
                sample.clients[index]["source"] = "Referral"

    # ------------------------------------------------------- the campaigns
    campaigns = [
        dict(name="New client, 20% off first visit", channel="Instagram",
             offer="20% off any first service", code="HELLO20",
             start=d(today, -112), end=d(today, 30), spend=180.00,
             target_rev=3000, status="Running",
             notes="Steady trickle. Most of them rebook, which is the "
                   "only reason it pays."),
        dict(name="Refer a friend", channel="Referral scheme",
             offer="$10 credit each when a friend books",
             code="FRIEND10", start=d(today, -240), end=None,
             spend=0.00, target_rev=2400, status="Running",
             notes="Costs nothing up front and brings the best "
                   "clients. Credit tracked on the loyalty tab."),
        dict(name="Tuesday blow dry club", channel="Email",
             offer="$25 blow dry, Tuesdays only", code="TUES25",
             start=d(today, -84), end=d(today, -14), spend=40.00,
             target_rev=900, status="Finished",
             notes="Filled the quiet morning but at a thin margin. "
                   "Worth repeating only in January."),
        dict(name="Balayage season", channel="Instagram",
             offer="Free bond treatment with any balayage",
             code="BONDFREE", start=d(today, -56), end=d(today, 14),
             spend=120.00, target_rev=3600, status="Running",
             notes="The treatment costs $9 and sells a $168 service. "
                   "Best offer we run."),
        dict(name="Lash launch", channel="TikTok",
             offer="$15 off a first classic lash set", code="LASH15",
             start=d(today, -42), end=d(today, -7), spend=95.00,
             target_rev=1200, status="Finished",
             notes="Brought in eight new lash clients. Five have "
                   "rebooked for infills."),
        dict(name="Birthday treat", channel="Email",
             offer="Free bond treatment in your birthday month",
             code="BIRTHDAY", start=d(today, -300), end=None,
             spend=0.00, target_rev=600, status="Running",
             notes="Automatic from the birthday column on the "
                   "clients tab. Costs product, buys goodwill."),
        dict(name="Google local ads", channel="Google Ads",
             offer="No offer - visibility only", code="GOOGLE",
             start=d(today, -70), end=d(today, 20), spend=240.00,
             target_rev=2000, status="Running",
             notes="Expensive per click. Watch the cost per new "
                   "client before renewing."),
        dict(name="Christmas gift vouchers", channel="Window & signage",
             offer="$100 voucher for $90", code="GIFT90",
             start=d(today, 40), end=d(today, 80), spend=60.00,
             target_rev=1500, status="Planned",
             notes="Print the cards in November. Last year sold 22."),
    ]
    sample.marketing = campaigns

    # --------------------------------------------------- the diary & till --
    appointments: List[Dict[str, Any]] = []
    sales: List[Dict[str, Any]] = []
    payments: List[Dict[str, Any]] = []

    service_weight = list(SERVICE_WEIGHT)
    service_indexes = list(range(len(SERVICES)))
    regulars = list(range(0, 86))        # the ones who come back

    # who can do what
    doers: Dict[int, List[int]] = {}
    for index, row in enumerate(SERVICES):
        doers[index] = list(row[9])
        if (row[1].startswith("Hair") and row[6] != "Master stylist"
                and PART_TIME not in doers[index]):
            doers[index].append(PART_TIME)

    def pick_service() -> int:
        return rng.choices(service_indexes, weights=service_weight)[0]

    last_seen: Dict[int, _dt.date] = {}

    for offset in range(-HISTORY_DAYS, FUTURE_DAYS + 1):
        day = d(today, offset)
        if day.weekday() not in OPEN_WEEKDAYS:
            continue
        future = day > today
        planned = PER_WEEKDAY[day.weekday()]
        if future:
            planned = max(3, int(planned * (0.78 if offset <= 7 else 0.45)))
        cursors = {index: 9.0 for index in range(len(STAFF))}
        cursors[NAILS] = 10.0
        cursors[RENTER] = 10.0

        for _slot in range(planned):
            s_index = pick_service()
            who_can = [w for w in doers[s_index] if w != DESK]
            who = rng.choice(who_can)
            minutes = SERVICES[s_index][2] + SERVICES[s_index][3]
            start_hour = cursors[who]
            if start_hour + minutes / 60 > 18.5:
                continue
            cursors[who] = start_hour + minutes / 60 + rng.choice(
                [0, 0, 0.25])
            hour = int(start_hour)
            minute = int(round((start_hour - hour) * 60 / 5) * 5)
            if minute >= 60:
                hour, minute = hour + 1, 0

            # who is in the chair
            if rng.random() < 0.72:
                client_index = rng.choice(regulars)
            else:
                client_index = rng.randrange(len(names))

            price = SERVICES[s_index][4]
            renter = who == RENTER
            if future:
                status = rng.choices(
                    ["Booked", "Confirmed", "Enquiry"],
                    weights=[6, 5, 1])[0]
            else:
                status = rng.choices(
                    ["Completed", "Cancelled", "No show", "Rescheduled"],
                    weights=[88, 6, 3, 3])[0]

            source = (rng.choice(["Repeat client", "Repeat client",
                                  "Repeat client", "Phone", "Instagram",
                                  "Walk-in", "Booking site", "Referral",
                                  "Google", "Website"])
                      if client_index in last_seen
                      else rng.choice(sources))

            appointment_no = len(appointments) + 1
            appointment = dict(
                date=day, start=t(hour, minute),
                client=client_id[client_index],
                service=service_id[s_index],
                staff=staff_id[who],
                minutes=None,
                price=0 if renter else None,
                status=status,
                source=source,
                deposit_due=None,
                notes=("Chair renter - bills her own clients"
                       if renter else ""),
                cancel_reason="",
                fee_charged=None,
                logged=d(day, -rng.randint(1, 24)),
            )
            if status in ("Cancelled", "Rescheduled"):
                appointment["cancel_reason"] = rng.choice(
                    ["Client cancelled", "Client rescheduled",
                     "Client unwell", "Stylist unavailable"])
            if status == "No show":
                appointment["cancel_reason"] = "No show"
                if rng.random() < 0.4:
                    appointment["fee_charged"] = cfg.no_show_fee
            appointments.append(appointment)
            appointment_ref = rid("Appointment", appointment_no)

            if status == "Completed":
                last_seen[client_index] = day
            if status != "Completed" or renter:
                continue

            # ---------------------------------------------- the sale --
            promo = ""
            discount_pct = None
            discount_amt = None
            category = SERVICES[s_index][1]
            offers = ["HELLO20", "FRIEND10"]
            if category == "Hair - styling" and price <= 45:
                offers.append("TUES25")
            if category == "Lashes":
                offers.append("LASH15")
            if category == "Hair - colour" and price >= 85:
                offers.append("BONDFREE")
            if price >= 55:
                offers.append("BIRTHDAY")
            if rng.random() < 0.13:
                promo = rng.choice(offers)
                if promo == "HELLO20":
                    discount_pct = 0.20
                elif promo == "TUES25":
                    discount_amt = money(max(0, price - 25))
                elif promo == "LASH15":
                    discount_amt = 15.00
                elif promo == "FRIEND10":
                    discount_amt = 10.00
                elif promo == "BONDFREE":
                    discount_amt = 35.00
                elif promo == "BIRTHDAY":
                    discount_amt = 20.00
            refund = None
            if rng.random() < 0.006:
                refund = money(price * 0.5)

            sale_no = len(sales) + 1
            sales.append(dict(
                date=day, appointment=appointment_ref,
                client=client_id[client_index], staff=staff_id[who],
                kind="Service", item=service_id[s_index],
                description="", qty=None, unit_price=None,
                discount_pct=discount_pct, discount_amt=discount_amt,
                promo=promo, refund=refund,
                notes=("Refunded - colour did not take, redo booked"
                       if refund else ""),
            ))
            gross = price
            discount = money((gross * (discount_pct or 0)) +
                             (discount_amt or 0))
            net_service = money(gross - discount - (refund or 0))
            sale_ref = rid("Sale", sale_no)
            service_tax = (money(net_service * cfg.tax_rate)
                           if cfg.tax_on_services == "Yes" else 0.0)
            service_total = money(net_service + service_tax)

            # ------------------------------------------- retail with it --
            retail_total = 0.0
            retail_sale_ref: Optional[str] = None
            attach = 0.30 if who == STYLIST else 0.17
            if rng.random() < attach:
                product_index = rng.choice(retail_indexes)
                qty = 1 if rng.random() < 0.85 else 2
                retail_no = len(sales) + 1
                sales.append(dict(
                    date=day, appointment=appointment_ref,
                    client=client_id[client_index], staff=staff_id[who],
                    kind="Retail", item=product_id[product_index],
                    description="", qty=qty if qty > 1 else None,
                    unit_price=None, discount_pct=None,
                    discount_amt=None, promo="", refund=None, notes="",
                ))
                retail_net = money(PRODUCTS[product_index][5] * qty)
                retail_tax = (money(retail_net * cfg.tax_rate)
                              if cfg.tax_on_retail == "Yes" else 0.0)
                retail_total = money(retail_net + retail_tax)
                retail_sale_ref = rid("Sale", retail_no)

            # --------------------------------------------- the money in --
            banked = "No" if (today - day).days <= 4 else "Yes"
            method = rng.choices(
                ["Card", "Card", "Card", "Cash", "Mobile wallet",
                 "Bank transfer"], weights=[40, 20, 12, 14, 10, 4])[0]
            deposit = 0.0
            if price >= cfg.deposit_threshold and rng.random() < 0.75:
                deposit = money(min(price * cfg.deposit_percent,
                                    service_total))
                payments.append(dict(
                    date=d(day, -rng.randint(2, 12)),
                    client=client_id[client_index], kind="Deposit",
                    method="Online deposit", amount=deposit, sale="",
                    appointment=appointment_ref, staff="",
                    reference=f"DEP{appointment_no:04d}", banked="Yes",
                    notes="",
                ))
            owing = money(service_total - deposit)
            unpaid = rng.random() < 0.012
            if not unpaid and owing > 0:
                payments.append(dict(
                    date=day, client=client_id[client_index],
                    kind="Balance" if deposit else "Full payment",
                    method=method, amount=owing, sale=sale_ref,
                    appointment=appointment_ref, staff=staff_id[who],
                    reference="", banked=banked, notes="",
                ))
            elif unpaid:
                appointments[-1]["notes"] = (
                    "Left without paying the balance - invoice sent")
            if retail_sale_ref and not unpaid:
                payments.append(dict(
                    date=day, client=client_id[client_index],
                    kind="Retail only", method=method,
                    amount=retail_total, sale=retail_sale_ref,
                    appointment=appointment_ref, staff=staff_id[who],
                    reference="", banked=banked, notes="",
                ))
            if refund:
                payments.append(dict(
                    date=d(day, rng.randint(1, 4)),
                    client=client_id[client_index], kind="Refund",
                    method=method, amount=refund, sale=sale_ref,
                    appointment=appointment_ref, staff=staff_id[who],
                    reference="", banked="Yes",
                    notes="Goodwill refund, redo booked"))
            if rng.random() < 0.34:
                tip = money(max(2, round(net_service * rng.uniform(
                    0.08, 0.18) * 2) / 2))
                payments.append(dict(
                    date=day, client=client_id[client_index], kind="Tip",
                    method=method, amount=tip, sale="",
                    appointment=appointment_ref, staff=staff_id[who],
                    reference="", banked=banked, notes=""))

    # -------------------------------------- chair rent, vouchers, courses --
    rent_day = today
    while rent_day.weekday() != 4:
        rent_day = d(rent_day, -1)
    for week in range(10):
        day = d(rent_day, -7 * week)
        sale_no = len(sales) + 1
        sales.append(dict(
            date=day, appointment="", client="", staff="", kind="Other",
            item="", description="Chair rent - Sofia Marchetti",
            qty=None, unit_price=cfg.chair_rent_week, discount_pct=None,
            discount_amt=None, promo="", refund=None,
            notes="Weekly chair rent. Income to the salon, not a "
                  "commission-earning sale."))
        payments.append(dict(
            date=day, client="", kind="Full payment",
            method="Bank transfer", amount=money(cfg.chair_rent_week),
            sale=rid("Sale", sale_no), appointment="", staff="",
            reference="RENT", banked="Yes", notes=""))

    for index in range(7):
        day = d(today, -rng.randint(4, 64))
        client_index = rng.randrange(len(names))
        value = rng.choice([50, 75, 100, 100, 150])
        sale_no = len(sales) + 1
        sales.append(dict(
            date=day, appointment="", client=client_id[client_index],
            staff=staff_id[DESK], kind="Gift certificate", item="",
            description=f"Gift certificate, ${value}", qty=None,
            unit_price=float(value), discount_pct=None,
            discount_amt=None, promo="", refund=None,
            notes="Sold at the desk. Counted as income when sold; "
                  "redeem it as a payment method later."))
        payments.append(dict(
            date=day, client=client_id[client_index],
            kind="Gift certificate sale", method="Card",
            amount=float(value), sale=rid("Sale", sale_no),
            appointment="", staff=staff_id[DESK], reference="",
            banked="Yes", notes=""))

    for index in range(4):
        day = d(today, -rng.randint(10, 60))
        client_index = rng.choice(regulars)
        value = rng.choice([180, 240, 300])
        sale_no = len(sales) + 1
        sales.append(dict(
            date=day, appointment="", client=client_id[client_index],
            staff=staff_id[OWNER], kind="Package", item="",
            description="Course of six blow dries", qty=None,
            unit_price=float(value), discount_pct=None,
            discount_amt=None, promo="", refund=None,
            notes="Paid up front, redeemed visit by visit."))
        payments.append(dict(
            date=day, client=client_id[client_index], kind="Package sale",
            method="Card", amount=float(value), sale=rid("Sale", sale_no),
            appointment="", staff=staff_id[OWNER], reference="",
            banked="Yes", notes=""))

    sample.appointments = appointments
    sample.sales = sales
    sample.payments = payments

    # ------------------------------------------------------- commissions --
    months = sorted({month_key(row["date"]) for row in sales})[-3:]
    totals: Dict[Tuple[str, int], Dict[str, float]] = {}
    for row in sales:
        key = month_key(row["date"])
        who = row["staff"]
        if not who:
            continue
        index = int(who.split("-")[1]) - 1
        bucket = totals.setdefault((key, index), dict(
            service=0.0, retail=0.0, cost=0.0, count=0))
        price = 0.0
        cost = 0.0
        if row["kind"] == "Service":
            s_index = int(row["item"].split("-")[1]) - 1
            price = SERVICES[s_index][4]
            cost = SERVICES[s_index][5]
        elif row["kind"] == "Retail":
            p_index = int(row["item"].split("-")[1]) - 1
            qty = row["qty"] or 1
            price = PRODUCTS[p_index][5] * qty
            cost = PRODUCTS[p_index][4] * qty
        else:
            price = row["unit_price"] or 0.0
        discount = money((price * (row["discount_pct"] or 0)) +
                         (row["discount_amt"] or 0))
        net = money(price - discount - (row["refund"] or 0))
        if row["kind"] == "Service":
            bucket["service"] += net
            bucket["cost"] += money(cost)
            bucket["count"] += 1
        elif row["kind"] == "Retail":
            bucket["retail"] += net

    def tier_rate(base: float) -> float:
        rate = COMMISSION_TIERS[0][1]
        for floor, value in COMMISSION_TIERS:
            if base >= floor:
                rate = value
        return rate

    commission_due: Dict[Tuple[str, int], float] = {}
    for key in months:
        for index, person in enumerate(STAFF):
            bucket = totals.get((key, index))
            if not bucket:
                continue
            method = person["method"]
            service = money(bucket["service"])
            retail = money(bucket["retail"])
            cost = money(bucket["cost"])
            rate = person["rate_service"] or cfg.service_commission
            retail_rate = person["rate_retail"] or cfg.retail_commission
            if method == "Percent of service":
                value = money(service * rate)
            elif method == "Percent of service and retail":
                value = money(money(service * rate) +
                              money(retail * retail_rate))
            elif method == "Percent after product cost":
                value = money(max(0.0, service - cost) * rate)
            elif method == "Tiered on monthly service revenue":
                value = money(service * tier_rate(service))
            elif method == "Flat amount per service":
                value = money(bucket["count"] * (person["flat_rate"] or 0))
            else:
                value = 0.0
            if value <= 0 and method in ("None", "Hourly only",
                                         "Rent a chair"):
                continue
            adjust = None
            if rng.random() < 0.12:
                adjust = rng.choice([-25.0, -15.0, 20.0, 35.0])
            current = key == month_key(today)
            paid_day = None
            if not current:
                year, month = (int(part) for part in key.split("-"))
                last = _dt.date(year + (month == 12),
                                1 if month == 12 else month + 1, 1)
                paid_day = d(last, -1)
            commission_due[(key, index)] = money(
                value + (adjust or 0))
            sample.commissions.append(dict(
                month=key, staff=staff_id[index],
                adjust=adjust, bonus=(50.0 if rng.random() < 0.08
                                      else None),
                approved="No" if current else "Yes",
                paid_on=paid_day,
                notes=("Current month, still running"
                       if current else ""),
            ))

    # ----------------------------------------------------------- payroll --
    tips_by: Dict[Tuple[_dt.date, int], float] = {}
    for row in payments:
        if row["kind"] != "Tip" or not row["staff"]:
            continue
        index = int(row["staff"].split("-")[1]) - 1
        week_end = row["date"]
        while week_end.weekday() != 4:
            week_end = d(week_end, 1)
        tips_by[(week_end, index)] = money(
            tips_by.get((week_end, index), 0.0) + row["amount"])

    friday = today
    while friday.weekday() != 4:
        friday = d(friday, -1)
    weekly_people = [(APPRENTICE, 30), (DESK, 24)]
    for week in range(10):
        period_end = d(friday, -7 * week)
        for index, hours in weekly_people:
            person = STAFF[index]
            worked = hours + rng.choice([-4, -2, 0, 0, 2, 3])
            commission = None
            if index == APPRENTICE and period_end.day >= 22:
                key = month_key(period_end)
                commission = commission_due.get((key, index))
            sample.payroll.append(dict(
                period_end=period_end, period="Weekly",
                staff=staff_id[index], hours=worked,
                rate=person["hourly"], commission=commission,
                tips=tips_by.get((period_end, index)),
                bonus=None, holiday=None,
                deductions=(12.0 if rng.random() < 0.08 else None),
                paid="Yes", paid_on=period_end, method="Bank transfer",
                notes="",
            ))

    for key in months:
        year, month = (int(part) for part in key.split("-"))
        last_day = _dt.date(year + (month == 12),
                            1 if month == 12 else month + 1, 1)
        period_end = min(d(last_day, -1), today)
        current = key == month_key(today)
        for index in (COLOURIST, STYLIST, NAILS, PART_TIME):
            value = commission_due.get((key, index))
            if value is None:
                continue
            tips = money(sum(
                row["amount"] for row in payments
                if row["kind"] == "Tip" and row["staff"] == staff_id[index]
                and month_key(row["date"]) == key))
            if current:
                continue        # nothing released yet - no row to file
            sample.payroll.append(dict(
                period_end=period_end, period="Monthly",
                staff=staff_id[index], hours=None, rate=None,
                commission=value,
                tips=tips or None, bonus=None, holiday=None,
                deductions=None, paid="Yes", paid_on=period_end,
                method="Bank transfer",
                notes="Commission released at the end of the month",
            ))

    # ---------------------------------------------------------- expenses --
    recurring = [
        ("Rent", "Juniper Street Holdings", "Salon rent, 118 Juniper St",
         2400.00, "Monthly", "Bank transfer"),
        ("Utilities", "Lone Star Power", "Electricity and water",
         310.00, "Monthly", "Bank transfer"),
        ("Insurance", "Shield Business Cover",
         "Public liability and treatment cover", 95.00, "Monthly",
         "Card"),
        ("Software & subscriptions", "BookEase",
         "Online booking and reminders", 79.00, "Monthly", "Card"),
        ("Telephone & internet", "Austin Fibre", "Salon line and wifi",
         65.00, "Monthly", "Bank transfer"),
        ("Accounting & legal", "Reed & Co Accountants",
         "Monthly bookkeeping", 120.00, "Monthly", "Bank transfer"),
        ("Loan repayment", "Capital Union Bank", "Fit-out loan",
         420.00, "Monthly", "Bank transfer"),
        ("Marketing & advertising", "Meta & Google", "Local ads",
         180.00, "Monthly", "Card"),
        ("Bank charges", "Capital Union Bank", "Business account fee",
         18.00, "Monthly", "Bank transfer"),
    ]
    weekly = [
        ("Laundry", "Fresh Fold Laundry", "Towels and gowns", 38.00),
        ("Cleaning", "Bright Spark Cleaning", "Evening clean, 3 nights",
         55.00),
        ("Other", "Juniper Street Market", "Coffee, milk, biscuits",
         24.00),
    ]
    for months_back in range(3):
        base_month = _dt.date(today.year, today.month, 1)
        for _ in range(months_back):
            base_month = d(base_month, -1).replace(day=1)
        for category, payee, description, amount, frequency, method \
                in recurring:
            day = min(base_month.replace(
                day=min(28, 3 + (hash(category) % 20))), today)
            paid = day <= today
            sample.expenses.append(dict(
                date=day, category=category, payee=payee,
                description=description,
                amount=amount + (rng.choice([0, 0, 0, 8, -6])
                                 if category == "Utilities" else 0),
                tax=None, method=method, recurring="Yes",
                frequency=frequency, due=day,
                paid="Yes" if paid else "No",
                paid_on=day if paid else None,
                reference=f"{category[:3].upper()}-{day:%y%m}",
                notes="",
            ))
    week_start = today
    while week_start.weekday() != 0:
        week_start = d(week_start, -1)
    for week in range(10):
        day = d(week_start, -7 * week)
        for category, payee, description, amount in weekly:
            sample.expenses.append(dict(
                date=day, category=category, payee=payee,
                description=description, amount=amount, tax=None,
                method="Card", recurring="Yes", frequency="Weekly",
                due=day, paid="Yes", paid_on=day,
                reference="", notes="",
            ))
    one_offs = [
        ("Training & education", "Nova Colour Academy",
         "Balayage masterclass, two stylists", 340.00, -38),
        ("Equipment", "Salon Fit Direct", "Replacement backwash tap",
         185.00, -52),
        ("Repairs & maintenance", "Ace Electrical",
         "Lighting repair, colour bar", 120.00, -21),
        ("Licences & memberships", "City of Austin",
         "Salon licence renewal", 240.00, -64),
        ("Card processing fees", "SwiftPay",
         "Card processing, last month", 318.00, -12),
        ("Card processing fees", "SwiftPay",
         "Card processing, month before", 296.00, -42),
        ("Marketing & advertising", "Barker Print",
         "Loyalty cards and price lists", 142.00, -30),
        ("Equipment", "Salon Fit Direct", "Second lash bed", 760.00, -7),
        ("Training & education", "Lume Lash Co",
         "Lash lift certification", 295.00, -25),
        ("Other", "Blooms on Fifth", "Reception flowers, monthly",
         45.00, -9),
    ]
    for category, payee, description, amount, offset in one_offs:
        day = d(today, offset)
        sample.expenses.append(dict(
            date=day, category=category, payee=payee,
            description=description, amount=amount, tax=None,
            method=rng.choice(["Card", "Bank transfer"]),
            recurring="No", frequency="One off", due=day,
            paid="Yes", paid_on=day, reference="", notes="",
        ))
    # two that are still waiting, so the chase list is not empty
    sample.expenses.append(dict(
        date=d(today, -5), category="Training & education",
        payee="Nova Colour Academy",
        description="Autumn colour course, two places",
        amount=486.00, tax=None, method="Bank transfer", recurring="No",
        frequency="One off", due=d(today, 9), paid="No", paid_on=None,
        reference="NCA-8841",
        notes="Booked, invoice due in nine days."))
    sample.expenses.append(dict(
        date=d(today, -16), category="Repairs & maintenance",
        payee="Ace Electrical", description="Emergency callout, dryer",
        amount=95.00, tax=None, method="Bank transfer", recurring="No",
        frequency="One off", due=d(today, -2), paid="No", paid_on=None,
        reference="AE-2214", notes="Chase - invoice is past due."))

    # --------------------------------------------------------- purchases --
    order_plan = [
        (0, 6, "Lumiere Wholesale", -62, "Received"),
        (16, 12, "Lumiere Wholesale", -62, "Received"),
        (18, 4, "Lumiere Wholesale", -62, "Received"),
        (20, 2, "Lumiere Wholesale", -62, "Received"),
        (1, 6, "Lumiere Wholesale", -48, "Received"),
        (4, 6, "Lumiere Wholesale", -48, "Received"),
        (24, 10, "Pollen Nail Supply", -45, "Received"),
        (25, 2, "Pollen Nail Supply", -45, "Received"),
        (26, 6, "Lume Lash Co", -40, "Received"),
        (27, 2, "Lume Lash Co", -40, "Received"),
        (2, 6, "Lumiere Wholesale", -34, "Received"),
        (5, 6, "Lumiere Wholesale", -34, "Received"),
        (29, 4, "Austin Salon Supplies", -30, "Received"),
        (30, 6, "Austin Salon Supplies", -30, "Received"),
        (31, 4, "Austin Salon Supplies", -30, "Received"),
        (17, 6, "Lumiere Wholesale", -24, "Received"),
        (19, 3, "Lumiere Wholesale", -24, "Received"),
        (21, 2, "Lumiere Wholesale", -24, "Received"),
        (3, 6, "Lumiere Wholesale", -18, "Received"),
        (9, 6, "Lumiere Wholesale", -18, "Received"),
        (28, 3, "Austin Salon Supplies", -16, "Received"),
        (10, 12, "Pollen Nail Supply", -12, "Received"),
        (11, 10, "Pollen Nail Supply", -12, "Received"),
        (22, 2, "Lumiere Wholesale", -9, "Part received"),
        (23, 1, "Lumiere Wholesale", -9, "Received"),
        (6, 6, "Lumiere Wholesale", -6, "Received"),
        (7, 6, "Lumiere Wholesale", -6, "Received"),
        (12, 6, "Lume Lash Co", -4, "Ordered"),
        (13, 6, "Lume Lash Co", -4, "Ordered"),
        (14, 12, "Austin Salon Supplies", -2, "Ordered"),
        (8, 6, "Lumiere Wholesale", -1, "Draft"),
    ]
    for p_index, qty, supplier, offset, status in order_plan:
        day = d(today, offset)
        cost = PRODUCTS[p_index][4]
        bumped = rng.random() < 0.18
        received_on = None
        qty_received = None
        if status == "Received":
            received_on = d(day, rng.randint(1, 5))
        elif status == "Part received":
            received_on = d(day, 3)
            qty_received = max(1, qty - 1)
        sample.purchases.append(dict(
            date=day, supplier=supplier, product=product_id[p_index],
            qty=qty,
            unit_cost=money(cost * 1.06) if bumped else None,
            shipping=(0.00 if supplier == "Lumiere Wholesale"
                      else rng.choice([0, 6.50, 9.00])),
            tax=None, status=status,
            expected=d(day, 3), received_on=received_on,
            qty_received=qty_received,
            invoice=(f"INV-{8000 + p_index * 13 + abs(offset)}"
                     if status != "Draft" else ""),
            paid="Yes" if offset < -12 else "No",
            notes=("Price up 6% on the last order" if bumped else ""),
        ))

    # ----------------------------------------------------------- loyalty --
    spend_by_client: Dict[str, float] = {}
    for row in sales:
        if not row["client"]:
            continue
        price = 0.0
        if row["kind"] == "Service":
            price = SERVICES[int(row["item"].split("-")[1]) - 1][4]
        elif row["kind"] == "Retail":
            p_index = int(row["item"].split("-")[1]) - 1
            price = PRODUCTS[p_index][5] * (row["qty"] or 1)
        else:
            price = row["unit_price"] or 0.0
        spend_by_client[row["client"]] = money(
            spend_by_client.get(row["client"], 0.0) + price)
    ranked = sorted(spend_by_client.items(), key=lambda item: -item[1])
    for client_ref, spend in ranked[:46]:
        points = int(spend * cfg.loyalty_points_per)
        redeemed = 0
        if points > cfg.loyalty_reward_at * 1.3 and rng.random() < 0.55:
            redeemed = cfg.loyalty_reward_at
        sample.loyalty.append(dict(
            client=client_ref,
            joined=d(today, -rng.randint(30, 500)),
            card=f"GG{rng.randint(10000, 99999)}",
            bonus_points=rng.choice([None, None, None, 25, 50]),
            redeemed=redeemed or None,
            last_reward=(d(today, -rng.randint(10, 120))
                         if redeemed else None),
            reward_type=rng.choice(["Amount off", "Free add-on",
                                    "Free service", "Percent off"]),
            referral_paid=None,
            active="Yes",
            notes="",
        ))

    # ----------------------------------- make the stock room add up --
    # Opening stock is set last, from what actually moved, so no
    # product on the demo tab finishes the period at a nonsense
    # number.  Three lines are left deliberately short so the reorder
    # alerts have something to say.
    short = {"Lash adhesive": 0, "Bleach powder": 1,
             "Backbar conditioner": 1, "Blonde toning shampoo": 2}
    received_qty: Dict[str, float] = {}
    for row in sample.purchases:
        if row["status"] == "Received":
            qty = row["qty_received"] or row["qty"] or 0
        elif row["status"] == "Part received":
            qty = row["qty_received"] or 0
        else:
            qty = 0
        received_qty[row["product"]] = received_qty.get(
            row["product"], 0) + qty
    sold_qty: Dict[str, float] = {}
    for row in sample.sales:
        if row["kind"] == "Retail":
            sold_qty[row["item"]] = sold_qty.get(row["item"], 0) + (
                row["qty"] or 1)
    for index, row in enumerate(product_rows):
        ref = product_id[index]
        went_out = ((row["used"] or 0) + sold_qty.get(ref, 0) +
                    (row["waste"] or 0) - (row["adjust"] or 0))
        cover = short.get(row["name"], rng.randint(4, 11))
        needed = went_out - received_qty.get(ref, 0) + cover
        row["opening"] = max(0, int(round(max(row["opening"], needed))))

    return sample
