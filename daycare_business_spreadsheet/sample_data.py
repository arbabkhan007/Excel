"""Demo data: ten weeks at Little Acorns Childcare.

A forty-two place setting in Columbus, Ohio.  Four rooms, twelve staff,
twenty families and thirty children - some of them siblings, which is
the whole point.  One family pays by the month, one by the hour, one
is on a government subsidy, one is three weeks late and one left in
March.  There is a child on an adrenaline pen, a staff member whose
first aid certificate runs out in a fortnight, a bumped head that was
reported to the parent in eleven minutes, and a room sitting at ninety
per cent of its licensed places.

Every figure hangs together.  The register feeds the billing, the
billing feeds the ledger, the rota feeds the payroll and the payroll
lands in the profit and loss - so when you change one thing you can
watch the rest of the workbook move.

Everything is dated relative to the build date, so the file never
looks stale, and `--blank` strips the lot out for your own children.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from .config import Config, rid

SEED = 20261031


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def t(hour: int, minute: int = 0) -> _dt.time:
    return _dt.time(hour, minute)


def money(value: float) -> float:
    return round(value + 1e-9, 2)


def month_key(day: _dt.date) -> str:
    return f"{day.year:04d}-{day.month:02d}"


def month_name(day: _dt.date) -> str:
    return day.strftime("%B %Y")


def _month_start(day: _dt.date, back: int = 0) -> _dt.date:
    year, month = day.year, day.month - back
    while month <= 0:
        month += 12
        year -= 1
    return _dt.date(year, month, 1)


def _month_end(first: _dt.date) -> _dt.date:
    if first.month == 12:
        return _dt.date(first.year, 12, 31)
    return _dt.date(first.year, first.month + 1, 1) - _dt.timedelta(days=1)


def _weekdays(start: _dt.date, end: _dt.date) -> List[_dt.date]:
    out: List[_dt.date] = []
    day = start
    while day <= end:
        if day.weekday() < 5:
            out.append(day)
        day += _dt.timedelta(days=1)
    return out


def _months_ago(day: _dt.date, months: int) -> _dt.date:
    """`day` shifted back whole months, clamped to a valid day."""
    year, month = day.year, day.month - months
    while month <= 0:
        month += 12
        year -= 1
    last = _month_end(_dt.date(year, month, 1)).day
    return _dt.date(year, month, min(day.day, last))


@dataclass
class Sample:
    classrooms: List[Dict[str, Any]] = field(default_factory=list)
    children: List[Dict[str, Any]] = field(default_factory=list)
    parents: List[Dict[str, Any]] = field(default_factory=list)
    waitlist: List[Dict[str, Any]] = field(default_factory=list)
    attendance: List[Dict[str, Any]] = field(default_factory=list)
    fees: List[Dict[str, Any]] = field(default_factory=list)
    ledger: List[Dict[str, Any]] = field(default_factory=list)
    daily_reports: List[Dict[str, Any]] = field(default_factory=list)
    meals: List[Dict[str, Any]] = field(default_factory=list)
    medical: List[Dict[str, Any]] = field(default_factory=list)
    incidents: List[Dict[str, Any]] = field(default_factory=list)
    staff: List[Dict[str, Any]] = field(default_factory=list)
    schedule: List[Dict[str, Any]] = field(default_factory=list)
    payroll: List[Dict[str, Any]] = field(default_factory=list)
    expenses: List[Dict[str, Any]] = field(default_factory=list)
    supplies: List[Dict[str, Any]] = field(default_factory=list)
    messages: List[Dict[str, Any]] = field(default_factory=list)
    documents: List[Dict[str, Any]] = field(default_factory=list)


_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg) if cfg.include_sample_data else Sample()
    return _CACHE[key]


# ---------------------------------------------------------------------------
# The rooms.  Capacities add up to the forty-two on the licence.
# ---------------------------------------------------------------------------

ROOMS: List[Dict[str, Any]] = [
    dict(name="Acorns", age_band="Infant", capacity=8, ratio_set=4,
         teacher=2, assistant=6, room_status="Open",
         notes="Babies from twelve weeks. Sleep room off the back, "
               "two cots short of full."),
    dict(name="Saplings", age_band="Young toddler", capacity=10,
         ratio_set=5, teacher=3, assistant=7, room_status="Open",
         notes="Walkers and early talkers. Busiest room in the "
               "building and the one the waitlist is for."),
    dict(name="Oaks", age_band="Toddler", capacity=12, ratio_set=6,
         teacher=4, assistant=8, room_status="Open",
         notes="Two to three year olds. Direct door to the garden."),
    dict(name="Willows", age_band="Preschool", capacity=12,
         ratio_set=10, teacher=5, assistant=9, room_status="Open",
         notes="School readiness, plus the before and after school "
               "children in term time."),
]


# ---------------------------------------------------------------------------
# The families.  Four of them have two children here, one has three -
# which is how the sibling discount and the shared invoice get a real
# workout rather than a demonstration.
# ---------------------------------------------------------------------------

FAMILIES: List[Tuple[Any, ...]] = [
    # surname, parent1, rel1, parent2, rel2, street, language, contact
    ("Okonkwo", "Adaeze Okonkwo", "Mother", "Chidi Okonkwo", "Father",
     "12 Bellflower Lane", "English", "App message"),
    ("Hartley", "Megan Hartley", "Mother", "Tom Hartley", "Father",
     "8 Kestrel Court", "English", "Email"),
    ("Nowak", "Kasia Nowak", "Mother", "", "",
     "214 Mulberry Street, Apt 3", "Polish", "Text message"),
    ("Delgado", "Rosa Delgado", "Mother", "Javier Delgado", "Father",
     "61 Sycamore Drive", "Spanish", "Phone call"),
    ("Brennan", "Claire Brennan", "Mother", "Niall Brennan", "Father",
     "19 Larkspur Way", "English", "App message"),
    ("Ahmed", "Yusra Ahmed", "Mother", "Imran Ahmed", "Father",
     "33 Chestnut Row", "Urdu", "WhatsApp"),
    ("Whitfield", "Jo Whitfield", "Mother", "Sam Whitfield",
     "Step-father", "7 Fieldfare Close", "English", "Email"),
    ("Petrova", "Elena Petrova", "Mother", "", "",
     "102 Hawthorn Avenue", "Russian", "Email"),
    ("Mbeki", "Thandi Mbeki", "Mother", "Sipho Mbeki", "Father",
     "45 Juniper Place", "English", "App message"),
    ("Clarke", "Hannah Clarke", "Mother", "Dan Clarke", "Father",
     "28 Rowan Gardens", "English", "Text message"),
    ("Nguyen", "Linh Nguyen", "Mother", "Minh Nguyen", "Father",
     "5 Alder Terrace", "Vietnamese", "App message"),
    ("O'Sullivan", "Aoife O'Sullivan", "Mother", "", "",
     "77 Primrose Hill", "English", "Phone call"),
    ("Ferreira", "Bianca Ferreira", "Mother", "Paulo Ferreira",
     "Father", "16 Magnolia Street", "Portuguese", "WhatsApp"),
    ("Lindqvist", "Anna Lindqvist", "Mother", "Erik Lindqvist",
     "Father", "90 Birchwood Lane", "English", "Email"),
    ("Grant", "Sara Grant", "Mother", "Michael Grant", "Father",
     "23 Foxglove Road", "English", "App message"),
    ("Bello", "Funmi Bello", "Mother", "Tunde Bello", "Father",
     "4 Bramble Court", "English", "Text message"),
    ("Kowalczyk", "Marta Kowalczyk", "Grandmother", "", "",
     "58 Cedar Walk", "Polish", "Phone call"),
    ("Reyes", "Camila Reyes", "Mother", "Diego Reyes", "Father",
     "131 Poplar Street", "Spanish", "App message"),
    ("Thornton", "Beth Thornton", "Mother", "Craig Thornton",
     "Father", "11 Heather Rise", "English", "Email"),
    ("Achebe", "Ngozi Achebe", "Mother", "", "",
     "66 Willow Grove", "English", "In person at pickup"),
]

_EMERGENCY = [
    ("Grace Mensah", "Grandmother"), ("Peter Hartley", "Grandfather"),
    ("Danuta Nowak", "Grandmother"), ("Luisa Delgado", "Aunt"),
    ("Mary Brennan", "Grandmother"), ("Sofia Ahmed", "Aunt"),
    ("Pauline Reed", "Family friend"), ("Irina Petrova", "Grandmother"),
    ("Nomsa Mbeki", "Aunt"), ("Jean Clarke", "Grandmother"),
    ("Hoa Nguyen", "Grandmother"), ("Dervla Cassidy", "Family friend"),
    ("Rita Ferreira", "Grandmother"), ("Lars Lindqvist", "Grandfather"),
    ("Pam Grant", "Grandmother"), ("Kehinde Bello", "Aunt"),
    ("Jerzy Kowalczyk", "Grandfather"), ("Ana Reyes", "Grandmother"),
    ("Lorna Thornton", "Grandmother"), ("Chika Achebe", "Aunt"),
]

_EMPLOYERS = [
    "Nationwide Children's Hospital", "Ohio State University",
    "Columbus City Schools", "Huntington Bank", "Self employed",
    "Kroger distribution", "COTA", "Grange Insurance",
    "Mount Carmel Health", "Honda of America", "Cardinal Health",
    "Freelance designer", "The Columbus Dispatch", "L Brands",
    "Franklin County", "Safelite", "Abbott Nutrition",
    "OhioHealth", "Bath & Body Works", "Donatos",
]

_DOCTORS = [
    ("Dr Imani Walker, Grove Family Practice", "(614) 555 0310"),
    ("Dr Peter Lang, Willow Road Surgery", "(614) 555 0344"),
    ("Dr Nina Castellano, Mulberry Clinic", "(614) 555 0377"),
    ("Dr Omar Haddad, Eastside Paediatrics", "(614) 555 0392"),
]


# ---------------------------------------------------------------------------
# The children.  `fam` is a 1 based index into FAMILIES, `months` is
# their age today, `days` the weekdays they attend (0 = Monday).
# ---------------------------------------------------------------------------

CHILDREN: List[Dict[str, Any]] = [
    # ---- Acorns, the infant room ------------------------------------
    dict(name="Obiageli Okonkwo", preferred="Obi", fam=1, months=9,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-210, diet="No restrictions"),
    dict(name="Felix Hartley", preferred="Felix", fam=2, months=6,
         gender="Boy", days=(1, 2, 3), schedule="3 days", hours=9.5,
         basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-95, diet="Dairy free"),
    dict(name="Zofia Nowak", preferred="Zosia", fam=3, months=11,
         gender="Girl", days=(0, 1, 2, 3), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-150, diet="No restrictions"),
    dict(name="Mateo Delgado", preferred="Mateo", fam=4, months=13,
         gender="Boy", days=(0, 2, 4), schedule="3 days", hours=8.5,
         basis="Daily rate", status="Active", room="Acorns",
         program="Infant", start=-120, diet="No restrictions"),
    dict(name="Saoirse Brennan", preferred="Saoirse", fam=5, months=8,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.5, basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-70, diet="No restrictions"),
    dict(name="Bilal Ahmed", preferred="Bilal", fam=6, months=10,
         gender="Boy", days=(1, 3), schedule="2 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-55, diet="Halal"),
    # ---- Saplings, young toddlers (the busy room) --------------------
    dict(name="Ivy Whitfield", preferred="Ivy", fam=7, months=18,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active",
         room="Saplings", program="Young toddler", start=-320,
         diet="No restrictions"),
    dict(name="Nikolai Petrov", preferred="Niko", fam=8, months=20,
         gender="Boy", days=(0, 1, 2, 3), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-280, diet="No restrictions"),
    dict(name="Ayanda Mbeki", preferred="Aya", fam=9, months=15,
         gender="Girl", days=(1, 2, 3, 4), schedule="4 days", hours=9.5,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-175, diet="Vegetarian"),
    dict(name="Rory Clarke", preferred="Rory", fam=10, months=22,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Monthly rate", status="Active",
         room="Saplings", program="Young toddler", start=-400,
         diet="Nut free"),
    dict(name="An Nguyen", preferred="An", fam=11, months=17,
         gender="Girl", days=(0, 2, 4), schedule="3 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-140, diet="No restrictions"),
    dict(name="Cillian O'Sullivan", preferred="Cillian", fam=12,
         months=23, gender="Boy", days=(0, 1, 2, 3, 4),
         schedule="Full-time 5 days", hours=9.5, basis="Weekly rate",
         status="Active", room="Saplings", program="Young toddler",
         start=-365, diet="Gluten free"),
    dict(name="Luca Ferreira", preferred="Luca", fam=13, months=19,
         gender="Boy", days=(1, 2, 3), schedule="3 days", hours=8.0,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-110, diet="No restrictions"),
    dict(name="Elsa Lindqvist", preferred="Elsa", fam=14, months=16,
         gender="Girl", days=(0, 1, 3, 4), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-200, diet="No restrictions"),
    dict(name="Noah Grant", preferred="Noah", fam=15, months=21,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active",
         room="Saplings", program="Young toddler", start=-240,
         diet="No restrictions"),
    # ---- Oaks, toddlers ----------------------------------------------
    dict(name="Tobi Bello", preferred="Tobi", fam=16, months=28,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.5, basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-430, diet="No restrictions"),
    dict(name="Amaka Okonkwo", preferred="Amaka", fam=1, months=31,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-560, diet="No restrictions"),
    dict(name="Jakub Kowalczyk", preferred="Kuba", fam=17, months=26,
         gender="Boy", days=(1, 2, 3, 4), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-190, diet="No restrictions"),
    dict(name="Valentina Reyes", preferred="Valen", fam=18, months=33,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.0, basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-300, diet="No restrictions"),
    dict(name="Harvey Thornton", preferred="Harvey", fam=19, months=29,
         gender="Boy", days=(0, 2, 4), schedule="3 days", hours=8.5,
         basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-160, diet="Egg free"),
    dict(name="Chiamaka Achebe", preferred="Chi", fam=20, months=34,
         gender="Girl", days=(0, 1, 2, 3), schedule="4 days", hours=9.5,
         basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-250, diet="No restrictions"),
    dict(name="Seren Hartley", preferred="Seren", fam=2, months=27,
         gender="Girl", days=(1, 2, 3), schedule="3 days", hours=9.5,
         basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-320, diet="Dairy free"),
    dict(name="Oscar Whitfield", preferred="Oscar", fam=7, months=30,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-520, diet="No restrictions"),
    # ---- Willows, preschool and school age ----------------------------
    dict(name="Isla Brennan", preferred="Isla", fam=5, months=44,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.0, basis="Weekly rate", status="Active",
         room="Willows", program="Preschool", start=-700,
         diet="No restrictions"),
    dict(name="Dylan Grant", preferred="Dylan", fam=15, months=41,
         gender="Boy", days=(0, 1, 2, 3), schedule="4 days", hours=8.5,
         basis="Weekly rate", status="Active", room="Willows",
         program="Preschool", start=-480, diet="No restrictions"),
    dict(name="Priya Nguyen", preferred="Priya", fam=11, months=38,
         gender="Girl", days=(1, 2, 3, 4), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Willows",
         program="Preschool", start=-330, diet="Vegetarian"),
    dict(name="Emilia Petrova", preferred="Emilia", fam=8, months=52,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Half day 5 days",
         hours=5.0, basis="Hourly rate", status="Active",
         room="Willows", program="Pre-K", start=-600,
         diet="No restrictions"),
    dict(name="Theo Delgado", preferred="Theo", fam=4, months=47,
         gender="Boy", days=(0, 2, 4), schedule="3 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Willows",
         program="Preschool", start=-410, diet="No restrictions"),
    dict(name="Jonah Okonkwo", preferred="Jonah", fam=1, months=68,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Before & after",
         hours=3.5, basis="Per session", status="Active",
         room="Willows", program="Before & after", start=-430,
         diet="No restrictions"),
    dict(name="Maya Reyes", preferred="Maya", fam=18, months=57,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.5, basis="Weekly rate", status="Active",
         room="Willows", program="Pre-K", start=-640,
         diet="No restrictions"),
    # ---- not currently attending --------------------------------------
    dict(name="Arthur Lindqvist", preferred="Artie", fam=14, months=63,
         gender="Boy", days=(), schedule="Full-time 5 days", hours=9.0,
         basis="Weekly rate", status="Graduated", room="Willows",
         program="Pre-K", start=-900, end=-45,
         leaving="Aged out / started school", diet="No restrictions"),
    dict(name="Kofi Mbeki", preferred="Kofi", fam=9, months=36,
         gender="Boy", days=(), schedule="3 days", hours=9.0,
         basis="Weekly rate", status="Withdrawn", room="Oaks",
         program="Toddler", start=-380, end=-70, leaving="Moved away",
         diet="No restrictions"),
    dict(name="Esme Thornton", preferred="Esme", fam=19, months=7,
         gender="Girl", days=(), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Pending start", room="Acorns",
         program="Infant", start=21, diet="No restrictions"),
    dict(name="Rafael Ferreira", preferred="Rafa", fam=13, months=5,
         gender="Boy", days=(), schedule="3 days", hours=9.0,
         basis="Weekly rate", status="On hold", room="Acorns",
         program="Infant", start=-30, diet="No restrictions"),
]

#: Weekly full-time rate by room, before any pro-rata or discount.
_ROOM_RATE = {"Acorns": 325.0, "Saplings": 305.0, "Oaks": 285.0,
              "Willows": 265.0}


# ---------------------------------------------------------------------------
# The team.  Between them they use every pay type the workbook knows.
# ---------------------------------------------------------------------------

STAFF: List[Dict[str, Any]] = [
    dict(name="Marianne Okafor", role="Owner / Director", room="",
         pay_type="Salary", salary=64000, hourly=0, hours=40,
         status="Active", hire=-3100, qual="EdD Early Childhood",
         shift="Admin"),
    dict(name="Dawn Pritchard", role="Centre manager", room="",
         pay_type="Salary", salary=51000, hourly=0, hours=40,
         status="Active", hire=-1850, qual="BA Early Years, Level 5",
         shift="Full day"),
    dict(name="Grace Lim", role="Room leader", room="Acorns",
         pay_type="Hourly", salary=0, hourly=21.5, hours=38,
         status="Active", hire=-1420, qual="Level 3 Early Years",
         shift="Opening"),
    dict(name="Tasha Boyd", role="Room leader", room="Saplings",
         pay_type="Hourly", salary=0, hourly=20.75, hours=38,
         status="Active", hire=-1090, qual="Level 3 Early Years",
         shift="Early"),
    dict(name="Priyanka Rao", role="Room leader", room="Oaks",
         pay_type="Hourly", salary=0, hourly=20.5, hours=36,
         status="Active", hire=-880, qual="Level 3 Early Years",
         shift="Morning"),
    dict(name="Beth Carmody", role="Lead teacher", room="Willows",
         pay_type="Hourly", salary=0, hourly=22.0, hours=38,
         status="Active", hire=-1660, qual="BSc Child Development",
         shift="Full day"),
    dict(name="Hollie Sanderson", role="Assistant teacher",
         room="Acorns", pay_type="Hourly", salary=0, hourly=17.25,
         hours=34, status="Active", hire=-640,
         qual="Level 2 Early Years", shift="Mid"),
    dict(name="Danny Echeverria", role="Assistant teacher",
         room="Saplings", pay_type="Hourly", salary=0, hourly=17.0,
         hours=32, status="Active", hire=-420,
         qual="Level 2 Early Years", shift="Afternoon"),
    dict(name="Keisha Monroe", role="Assistant teacher", room="Oaks",
         pay_type="Hourly", salary=0, hourly=16.75, hours=30,
         status="Probation", hire=-70, qual="Level 2 in progress",
         shift="Late"),
    dict(name="Aiden Walsh", role="Apprentice / trainee",
         room="Willows", pay_type="Hourly", salary=0, hourly=13.5,
         hours=30, status="Active", hire=-240,
         qual="Apprenticeship, year 1", shift="Closing"),
    dict(name="Rosemary Hicks", role="Cook", room="",
         pay_type="Hourly", salary=0, hourly=18.0, hours=25,
         status="Active", hire=-1300, qual="Level 2 Food Safety",
         shift="Morning"),
    dict(name="Jess Turnbull", role="Substitute / supply", room="",
         pay_type="Daily rate", salary=0, hourly=0, day_rate=132,
         hours=0, status="Casual / bank", hire=-500,
         qual="Level 3 Early Years", shift="Sub cover"),
]

_STAFF_ROOM_ORDER = ["Acorns", "Saplings", "Oaks", "Willows"]


# ---------------------------------------------------------------------------
# Supplies
# ---------------------------------------------------------------------------

SUPPLIES: List[Tuple[Any, ...]] = [
    # name, category, unit, pack, supplier, lead, opening, recv, used,
    # waste, cost, min
    ("Nappies size 3", "Nappies & wipes", "Box of 150", 150,
     "Columbus Care Supply", 3, 9, 6, 11, 0, 38.90, 3),
    ("Nappies size 5", "Nappies & wipes", "Box of 120", 120,
     "Columbus Care Supply", 3, 6, 4, 7, 0, 41.50, 2),
    ("Baby wipes", "Nappies & wipes", "Case of 12", 12,
     "Columbus Care Supply", 3, 14, 8, 17, 1, 22.40, 5),
    ("Nappy sacks", "Nappies & wipes", "Roll pack", 200,
     "Columbus Care Supply", 3, 11, 0, 5, 0, 6.80, 4),
    ("Barrier cream", "Hygiene & bathroom", "Tub 500g", 1,
     "Ohio Nursery Direct", 5, 7, 2, 3, 0, 9.25, 3),
    ("Hand soap refill", "Hygiene & bathroom", "5 litre", 1,
     "Buckeye Janitorial", 4, 5, 3, 6, 0, 17.60, 2),
    ("Paper towels", "Hygiene & bathroom", "Case of 24", 24,
     "Buckeye Janitorial", 4, 8, 6, 12, 0, 29.95, 3),
    ("Toilet rolls", "Hygiene & bathroom", "Pack of 36", 36,
     "Buckeye Janitorial", 4, 6, 4, 8, 0, 24.50, 3),
    ("Disinfectant spray", "Cleaning & sanitising", "750ml", 1,
     "Buckeye Janitorial", 4, 16, 6, 14, 1, 4.95, 8),
    ("Surface sanitiser concentrate", "Cleaning & sanitising",
     "5 litre", 1, "Buckeye Janitorial", 4, 3, 2, 3, 0, 31.00, 2),
    ("Mop heads", "Cleaning & sanitising", "Each", 1,
     "Buckeye Janitorial", 7, 6, 0, 2, 0, 7.20, 3),
    ("Bin liners", "Cleaning & sanitising", "Roll of 50", 50,
     "Buckeye Janitorial", 4, 12, 4, 10, 0, 11.30, 5),
    ("Infant formula stage 1", "Drinks & formula", "800g tin", 1,
     "Ohio Nursery Direct", 5, 5, 4, 7, 0, 26.75, 3),
    ("Whole milk", "Drinks & formula", "Gallon", 1, "Kroger wholesale",
     1, 10, 32, 38, 2, 3.85, 8),
    ("Oat milk", "Drinks & formula", "Litre", 1, "Kroger wholesale",
     1, 6, 8, 11, 0, 2.95, 4),
    ("Porridge oats", "Food & kitchen", "3kg bag", 1,
     "Kroger wholesale", 2, 4, 3, 5, 0, 7.40, 2),
    ("Wholemeal bread", "Food & kitchen", "Loaf", 1,
     "Kroger wholesale", 1, 8, 26, 31, 3, 2.60, 6),
    ("Pasta", "Food & kitchen", "5kg", 1, "Kroger wholesale", 2, 3,
     2, 4, 0, 9.80, 2),
    ("Frozen mixed vegetables", "Food & kitchen", "2.5kg", 1,
     "Kroger wholesale", 2, 7, 6, 10, 0, 8.15, 4),
    ("Gluten free pasta", "Food & kitchen", "500g", 1,
     "Kroger wholesale", 2, 4, 2, 3, 0, 3.90, 3),
    ("Paint, poster, assorted", "Arts & crafts", "Pack of 6", 6,
     "Ohio Nursery Direct", 6, 5, 2, 4, 0, 18.40, 2),
    ("Sugar paper A2", "Arts & crafts", "Ream of 250", 250,
     "Ohio Nursery Direct", 6, 4, 1, 2, 0, 21.00, 2),
    ("Glue sticks", "Arts & crafts", "Pack of 12", 12,
     "Ohio Nursery Direct", 6, 7, 2, 5, 1, 9.60, 3),
    ("Playdough", "Arts & crafts", "Tub", 1, "Ohio Nursery Direct",
     6, 12, 0, 6, 2, 2.20, 6),
    ("First aid refill kit", "First aid", "Kit", 1,
     "Midwest Safety Co", 7, 3, 1, 1, 0, 34.00, 2),
    ("Disposable gloves", "Safety & PPE", "Box of 200", 200,
     "Midwest Safety Co", 5, 9, 6, 13, 0, 12.75, 4),
    ("Cot sheets", "Linen & bedding", "Each", 1, "Ohio Nursery Direct",
     10, 24, 0, 0, 2, 6.40, 12),
    ("Printer paper A4", "Office & stationery", "Ream", 1,
     "Columbus Office Hub", 3, 6, 2, 4, 0, 5.90, 3),
]


# ---------------------------------------------------------------------------
# The builder
# ---------------------------------------------------------------------------

def _build(cfg: Config) -> Sample:
    rng = random.Random(SEED)
    today = cfg.today
    s = Sample()

    # Monday of the current week - every weekly thing hangs off this.
    monday = today - _dt.timedelta(days=today.weekday())
    history_start = monday - _dt.timedelta(weeks=10)
    all_days = _weekdays(history_start, today)

    months = [_month_start(today, back) for back in (3, 2, 1, 0)]

    # ------------------------------------------------------- classrooms --
    staff_id = {i + 1: rid("Staff", i + 1) for i in range(len(STAFF))}
    for room in ROOMS:
        s.classrooms.append(dict(
            name=room["name"], age_band=room["age_band"],
            capacity=room["capacity"], ratio_set=room["ratio_set"],
            teacher=staff_id[room["teacher"]],
            assistant=staff_id[room["assistant"]],
            room_status=room["room_status"], notes=room["notes"]))

    # ---------------------------------------------------------- families --
    for index, fam in enumerate(FAMILIES, start=1):
        (surname, p1, rel1, p2, rel2, street, language, contact) = fam
        em_name, em_rel = _EMERGENCY[index - 1]
        doctor, doc_phone = _DOCTORS[index % len(_DOCTORS)]
        pickup = p1 + (f", {p2}" if p2 else "") + f", {em_name}"
        s.parents.append(dict(
            family_name=f"The {surname} family",
            parent1=p1, relation1=rel1,
            phone1=f"(614) 555 {1000 + index * 7:04d}",
            email1=f"{p1.split()[0].lower()}.{surname.lower()}"
                   f"@example.com".replace("'", ""),
            employer1=_EMPLOYERS[index - 1],
            work_phone1=f"(614) 555 {2000 + index * 3:04d}",
            parent2=p2, relation2=rel2,
            phone2=(f"(614) 555 {3000 + index * 5:04d}" if p2 else ""),
            email2=(f"{p2.split()[0].lower()}.{surname.lower()}"
                    f"@example.com".replace("'", "") if p2 else ""),
            address=street, city="Columbus",
            postcode=f"432{index % 10}{(index * 3) % 10}",
            contact_by=contact, language=language,
            em1_name=em_name,
            em1_phone=f"(614) 555 {4000 + index * 11:04d}",
            em1_relation=em_rel,
            em2_name=(_EMERGENCY[(index + 4) % len(_EMERGENCY)][0]
                      if index % 3 else ""),
            em2_phone=(f"(614) 555 {5000 + index * 13:04d}"
                       if index % 3 else ""),
            doctor=doctor, doctor_phone=doc_phone,
            pickup_list=pickup,
            pickup_ok="Yes" if index != 17 else "No",
            custody_notes=("Father has alternate weekends. Court order "
                           "on file - collection by mother only on "
                           "weekdays."
                           if index == 12 else ""),
            joined=_months_ago(today, 24 - (index % 20)),
            notes=("Grandmother is the main contact and does all "
                   "pickups." if index == 17 else
                   "Second child starting in the new year."
                   if index == 2 else "")))

    family_id = {i: rid("Family", i) for i in range(1, len(FAMILIES) + 1)}

    # ---------------------------------------------------------- children --
    seen_family: Dict[int, int] = {}
    kids: List[Dict[str, Any]] = []
    for index, kid in enumerate(CHILDREN, start=1):
        fam = kid["fam"]
        seen_family[fam] = seen_family.get(fam, 0) + 1
        seq = seen_family[fam]
        dob = _months_ago(today, kid["months"]) - _dt.timedelta(
            days=(index * 3) % 25)
        start = d(today, kid["start"])
        enrolled = start - _dt.timedelta(days=21 + (index % 30))
        days = kid["days"]
        per_week = len(days) if days else {
            "Full-time 5 days": 5, "4 days": 4, "3 days": 3,
            "2 days": 2, "Half day 5 days": 5,
            "Before & after": 5}.get(kid["schedule"], 3)
        base = _ROOM_RATE.get(kid["room"], 265.0)
        weekly = base * per_week / 5.0
        if kid["basis"] == "Weekly rate":
            rate = round(weekly / 5.0) * 5.0
        elif kid["basis"] == "Monthly rate":
            rate = round(weekly * 52 / 12 / 5) * 5.0
        elif kid["basis"] == "Daily rate":
            rate = round(base / 5.0 / 5) * 5.0
        elif kid["basis"] == "Hourly rate":
            rate = 9.50
        elif kid["basis"] == "Per session":
            rate = 14.00
        else:
            rate = round(weekly / 5.0) * 5.0
        usual_in = t(7, 30) if per_week >= 4 else t(8, 15)
        usual_out = t(17, 30) if kid["hours"] >= 9 else t(13, 0)
        if kid["schedule"] == "Before & after":
            usual_in, usual_out = t(7, 15), t(17, 45)
        parent = FAMILIES[fam - 1][1]
        partner = FAMILIES[fam - 1][3]
        pickup = parent + (f", {partner}" if partner else "") + \
            f", {_EMERGENCY[fam - 1][0]}"
        record = dict(
            name=kid["name"], preferred=kid["preferred"], dob=dob,
            gender=kid["gender"], family=family_id[fam],
            program=kid["program"], room=kid["room"],
            schedule=kid["schedule"], days_week=float(per_week),
            hours_day=kid["hours"], usual_in=usual_in,
            usual_out=usual_out, enrolled_date=enrolled,
            start_date=start,
            end_date=(d(today, kid["end"]) if kid.get("end") else None),
            leaving_reason=kid.get("leaving", ""),
            status=kid["status"], basis=kid["basis"], rate=money(rate),
            discount_pct=(0.1 if index == 20 else 0),
            subsidy_week=(95.0 if index in (3, 17) else
                          120.0 if index == 11 else 0),
            diet=kid["diet"],
            medical_notes="", pickup_list=pickup,
            pickup_password=f"{kid['preferred'].lower()}"
                            f"{(index * 7) % 90 + 10}",
            photo_ok="Yes" if index not in (12, 26) else "No",
            sunscreen_ok="Yes" if index != 19 else "No",
            outings_ok="Yes" if index not in (4, 26) else "No",
            contract="Signed" if index not in (31, 32) else "Requested",
            notes="")
        # a few care notes, so the register is not uniformly bland
        if index == 2:
            record["medical_notes"] = ("Dairy free - soya formula "
                                       "supplied by parent, labelled.")
        if index == 10:
            record["medical_notes"] = ("Severe nut allergy. Adrenaline "
                                       "pen in the red box by the "
                                       "Saplings door. Allergy plan on "
                                       "file.")
        if index == 12:
            record["medical_notes"] = "Coeliac - strictly gluten free."
        if index == 22:
            record["medical_notes"] = ("Mild asthma. Blue inhaler and "
                                       "spacer in room bag.")
        kids.append(record)
        record["_seq"] = seq
        record["_fam"] = fam
        record["_days"] = days
        record["_weekly"] = money(weekly)
    for item in kids:
        s.children.append({k: v for k, v in item.items()
                           if not k.startswith("_")})

    child_id = {i + 1: rid("Child", i + 1) for i in range(len(kids))}
    for index, item in enumerate(kids, start=1):
        item["_id"] = child_id[index]

    active = [k for k in kids if k["status"] == "Active"]
    by_id = {k["_id"]: k for k in kids}

    # ---------------------------------------------------------- waitlist --
    wl_names = [
        ("Nora Bianchi", "Giulia Bianchi", "Saplings", "Young toddler"),
        ("Caleb Osei", "Abena Osei", "Acorns", "Infant"),
        ("Robin Fraser", "Mhairi Fraser", "Saplings", "Young toddler"),
        ("Amira Haddad", "Layla Haddad", "Acorns", "Infant"),
        ("Jude Whitfield", "Jo Whitfield", "Acorns", "Infant"),
        ("Sofia Marchetti", "Elisa Marchetti", "Oaks", "Toddler"),
        ("Leo Kaminski", "Ola Kaminski", "Saplings", "Young toddler"),
        ("Zara Patel", "Nisha Patel", "Willows", "Preschool"),
        ("Henry Boateng", "Akosua Boateng", "Acorns", "Infant"),
        ("Mila Novak", "Petra Novak", "Saplings", "Young toddler"),
        ("Ethan Grant", "Sara Grant", "Acorns", "Infant"),
        ("Aarav Shah", "Devi Shah", "Oaks", "Toddler"),
        ("Freya Donnelly", "Roisin Donnelly", "Saplings",
         "Young toddler"),
        ("Malik Johnson", "Tanya Johnson", "Willows", "Preschool"),
        ("Clara Weiss", "Ingrid Weiss", "Oaks", "Toddler"),
        ("Rosa Delgado jr", "Rosa Delgado", "Acorns", "Infant"),
    ]
    wl_status = ["Enrolled", "Offer accepted", "Place offered", "Toured",
                 "Tour booked", "Information sent", "New enquiry",
                 "Waiting - no space", "Waiting - future date",
                 "Offer declined", "Lost"]
    sources = ["Parent referral", "Google search", "Sibling already here",
               "Word of mouth", "Facebook", "Sign outside",
               "Google Maps", "Walk in", "Employer scheme",
               "Returning family", "Local directory"]
    for index, (child_name, parent, room, program) in enumerate(
            wl_names, start=1):
        enquiry = d(today, -(130 - index * 7))
        status = wl_status[index % len(wl_status)]
        priority = ("Sibling" if "Whitfield" in child_name
                    or "Grant" in child_name or "Delgado" in child_name
                    else "Staff child" if index == 6
                    else "Returning family" if index == 10
                    else "Low" if index == 14 else "Standard")
        sib = ""
        if priority == "Sibling":
            for item in kids:
                if item["name"].split()[-1] == child_name.split()[-1]:
                    sib = item["_id"]
                    break
        toured = status in ("Toured", "Place offered", "Offer accepted",
                            "Enrolled", "Offer declined")
        offered = status in ("Place offered", "Offer accepted",
                             "Enrolled", "Offer declined")
        s.waitlist.append(dict(
            enquiry_date=enquiry, child_name=child_name,
            dob=_months_ago(today, 4 + (index * 5) % 40),
            parent_name=parent,
            phone=f"(614) 555 {6000 + index * 9:04d}",
            email=f"{parent.split()[0].lower()}@example.com",
            source=sources[index % len(sources)], program=program,
            room=room, days_wanted=float(3 + index % 3),
            start_wanted=d(today, -20 + index * 11),
            priority=priority, sibling=sib, status=status,
            info_sent=d(enquiry, 1),
            tour_date=(d(enquiry, 9) if toured else None),
            offered_date=(d(enquiry, 20) if offered else None),
            response_date=(d(enquiry, 24)
                           if status in ("Offer accepted", "Enrolled",
                                         "Offer declined") else None),
            enrolled_child=("" if status != "Enrolled"
                            else kids[index % len(active)]["_id"]),
            followed_up=d(enquiry, 30 + index),
            notes=("Wants the same days as her brother."
                   if priority == "Sibling" else
                   "Needs a September start, happy to wait."
                   if status == "Waiting - future date" else
                   "Chose a setting closer to work."
                   if status in ("Lost", "Offer declined") else "")))

    # -------------------------------------------------------- attendance --
    floor_staff = [staff_id[i] for i in (3, 4, 5, 6, 7, 8, 9, 10)]
    closed_day = None
    for day in reversed(all_days):
        if day.weekday() == 0 and (today - day).days > 20:
            closed_day = day
            break

    reasons_sick = ["Temperature overnight", "Chickenpox",
                    "Sickness bug - 48 hour rule", "Heavy cold",
                    "Ear infection", "Hand, foot and mouth"]
    reasons_family = ["Family holiday", "Grandparents visiting",
                      "Wedding", "Moving house", "Parent day off"]

    for day in all_days:
        if closed_day and day == closed_day:
            for kid in active:
                s.attendance.append(dict(
                    date=day, child=kid["_id"], status="Centre closed",
                    in_time=None, out_time=None, fee_raised="No",
                    dropped_by="", collected_by="",
                    reason="Staff training day - notified in August",
                    recorded_by=staff_id[2], temp_check="",
                    notes=""))
            continue
        for kid in active:
            if day.weekday() not in kid["_days"]:
                continue
            if day < kid["start_date"]:
                continue
            roll = rng.random()
            status = "Present"
            if roll > 0.975:
                status = "Absent - sick"
            elif roll > 0.955:
                status = "Absent - family"
            elif roll > 0.948:
                status = "Absent - unexplained"
            elif roll > 0.935:
                status = "Holiday - child"
            elif roll > 0.885:
                status = "Late"
            elif roll > 0.862:
                status = "Early pickup"
            present = status in ("Present", "Late", "Early pickup")
            in_t = out_t = None
            dropped = collected = ""
            parent = FAMILIES[kid["_fam"] - 1][1]
            partner = FAMILIES[kid["_fam"] - 1][3]
            if present:
                base_in = kid["usual_in"]
                jitter = rng.choice([-10, -5, 0, 0, 5, 8])
                if status == "Late":
                    jitter = rng.choice([35, 42, 55, 70])
                minutes = base_in.hour * 60 + base_in.minute + jitter
                in_t = t(minutes // 60, minutes % 60)
                base_out = kid["usual_out"]
                jout = rng.choice([-8, -3, 0, 0, 4, 9, 22])
                if status == "Early pickup":
                    jout = -rng.choice([75, 95, 120])
                mo = base_out.hour * 60 + base_out.minute + jout
                out_t = t(mo // 60, mo % 60)
                dropped = parent if rng.random() < 0.7 else (
                    partner or _EMERGENCY[kid["_fam"] - 1][0])
                collected = partner or parent
                if rng.random() < 0.12:
                    collected = _EMERGENCY[kid["_fam"] - 1][0]
            reason = ""
            if status == "Absent - sick":
                reason = rng.choice(reasons_sick)
            elif status == "Absent - family":
                reason = rng.choice(reasons_family)
            elif status == "Holiday - child":
                reason = "Booked holiday - two weeks' notice given"
            s.attendance.append(dict(
                date=day, child=kid["_id"], status=status,
                in_time=in_t, out_time=out_t,
                fee_raised="No" if status == "Holiday - child" else "Yes",
                dropped_by=dropped, collected_by=collected,
                reason=reason, recorded_by=rng.choice(floor_staff),
                temp_check=("36.6" if present and rng.random() < 0.4
                            else ""),
                notes=""))

    # a handful of deliberate loose ends on the most recent days
    recent = [r for r in s.attendance
              if r["date"] == all_days[-1] and r["status"] == "Present"]
    for row in recent[:2]:
        row["out_time"] = None
        row["notes"] = "Still here at close - check the book."
    if len(recent) > 4:
        recent[4]["collected_by"] = "Neighbour - Mrs Dunn"
        recent[4]["notes"] = ("Mum phoned at 2pm to authorise. Added to "
                              "the list the next morning.")

    # -------------------------------------------------- charges & ledger --
    # Work out every charge first, group it by family and month, then
    # raise one invoice per family per month and pay most of them.
    charges: List[Dict[str, Any]] = []

    def weeks_in(first: _dt.date, last: _dt.date) -> float:
        return round(len(_weekdays(first, last)) / 5.0, 2)

    for first in months:
        last = min(_month_end(first), today)
        label = month_name(first)
        span = weeks_in(first, last)
        for kid in active:
            if kid["start_date"] > last:
                continue
            basis = kid["basis"]
            if basis == "Monthly rate":
                qty, rate = 1.0, kid["rate"]
            elif basis == "Daily rate":
                qty = float(len([x for x in _weekdays(first, last)
                                 if x.weekday() in kid["_days"]]))
                rate = kid["rate"]
            elif basis == "Hourly rate":
                qty = round(len([x for x in _weekdays(first, last)
                                 if x.weekday() in kid["_days"]])
                            * kid["hours_day"], 1)
                rate = kid["rate"]
            elif basis == "Per session":
                qty = float(len([x for x in _weekdays(first, last)
                                 if x.weekday() in kid["_days"]]) * 2)
                rate = kid["rate"]
            else:
                qty, rate = span, kid["rate"]
            charges.append(dict(
                date=first, child=kid["_id"], family_set="",
                fee_type="Tuition",
                description=f"Childcare - {label}",
                period_from=first, period_to=last, qty=qty,
                rate=money(rate), discount_pct=0,
                auto_source="Contracted schedule", notes="",
                _fam=kid["_fam"], _month=first))
            # sibling discount on the second and later child
            if kid["_seq"] > 1:
                charges.append(dict(
                    date=first, child=kid["_id"], family_set="",
                    fee_type="Sibling discount",
                    description=f"Sibling discount - {label}",
                    period_from=first, period_to=last, qty=1.0,
                    rate=money(qty * rate * 0.1), discount_pct=0,
                    auto_source="Contracted schedule", notes="",
                    _fam=kid["_fam"], _month=first))
            if kid["subsidy_week"]:
                charges.append(dict(
                    date=first, child=kid["_id"], family_set="",
                    fee_type="Subsidy credit",
                    description=f"Publicly funded hours - {label}",
                    period_from=first, period_to=last, qty=span,
                    rate=money(kid["subsidy_week"]), discount_pct=0,
                    auto_source="Manual adjustment",
                    notes="Paid direct by the county scheme.",
                    _fam=kid["_fam"], _month=first))

    # extras: late pickups, meals, trips, registration
    late_rows = [r for r in s.attendance
                 if r["out_time"] and r["status"] in
                 ("Present", "Late", "Early pickup")]
    rng.shuffle(late_rows)
    for row in late_rows[:14]:
        kid = by_id[row["child"]]
        first = _month_start(row["date"])
        charges.append(dict(
            date=row["date"], child=kid["_id"], family_set="",
            fee_type="Late pickup fee",
            description="Collected after the contracted time",
            period_from=row["date"], period_to=row["date"], qty=1.0,
            rate=15.0, discount_pct=0, auto_source="Late pickup",
            notes="", _fam=kid["_fam"], _month=first))
    for index, kid in enumerate(active):
        if index % 3:
            continue
        first = months[-2]
        charges.append(dict(
            date=d(first, 9), child=kid["_id"], family_set="",
            fee_type="Field trip",
            description="Farm visit - coach and entry",
            period_from=d(first, 9), period_to=d(first, 9), qty=1.0,
            rate=12.0, discount_pct=0, auto_source="Activity or trip",
            notes="", _fam=kid["_fam"], _month=first))
    for kid in active[:6]:
        first = months[-1]
        charges.append(dict(
            date=d(first, 2), child=kid["_id"], family_set="",
            fee_type="Meal fee",
            description="Hot lunches - extra days",
            period_from=first, period_to=_month_end(first), qty=4.0,
            rate=4.5, discount_pct=0, auto_source="Meals", notes="",
            _fam=kid["_fam"], _month=first))
    for kid in kids:
        if kid["status"] != "Pending start":
            continue
        first = months[-1]
        charges.append(dict(
            date=d(today, -9), child=kid["_id"], family_set="",
            fee_type="Registration fee",
            description="One-off registration",
            period_from=d(today, -9), period_to=d(today, -9), qty=1.0,
            rate=75.0, discount_pct=0, auto_source="Registration",
            notes="Non refundable.", _fam=kid["_fam"], _month=first))

    # group by (family, month) and raise the invoices
    groups: Dict[Tuple[int, _dt.date], List[Dict[str, Any]]] = {}
    for charge in charges:
        groups.setdefault((charge["_fam"], charge["_month"]),
                          []).append(charge)

    def net_of(charge: Dict[str, Any]) -> float:
        gross = charge["qty"] * charge["rate"]
        credit = charge["fee_type"] in (
            "Sibling discount", "Staff discount", "Promotional discount",
            "Subsidy credit", "Scholarship credit", "Goodwill credit")
        return money(-gross if credit else gross)

    ledger: List[Dict[str, Any]] = []
    invoice_for: Dict[Tuple[int, _dt.date], str] = {}

    def add_ledger(row: Dict[str, Any]) -> str:
        ledger.append(row)
        kind = ("Receipt" if row["type"] in ("Payment", "Refund")
                else "Invoice")
        return rid(kind, len(ledger))

    for first in months:
        label = month_name(first)
        for fam in range(1, len(FAMILIES) + 1):
            lines = groups.get((fam, first))
            if not lines:
                continue
            issue = d(first, 0)
            if first == months[-1]:
                issue = min(d(first, 0), today)
            ref = add_ledger(dict(
                date=issue, type="Invoice", family=family_id[fam],
                child="", period=label,
                description=f"Childcare fees - {label}", amount=None,
                against="", method="", reference="",
                terms=cfg.payment_terms_days,
                sent="Yes", notes=""))
            invoice_for[(fam, first)] = ref
        # payments, once the invoices for the month exist
        for fam in range(1, len(FAMILIES) + 1):
            ref = invoice_for.get((fam, first))
            if not ref:
                continue
            total = money(sum(net_of(c) for c in groups[(fam, first)]))
            if total <= 0:
                continue
            is_current = first == months[-1]
            chase = fam in (3, 11, 17)
            if is_current and chase:
                continue                      # the three that are late
            if is_current and fam % 4 == 0:
                part = money(total * 0.5)
                add_ledger(dict(
                    date=d(first, 6), type="Payment",
                    family=family_id[fam], child="", period=label,
                    description="Part payment", amount=part,
                    against=ref, method="Card",
                    reference=f"CRD{fam:03d}{first.month:02d}", terms=0,
                    sent="Yes", notes="Second half agreed for the "
                                      "28th."))
                continue
            method = ("Direct debit" if fam % 3 == 0 else
                      "Bank transfer" if fam % 3 == 1 else "Card")
            if fam == 5:
                method = "Childcare voucher"
            if fam == 9:
                method = "Government subsidy"
            add_ledger(dict(
                date=d(first, 4 + (fam % 9)), type="Payment",
                family=family_id[fam], child="", period=label,
                description=f"Payment received - {label}",
                amount=total, against=ref, method=method,
                reference=f"{method[:3].upper()}{fam:03d}"
                          f"{first.month:02d}",
                terms=0, sent="Yes", notes=""))

    # a deposit, a credit note, a refund and a write-off
    pending = [k for k in kids if k["status"] == "Pending start"]
    if pending:
        add_ledger(dict(
            date=d(today, -8), type="Payment",
            family=pending[0]["family"], child=pending[0]["_id"],
            period="Deposit", description="Enrollment deposit",
            amount=150.0, against="", method="Bank transfer",
            reference="DEP0041", terms=0, sent="Yes",
            notes="Held against the first invoice in January."))
    first_prev = months[-2]
    ref_prev = invoice_for.get((2, first_prev))
    if ref_prev:
        add_ledger(dict(
            date=d(first_prev, 19), type="Credit note",
            family=family_id[2], child="", period=month_name(first_prev),
            description="Two days charged in error",
            amount=118.0, against=ref_prev, method="", reference="",
            terms=0, sent="Yes",
            notes="Seren was away on the agreed closure days."))
    ref_old = invoice_for.get((14, months[0]))
    if ref_old:
        add_ledger(dict(
            date=d(months[1], 3), type="Refund",
            family=family_id[14], child="", period=month_name(months[0]),
            description="Overpayment returned",
            amount=96.0, against=ref_old, method="Bank transfer",
            reference="RFD0007", terms=0, sent="Yes", notes=""))
    ref_gone = invoice_for.get((9, months[0]))
    if ref_gone:
        add_ledger(dict(
            date=d(months[1], 14), type="Write-off",
            family=family_id[9], child="", period=month_name(months[0]),
            description="Family moved away - unrecoverable",
            amount=142.5, against=ref_gone, method="", reference="",
            terms=0, sent="Yes",
            notes="Three letters, no forwarding address."))

    s.ledger = [{k: v for k, v in row.items() if not k.startswith("_")}
                for row in ledger]

    for charge in charges:
        charge["invoice"] = invoice_for.get(
            (charge["_fam"], charge["_month"]), "")
    s.fees = [{k: v for k, v in c.items() if not k.startswith("_")}
              for c in charges]

    # ------------------------------------------------------ daily reports --
    report_days = all_days[-10:]
    report_rooms = ("Acorns", "Saplings", "Oaks")
    present_key = {(r["child"], r["date"]) for r in s.attendance
                   if r["status"] in ("Present", "Late", "Early pickup")}
    moods = ["Happy", "Settled", "Playful", "Quiet", "Tired", "Excited",
             "Clingy"]
    doings = [
        "Water play in the garden, then story corner",
        "Painting with sponges and a long nap",
        "Music and movement, bubbles outside",
        "Sand tray, dough and a walk to the post box",
        "Construction blocks and a picnic lunch outdoors",
        "Messy play with cooked spaghetti - well worth the clean up",
        "Autumn leaves collage and puddle jumping",
        "Heuristic play baskets and singing",
    ]
    learning = ["Taking turns", "Early mark making", "Self feeding",
                "Listening and attention", "Gross motor - climbing",
                "Naming colours", "Sharing with a friend",
                "Pouring and scooping"]
    for day in report_days:
        for kid in active:
            if kid["room"] not in report_rooms:
                continue
            if (kid["_id"], day) not in present_key:
                continue
            infant = kid["room"] == "Acorns"
            nap_from = t(12, 15) if not infant else t(9, 30)
            nap_to = t(14, 0) if not infant else t(11, 0)
            s.daily_reports.append(dict(
                date=day, child=kid["_id"], mood=rng.choice(moods),
                behaviour=rng.choice([
                    "Settled straight away at drop off.",
                    "A bit tearful at first, fine within ten minutes.",
                    "Full of energy all morning.",
                    "Quiet after lunch, wanted a cuddle.",
                    "Very independent today.",
                ]),
                activities=rng.choice(doings),
                learning=rng.choice(learning),
                outdoor_min=rng.choice([45, 60, 75, 90, 30]),
                nap_from=nap_from, nap_to=nap_to,
                nap2_min=(rng.choice([0, 30, 40]) if infant else 0),
                nappies_wet=(rng.choice([3, 4, 5]) if infant
                             else rng.choice([0, 1, 2])),
                nappies_soiled=(rng.choice([1, 1, 2]) if infant
                                else rng.choice([0, 0, 1])),
                toileting=("Nappy - wet" if infant
                           else rng.choice(["Potty - success",
                                            "Toilet", "Potty - tried",
                                            "Nappy - dry"])),
                potty_notes=("" if infant else
                             rng.choice(["Asked to go twice today.",
                                         "Dry all afternoon.", "", ""])),
                water=rng.choice([3, 4, 5, 6]),
                bottles=(rng.choice([180, 210, 240]) if infant else 0),
                teacher=rng.choice(floor_staff),
                observations=rng.choice([
                    "Stacked four blocks without help - first time.",
                    "Used two words together at snack.",
                    "Shared the trike without being asked.",
                    "Walked the length of the garden unaided.",
                    "Named three colours correctly.",
                ]),
                parent_notes=rng.choice([
                    "Could you send in a spare jumper tomorrow?",
                    "Nappies are running low.",
                    "", "", "A lovely day - see you in the morning.",
                ]),
                shared="Yes"))

    # ----------------------------------------------------------- meals ----
    meal_days = all_days[-4:]
    menu = {
        "Breakfast": ["Porridge with banana", "Toast and fruit",
                      "Weetabix and milk", "Scrambled egg on toast"],
        "Lunch": ["Chicken casserole, rice, green beans",
                  "Pasta bolognese and salad",
                  "Fish pie, peas and sweetcorn",
                  "Lentil curry, rice and naan",
                  "Shepherd's pie and broccoli"],
        "Afternoon snack": ["Apple slices and crackers",
                            "Yoghurt and berries",
                            "Breadsticks and hummus",
                            "Banana and rice cakes"],
    }
    portions = ["All", "Most", "About half", "A little", "None",
                "Refused"]
    weights = [38, 30, 16, 9, 4, 3]
    for day in meal_days:
        for sitting in ("Breakfast", "Lunch", "Afternoon snack"):
            food = rng.choice(menu[sitting])
            for kid in active:
                if (kid["_id"], day) not in present_key:
                    continue
                if sitting == "Breakfast" and rng.random() < 0.45:
                    continue
                infant = kid["room"] == "Acorns"
                alt = ""
                if kid["diet"] == "Gluten free":
                    alt = "Gluten free alternative"
                elif kid["diet"] == "Dairy free" and "milk" in food.lower():
                    alt = "Oat milk"
                elif kid["diet"] == "Vegetarian" and sitting == "Lunch":
                    alt = "Vegetarian portion"
                elif kid["diet"] == "Halal" and sitting == "Lunch":
                    alt = "Halal portion"
                s.meals.append(dict(
                    date=day, sitting=sitting, child=kid["_id"],
                    food=food,
                    portion=rng.choices(portions, weights=weights)[0],
                    alternative=alt,
                    drink="Milk" if infant else rng.choice(
                        ["Water", "Milk", "Water"]),
                    water_ml=rng.choice([120, 150, 180, 200]),
                    own_food="No", logged_by=rng.choice(floor_staff),
                    notes=""))
    # one bottle round for the babies
    for kid in active:
        if kid["room"] != "Acorns":
            continue
        if (kid["_id"], meal_days[-1]) not in present_key:
            continue
        s.meals.append(dict(
            date=meal_days[-1], sitting="Bottle", child=kid["_id"],
            food="Formula, parent supplied", portion="All",
            alternative="", drink="Formula", water_ml=210,
            own_food="Yes", logged_by=staff_id[3],
            notes="Made up fresh at 11:15."))

    # --------------------------------------------------------- medical ----
    med: List[Dict[str, Any]] = []

    def add_med(**row):
        base = dict(child="", record_type="", detail="", severity="",
                    reaction="", action="", medication="", dose="",
                    stored="", trained="", consent="", consent_date=None,
                    prescriber="", doctor="", doctor_phone="",
                    reviewed=None, expires=None, given_date=None,
                    given_time=None, given_by="", witness="", notes="")
        base.update(row)
        med.append(base)

    def kid_by_name(fragment: str) -> Optional[Dict[str, Any]]:
        for item in kids:
            if fragment in item["name"]:
                return item
        return None

    nut = kid_by_name("Rory Clarke")
    asthma = kid_by_name("Seren Hartley")
    coeliac = kid_by_name("Cillian")
    dairy = kid_by_name("Felix Hartley")
    eczema = kid_by_name("Noah Grant")
    egg = kid_by_name("Harvey Thornton")

    if nut:
        add_med(child=nut["_id"], record_type="Allergy",
                detail="Peanuts and all tree nuts",
                severity="Anaphylaxis risk",
                reaction="Hives within minutes, lip swelling, "
                         "difficulty breathing.",
                action="Give adrenaline pen immediately, call 911, "
                       "then phone the parents. Do not wait to see if "
                       "it settles.",
                medication="Adrenaline auto-injector 0.15mg",
                dose="0.15mg, outer thigh, through clothing if needed",
                stored="Red allergy box, Saplings door, unlocked",
                trained="Grace Lim, Tasha Boyd, Dawn Pritchard, "
                        "Beth Carmody",
                consent="Yes", consent_date=d(today, -300),
                prescriber="Dr Imani Walker",
                doctor="Grove Family Practice",
                doctor_phone="(614) 555 0310",
                reviewed=d(today, -120), expires=d(today, 95),
                notes="Allergy action plan with photograph on the "
                      "kitchen wall and in the room file.")
        add_med(child=nut["_id"], record_type="Immunisation",
                detail="Routine schedule up to date",
                severity="None recorded", consent="Yes",
                consent_date=d(today, -300),
                doctor="Grove Family Practice",
                doctor_phone="(614) 555 0310",
                reviewed=d(today, -300), notes="Copy on file.")
    if asthma:
        add_med(child=asthma["_id"], record_type="Asthma",
                detail="Mild, triggered by colds and cold air",
                severity="Moderate",
                reaction="Wheeze and a tight cough, usually after "
                         "running outside.",
                action="Two puffs through the spacer, sit her up, wait "
                       "five minutes. Phone home if it does not settle.",
                medication="Salbutamol inhaler 100mcg",
                dose="2 puffs via spacer, repeat once after 5 minutes",
                stored="Oaks room medical bag",
                trained="Priyanka Rao, Keisha Monroe, Dawn Pritchard",
                consent="Yes", consent_date=d(today, -210),
                prescriber="Dr Peter Lang",
                doctor="Willow Road Surgery",
                doctor_phone="(614) 555 0344",
                reviewed=d(today, -60), expires=d(today, 18),
                notes="Inhaler expires soon - new one requested.")
        for offset, hour in ((-12, 11), (-5, 14), (-2, 10)):
            add_med(child=asthma["_id"], record_type="Medication given",
                    detail="Wheezy after outdoor play",
                    severity="Moderate",
                    medication="Salbutamol inhaler 100mcg",
                    dose="2 puffs via spacer", consent="Yes",
                    consent_date=d(today, -210),
                    given_date=d(today, offset), given_time=t(hour, 20),
                    given_by=staff_id[5], witness="Keisha Monroe",
                    notes="Settled within ten minutes. Parent told at "
                          "pickup.")
    if coeliac:
        add_med(child=coeliac["_id"], record_type="Intolerance",
                detail="Coeliac disease - strictly gluten free",
                severity="Severe",
                reaction="Stomach pain and diarrhoea for days.",
                action="Separate chopping board and toaster. Check "
                       "every label, including playdough.",
                consent="Yes", consent_date=d(today, -340),
                doctor="Eastside Paediatrics",
                doctor_phone="(614) 555 0392",
                reviewed=d(today, -95),
                notes="Kitchen has the full list. Rosemary checks "
                      "deliveries.")
    if dairy:
        add_med(child=dairy["_id"], record_type="Dietary requirement",
                detail="Cow's milk protein allergy",
                severity="Moderate",
                reaction="Rash and reflux.",
                action="Soya formula supplied by the parent, labelled "
                       "and dated in the fridge.",
                consent="Yes", consent_date=d(today, -80),
                doctor="Mulberry Clinic",
                doctor_phone="(614) 555 0377",
                reviewed=d(today, -80), notes="")
    if eczema:
        add_med(child=eczema["_id"], record_type="Eczema",
                detail="Dry patches behind the knees and elbows",
                severity="Mild",
                action="Emollient after water play and before nap.",
                medication="Emollient cream", dose="Thin layer",
                stored="Saplings changing shelf",
                trained="Tasha Boyd, Danny Echeverria",
                consent="Yes", consent_date=d(today, -150),
                reviewed=d(today, -150), expires=d(today, -5),
                notes="Consent needs renewing - parent asked at "
                      "pickup.")
        add_med(child=eczema["_id"], record_type="Medication given",
                detail="Emollient after water play",
                medication="Emollient cream", dose="Thin layer",
                consent="Yes", consent_date=d(today, -150),
                given_date=d(today, -3), given_time=t(15, 10),
                given_by=staff_id[8], witness="Tasha Boyd", notes="")
    if egg:
        add_med(child=egg["_id"], record_type="Allergy",
                detail="Egg - baked egg tolerated",
                severity="Mild",
                reaction="Mouth itching, occasional hives.",
                action="Avoid loose egg. Baked goods are fine.",
                consent="Yes", consent_date=d(today, -130),
                doctor="Grove Family Practice",
                doctor_phone="(614) 555 0310",
                reviewed=d(today, -130), notes="")
    # a scattering of ordinary records so the tab is not all drama
    for index, kid in enumerate(active[:10]):
        add_med(child=kid["_id"], record_type="Immunisation",
                detail="Routine schedule up to date",
                severity="None recorded", consent="Yes",
                consent_date=d(today, -200 - index * 9),
                doctor=_DOCTORS[index % len(_DOCTORS)][0],
                doctor_phone=_DOCTORS[index % len(_DOCTORS)][1],
                reviewed=d(today, -200 - index * 9),
                expires=(d(today, 40) if index == 3 else None),
                notes="")
    s.medical = med

    # -------------------------------------------------------- incidents ---
    incident_rows = [
        ("Bump to head", "Playground", "Minor",
         "Tripped on the edge of the soft surface while running to the "
         "trikes and went down onto her hands and forehead.",
         "Small red mark to the left forehead, no swelling.",
         "Cold compress for five minutes, comforted, observed for the "
         "rest of the session.", 11, True, -2, 10, 45),
        ("Bite", "Toddler room", "Minor",
         "Two children reached for the same truck. One bit the other "
         "on the forearm.",
         "Oval mark to the right forearm, skin not broken.",
         "Washed with soap and water, cold compress, both children "
         "supported.", 14, True, -5, 15, 30),
        ("Cut or graze", "Garden", "Minor",
         "Slipped on wet decking coming down the slide.",
         "Graze to the right knee, approximately 2cm.",
         "Cleaned with saline, plaster applied with parental consent "
         "on file.", 9, True, -8, 11, 20),
        ("Fall", "Hallway", "Minor",
         "Running from the cloakroom and lost footing on the mat.",
         "No visible injury, shaken.",
         "Comforted, checked over, back to play within five minutes.",
         7, True, -12, 13, 55),
        ("Allergic reaction", "Dining area", "Serious",
         "Developed hives on the neck and chest within two minutes of "
         "sitting down for lunch. Food had been checked but a visiting "
         "grandparent had brought in a cake.",
         "Hives to neck and upper chest, no breathing difficulty.",
         "Removed from the table, antihistamine not given (no consent "
         "on file), parents phoned immediately, monitored one to one "
         "until collected.", 4, True, -17, 8, 160),
        ("Illness at the setting", "Preschool room", "No injury",
         "Temperature of 38.4 at the 11am check, flushed and "
         "lethargic.",
         "None.",
         "Moved to the quiet corner, fluids offered, parents called.",
         6, True, -21, 12, 35),
        ("Asthma episode", "Playground", "Moderate",
         "Wheezy and coughing after ten minutes of running outside on "
         "a cold morning.",
         "None - respiratory.",
         "Two puffs of her own inhaler via spacer, sat quietly, "
         "settled within ten minutes.", 22, True, -12, 10, 25),
        ("Near miss", "Entrance", "No injury",
         "The outer door was found on the latch rather than locked "
         "when the cleaner left. No child was in the lobby.",
         "None.",
         "Door secured immediately, lock checked, cleaning contractor "
         "briefed, entry procedure re-read at the staff meeting.",
         0, False, -26, 17, 0),
        ("Bruise", "Toddler room", "Minor",
         "Walked into the corner of the book trolley.",
         "Small bruise to the left shin.",
         "Cold compress and a cuddle.", 17, True, -30, 14, 50),
        ("Minor injury", "Nap room", "Minor",
         "Rolled off the low mat while waking from a nap.",
         "No mark, startled.",
         "Comforted and checked over.", 2, True, -34, 13, 70),
        ("Cut or graze", "Outing", "Minor",
         "Caught a finger on the gate latch at the farm.",
         "Small graze to the index finger.",
         "Cleaned with a wipe from the outing bag, plaster applied.",
         24, True, -38, 11, 240),
        ("Behaviour incident", "Toddler room", "No injury",
         "Repeated throwing of resources at tidy up time, two "
         "occasions in one morning.",
         "None.",
         "Supported with a quiet activity, strategy agreed with the "
         "parent at pickup.", 10, True, -41, 16, 90),
        ("Bump to head", "Classroom", "Moderate",
         "Climbed onto a chair to reach the window and fell backwards "
         "onto the carpet.",
         "Swelling to the back of the head, no loss of consciousness.",
         "Cold compress, head injury observation sheet started, "
         "parents phoned at once and advised to collect.",
         19, True, -46, 9, 8),
        ("Medication error", "Preschool room", "No injury",
         "Inhaler given at 13:40 rather than the 13:00 noted on the "
         "plan - clock in the room was slow.",
         "None - no adverse effect.",
         "Parent and manager informed the same afternoon, room clock "
         "replaced, medication times moved to the kitchen board.",
         22, True, -52, 15, 20),
        ("Lost or found property", "Entrance", "No injury",
         "A coat went home with the wrong family.",
         "None.",
         "Both families phoned, coat swapped the next morning, name "
         "labels reminder sent out.", 13, False, -58, 17, 0),
        ("Fall", "Garden", "Minor",
         "Jumped from the second step of the climbing frame.",
         "Grazed palms, both hands.",
         "Washed and dried, no plaster needed.", 15, True, -63, 10, 60),
        ("Scratch", "Infant room", "Minor",
         "Another baby reached out while they were both on the mat.",
         "Two short scratches to the cheek.",
         "Cleaned, parents shown at pickup, nails checked.",
         1, True, -69, 14, 110),
        ("Illness at the setting", "Toddler room", "No injury",
         "Vomited once after lunch, no temperature.",
         "None.",
         "Parents called, 48 hour exclusion explained, area cleaned "
         "and sanitised.", 20, True, -75, 13, 40),
    ]
    for index, row in enumerate(incident_rows, start=1):
        (category, location, severity, description, injury, first_aid,
         kid_index, notify, day_offset, hour, notify_mins) = row
        when = d(today, day_offset)
        while when.weekday() > 4:
            when -= _dt.timedelta(days=1)
        kid = kids[kid_index] if kid_index else None
        notify_at = (hour * 60 + 15 + notify_mins)
        first_aider = rng.choice(floor_staff)
        s.incidents.append(dict(
            date=when, time=t(hour, 15),
            child=(kid["_id"] if kid else ""), location=location,
            category=category, severity=severity,
            description=description, injury=injury,
            first_aid=first_aid, first_aider=first_aider,
            staff=rng.choice(floor_staff),
            witness=rng.choice(["Dawn Pritchard", "Grace Lim",
                                "Beth Carmody", "Hollie Sanderson"]),
            notified="Yes" if notify else "No",
            notified_date=when if notify else None,
            notified_time=(t(min(18, notify_at // 60), notify_at % 60)
                           if notify else None),
            notified_how=("Phone call" if notify_mins < 60
                          else "In person at pickup") if notify else "",
            notified_by=first_aider if notify else "",
            signed="Yes" if notify and index > 2 else "No",
            signed_date=(d(when, 1) if notify and index > 2 else None),
            follow_up=("Head injury sheet completed and sent home. "
                       "Parent to call if drowsy overnight."
                       if "head" in category.lower() else
                       "Review the allergy procedure for visitor food "
                       "at the next staff meeting."
                       if category == "Allergic reaction" else
                       "Lock checked daily for a fortnight."
                       if category == "Near miss" else ""),
            follow_by=(d(when, 7) if index % 3 == 0 else None),
            follow_done=("Yes" if index % 3 == 0 and index > 6
                         else "No" if index % 3 == 0 else ""),
            reported_out=("County licensing - notified within 24 hours"
                          if severity == "Serious" else ""),
            medical_sought=("Parent took her to the GP the same "
                            "afternoon as a precaution"
                            if severity == "Serious" else ""),
            notes=""))

    # ------------------------------------------------------------ staff ---
    for index, person in enumerate(STAFF, start=1):
        hire = d(today, person["hire"])
        first_aid_exp = d(today, 14 if index == 7 else
                          -20 if index == 10 else
                          120 + index * 23)
        s.staff.append(dict(
            name=person["name"], role=person["role"],
            room=person["room"],
            phone=f"(614) 555 {7000 + index * 13:04d}",
            email=f"{person['name'].split()[0].lower()}"
                  f"@littleacorns.example",
            address=f"{index * 7 + 3} "
                    f"{['Oak', 'Elm', 'Vine', 'Maple', 'Cedar'][index % 5]}"
                    f" Street, Columbus",
            dob=_months_ago(today, 260 + index * 17),
            hire_date=hire, status=person["status"],
            end_date=None, pay_type=person["pay_type"],
            hourly=money(person["hourly"]) if person["hourly"] else None,
            salary=money(person["salary"]) if person["salary"] else None,
            contract_hours=float(person["hours"]) if person["hours"]
            else None,
            day_rate=(money(person["day_rate"])
                      if person.get("day_rate") else None),
            first_aid_exp=first_aid_exp,
            cpr_exp=d(today, 90 + index * 19),
            safeguard_exp=d(today, 45 + index * 31),
            check_exp=d(today, 200 + index * 27),
            check_done=d(hire, -14),
            qualification=person["qual"],
            training=rng.choice([
                "Paediatric first aid, food hygiene level 2",
                "Safeguarding level 2, allergy awareness",
                "SEND awareness, behaviour support",
                "Food hygiene level 2, fire marshal",
            ]),
            em_name=rng.choice(["Partner - see file", "Mother",
                                "Sister", "Brother"]),
            em_phone=f"(614) 555 {8000 + index * 7:04d}",
            notes=("First aid certificate expires in two weeks - "
                   "course booked." if index == 7 else
                   "First aid certificate has lapsed. Off the ratio "
                   "count until renewed." if index == 10 else
                   "Probation review due at three months."
                   if person["status"] == "Probation" else
                   "Books through the agency, usually a day's notice."
                   if person["pay_type"] == "Daily rate" else "")))

    # --------------------------------------------------------- schedule ---
    sched_days = [day for day in all_days if (today - day).days <= 35]
    shift_times = {
        "Opening": (t(6, 45), t(15, 15)),
        "Early": (t(7, 0), t(15, 30)),
        "Morning": (t(8, 0), t(16, 0)),
        "Mid": (t(8, 30), t(16, 30)),
        "Afternoon": (t(9, 30), t(18, 0)),
        "Late": (t(9, 45), t(18, 15)),
        "Closing": (t(10, 0), t(18, 15)),
        "Full day": (t(7, 30), t(17, 30)),
        "Admin": (t(9, 0), t(15, 0)),
        "Sub cover": (t(8, 0), t(17, 0)),
    }
    leave_block = {
        4: (-16, -12, "Annual leave"),
        8: (-6, -5, "Sick leave"),
        11: (-23, -22, "Training"),
    }
    for day in sched_days:
        for index, person in enumerate(STAFF, start=1):
            if person["status"] == "Casual / bank":
                # the sub only works when somebody is away
                continue
            shift = person["shift"]
            leave = leave_block.get(index)
            on_leave = False
            if leave:
                low, high, kind = leave
                if d(today, low) <= day <= d(today, high):
                    on_leave = True
            if on_leave:
                s.schedule.append(dict(
                    date=day, staff=staff_id[index], room="",
                    shift="Leave", leave_type=leave[2],
                    start=None, finish=None, break_min=0,
                    clock_in=None, clock_out=None, substitute="",
                    notes=""))
                continue
            if index == 11 and day.weekday() == 4:
                shift = "Off"
            if index == 9 and day.weekday() in (3, 4):
                shift = "Off"
            if shift == "Off":
                s.schedule.append(dict(
                    date=day, staff=staff_id[index], room="",
                    shift="Off", leave_type="", start=None,
                    finish=None, break_min=0, clock_in=None,
                    clock_out=None, substitute="", notes=""))
                continue
            start, finish = shift_times.get(shift, (t(8, 0), t(16, 0)))
            jitter_in = rng.choice([-4, -2, 0, 0, 1, 3, 6])
            jitter_out = rng.choice([-3, 0, 0, 2, 5, 11])
            mi = start.hour * 60 + start.minute + jitter_in
            mo = finish.hour * 60 + finish.minute + jitter_out
            s.schedule.append(dict(
                date=day, staff=staff_id[index],
                room=person["room"], shift=shift, leave_type="",
                start=start, finish=finish,
                break_min=30 if person["hours"] >= 30 else 15,
                clock_in=t(mi // 60, mi % 60),
                clock_out=t(mo // 60, mo % 60),
                substitute="", notes=""))
    # the substitute covers the week of annual leave
    low, high, _kind = leave_block[4]
    for day in sched_days:
        if d(today, low) <= day <= d(today, high):
            s.schedule.append(dict(
                date=day, staff=staff_id[12], room="Saplings",
                shift="Sub cover", leave_type="", start=t(8, 0),
                finish=t(17, 0), break_min=30, clock_in=t(7, 58),
                clock_out=t(17, 4), substitute=staff_id[4],
                notes="Covering Tasha's annual leave."))

    # ---------------------------------------------------------- payroll ---
    period_end = monday - _dt.timedelta(days=3)      # last Friday
    for run in range(3):
        end = period_end - _dt.timedelta(days=14 * run)
        start = end - _dt.timedelta(days=13)
        for index, person in enumerate(STAFF, start=1):
            if person["status"] == "Casual / bank" and run != 1:
                continue
            monthly = person["pay_type"] == "Salary"
            s.payroll.append(dict(
                period_end=end,
                period_start=start,
                pay_period="Monthly" if monthly else "Fortnightly",
                staff=staff_id[index],
                bonus=(150.0 if index == 3 and run == 1 else 0),
                adjust=(-24.0 if index == 10 and run == 0 else 0),
                deductions=money(
                    (person["salary"] / 12 * 0.21) if monthly else
                    (person["hourly"] * person["hours"] * 2 * 0.19)
                    if person["hourly"] else 0),
                paid="No" if run == 0 else "Yes",
                paid_date=None if run == 0 else d(end, 5),
                method="Bank transfer",
                notes=("Long service bonus." if index == 3 and run == 1
                       else "Correction for a short shift in the "
                            "previous run."
                       if index == 10 and run == 0 else
                       "Agency cover - invoiced separately."
                       if person["pay_type"] == "Daily rate" else "")))

    # --------------------------------------------------------- expenses ---
    recurring = [
        ("Rent", "Monthly rent - 47 Willow Grove", "Grove Property LLC",
         3200.0, 0.0, "Bank transfer", "Monthly", 1),
        ("Utilities", "Gas and electricity", "AEP Ohio", 612.40, 0.0,
         "Direct debit", "Monthly", 8),
        ("Telephone & internet", "Broadband and phone line",
         "Spectrum Business", 109.00, 0.0, "Direct debit", "Monthly",
         8),
        ("Insurance", "Public liability and employers liability",
         "Grange Insurance", 308.75, 0.0, "Direct debit", "Monthly", 2),
        ("Software & subscriptions", "Parent app and payroll software",
         "Brightwheel", 94.00, 0.0, "Card", "Monthly", 5),
        ("Waste & recycling", "Commercial waste collection",
         "Rumpke", 68.50, 0.0, "Direct debit", "Monthly", 12),
        ("Cleaning supplies", "Monthly janitorial order",
         "Buckeye Janitorial", 243.80, 19.50, "Card", "Monthly", 10),
        ("Nappies & hygiene", "Nappies, wipes and creams",
         "Columbus Care Supply", 418.60, 33.50, "Card", "Monthly", 6),
    ]
    for first in months:
        last = min(_month_end(first), today)
        for (category, description, vendor, amount, tax, method,
             freq, day_of) in recurring:
            when = d(first, day_of - 1)
            if when > last:
                continue
            s.expenses.append(dict(
                date=when, category=category, description=description,
                vendor=vendor, amount=money(amount), tax=money(tax),
                method=method, receipt="Yes",
                receipt_ref=f"{vendor[:3].upper()}-{when:%y%m}",
                frequency=freq, notes=""))
        # weekly food shop
        for week_start in _weekdays(first, last)[::5]:
            s.expenses.append(dict(
                date=week_start, category="Food & groceries",
                description="Weekly food order",
                vendor="Kroger wholesale",
                amount=money(rng.uniform(380, 520)), tax=0.0,
                method="Card", receipt="Yes",
                receipt_ref=f"KRG-{week_start:%y%m%d}",
                frequency="Weekly", notes=""))

    one_offs = [
        ("Toys & learning materials", "Replacement home corner set",
         "Ohio Nursery Direct", 486.00, 38.90, -82),
        ("Repairs & maintenance", "Boiler service and two radiators",
         "Hewitt Heating", 612.00, 0.0, -70),
        ("Training & qualifications", "Paediatric first aid, four staff",
         "Midwest Safety Co", 560.00, 0.0, -61),
        ("Marketing & advertising", "Local parenting magazine, autumn",
         "Columbus Family Press", 240.00, 19.20, -55),
        ("Licensing & registration", "Annual licence renewal",
         "Ohio Department of Job and Family Services", 375.00, 0.0,
         -49),
        ("Furniture & equipment", "Six stacking cots",
         "Ohio Nursery Direct", 918.00, 73.44, -44),
        ("Arts & crafts", "Autumn term craft order",
         "Ohio Nursery Direct", 186.40, 14.90, -38),
        ("Field trips", "Farm visit coach hire", "Buckeye Coaches",
         320.00, 0.0, -34),
        ("Professional fees", "Quarterly bookkeeping",
         "Lane & Co Accountants", 420.00, 0.0, -30),
        ("Inspection & compliance", "Fire extinguisher service",
         "Midwest Safety Co", 148.00, 11.84, -27),
        ("Recruitment", "Job advert - assistant teacher",
         "Indeed", 149.00, 0.0, -24),
        ("Office & admin", "Printer toner and stationery",
         "Columbus Office Hub", 96.30, 7.70, -20),
        ("Repairs & maintenance", "Garden fence panel replaced",
         "Hewitt Property Services", 275.00, 0.0, -16),
        ("Bank & card fees", "Card processing - last month",
         "Stripe", 138.45, 0.0, -12),
        ("Educational supplies", "Phonics and story books",
         "Ohio Nursery Direct", 212.80, 17.02, -9),
        ("Staff benefits", "Team lunch - end of term",
         "Donatos", 145.00, 11.60, -6),
        ("Security", "Door entry system call-out", "SafeLock Ohio",
         185.00, 0.0, -4),
        ("Transport & vehicle", "Minibus fuel and check",
         "Shell / Hall's Garage", 164.20, 0.0, -3),
        # the two wage rows that prove the P&L does not double count
        ("Staff wages", "October payroll - posted from the bank",
         "Payroll", 18940.00, 0.0, -2),
        ("Payroll taxes & on-costs", "Employer taxes - October",
         "IRS / Ohio", 2651.60, 0.0, -2),
    ]
    for (category, description, vendor, amount, tax, offset) in one_offs:
        s.expenses.append(dict(
            date=d(today, offset), category=category,
            description=description, vendor=vendor,
            amount=money(amount), tax=money(tax),
            method=rng.choice(["Card", "Bank transfer", "Card"]),
            receipt="Yes" if offset % 7 else "No",
            receipt_ref=(f"{vendor[:3].upper()}-{abs(offset):04d}"
                         if offset % 7 else ""),
            frequency="One off",
            notes=("Counted from the Payroll Tracker, not from here - "
                   "see the Settings switch."
                   if category in ("Staff wages",
                                   "Payroll taxes & on-costs") else "")))

    # --------------------------------------------------------- supplies ---
    for index, item in enumerate(SUPPLIES, start=1):
        (name, category, unit, pack, supplier, lead, opening, recv,
         used, waste, cost, minimum) = item
        s.supplies.append(dict(
            name=name, category=category, unit=unit, pack_size=pack,
            supplier=supplier,
            supplier_phone=f"(614) 555 {9000 + index * 3:04d}",
            lead_days=lead, opening=float(opening),
            received=float(recv), used=float(used), waste=float(waste),
            adjust=(-1.0 if index == 9 else 0.0), cost=money(cost),
            min_level=float(minimum),
            last_ordered=d(today, -(7 + index * 2)),
            last_counted=d(today, -(index % 7)),
            notes=("Count was two out at the last stock take."
                   if index == 9 else
                   "Parent supplies her own - do not reorder."
                   if index == 13 else "")))

    # ---------------------------------------------------------- messages --
    message_rows = [
        (1, "Billing & fees", "Outgoing", "Email",
         "October invoice", "Sent the October invoice and the "
         "breakdown of the two extra days.", "Closed", -27, 0, None),
        (3, "Billing & fees", "Outgoing", "Phone call",
         "Overdue balance", "Left a voicemail about the September "
         "balance. Promised a call back.", "Waiting on parent", -9, 0,
         None),
        (3, "Billing & fees", "Outgoing", "Email",
         "Overdue balance - second note", "Emailed a statement and "
         "offered a payment plan.", "Waiting on parent", -3, 7, None),
        (11, "Billing & fees", "Outgoing", "Text message",
         "October balance", "Texted a reminder. Parent says the "
         "subsidy payment is late at their end.", "Open", -5, 4, None),
        (17, "Billing & fees", "Incoming", "Phone call",
         "Payment arrangement", "Grandmother called to agree a "
         "fortnightly arrangement from November.", "Waiting on us", -2,
         1, None),
        (2, "Settling in", "Outgoing", "App message",
         "Felix's first fortnight", "Sent a short note and two photos "
         "about how the settling in is going.", "Closed", -34, 0, None),
        (5, "Health & illness", "Incoming", "Phone call",
         "Temperature overnight", "Mum phoned to say Saoirse had a "
         "temperature and would not be in.", "Closed", -6, 0, None),
        (10, "Allergy", "Incoming", "In person at pickup",
         "Nut allergy plan review", "Dad asked to go through the "
         "allergy plan again before the Christmas party.",
         "Meeting booked", -8, 5, 6),
        (10, "Accident / incident", "Outgoing", "Phone call",
         "Hives at lunch", "Phoned straight away about the reaction. "
         "Full report sent the same afternoon.", "Closed", -17, 0,
         None),
        (19, "Accident / incident", "Outgoing", "Phone call",
         "Bump to the head", "Phoned within eight minutes. Parent "
         "collected at 10:30.", "Closed", -46, 0, None),
        (4, "Schedule change", "Incoming", "Email",
         "Extra day from January", "Asked whether Mateo could move "
         "from three days to four.", "Waiting on us", -4, 2, None),
        (7, "Enrollment change", "Incoming", "App message",
         "Second child", "Asked about a place for Jude in the Acorns "
         "room next spring.", "Waiting on us", -1, 3, None),
        (13, "Late pickup", "Outgoing", "Phone call",
         "Late collection", "Phoned at 18:05. Dad arrived 18:22. "
         "Late fee explained.", "Closed", -13, 0, None),
        (6, "Policy notice", "Outgoing", "Email",
         "Fee increase from January", "Sent the annual letter - 4% "
         "from 1 January with eight weeks' notice.", "Closed", -40, 0,
         None),
        (8, "Progress update", "Outgoing", "App message",
         "Emilia's half-term summary", "Shared the learning journal "
         "and three observations.", "Closed", -21, 0, None),
        (9, "Complaint", "Incoming", "Email",
         "Lost coat", "Unhappy that a coat went home with another "
         "family. Apologised, coat returned next morning, name labels "
         "reminder sent to all.", "Closed", -57, 0, None),
        (15, "Compliment", "Incoming", "Email",
         "Thank you", "Lovely note about how Noah has settled with "
         "Tasha.", "No reply needed", -11, 0, None),
        (20, "Compliment", "Incoming", "In person at pickup",
         "Thank you for the trip", "Said the farm trip photos made "
         "her week.", "No reply needed", -31, 0, None),
        (12, "Attendance", "Outgoing", "Phone call",
         "Unexplained absence", "Called after two days with no "
         "message. All fine - family were away.", "Closed", -19, 0,
         None),
        (18, "Meeting request", "Incoming", "Email",
         "Parent evening slot", "Asked for a later slot on the "
         "Thursday.", "Meeting booked", -7, 2, 4),
        (14, "Health & illness", "Outgoing", "Phone call",
         "Sickness exclusion", "Explained the 48 hour rule after a "
         "vomiting episode.", "Closed", -74, 0, None),
        (16, "General update", "Outgoing", "Email",
         "Christmas closure dates", "Sent the closure dates and the "
         "January return date.", "Closed", -15, 0, None),
        (1, "Photo consent", "Outgoing", "App message",
         "Photo consent renewal", "Asked for the annual photo consent "
         "to be re-signed.", "Waiting on parent", -10, 5, None),
        (2, "Medication", "Outgoing", "Phone call",
         "Inhaler expiry", "Asked for a replacement inhaler before "
         "the current one expires.", "Waiting on parent", -4, 3, None),
        (20, "Behaviour", "Outgoing", "In person at pickup",
         "Throwing at tidy up", "Agreed a consistent approach at home "
         "and here. Review in a fortnight.", "Open", -41, 14, None),
        (5, "Closure notice", "Outgoing", "Email",
         "Staff training day", "Reminder of the training day closure.",
         "Closed", -26, 0, None),
    ]
    staff_logger = [staff_id[2], staff_id[1], staff_id[4], staff_id[3],
                    staff_id[6]]
    for index, row in enumerate(message_rows, start=1):
        (fam, category, direction, method, subject, text, status,
         offset, follow, meeting) = row
        when = d(today, offset)
        kid_ref = ""
        for item in kids:
            if item["_fam"] == fam:
                kid_ref = item["_id"]
                break
        closed = status in ("Closed", "No reply needed")
        s.messages.append(dict(
            date=when, time=t(9 + index % 8, (index * 7) % 60),
            family=family_id[fam],
            child=kid_ref if index % 3 else "",
            direction=direction, method=method, category=category,
            subject=subject, notes=text,
            staff=staff_logger[index % len(staff_logger)],
            follow_date=(d(when, follow) if follow else None),
            follow_action=("Call back with an answer."
                           if status == "Waiting on us" else
                           "Chase if nothing by the date."
                           if status == "Waiting on parent" else ""),
            status=status,
            closed_date=(d(when, 1) if closed else None),
            meeting=(d(today, meeting) if meeting else None)))

    # --------------------------------------------------------- documents --
    child_docs = [
        ("Enrollment agreement", "Yes", None),
        ("Emergency contact form", "Yes", 365),
        ("Immunisation record", "Yes", None),
        ("Photo & video consent", "No", 365),
        ("Pickup authorization", "Yes", 365),
    ]
    for index, kid in enumerate(kids, start=1):
        for offset, (doc, required, life) in enumerate(child_docs):
            if offset == 4 and index % 3:
                continue
            if offset == 3 and index % 2:
                continue
            received = kid["enrolled_date"]
            status = "On file"
            expires = None
            if life:
                expires = d(received, life)
                if expires < today:
                    status = "Expired"
            if index in (31, 32) and offset in (0, 2):
                status = "Requested"
                received = None
                expires = None
            if index == 24 and offset == 2:
                status = "Chased"
                received = None
            s.documents.append(dict(
                holder="Child", who=kid["_id"], doc_type=doc,
                required=required, status=status,
                requested=d(received or today, -10),
                received=received, expires=expires, renewed=received,
                renew_every=(12 if life else None),
                owner=staff_id[2],
                location="Child files, cabinet 2",
                reference="", notes=""))
    staff_docs = [
        ("Staff contract", "Yes", None, "hire"),
        ("Background check", "Yes", None, "check_exp"),
        ("First aid certificate", "Yes", None, "first_aid_exp"),
        ("Child protection training", "Yes", None, "safeguard_exp"),
        ("Qualification certificate", "No", None, None),
    ]
    for index, person in enumerate(STAFF, start=1):
        record = s.staff[index - 1]
        for doc, required, _life, source in staff_docs:
            expires = record.get(source) if source and source != "hire" \
                else None
            status = "On file"
            if expires and expires < today:
                status = "Expired"
            if index == 9 and doc == "Child protection training":
                status = "Requested"
                expires = None
            s.documents.append(dict(
                holder="Staff", who=staff_id[index], doc_type=doc,
                required=required, status=status,
                requested=d(record["hire_date"], -14),
                received=record["hire_date"] if status != "Requested"
                else None,
                expires=expires, renewed=record["hire_date"],
                renew_every=(36 if doc == "First aid certificate"
                             else None),
                owner=staff_id[2], location="Staff files, locked drawer",
                reference="", notes=("Course booked for the 14th."
                                     if status == "Requested" else "")))
    business_docs = [
        ("Business licence", d(today, 205), "OH-CC-2017-04418",
         "Office, licence folder"),
        ("Insurance policy", d(today, 58), "GRG-884102-11",
         "Office, insurance folder"),
        ("Fire inspection", d(today, 26), "FI-2026-0881",
         "Office, compliance folder"),
        ("Health inspection", d(today, 141), "HI-2026-0442",
         "Office, compliance folder"),
        ("Risk assessment", d(today, -11), "RA-ROOMS-2026",
         "Shared drive, Policies"),
        ("Policy review", d(today, 74), "POL-REVIEW-2026",
         "Shared drive, Policies"),
        ("Food handling certificate", d(today, 312), "FH-RH-2026",
         "Kitchen folder"),
    ]
    for doc, expires, reference, where in business_docs:
        status = "Expired" if expires < today else "On file"
        s.documents.append(dict(
            holder="Business", who="", doc_type=doc, required="Yes",
            status=status, requested=None,
            received=_months_ago(expires, 12), expires=expires,
            renewed=_months_ago(expires, 12), renew_every=12,
            owner=staff_id[1], location=where, reference=reference,
            notes=("Overdue - the annual room review was not signed "
                   "off." if status == "Expired" else "")))

    return s
