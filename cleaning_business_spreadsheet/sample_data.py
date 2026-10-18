"""Demo data: a month in the life of Bright & Tidy Cleaning Co.

A two-van, eight-cleaner operation in Portland with fifteen clients, a
couple of commercial contracts, an Airbnb changeover round, one awkward
end-of-tenancy job and a complaint that had to be put right. Every tab
opens with believable numbers so the dashboard, the quote calculator and
the reports mean something the moment the file is opened - and `--blank`
strips the lot out for your own data.

Everything is dated relative to the build date, so the workbook never
looks stale.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .config import Config


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def week_monday(day: _dt.date) -> _dt.date:
    return day - _dt.timedelta(days=day.weekday())


def cid(n: int) -> str:
    return f"CL-{n:04d}"


def pid(n: int) -> str:
    return f"PR-{n:04d}"


def jid(n: int) -> str:
    return f"JOB-{n:04d}"


def qid(n: int) -> str:
    return f"Q-{n:04d}"


def iid(n: int) -> str:
    return f"INV-{n:04d}"


@dataclass
class Sample:
    clients: List[Dict[str, Any]] = field(default_factory=list)
    properties: List[Dict[str, Any]] = field(default_factory=list)
    leads: List[Dict[str, Any]] = field(default_factory=list)
    services: List[Dict[str, Any]] = field(default_factory=list)
    quotes: List[Dict[str, Any]] = field(default_factory=list)
    quote_lines: List[Dict[str, Any]] = field(default_factory=list)
    jobs: List[Dict[str, Any]] = field(default_factory=list)
    job_tasks: List[Dict[str, Any]] = field(default_factory=list)
    checklists: List[Dict[str, Any]] = field(default_factory=list)
    employees: List[Dict[str, Any]] = field(default_factory=list)
    timesheets: List[Dict[str, Any]] = field(default_factory=list)
    invoices: List[Dict[str, Any]] = field(default_factory=list)
    invoice_lines: List[Dict[str, Any]] = field(default_factory=list)
    payments: List[Dict[str, Any]] = field(default_factory=list)
    supplies: List[Dict[str, Any]] = field(default_factory=list)
    supply_usage: List[Dict[str, Any]] = field(default_factory=list)
    equipment: List[Dict[str, Any]] = field(default_factory=list)
    expenses: List[Dict[str, Any]] = field(default_factory=list)
    mileage: List[Dict[str, Any]] = field(default_factory=list)
    complaints: List[Dict[str, Any]] = field(default_factory=list)
    incidents: List[Dict[str, Any]] = field(default_factory=list)
    safety: List[Dict[str, Any]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# The client book
# ---------------------------------------------------------------------------
# name, type, contact, phone, email, address, status, source, since offset,
# frequency, day, slot, cleaner, rate, terms, contract, vip, notes
CLIENTS = [
    ("Harper Residence", "Residential", "Megan Harper", "(503) 555 0188",
     "megan.harper@example.com", "42 Larch Street, Portland OR 97215",
     "Regular", "Google search", -760, "Weekly", "Tuesday",
     "Morning (9-12)", "Marta Kovacs", 45, 7, "Rolling monthly", "No",
     "Two cats. Hates the smell of bleach - use the eco spray."),
    ("Okafor Family", "Residential", "Tobi Okafor", "(503) 555 0231",
     "tobi.okafor@example.com", "9 Elmwood Court, Portland OR 97212",
     "Regular", "Referral - client", -540, "Fortnightly", "Thursday",
     "Afternoon (2-5)", "Jade Ellis", 45, 7, "Rolling monthly", "No",
     "Twins under three - toys everywhere, allow extra tidy time."),
    ("Riverbend Dental", "Medical & dental", "Dr. Alan Reyes",
     "(503) 555 0410", "practice@riverbenddental.example",
     "118 Riverbend Plaza, Suite 3, Portland OR 97209", "Active",
     "Directory listing", -430, "Weekly", "Friday", "Late (5-8)",
     "Tomas Ruiz", 52, 14, "12 months", "Yes",
     "Clinical waste stays with the practice. Sign in at reception."),
    ("Cedar & Co Accountants", "Commercial office", "Priya Shah",
     "(503) 555 0366", "priya@cedarandco.example",
     "77 Cedar Avenue, Floor 2, Portland OR 97204", "Regular",
     "Referral - client", -620, "Twice a week", "Monday", "Late (5-8)",
     "Tomas Ruiz", 50, 14, "12 months", "Yes",
     "Out of hours only. Alarm code is on the Properties tab."),
    ("Lakeview Lettings", "Letting agent", "Sam Whitlock",
     "(503) 555 0177", "sam@lakeviewlettings.example",
     "3 Lakeview Road, Portland OR 97202", "Active", "Letting agent",
     "", "One-off", "Flexible", "Flexible", "Priya Raman", 48, 21,
     "No contract", "No",
     "End of tenancy work. Always wants photos and a receipt."),
    ("The Maple Cafe", "Restaurant & cafe", "Gio Bellini",
     "(503) 555 0145", "gio@maplecafe.example",
     "210 Maple Street, Portland OR 97214", "Active", "Word of mouth",
     -300, "Weekly", "Sunday", "Early (7-9)", "Amara Diallo", 55, 7,
     "6 months", "No",
     "Before opening. Kitchen deep clean quarterly, billed separately."),
    ("Bright Start Nursery", "School & nursery", "Carla Jensen",
     "(503) 555 0122", "carla@brightstart.example",
     "15 Juniper Lane, Portland OR 97213", "Regular", "Google search",
     -410, "Weekly", "Friday", "Late (5-8)", "Amara Diallo", 50, 14,
     "12 months", "Yes",
     "Background checks required for everyone on site. No harsh "
     "chemicals in the baby room."),
    ("Nguyen Household", "Residential", "Lien Nguyen", "(503) 555 0259",
     "lien.nguyen@example.com", "88 Alder Way, Portland OR 97211",
     "Active", "Facebook", -210, "Monthly", "Wednesday",
     "Morning (9-12)", "Jade Ellis", 45, 7, "Rolling monthly", "No",
     "Shoes off at the door. Key is in the lockbox."),
    ("Summit Fitness Studio", "Gym & studio", "Dev Patel",
     "(503) 555 0333", "dev@summitfitness.example",
     "501 Summit Drive, Portland OR 97201", "Active", "Instagram",
     -180, "Twice a week", "Monday", "Early (7-9)", "Luke Bennett", 52,
     14, "6 months", "No",
     "Mirrors and mats matter more than floors here."),
    ("Hartley Holiday Lets", "Holiday let / Airbnb", "Rosie Hartley",
     "(503) 555 0288", "rosie@hartleylets.example",
     "6 Quay Street, Portland OR 97209", "Regular", "Nextdoor", -330,
     "Weekly", "Saturday", "Midday (12-2)", "Jade Ellis", 50, 7,
     "Rolling monthly", "Yes",
     "Changeover between 11 and 3. Linen is in the hall cupboard."),
    ("Donovan Residence", "Residential", "Paul Donovan", "(503) 555 0194",
     "paul.donovan@example.com", "27 Birch Close, Portland OR 97215",
     "Paused", "Van signage", -480, "Fortnightly", "Thursday",
     "Morning (9-12)", "Marta Kovacs", 45, 7, "Rolling monthly", "No",
     "Paused for the summer, back in September - call in August."),
    ("Willow Park Estates", "Property manager", "Grace Lin",
     "(503) 555 0401", "grace@willowpark.example",
     "Willow Park, Block A-D, Portland OR 97217", "Active",
     "Referral - staff", -150, "Monthly", "Wednesday",
     "Afternoon (2-5)", "Luke Bennett", 48, 30, "12 months", "No",
     "Four communal stairwells and the bin stores. Invoice monthly."),
    ("Byrne Residence", "Residential", "Aoife Byrne", "(503) 555 0276",
     "aoife.byrne@example.com", "14 Rosewood Terrace, Portland OR 97206",
     "New", "Google Business Profile", -20, "Fortnightly", "Tuesday",
     "Afternoon (2-5)", "Marta Kovacs", 45, 7, "Trial period", "No",
     "Came from a deep clean quote. Trialling a fortnightly visit."),
    ("Stonebridge Law", "Commercial office", "Marcus Webb",
     "(503) 555 0388", "m.webb@stonebridgelaw.example",
     "4 Stonebridge Court, Portland OR 97204", "Active", "Yelp", -95,
     "Weekly", "Friday", "Late (5-8)", "Tomas Ruiz", 52, 14,
     "Rolling monthly", "No",
     "Confidential paperwork - never move anything off a desk."),
    ("Alvarez Family", "Residential", "Ines Alvarez", "(503) 555 0215",
     "ines.alvarez@example.com", "63 Hawthorn Road, Portland OR 97203",
     "Inactive", "Flyer / leaflet", -700, "One-off", "Flexible",
     "Flexible", "Jade Ellis", 45, 7, "No contract", "No",
     "Moved out of the area in the spring. Left a five star review."),
]

# ---------------------------------------------------------------------------
# The properties we actually clean
# ---------------------------------------------------------------------------
# client index, name, type, address, beds, baths, rooms, area, condition,
# access, key, parking, pets, floors, hours, notes
PROPERTIES = [
    (1, "Harper - family home", "House",
     "42 Larch Street, Portland OR 97215", 3, 2, 4, 1450, "Normal",
     "We hold a key", "Held by us", "Driveway", "Dog & cat",
     "Mixed", 3.0, "Cat flap in the utility door. Dog is friendly."),
    (2, "Okafor - townhouse", "Townhouse",
     "9 Elmwood Court, Portland OR 97212", 4, 2, 5, 1700, "Needs attention",
     "Client at home", "Not applicable", "Free street parking", "None",
     "Carpet & tile", 3.5, "Playroom gets the worst of it."),
    (3, "Riverbend Dental - practice", "Clinic",
     "118 Riverbend Plaza, Suite 3, Portland OR 97209", 0, 3, 8, 2200,
     "Normal", "Reception sign-in", "Not applicable", "Visitor bay",
     "None", "Vinyl", 2.5,
     "Four surgeries, waiting room, two staff rooms. Clinical bins "
     "are not ours."),
    (4, "Cedar & Co - second floor", "Office suite",
     "77 Cedar Avenue, Floor 2, Portland OR 97204", 0, 2, 9, 3100,
     "Very clean - upkeep only", "Keypad code", "Not applicable",
     "Car park", "None", "Carpet", 2.0,
     "Alarm: enter code within 30 seconds, panel by the lift."),
    (5, "Lakeview - 12B Lakeview Road", "Apartment",
     "12B Lakeview Road, Portland OR 97202", 2, 1, 3, 780,
     "Heavily soiled", "Agent meets us", "Not applicable",
     "Metered street", "None", "Laminate", 6.0,
     "End of tenancy. Oven and extractor are the big jobs."),
    (6, "The Maple Cafe - front of house", "Restaurant",
     "210 Maple Street, Portland OR 97214", 0, 2, 4, 1200, "Normal",
     "We hold a key", "Held by us", "Loading bay", "None", "Tile", 2.0,
     "Before 8am. Chairs down, floors mopped, glass polished."),
    (7, "Bright Start Nursery - main building", "School",
     "15 Juniper Lane, Portland OR 97213", 0, 4, 10, 2600, "Normal",
     "We hold a key", "In the key safe", "Driveway", "None", "Vinyl",
     3.5, "Toy sanitising on Fridays. Log it on the checklist."),
    (8, "Nguyen - family home", "House",
     "88 Alder Way, Portland OR 97211", 4, 2, 5, 1900, "Normal",
     "Lockbox", "In the key safe", "Driveway", "Cat", "Hardwood", 3.5,
     "Lockbox code changes every quarter - check before you go."),
    (9, "Summit Fitness - studio", "Gym",
     "501 Summit Drive, Portland OR 97201", 0, 2, 5, 2400, "Normal",
     "Keypad code", "Not applicable", "Car park", "None", "Rubber",
     2.5, "Mats, mirrors, changing rooms. Before 7am."),
    (10, "Hartley - Quay Street loft", "Holiday let",
     "6 Quay Street, Portland OR 97209", 2, 1, 3, 950, "Normal",
     "Lockbox", "In the key safe", "Permit needed", "None", "Hardwood",
     2.0, "Full changeover: linen, towels, welcome pack, photos."),
    (10, "Hartley - Riverside studio", "Holiday let",
     "19 Riverside Walk, Portland OR 97209", 1, 1, 2, 540, "Normal",
     "Lockbox", "In the key safe", "Pay and display", "None",
     "Laminate", 1.5, "Small but fiddly. Same checklist as the loft."),
    (11, "Donovan - bungalow", "Bungalow",
     "27 Birch Close, Portland OR 97215", 2, 1, 4, 1100, "Normal",
     "We hold a key", "Held by us", "Driveway", "Dog", "Carpet", 3.0,
     "Paused until September."),
    (12, "Willow Park - Block A", "Communal area",
     "Willow Park Block A, Portland OR 97217", 0, 0, 3, 1600, "Normal",
     "Concierge / front desk", "Not applicable", "Visitor bay", "None",
     "Stone", 1.5, "Stairwell, lift lobby, bin store."),
    (12, "Willow Park - Block B", "Communal area",
     "Willow Park Block B, Portland OR 97217", 0, 0, 3, 1600, "Normal",
     "Concierge / front desk", "Not applicable", "Visitor bay", "None",
     "Stone", 1.5, "Same scope as Block A."),
    (13, "Byrne - terrace", "House",
     "14 Rosewood Terrace, Portland OR 97206", 3, 1, 4, 1250,
     "Needs attention", "Client at home", "Not applicable",
     "Free street parking", "Dog", "Mixed", 3.0,
     "First deep clean done, now on a fortnightly trial."),
    (14, "Stonebridge Law - ground floor", "Office suite",
     "4 Stonebridge Court, Portland OR 97204", 0, 2, 7, 2100,
     "Very clean - upkeep only", "Smart lock", "Not applicable",
     "Underground garage", "None", "Carpet", 2.0,
     "Desks are never to be cleared. Bins and floors only."),
]

# ---------------------------------------------------------------------------
# What we sell
# ---------------------------------------------------------------------------
# name, category, model, hours, price, unit, uplift, supplies, popular,
# description
SERVICES = [
    ("Regular home clean", "Regular domestic clean", "Per hour", 3.0, 45,
     "hour", 0, 7.50, "Yes",
     "Kitchen, bathrooms, floors, dusting and beds made. The weekly "
     "or fortnightly visit most homes book."),
    ("First deep clean", "Deep clean", "Per hour", 6.0, 45, "hour", 0.45,
     18.00, "Yes",
     "The one before the regular visits start: skirting boards, "
     "inside the oven, limescale, behind furniture."),
    ("End of tenancy clean", "End of tenancy", "Per area", 6.0, 0.09,
     "sq ft", 0.35, 24.00, "Yes",
     "Priced on floor area. Everything cleaned inside and out so the "
     "deposit comes back."),
    ("Move in clean", "Move in / move out", "Per area", 5.0, 0.08,
     "sq ft", 0.35, 20.00, "No",
     "Before the furniture arrives. Cupboards, appliances, floors."),
    ("Office clean - small", "Commercial & office", "Flat price per visit",
     2.0, 110, "visit", 0, 6.00, "Yes",
     "Up to about 2,000 sq ft: desks, kitchen, washrooms, bins, "
     "floors."),
    ("Office clean - large", "Commercial & office", "Flat price per visit",
     3.5, 185, "visit", 0, 10.00, "No",
     "2,000 to 4,000 sq ft, same scope as the small office clean."),
    ("Airbnb changeover", "Airbnb changeover", "Flat price per visit",
     2.0, 95, "visit", 0, 8.00, "Yes",
     "Full reset between guests: linen, towels, bathroom, kitchen, "
     "welcome pack and photos."),
    ("Communal area clean", "Commercial & office", "Flat price per visit",
     1.5, 75, "visit", 0, 4.50, "No",
     "Stairwells, lobbies and bin stores for blocks and estates."),
    ("Nursery & school clean", "Commercial & office",
     "Flat price per visit", 3.5, 175, "visit", 0, 11.00, "No",
     "Classroom floors, toilets, toy sanitising and touch points."),
    ("Restaurant & kitchen clean", "Commercial & office",
     "Flat price per visit", 2.0, 125, "visit", 0, 8.00, "No",
     "Front of house, kitchen floors, surfaces and the bin area. "
     "Before opening or after close."),
    ("Oven clean", "Add-on / extra", "Per item", 1.5, 85, "oven", 0,
     9.00, "Yes",
     "Dip tank treatment for the racks, full degrease of the cavity, "
     "door glass split and cleaned."),
    ("Interior window clean", "Add-on / extra", "Per item", 0.75, 45,
     "property", 0, 3.00, "No",
     "Glass, frames and sills inside. Exterior is priced separately."),
    ("Exterior window clean", "Window cleaning", "Per item", 1.0, 60,
     "property", 0, 2.00, "No",
     "Pole fed, up to three storeys. Subcontracted to Chris."),
    ("Carpet clean", "Carpet & upholstery", "Per room", 1.0, 55, "room",
     0, 7.00, "No",
     "Hot water extraction. Allow four hours drying before furniture "
     "goes back."),
    ("Upholstery clean", "Carpet & upholstery", "Per item", 1.0, 65,
     "item", 0, 6.00, "No",
     "Sofas and armchairs, fabric only. Leather is a different "
     "treatment."),
    ("Inside fridge", "Add-on / extra", "Per item", 0.5, 30, "item", 0,
     2.50, "No", "Shelves out, washed, dried and put back."),
    ("Inside kitchen cupboards", "Add-on / extra", "Per item", 1.0, 45,
     "kitchen", 0, 3.50, "No",
     "Emptied, wiped inside and out, lined if you supply the liner."),
    ("Spring clean", "Spring clean", "Per hour", 5.0, 45, "hour", 0.25,
     14.00, "No",
     "The annual going-over: windows, blinds, cupboards, skirtings "
     "and the backs of things."),
    ("Post-construction clean", "Post-construction", "Per area", 8.0,
     0.14, "sq ft", 0.35, 32.00, "No",
     "Dust everywhere, paint specks, sticker removal. Price it after "
     "a site visit."),
    ("Pressure wash patio", "Pressure washing", "Per area", 2.0, 0.35,
     "sq ft", 0, 4.00, "No",
     "Patios, paths and driveways. Weather dependent."),
    ("Laundry & ironing", "Add-on / extra", "Per hour", 1.0, 38, "hour",
     0, 1.50, "No", "Charged by the hour at the property."),
]

# ---------------------------------------------------------------------------
# The team
# ---------------------------------------------------------------------------
# name, role, type, phone, email, start offset, pay, hours/week, status,
# skills, checked, notes
EMPLOYEES = [
    ("Dana Whitfield", "Owner", "Owner", "(503) 555 0142",
     "dana@brightandtidy.co", -2190, 25.00, 20, "Active",
     "Quoting, deep cleans, everything else", "Yes",
     "Quotes, invoices and the jobs nobody else can fit in."),
    ("Marta Kovacs", "Team leader", "Full time", "(503) 555 0511",
     "marta@brightandtidy.co", -1460, 21.00, 38, "Active",
     "Residential, deep clean, training", "Yes",
     "Runs van 1 and trains every new starter."),
    ("Jade Ellis", "Cleaner", "Full time", "(503) 555 0522",
     "jade@brightandtidy.co", -980, 18.50, 37, "Active",
     "Residential, changeovers", "Yes", "Fast and reliable on homes."),
    ("Tomas Ruiz", "Cleaner", "Full time", "(503) 555 0533",
     "tomas@brightandtidy.co", -740, 18.50, 36, "Active",
     "Commercial, out of hours", "Yes",
     "Does the evening office round. Holds the alarm codes."),
    ("Priya Raman", "Deep clean specialist", "Full time",
     "(503) 555 0544", "priya@brightandtidy.co", -610, 20.00, 35,
     "Active", "Deep clean, end of tenancy, carpets", "Yes",
     "The one to send when the place is in a state."),
    ("Luke Bennett", "Cleaner", "Part time", "(503) 555 0555",
     "luke@brightandtidy.co", -300, 17.50, 20, "Active",
     "Commercial, communal areas", "Yes",
     "Mornings only - studies in the afternoon."),
    ("Amara Diallo", "Cleaner", "Full time", "(503) 555 0566",
     "amara@brightandtidy.co", -200, 18.50, 35, "Probation",
     "Residential, cafe, nursery", "Yes",
     "Three months in and already asked for by name."),
    ("Chris Boyd", "Window cleaner", "Subcontractor", "(503) 555 0577",
     "chris@clearviewwindows.example", -520, 28.00, 8, "Active",
     "Exterior windows, pressure washing", "Yes",
     "Invoices us per property. Has his own insurance - on file."),
]


# ---------------------------------------------------------------------------
# Per property extras: how far away, which checklist, how to get in
# ---------------------------------------------------------------------------
# property number -> distance, checklist (a service name), on-site supplies
PROPERTY_EXTRA = {
    1: (4.2, "Regular home clean", "No"),
    2: (6.0, "Regular home clean", "No"),
    3: (3.1, "Office clean - small", "Yes"),
    4: (2.4, "Office clean - small", "Yes"),
    5: (7.5, "End of tenancy clean", "No"),
    6: (5.0, "Restaurant & kitchen clean", "Yes"),
    7: (4.6, "Nursery & school clean", "Yes"),
    8: (8.2, "Regular home clean", "No"),
    9: (3.8, "Office clean - small", "Yes"),
    10: (2.9, "Airbnb changeover", "Yes"),
    11: (3.4, "Airbnb changeover", "Yes"),
    12: (6.8, "Regular home clean", "No"),
    13: (9.1, "Communal area clean", "No"),
    14: (9.1, "Communal area clean", "No"),
    15: (5.5, "Regular home clean", "No"),
    16: (2.2, "Office clean - small", "Yes"),
}

ENTRY_NOTES = {
    "We hold a key": "Key 12 on the board. Alarm panel inside the front "
                     "door, code on the client card.",
    "Client at home": "Someone is always in. Ring, do not knock - the "
                      "bell wakes the baby.",
    "Reception sign-in": "Sign in at reception, collect the visitor "
                         "badge, sign out on the way through.",
    "Keypad code": "Keypad code then the green button. Re-arm on the "
                   "way out or the alarm company calls the owner.",
    "Agent meets us": "The agent unlocks and leaves. Call them when "
                      "you are finished so they can lock up.",
    "Lockbox": "Lockbox on the gatepost. Put the key back before you "
               "leave, never in your pocket.",
    "Smart lock": "One-time code sent the morning of the visit. It "
                  "expires at midnight.",
    "Open access": "Communal doors are open during the day. Bin store "
                   "key is on the ring.",
    "Key collected each visit": "Collect the key from the office before "
                                "you set off.",
    "Concierge / front desk": "Concierge holds the key and will want "
                              "the signature book.",
    "Neighbour holds key": "Number 44 next door holds the key. Text "
                           "first, she is often out before nine.",
}

# ---------------------------------------------------------------------------
# Master checklists, one per service
# ---------------------------------------------------------------------------
# area, task, how it should look, minutes, how often, required, photo
CHECKLIST_TASKS = {
    "Regular home clean": [
        ("Kitchen", "Wipe down worktops and splashback",
         "Clear, dry and streak free - nothing left on the surface",
         10, "Every visit", "Yes", "No"),
        ("Kitchen", "Clean sink and taps, polish dry",
         "No limescale round the base of the tap", 8, "Every visit",
         "Yes", "No"),
        ("Kitchen", "Front of cupboards and appliances",
         "No finger marks, handles wiped", 10, "Every visit", "Yes",
         "No"),
        ("Kitchen", "Hob and extractor front",
         "Degreased, no burnt-on spots", 10, "Every visit", "Yes",
         "No"),
        ("Kitchen", "Inside the microwave",
         "Wiped out, turntable washed", 5, "Weekly", "No", "No"),
        ("Bathroom", "Descale and polish the shower screen",
         "Clear glass, no water spots", 12, "Every visit", "Yes",
         "Yes"),
        ("Bathroom", "Clean and disinfect the toilet",
         "Inside, outside, behind and the seat hinges", 8,
         "Every visit", "Yes", "No"),
        ("Bathroom", "Basin, taps and mirror",
         "Polished, no splash marks on the mirror", 8, "Every visit",
         "Yes", "No"),
        ("Bathroom", "Floor washed, edges included",
         "Hair cleared from corners and behind the door", 8,
         "Every visit", "Yes", "No"),
        ("Bedroom", "Change the bed linen if left out",
         "Bed made properly, old linen in the basket", 8, "Weekly",
         "No", "No"),
        ("Bedroom", "Dust surfaces and headboard",
         "Nothing moved out of place, everything put back", 8,
         "Every visit", "Yes", "No"),
        ("Living room", "Dust and polish all surfaces",
         "Including skirtings, sills and the top of the frames", 12,
         "Every visit", "Yes", "No"),
        ("Living room", "Vacuum carpets and under the cushions",
         "Lines left in the carpet, cushions plumped", 12,
         "Every visit", "Yes", "No"),
        ("Hallway & stairs", "Vacuum stairs and hall, wipe bannister",
         "Edges done, no dust on the spindles", 10, "Every visit",
         "Yes", "No"),
        ("Whole property", "Empty bins and replace liners",
         "Bins wiped if sticky, rubbish taken to the outside bin", 6,
         "Every visit", "Yes", "No"),
        ("Whole property", "Final walk through and lock up",
         "Lights off, windows shut, door locked, key where it should "
         "be", 5, "Every visit", "Yes", "No"),
    ],
    "Office clean - small": [
        ("Reception", "Reception desk and glass door",
         "No finger marks on the glass at eye height", 8,
         "Every visit", "Yes", "No"),
        ("Open plan office", "Empty every bin, replace liners",
         "Including the recycling under each desk", 10, "Every visit",
         "Yes", "No"),
        ("Open plan office", "Wipe desks around equipment",
         "Nothing moved, no papers disturbed", 12, "Every visit",
         "Yes", "No"),
        ("Open plan office", "Vacuum throughout",
         "Under the desks as well as the walkways", 15, "Every visit",
         "Yes", "No"),
        ("Kitchenette", "Worktops, sink and the front of the fridge",
         "Dry, no crumbs, no mugs left in the sink", 10,
         "Every visit", "Yes", "No"),
        ("Toilets", "Clean and disinfect, restock consumables",
         "Paper, soap and towels topped up", 15, "Every visit", "Yes",
         "No"),
        ("Meeting room", "Table, chairs and whiteboard",
         "Board wiped properly, chairs pushed in", 8, "Every visit",
         "Yes", "No"),
        ("Whole property", "Touch points - handles and switches",
         "Every handle, switch and lift button", 8, "Every visit",
         "Yes", "No"),
        ("Whole property", "Set the alarm and sign out",
         "Alarm armed, badge returned, log signed", 4, "Every visit",
         "Yes", "No"),
    ],
    "Airbnb changeover": [
        ("Whole property", "Strip beds and start the laundry",
         "Linen counted - report anything stained or missing", 10,
         "Every visit", "Yes", "Yes"),
        ("Bedroom", "Make up beds, hotel finish",
         "Crisp, square corners, cushions set out", 15, "Every visit",
         "Yes", "Yes"),
        ("Bathroom", "Full clean and fresh towels",
         "Towels folded the way the host likes them", 15,
         "Every visit", "Yes", "Yes"),
        ("Kitchen", "Empty, clean and restock",
         "Fridge emptied of guest food, dishwasher run and emptied",
         15, "Every visit", "Yes", "No"),
        ("Whole property", "Welcome pack and consumables",
         "Coffee, tea, milk, paper and soap replaced", 8,
         "Every visit", "Yes", "No"),
        ("Whole property", "Check for damage and lost property",
         "Photograph anything at all - before the next guest arrives",
         8, "Every visit", "Yes", "Yes"),
        ("Living room", "Reset the room to the photos",
         "Furniture and cushions exactly as the listing shows", 10,
         "Every visit", "Yes", "Yes"),
        ("Whole property", "Vacuum and mop through",
         "No sand, no crumbs, no hair", 15, "Every visit", "Yes",
         "No"),
        ("Whole property", "Bins out, recycling sorted",
         "Collection day is Tuesday - put them out the night before",
         5, "Every visit", "Yes", "No"),
        ("Whole property", "Photos to the host",
         "Six photos: each room ready for the guest", 5,
         "Every visit", "Yes", "Yes"),
    ],
    "End of tenancy clean": [
        ("Kitchen", "Oven, racks and extractor",
         "Degreased to bare metal - this is what the deposit turns "
         "on", 60, "Every visit", "Yes", "Yes"),
        ("Kitchen", "Inside every cupboard and drawer",
         "Emptied, washed, dried, left open to air", 35,
         "Every visit", "Yes", "No"),
        ("Kitchen", "Fridge and freezer, defrosted",
         "Switched off, propped open, shelves washed", 30,
         "Every visit", "Yes", "Yes"),
        ("Kitchen", "Behind and beside the appliances",
         "Pull them out if you safely can", 20, "Every visit", "Yes",
         "No"),
        ("Bathroom", "Descale everything",
         "Tiles, screen, taps, shower head soaked", 35, "Every visit",
         "Yes", "Yes"),
        ("Whole property", "Interior windows, frames and sills",
         "Including the runners", 30, "Every visit", "Yes", "No"),
        ("Whole property", "Skirtings, doors and frames",
         "Washed, not just dusted", 30, "Every visit", "Yes", "No"),
        ("Whole property", "Light fittings and switches",
         "Shades dusted, switches wiped", 15, "Every visit", "Yes",
         "No"),
        ("Whole property", "Carpets vacuumed and spot treated",
         "Note any stain that will not come out", 25, "Every visit",
         "Yes", "Yes"),
        ("Whole property", "Hard floors washed",
         "Edges and corners by hand", 20, "Every visit", "Yes", "No"),
        ("Whole property", "Cobwebs, vents and extractor covers",
         "Look up - inventory clerks do", 15, "Every visit", "Yes",
         "No"),
        ("Whole property", "Photograph every room finished",
         "Date stamped. It settles deposit arguments", 10,
         "Every visit", "Yes", "Yes"),
    ],
    "Communal area clean": [
        ("Communal area", "Vacuum stairs and landings",
         "Including the edges and the corners", 20, "Every visit",
         "Yes", "No"),
        ("Communal area", "Mop the entrance hall",
         "Wet floor sign out while it dries", 15, "Every visit",
         "Yes", "No"),
        ("Communal area", "Glass in the entrance doors",
         "Both sides, no smears", 10, "Every visit", "Yes", "No"),
        ("Communal area", "Bin store - sweep and deodorise",
         "Bins pushed back, lids closed", 15, "Every visit", "Yes",
         "Yes"),
        ("Communal area", "Dust handrails, sills and notice boards",
         "Old notices taken down", 10, "Every visit", "Yes", "No"),
        ("Communal area", "Report anything broken or dumped",
         "Photograph it and send it to the managing agent", 5,
         "Every visit", "Yes", "Yes"),
    ],
    "Nursery & school clean": [
        ("Whole property", "Sanitise all touch points",
         "Handles, switches, rails, table edges - approved product "
         "only", 20, "Every visit", "Yes", "No"),
        ("Whole property", "Toy surfaces and play mats",
         "Child-safe sanitiser, left to air dry", 25, "Every visit",
         "Yes", "No"),
        ("Toilets", "Child toilets and low basins",
         "Disinfected, steps wiped, paper restocked", 25,
         "Every visit", "Yes", "No"),
        ("Kitchen", "Food prep surfaces",
         "Food-safe sanitiser, separate cloths", 15, "Every visit",
         "Yes", "No"),
        ("Whole property", "Floors vacuumed and mopped",
         "Under the tables, along the skirtings", 30, "Every visit",
         "Yes", "No"),
        ("Whole property", "Bins emptied, nappy bin sealed",
         "Nappy waste goes in the designated bin only", 10,
         "Every visit", "Yes", "No"),
        ("Whole property", "Sign the cleaning record",
         "The inspector will ask for it", 5, "Every visit", "Yes",
         "No"),
    ],
    "Restaurant & kitchen clean": [
        ("Whole property", "Sweep and mop front of house",
         "Chairs up, under the tables, back down", 25, "Every visit",
         "Yes", "No"),
        ("Whole property", "Tables, booths and high chairs",
         "Sanitised, no sticky edges", 15, "Every visit", "Yes",
         "No"),
        ("Kitchen", "Kitchen floor degreased",
         "Including under the pass and behind the fryer", 25,
         "Every visit", "Yes", "No"),
        ("Kitchen", "Stainless steel surfaces polished",
         "With the grain, no smears", 15, "Every visit", "Yes", "No"),
        ("Toilets", "Customer toilets, full clean and restock",
         "Checked twice - this is what reviews mention", 20,
         "Every visit", "Yes", "No"),
        ("Whole property", "Bin area and bottle store",
         "Hosed, deodorised, bins returned", 10, "Every visit", "Yes",
         "No"),
        ("Whole property", "Entrance glass and menu board",
         "First thing a customer sees", 10, "Every visit", "Yes",
         "No"),
    ],
}

# ---------------------------------------------------------------------------
# The cupboard
# ---------------------------------------------------------------------------
# name, category, supplier, unit, pack size, pack cost, stock, reorder,
# order qty, eco, coshh, notes
SUPPLIES = [
    ("Multi-surface cleaner", "Chemicals & detergents", "CleanCo Supplies",
     "5L drum", 5, 18.40, 6, 3, 6, "Yes", "Yes",
     "Diluted 1:40 in the trigger bottles. Data sheet on file."),
    ("Bathroom descaler", "Chemicals & detergents", "CleanCo Supplies",
     "5L drum", 5, 22.80, 3, 2, 4, "No", "Yes",
     "Acidic - never mix with bleach. Gloves and goggles."),
    ("Kitchen degreaser", "Chemicals & detergents", "CleanCo Supplies",
     "5L drum", 5, 24.50, 2, 2, 4, "No", "Yes",
     "The cafe and the oven work eat this."),
    ("Glass cleaner", "Chemicals & detergents", "CleanCo Supplies",
     "Bottle", 1, 3.10, 14, 6, 12, "Yes", "No",
     "Buy the concentrate, it is a third of the price."),
    ("Toilet cleaner", "Chemicals & detergents", "Wholesale Direct",
     "Bottle", 1, 2.40, 11, 6, 12, "No", "Yes", ""),
    ("Disinfectant - nursery safe", "Specialist products",
     "Hygiene Partners", "5L drum", 5, 31.00, 2, 1, 2, "Yes", "Yes",
     "The only product allowed in the baby room."),
    ("Microfibre cloths - colour coded", "Cloths & sponges",
     "Wholesale Direct", "Pack", 10, 9.90, 9, 4, 8, "Yes", "No",
     "Red bathroom, blue general, green kitchen, yellow glass."),
    ("Scourer sponges", "Cloths & sponges", "Wholesale Direct", "Pack",
     10, 4.20, 7, 3, 6, "No", "No", ""),
    ("Mop heads", "Mops & buckets", "CleanCo Supplies", "Each", 1, 6.50,
     5, 4, 8, "No", "No", "Washed at 60 after every commercial job."),
    ("Bin bags - heavy duty", "Bin bags & waste", "Wholesale Direct",
     "Roll", 50, 7.80, 8, 4, 10, "No", "No", ""),
    ("Nitrile gloves", "Gloves & PPE", "Hygiene Partners", "Box", 100,
     11.50, 4, 3, 6, "No", "No", "Powder free. Two sizes in the van."),
    ("Toilet rolls - commercial", "Paper & consumables",
     "Wholesale Direct", "Case", 36, 28.00, 2, 2, 4, "Yes", "No",
     "Recharged to the office clients at cost plus 10%."),
    ("Hand soap refill", "Paper & consumables", "Wholesale Direct",
     "Refill", 1, 4.60, 6, 3, 6, "Yes", "No", ""),
    ("Vacuum bags - Sebo", "Vacuum bags & filters", "Vac Spares Online",
     "Pack", 10, 12.40, 3, 2, 4, "No", "No",
     "Only fits the two uprights, not the backpack."),
    ("Laundry detergent", "Laundry", "Wholesale Direct", "5L drum", 5,
     14.20, 2, 1, 2, "Yes", "No", "For the changeover linen."),
    ("Oven cleaning gel", "Specialist products", "Oven Pro Supplies",
     "Bottle", 1, 16.90, 3, 2, 4, "No", "Yes",
     "Caustic. Full PPE, windows open, never on a hot oven."),
]

# ---------------------------------------------------------------------------
# The kit
# ---------------------------------------------------------------------------
# name, type, make, serial, ownership, status, assigned, location,
# bought, cost, life, warranty, last service, interval, service cost,
# safety test, insured, notes
EQUIPMENT = [
    ("Sebo upright - van 1", "Upright vacuum", "Sebo Automatic X7",
     "SB-774210", "Owned outright", "In service", "Marta Kovacs",
     "Van 1", -980, 385.00, 7, -250, -120, 180, 45.00, -120, "Yes",
     "Workhorse. Belt changed every spring."),
    ("Sebo upright - van 2", "Upright vacuum", "Sebo Automatic X7",
     "SB-774588", "Owned outright", "In service", "Tomas Ruiz",
     "Van 2", -610, 385.00, 7, -245, -200, 180, 45.00, -200, "Yes",
     ""),
    ("Numatic backpack", "Backpack vacuum", "Numatic RSB150",
     "NM-220914", "Owned outright", "In use", "Luke Bennett", "Van 2",
     -420, 299.00, 6, -55, -95, 180, 35.00, -95, "Yes",
     "For the stairwells - saves twenty minutes a block."),
    ("Wet and dry vacuum", "Wet & dry vacuum", "Karcher WD6",
     "KR-556102", "Owned outright", "Spare", "", "Store cupboard",
     -730, 189.00, 6, "", -310, 365, 30.00, -310, "Yes",
     "Comes out for floods and post-construction work."),
    ("Carpet extraction machine", "Carpet machine", "Prochem Steempro",
     "PC-100483", "Financed", "In service", "Priya Raman", "Van 1",
     -520, 1850.00, 8, -155, -40, 180, 120.00, -40, "Yes",
     "On finance until next March. Earns its keep twice over."),
    ("Steam cleaner", "Steam cleaner", "Karcher SC4", "KR-889321",
     "Owned outright", "In repair", "", "Repair shop", -400, 245.00,
     5, "", -210, 365, 55.00, -210, "Yes",
     "Boiler seal gone. Quoted 60 to fix, 245 to replace."),
    ("Pressure washer", "Pressure washer", "Nilfisk C135",
     "NF-443027", "Owned outright", "In service", "Chris Boyd",
     "Van 2", -300, 320.00, 6, -120, -150, 365, 40.00, -150, "Yes",
     "Patio and bin store work only."),
    ("Window pole system", "Window pole system", "Unger HiFlo",
     "UG-770145", "Owned outright", "In service", "Chris Boyd",
     "Van 2", -510, 640.00, 8, "", -180, 365, 25.00, -180, "Yes",
     "Pure water system, three storey reach."),
    ("Step ladder - 5 tread", "Ladder & steps", "Youngman 5",
     "YM-003311", "Owned outright", "In service", "", "Van 1", -900,
     95.00, 10, "", -75, 365, 0.00, -75, "Yes",
     "Inspected quarterly - log it here."),
    ("Cleaning trolley", "Trolley & caddy", "Numatic NX1",
     "NM-551209", "Owned outright", "In service", "Tomas Ruiz",
     "Cedar & Co site", -350, 210.00, 8, "", "", 365, 0.00, "", "Yes",
     "Lives at the office site - saves carrying."),
    ("Van 1 - Transit Connect", "Vehicle", "Ford Transit Connect",
     "VIN-9921", "Financed", "In service", "Marta Kovacs", "Van 1",
     -1100, 18500.00, 8, -370, -85, 180, 320.00, -85, "Yes",
     "Serviced every six months, MOT in the autumn."),
    ("Van 2 - Berlingo", "Vehicle", "Citroen Berlingo", "VIN-7740",
     "Financed", "In service", "Tomas Ruiz", "Van 2", -700, 16200.00,
     8, -340, -170, 180, 295.00, -170, "Yes", ""),
]

# ---------------------------------------------------------------------------
# The paperwork that has to stay in date
# ---------------------------------------------------------------------------
# document, applies to, employee, reference, issuer, issued, expires,
# cost, holder, required, reviewed, notes
SAFETY_DOCS = [
    ("Public liability insurance", "The business", "", "PL-884213",
     "Hiscox", -300, 65, 486.00, "Insurance folder, office", "Yes",
     -300, "2,000,000 cover. Commercial clients ask for this first."),
    ("Employer liability insurance", "The business", "", "EL-884214",
     "Hiscox", -300, 65, 310.00, "Insurance folder, office", "Yes",
     -300, "Legally required once you employ anyone."),
    ("Vehicle insurance", "Van 1", "", "MV-220714", "Admiral Business",
     -200, 165, 1180.00, "Van document wallet", "Yes", -200, ""),
    ("Vehicle insurance", "Van 2", "", "MV-220715", "Admiral Business",
     -200, 165, 1090.00, "Van document wallet", "Yes", -200, ""),
    ("Background check", "Staff", "Amara Diallo", "BC-55210",
     "CheckPoint", -195, 900, 44.00, "Staff file", "Yes", -195,
     "Required for the nursery contract."),
    ("Background check", "Staff", "Marta Kovacs", "BC-41088",
     "CheckPoint", -1450, 180, 44.00, "Staff file", "Yes", -1450,
     "Due for renewal - the nursery audits this every year."),
    ("Background check", "Staff", "Jade Ellis", "BC-47720",
     "CheckPoint", -970, 740, 44.00, "Staff file", "Yes", -970, ""),
    ("Training certificate", "Staff", "Priya Raman", "TR-COSHH-220",
     "BICSc", -610, 455, 120.00, "Staff file", "Yes", -610,
     "COSHH handling. Renew every three years."),
    ("First aid certificate", "Staff", "Marta Kovacs", "FA-77120",
     "Red Cross", -720, 375, 95.00, "Staff file", "No", -720,
     "One trained first aider is enough for a team this size."),
    ("COSHH data sheet", "Oven cleaning gel", "", "SDS-OVEN-04",
     "Oven Pro Supplies", -400, 330, 0.00, "Chemical folder, van 1",
     "Yes", -120, "Must travel with the product, not stay in the "
     "office."),
    ("COSHH data sheet", "Bathroom descaler", "", "SDS-DESC-11",
     "CleanCo Supplies", -380, 350, 0.00, "Chemical folder, van 1",
     "Yes", -120, ""),
    ("Risk assessment", "Working at height", "", "RA-004",
     "Dana Whitfield", -390, 340, 0.00, "Health and safety folder",
     "Yes", -390,
     "Covers steps and pole work. Reviewed annually - this one is "
     "nearly due."),
    ("Risk assessment", "Lone working", "", "RA-007",
     "Dana Whitfield", -210, 520, 0.00, "Health and safety folder",
     "Yes", -210, "Check-in text at the end of every evening job."),
    ("Method statement", "End of tenancy cleaning", "", "MS-002",
     "Dana Whitfield", -260, 470, 0.00, "Health and safety folder",
     "No", -260, "Letting agents ask for this with the quote."),
    ("Electrical safety test", "Equipment", "", "PAT-2026-01",
     "SafeTest Ltd", -85, 280, 140.00, "Equipment folder", "Yes",
     -85, "All twelve items tested. Stickers on each machine."),
    ("Waste carrier licence", "The business", "", "WCL-99120",
     "Environment Agency", -500, 230, 154.00, "Office folder", "Yes",
     -500, "Needed to take trade waste away from commercial jobs."),
]

# ---------------------------------------------------------------------------
# Enquiries in the pipeline
# ---------------------------------------------------------------------------
# days ago, name, phone, email, address, type, source, service,
# frequency, stage, value, quote no, visit offset, contacted,
# next action, next offset, client no, lost reason, notes
LEADS = [
    (-2, "Ellis Shaw", "(503) 555 0612", "ellis.shaw@example.com",
     "31 Fir Street, Portland OR 97215", "Residential", "Google search",
     "Regular home clean", "Fortnightly", "New enquiry", 150, "", "",
     "No", "Call back and book a quote visit", 1, "", "",
     "Rang while we were on a job. Three bed, two bath, one dog."),
    (-3, "Priya Nair", "(503) 555 0633", "priya.nair@example.com",
     "8 Aspen Court, Portland OR 97212", "Residential",
     "Referral - client", "Regular home clean", "Weekly", "Contacted",
     220, "", "", "Yes", "Send the quote", 1, "", "",
     "Referred by the Okafors. Wants the same cleaner every week."),
    (-6, "Marsh & Fletcher Dental", "(503) 555 0655",
     "office@marshfletcher.example",
     "202 Bridge Road, Portland OR 97209", "Medical & dental", "Yelp",
     "Office clean - small", "Weekly", "Visit booked", 480, "", 2,
     "Yes", "Site visit Thursday at 6pm", 2, "", "",
     "Second practice in the area - could lead to both."),
    (-9, "Tom Ridley", "(503) 555 0677", "t.ridley@example.com",
     "77 Chestnut Avenue, Portland OR 97206", "Landlord",
     "Google Business Profile", "End of tenancy clean", "One-off",
     "Quote sent", 395, "Q-0004", "", "Yes", "Chase the quote", 0,
     "", "", "Tenant leaves on the 28th. Needs it done that week."),
    (-14, "Beacon Coworking", "(503) 555 0699",
     "hello@beaconcowork.example",
     "14 Beacon Street, Portland OR 97204", "Commercial office",
     "Nextdoor", "Office clean - large", "Twice a week",
     "Following up", 1240, "Q-0005", -8, "Yes",
     "Second follow-up call", -1, "", "",
     "Biggest opportunity this quarter. Decision maker is Ana."),
    (-21, "Harriet Vance", "(503) 555 0701", "h.vance@example.com",
     "5 Willow Bank, Portland OR 97202", "Residential", "Facebook",
     "First deep clean", "One-off", "Won", 320, "Q-0003", -18, "Yes",
     "Booked in", "", "", "",
     "Became the Byrne-style trial - booked a deep clean first."),
    (-26, "Sunnyside Apartments", "(503) 555 0722",
     "manager@sunnyside.example",
     "Sunnyside Block C, Portland OR 97217", "Property manager",
     "Referral - staff", "Communal area clean", "Monthly", "Lost",
     320, "Q-0006", -22, "Yes", "", "", "", "Price too high",
     "Went with a cheaper outfit. Worth a call in six months."),
    (-32, "Dalton Residence", "(503) 555 0744",
     "dalton.home@example.com",
     "19 Poplar Drive, Portland OR 97203", "Residential",
     "Flyer / leaflet", "Regular home clean", "Monthly", "Lost", 110,
     "", "", "Yes", "", "", "", "Out of our area",
     "Forty minutes each way. Politely declined."),
    (-5, "Craig Mellor", "(503) 555 0766", "craig.mellor@example.com",
     "44 Hazel Road, Portland OR 97213", "Residential", "Instagram",
     "Oven clean", "One-off", "New enquiry", 85, "", "", "No",
     "Price it and text him back", 0, "", "",
     "Saw the before and after photos on Instagram."),
    (-11, "Northgate Physio", "(503) 555 0788",
     "clinic@northgatephysio.example",
     "60 Northgate Way, Portland OR 97211", "Medical & dental",
     "Directory listing", "Office clean - small", "Twice a week",
     "Negotiating", 620, "Q-0007", -7, "Yes",
     "They want 10% off for a 12 month contract", 2, "", "",
     "Worth doing at 5% if they sign for a year."),
]

# ---------------------------------------------------------------------------
# Quotes raised
# ---------------------------------------------------------------------------
# key, date offset, client no, prospect, property no, address, service,
# frequency, status, sent offset, followed up offset, decision offset,
# discount, job key, decline reason, notes, lines
QUOTES = [
    ("Q1", -34, 13, "", 15, "", "Regular home clean", "Fortnightly",
     "Accepted", -34, "", -31, 0.0, "", "",
     "Trial fortnightly visit after the deep clean.",
     [("Regular home clean", 1, None, 0.0, "No",
       "Three bed terrace, fortnightly visit."),
      ("Interior window clean", 1, None, 0.0, "Yes",
       "Optional - they took it on the first visit only.")]),
    ("Q2", -28, 9, "", 9, "", "Office clean - small", "Twice a week",
     "Accepted", -28, "", -24, 0.05, "", "",
     "Studio wanted mirrors and mats done properly.",
     [("Office clean - small", 2, None, 0.0, "No",
       "Two evening visits a week, mirrors and mats included.")]),
    ("Q3", -21, "", "Harriet Vance", "", "5 Willow Bank, Portland OR "
     "97202", "First deep clean", "One-off", "Accepted", -21, -19,
     -18, 0.0, "", "", "Deep clean first, then talk about regular.",
     [("First deep clean", 1, None, 0.0, "No",
       "Whole house deep clean, two cleaners, one day.")]),
    ("Q4", -9, "", "Tom Ridley", "", "77 Chestnut Avenue, Portland OR "
     "97206", "End of tenancy clean", "One-off", "Sent", -9, "", "",
     0.0, "", "",
     "Two bed flat, tenant moving out on the 28th.",
     [("End of tenancy clean", 1, None, 0.0, "No",
       "Full end of tenancy clean to inventory standard."),
      ("Oven clean", 1, None, 0.0, "No",
       "Oven, racks and extractor - dip tank treatment."),
      ("Carpet clean", 3, None, 0.0, "Yes",
       "Three rooms of carpet, hot water extraction.")]),
    ("Q5", -14, "", "Beacon Coworking", "", "14 Beacon Street, "
     "Portland OR 97204", "Office clean - large", "Twice a week",
     "Followed up", -14, -6, "", 0.0, "", "",
     "Coworking space, 4,000 sq ft, two evenings a week.",
     [("Office clean - large", 2, None, 0.0, "No",
       "Two evening visits a week across the whole floor."),
      ("Interior window clean", 1, None, 0.0, "Yes",
       "Monthly internal glass - optional extra.")]),
    ("Q6", -26, "", "Sunnyside Apartments", "", "Sunnyside Block C, "
     "Portland OR 97217", "Communal area clean", "Monthly",
     "Declined", -26, -20, -18, 0.0, "", "Price too high",
     "Lost on price. They took a 55 a visit offer.",
     [("Communal area clean", 4, None, 0.0, "No",
       "Four blocks, monthly visit each.")]),
    ("Q7", -11, "", "Northgate Physio", "", "60 Northgate Way, "
     "Portland OR 97211", "Office clean - small", "Twice a week",
     "Sent", -11, -4, "", 0.05, "", "",
     "They are pushing for 10%. Hold at 5% with a 12 month term.",
     [("Office clean - small", 2, None, 0.0, "No",
       "Two visits a week, clinical waste excluded.")]),
    ("Q8", -45, 5, "", 5, "", "End of tenancy clean", "One-off",
     "Expired", -45, -38, "", 0.0, "", "",
     "Agent never came back on this one - chase the branch.",
     [("End of tenancy clean", 1, 420.00, 0.0, "No",
       "Two bed apartment, end of tenancy to inventory standard.")]),
]

# ---------------------------------------------------------------------------
# Regular rounds: client, property, service, weekday, start, every n days
# ---------------------------------------------------------------------------
ROUND = [
    (1, 1, "Regular home clean", 1, (9, 0), 7, "Marta Kovacs", "", 1,
     "Weekly"),
    (2, 2, "Regular home clean", 3, (14, 0), 14, "Jade Ellis", "", 1,
     "Fortnightly"),
    (3, 3, "Office clean - small", 4, (18, 0), 7, "Tomas Ruiz", "", 1,
     "Weekly"),
    (4, 4, "Office clean - small", 0, (18, 30), 7, "Tomas Ruiz", "", 1,
     "Twice a week"),
    (4, 4, "Office clean - small", 3, (18, 30), 7, "Tomas Ruiz", "", 1,
     "Twice a week"),
    (6, 6, "Restaurant & kitchen clean", 6, (7, 0), 7, "Amara Diallo",
     "", 1, "Weekly"),
    (7, 7, "Nursery & school clean", 4, (17, 30), 7, "Amara Diallo",
     "Luke Bennett", 2, "Weekly"),
    (8, 8, "Regular home clean", 2, (9, 30), 28, "Jade Ellis", "", 1,
     "Monthly"),
    (9, 9, "Office clean - small", 0, (7, 0), 7, "Luke Bennett", "", 1,
     "Twice a week"),
    (9, 9, "Office clean - small", 3, (7, 0), 7, "Luke Bennett", "", 1,
     "Twice a week"),
    (10, 10, "Airbnb changeover", 5, (11, 0), 7, "Jade Ellis", "", 1,
     "Weekly"),
    (10, 11, "Airbnb changeover", 5, (14, 0), 7, "Jade Ellis", "", 1,
     "Weekly"),
    (12, 13, "Communal area clean", 2, (14, 0), 28, "Luke Bennett", "",
     1, "Monthly"),
    (12, 14, "Communal area clean", 2, (15, 30), 28, "Luke Bennett",
     "", 1, "Monthly"),
    (13, 15, "Regular home clean", 1, (14, 0), 14, "Marta Kovacs", "",
     1, "Fortnightly"),
    (14, 16, "Office clean - small", 4, (20, 0), 7, "Tomas Ruiz", "",
     1, "Weekly"),
]

#: client 11 is paused for the summer - jobs stop forty days ago
ROUND_STOPS = {(11, 12): -40}

# ---------------------------------------------------------------------------
# Jobs that are not part of a round
# ---------------------------------------------------------------------------
# key, day offset, client, property, service, type, start, status,
# employee, employee2, cleaners, hours override, price override,
# rating, feedback, notes
ONE_OFFS = [
    ("J_tenancy", -8, 5, 5, "End of tenancy clean", "One-off clean",
     (8, 30), "Completed", "Priya Raman", "Amara Diallo", 2, 6.5,
     395.00, 5, "Agent said it was the best handover all year.",
     "Oven was the worst we have seen. Photographed everything."),
    ("J_deep", -18, 13, 15, "First deep clean", "First deep clean",
     (9, 0), "Completed", "Priya Raman", "Marta Kovacs", 2, 6.0, None,
     5, "Lovely job, booked a fortnightly after this.",
     "Converted straight to a fortnightly visit."),
    ("J_oven", -10, 1, 1, "Oven clean", "One-off clean", (13, 0),
     "Completed", "Priya Raman", "", 1, None, None, 5, "", 
     "Added on at the end of the regular visit."),
    ("J_carpet", 6, 4, 4, "Carpet clean", "One-off clean", (18, 0),
     "Booked", "Priya Raman", "", 1, 3.0, None, "", "",
     "Four offices of carpet, quoted separately from the contract."),
    ("J_spring", 9, 8, 8, "Spring clean", "One-off clean", (9, 0),
     "Booked", "Marta Kovacs", "Jade Ellis", 2, 5.0, None, "", "",
     "Annual going-over. Blinds, windows and cupboards."),
    ("J_noaccess", -13, 8, 8, "Regular home clean", "Recurring visit",
     (9, 30), "No access", "Jade Ellis", "", 1, None, 0.00, "", "",
     "Nobody home and the lockbox was empty. Charged nothing, "
     "rebooked."),
    ("J_cancel", -4, 2, 2, "Regular home clean", "Recurring visit",
     (14, 0), "Cancelled", "Jade Ellis", "", 1, None, None, "", "",
     "Family away. Cancelled the night before, no charge."),
    ("J_reclean", -5, 6, 6, "Restaurant & kitchen clean",
     "Re-clean (free)", (7, 0), "Completed", "Marta Kovacs", "", 1,
     1.5, 0.00, 4, "Sorted, thank you for coming back.",
     "Free re-clean after the complaint about the kitchen floor."),
    ("J_quote", 2, "", "", "", "Quote visit", (17, 0), "Booked",
     "Dana Whitfield", "", 1, 0.5, 0.00, "", "",
     "Site visit at Marsh & Fletcher Dental - take the method "
     "statement."),
]

# ---------------------------------------------------------------------------
# Money going out
# ---------------------------------------------------------------------------
# category, type, description, supplier, amount, tax, method, receipt,
# recurring, deductible, day of month, months back
#: service name -> the hours one cleaner is expected to take, so the
#: demo timesheets can be built from the same figure the price uses
SERVICE_HOURS = {row[0]: row[3] for row in SERVICES}

MONTHLY_COSTS = [
    ("Business insurance", "Fixed monthly", "Public and employer "
     "liability", "Hiscox", 66.00, 0.0, "Direct debit", "Yes", "Yes",
     "Yes", 1),
    ("Vehicle insurance", "Fixed monthly", "Van 1 and van 2 fleet "
     "policy", "Admiral Business", 150.00, 0.0, "Direct debit", "Yes",
     "Yes", "Yes", 1),
    ("Software & apps", "Fixed monthly", "Accounting software and "
     "cloud storage", "Ledgerly", 32.00, 0.0, "Card", "Yes", "Yes",
     "Yes", 3),
    ("Phone & internet", "Fixed monthly", "Two business mobiles",
     "Telecom NW", 58.00, 0.0, "Direct debit", "Yes", "Yes", "Yes", 5),
    ("Marketing & advertising", "Fixed monthly", "Google Ads - local "
     "cleaning searches", "Google", 120.00, 0.0, "Card", "Yes", "Yes",
     "Yes", 7),
    ("Office & rent", "Fixed monthly", "Storage unit for stock and "
     "machines", "SpaceStore", 75.00, 0.0, "Direct debit", "Yes",
     "Yes", "Yes", 2),
    ("Accounting & legal", "Fixed monthly", "Bookkeeping",
     "Cedar & Co Accountants", 85.00, 0.0, "Bank transfer", "Yes",
     "Yes", "Yes", 12),
    ("Vehicle fuel", "Variable", "Van 1 fuel", "Shell", 110.00, 0.0,
     "Card", "Yes", "No", "Yes", 9),
    ("Vehicle fuel", "Variable", "Van 2 fuel", "Shell", 95.00, 0.0,
     "Card", "Yes", "No", "Yes", 16),
    ("Cleaning supplies", "Variable", "Monthly stock order",
     "CleanCo Supplies", 165.00, 0.0, "Card", "Yes", "No", "Yes", 4),
    ("Laundry", "Variable", "Changeover linen service",
     "Quay Laundry", 64.00, 0.0, "Bank transfer", "Yes", "No", "Yes",
     18),
    ("Subcontractors", "Job related", "Window cleaning - Chris Boyd",
     "Clearview Windows", 180.00, 0.0, "Bank transfer", "Yes", "No",
     "Yes", 21),
]

ONE_OFF_COSTS = [
    (-6, "Equipment repair", "One-off", "Steam cleaner boiler seal",
     "Vac Spares Online", 62.00, 0.0, "Card", "Yes", "No", "Yes"),
    (-12, "Cleaning supplies", "Variable", "Emergency degreaser run",
     "Wholesale Direct", 48.90, 0.0, "Cash", "No", "No", "Yes"),
    (-19, "Training & certification", "One-off",
     "COSHH refresher for two cleaners", "BICSc", 190.00, 0.0,
     "Bank transfer", "Yes", "No", "Yes"),
    (-24, "Uniforms", "One-off", "Six polo shirts and two fleeces",
     "WorkWear Direct", 142.00, 0.0, "Card", "Yes", "No", "Yes"),
    (-31, "Equipment purchase", "One-off", "Replacement mop system",
     "CleanCo Supplies", 98.00, 0.0, "Card", "Yes", "No", "Yes"),
    (-38, "Marketing & advertising", "One-off",
     "Van signage refresh", "SignShop PDX", 380.00, 0.0,
     "Bank transfer", "Yes", "No", "Yes"),
    (-44, "Vehicle maintenance", "One-off", "Van 1 service and MOT",
     "Hilltop Garage", 318.00, 0.0, "Card", "Yes", "No", "Yes"),
    (-2, "Bank & card fees", "Fixed monthly", "Card processing fees",
     "SumUp", 18.40, 0.0, "Direct debit", "No", "Yes", "Yes"),
    (-27, "Waste disposal", "One-off", "Trade waste from the end of "
     "tenancy job", "City Waste", 45.00, 0.0, "Card", "Yes", "No",
     "Yes"),
]

# ---------------------------------------------------------------------------
# When it went wrong
# ---------------------------------------------------------------------------
# day offset, client, job key, employee, type, severity, description,
# status, owner, action, resolution, reclean key, cost, resolved
# offset, kept, rating after, prevention
COMPLAINTS = [
    (-7, 6, "", "Amara Diallo", "Missed area", "Medium",
     "Kitchen floor behind the fryer was missed two weeks running.",
     "Resolved", "Dana Whitfield",
     "Went back the same week and did it free.", "Free re-clean",
     "J_reclean", 68.00, -5, "Yes", 4,
     "Added 'behind the fryer' to the cafe checklist as a required "
     "task."),
    (-16, 10, "", "Jade Ellis", "Quality below standard", "Low",
     "Host said the bathroom mirror was smeared on a changeover.",
     "Closed - no action", "Dana Whitfield",
     "Spoke to Jade, sent the photo. Guest never mentioned it.",
     "Apology", "", 0.00, -15, "Yes", 5,
     "Glass cloth now colour coded yellow so it is never the wrong "
     "one."),
    (-3, 12, "", "Luke Bennett", "Arrived late", "Low",
     "Agent rang - we arrived at 3pm instead of 2pm for Block A.",
     "Investigating", "Dana Whitfield",
     "Checking the rota - the two blocks are booked too close "
     "together.", "", "", 0.00, "", "", "",
     "Move Block B to 4pm so there is travel time between them."),
    (-42, 2, "", "Jade Ellis", "Breakage", "Medium",
     "Glass vase knocked off a windowsill while cleaning behind it.",
     "Resolved", "Dana Whitfield",
     "Replaced it the next day and apologised in person.",
     "Item replaced", "", 38.00, -40, "Yes", 5,
     "Ornaments are now moved to the table, not balanced on the "
     "sill."),
]

# ---------------------------------------------------------------------------
# Safety incidents
# ---------------------------------------------------------------------------
# day offset, time, type, employee, job key, property, severity,
# description, injury, first aid, medical, damage, cost, insurance,
# reported to, action, prevention, status, closed offset
INCIDENTS = [
    (-23, (10, 15), "Slip, trip or fall", "Jade Ellis", "", 8, "Low",
     "Slipped on a wet kitchen floor that had just been mopped - no "
     "wet floor sign out.", "Bruised elbow", "Yes", "No", "No", 0.00,
     "No", "Dana Whitfield",
     "First aid applied, finished the shift.",
     "Wet floor signs now in both vans and on every checklist.",
     "Closed", -22),
    (-61, (16, 40), "Chemical splash", "Priya Raman", "", 5, "Medium",
     "Oven gel splashed onto the forearm while lifting a rack out of "
     "the dip tank.", "Minor chemical burn", "Yes", "No", "No", 0.00,
     "No", "Dana Whitfield",
     "Rinsed for fifteen minutes, checked at the pharmacy.",
     "Long gauntlets now mandatory for all dip tank work - written "
     "into the method statement.", "Closed", -58),
    (-11, (19, 5), "Near miss", "Tomas Ruiz", "", 4, "Low",
     "Trailing vacuum cable across a doorway in the office - nobody "
     "tripped, but somebody could have.", "None", "No", "No", "No",
     0.00, "No", "Dana Whitfield",
     "Reminded the team to work from the nearest socket.",
     "Cable routing added to the office checklist.", "Closed", -10),
]


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------

_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg)
    return _CACHE[key]


def _month_key(day: _dt.date) -> str:
    return day.strftime("%B %Y")


def _build(cfg: Config) -> Sample:
    rnd = random.Random(20261018)
    today = cfg.today
    monday = week_monday(today)
    s = Sample()

    # ------------------------------------------------------------ clients --
    for (name, kind, contact, phone, email, address, status, source,
         since, freq, day, slot, cleaner, rate, terms, contract, vip,
         notes) in CLIENTS:
        s.clients.append({
            "name": name, "type": kind, "contact": contact,
            "phone": phone, "email": email, "address": address,
            "status": status, "source": source,
            "since": d(today, since) if since != "" else d(today, -30),
            "frequency": freq, "pref_day": day, "pref_time": slot,
            "pref_cleaner": cleaner, "rate": rate, "terms": terms,
            "contract": contract, "vip": vip, "marketing": "Yes",
            "notes": notes,
        })

    # --------------------------------------------------------- properties --
    for index, (client, name, kind, address, beds, baths, rooms, area,
                condition, access, key, parking, pets, floors, hours,
                notes) in enumerate(PROPERTIES, 1):
        distance, checklist, onsite = PROPERTY_EXTRA.get(
            index, (5.0, "Regular home clean", "No"))
        s.properties.append({
            "client": cid(client), "name": name, "type": kind,
            "address": address, "beds": beds, "baths": baths,
            "rooms": rooms, "area": area, "condition": condition,
            "access": access, "key_status": key, "parking": parking,
            "pets": pets, "floors": floors, "hours": hours,
            "distance": distance, "checklist": checklist,
            "supplies_onsite": onsite, "active": "Yes",
            "entry_notes": ENTRY_NOTES.get(access, ""),
            "notes": notes,
        })

    # ----------------------------------------------------------- services --
    for (name, category, model, hours, price, unit, uplift, supplies,
         popular, description) in SERVICES:
        s.services.append({
            "name": name, "category": category, "model": model,
            "hours": hours, "price": price, "unit": unit,
            "uplift": uplift, "supplies": supplies, "popular": popular,
            "active": "Yes", "description": description,
        })

    # ---------------------------------------------------------- employees --
    for (name, role, kind, phone, email, start, pay, hours, status,
         skills, checked, notes) in EMPLOYEES:
        s.employees.append({
            "name": name, "role": role, "type": kind, "phone": phone,
            "email": email, "start": d(today, start), "pay": pay,
            "hours": hours, "status": status, "skills": skills,
            "checked": checked, "notes": notes,
            "licence": "Yes" if role != "Cleaner" else "No",
            "insured": "Yes" if kind == "Subcontractor" else "N/A",
        })

    # ----------------------------------------------------------- the diary --
    job_index: Dict[str, int] = {}
    rows: List[Dict[str, Any]] = []

    def add_job(**row) -> int:
        rows.append(row)
        return len(rows)

    first_day = monday - _dt.timedelta(days=35)
    last_day = monday + _dt.timedelta(days=27)

    for (client, prop, service, weekday, start, step, cleaner, mate,
         crew, frequency) in ROUND:
        day = first_day + _dt.timedelta(
            days=(weekday - first_day.weekday()) % 7)
        stop = ROUND_STOPS.get((client, prop))
        while day <= last_day:
            if stop is not None and day > d(today, stop):
                break
            if day < today:
                status = "Completed"
            elif day == today:
                status = "In progress"
            elif (day - today).days <= 5:
                status = "Confirmed"
            else:
                status = "Booked"
            rating = ""
            feedback = ""
            actual = None
            if status == "Completed" and rnd.random() < 0.45:
                rating = rnd.choice([5, 5, 5, 4])
            if status == "Completed" and rnd.random() < 0.6:
                actual = None
            distance = PROPERTY_EXTRA.get(prop, (5.0,))[0]
            add_job(
                date=day, start=_dt.time(*start), client=cid(client),
                property=pid(prop), service=service,
                type="Recurring visit", frequency=frequency,
                status=status, priority="Normal", employee=cleaner,
                employee2=mate, cleaners=crew, hours_input=None,
                price_input=None, travel=round(distance * 2, 1),
                recurring="Yes",
                confirmed="Yes" if status in ("Completed", "Confirmed",
                                              "In progress") else "No",
                actual_input=actual, rating=rating, feedback=feedback,
                notes="")
            day += _dt.timedelta(days=step)

    for (tag, offset, client, prop, service, kind, start, status,
         cleaner, mate, crew, hours, price, rating, feedback,
         notes) in ONE_OFFS:
        day = d(today, offset)
        distance = PROPERTY_EXTRA.get(prop, (4.0,))[0] if prop else 3.0
        n = add_job(
            date=day, start=_dt.time(*start),
            client=cid(client) if client else "",
            property=pid(prop) if prop else "", service=service,
            type=kind,
            frequency="One-off" if kind != "Recurring visit"
            else "Weekly",
            status=status,
            priority="High" if kind == "One-off clean" else "Normal",
            employee=cleaner, employee2=mate, cleaners=crew,
            hours_input=hours, price_input=price,
            travel=round(distance * 2, 1), recurring="No",
            confirmed="Yes" if status != "Booked" else "No",
            actual_input=None, rating=rating, feedback=feedback,
            notes=notes)
        job_index[tag] = n

    rows.sort(key=lambda r: (r["date"], str(r["start"])))
    s.jobs = rows
    # the sort moved everything, so rebuild the tag index
    job_index = {}
    for tag, offset, client, prop, service, kind, start, *_ in ONE_OFFS:
        day = d(today, offset)
        for i, row in enumerate(rows, 1):
            if (row["date"] == day and row["start"] == _dt.time(*start)
                    and row["service"] == service
                    and row["type"] == kind):
                job_index[tag] = i
                break

    def job_no(tag: str) -> str:
        n = job_index.get(tag)
        return jid(n) if n else ""

    # -------------------------------------------------------- timesheets --
    for n, row in enumerate(rows, 1):
        if row["status"] not in ("Completed", "In progress"):
            continue
        if not row["employee"]:
            continue
        if row["date"] < monday - _dt.timedelta(days=56):
            # every finished visit in the history gets its hours, so
            # the payroll, the utilisation and the job costing all
            # cover the same period
            continue
        # the hours a visit is planned at - the same number the Jobs
        # tab prices from, so the Timesheets tab cannot quietly
        # contradict it. Real visits run a little over or under.
        prop_hours = 2.5
        if row["property"]:
            prop_hours = PROPERTIES[int(row["property"][3:]) - 1][14]
        planned = (row["hours_input"] or SERVICE_HOURS.get(row["service"])
                   or prop_hours)
        planned = planned * rnd.uniform(0.92, 1.14)
        hours = round(planned / max(1, row["cleaners"]), 2)
        start = row["start"]
        begin = _dt.datetime.combine(row["date"], start)
        wobble = rnd.choice([0, 0, 0, 15, -15, 30])
        finish = begin + _dt.timedelta(hours=hours, minutes=wobble)
        late = rnd.random() < 0.08
        week_old = (today - row["date"]).days
        s.timesheets.append({
            "date": row["date"], "employee": row["employee"],
            "job": jid(n), "activity": "", "start": start,
            "finish": finish.time(),
            "break_min": 30 if hours >= 5 else 0,
            "hours_input": None, "overtime": None,
            "attendance": "Late" if late else "On time",
            "billable": "Yes", "approved": "Yes" if week_old > 7
            else "No", "paid": "Yes" if week_old > 14 else "No",
            "notes": "Traffic on the bridge." if late else "",
        })
        if row["employee2"]:
            s.timesheets.append({
                "date": row["date"], "employee": row["employee2"],
                "job": jid(n), "activity": "", "start": start,
                "finish": finish.time(), "break_min": 0,
                "hours_input": None, "overtime": None,
                "attendance": "On time", "billable": "Yes",
                "approved": "Yes" if week_old > 7 else "No",
                "paid": "Yes" if week_old > 14 else "No", "notes": "",
            })

    # a few non-cleaning rows so the utilisation figure is honest
    for offset, who, activity, start, hours, note in [
            (-9, "Dana Whitfield", "Quote visit", (17, 0), 1.5,
             "Two quote visits and the drive between them."),
            (-6, "Marta Kovacs", "Training", (9, 0), 3.0,
             "COSHH refresher with Amara."),
            (-4, "Luke Bennett", "Supply run", (8, 0), 1.0,
             "Stock collection from CleanCo."),
            (-2, "Dana Whitfield", "Admin errand", (14, 0), 2.0,
             "Invoicing and chasing two overdue accounts.")]:
        begin = _dt.datetime.combine(d(today, offset), _dt.time(*start))
        s.timesheets.append({
            "date": d(today, offset), "employee": who,
            "job": "", "activity": activity, "start": _dt.time(*start),
            "finish": (begin + _dt.timedelta(hours=hours)).time(),
            "break_min": 0, "hours_input": None, "overtime": None,
            "attendance": "On time", "billable": "No",
            "approved": "Yes", "paid": "No", "notes": note,
        })
    s.timesheets.append({
        "date": d(today, -13), "employee": "Amara Diallo", "job": "",
        "activity": "", "start": None, "finish": None,
        "break_min": 0, "hours_input": 0, "overtime": None,
        "attendance": "Sick", "billable": "No", "approved": "Yes",
        "paid": "Yes", "notes": "Covered by Marta at the nursery.",
    })

    # ----------------------------------------------------------- invoices --
    groups: Dict[Any, List[int]] = {}
    for n, row in enumerate(rows, 1):
        if row["status"] != "Completed":
            continue
        if row["price_input"] == 0:
            continue                      # free re-clean, nothing to bill
        if (today - row["date"]).days < 4:
            continue                      # done, not billed yet
        key = (row["client"], row["date"].year, row["date"].month)
        groups.setdefault(key, []).append(n)

    ordered = sorted(groups.items(),
                     key=lambda kv: (kv[0][1], kv[0][2], kv[0][0]))
    overdue_slots = {2, 5}
    part_paid_slot = 8
    for slot, ((client, year, month), job_nos) in enumerate(ordered):
        last_job = max(rows[n - 1]["date"] for n in job_nos)
        raised = min(last_job + _dt.timedelta(days=1), today)
        age = (today - raised).days
        client_no = int(client[3:])
        terms = CLIENTS[client_no - 1][14]
        commercial = CLIENTS[client_no - 1][1] != "Residential"
        if age > terms + 7 and slot in overdue_slots:
            status = "Overdue"
        elif slot == part_paid_slot and age > 5:
            status = "Part paid"
        elif age >= terms:
            status = "Paid"
        elif age >= 1:
            status = "Sent"
        else:
            status = "Draft"
        invoice_no = len(s.invoices) + 1
        s.invoices.append({
            "date": raised, "client": client,
            "period": _dt.date(year, month, 1).strftime("%B %Y"),
            "status": status, "sent_date": raised if status != "Draft"
            else None, "terms_input": None, "discount": None,
            "method": "Bank transfer" if commercial else "Card",
            "po": f"PO-{4000 + slot}" if commercial else "",
            "reminders": 2 if status == "Overdue" else
            (1 if status == "Part paid" else None),
            "last_reminder": d(today, -3) if status == "Overdue"
            else None,
            "notes": "Monthly account." if commercial else "",
        })
        for n in job_nos:
            s.invoice_lines.append({
                "invoice": iid(invoice_no), "job": jid(n),
                "description": "", "qty": 1, "price_input": None,
                "taxable": "Yes", "notes": "",
            })
            rows[n - 1]["status"] = "Invoiced"
        # ------------------------------------------------------ payments --
        if status in ("Paid", "Part paid"):
            paid_on = min(raised + _dt.timedelta(
                days=max(3, terms - rnd.choice([0, 1, 2, 5]))), today)
            s.payments.append({
                "date": paid_on, "invoice": iid(invoice_no),
                "amount": None, "fee": None,
                "method": "Bank transfer" if commercial else "Card",
                "reference": f"BT-{8800 + invoice_no}" if commercial
                else f"SUMUP-{3300 + invoice_no}",
                "deposit": "No", "banked": "Yes",
                "notes": "" if status == "Paid"
                else "Part payment - balance promised next week.",
            })

    # the payment amounts: full for Paid, 60% for Part paid
    for payment in s.payments:
        invoice_no = int(payment["invoice"][4:])
        invoice = s.invoices[invoice_no - 1]
        total = 0.0
        for line in s.invoice_lines:
            if line["invoice"] != payment["invoice"]:
                continue
            job = rows[int(line["job"][4:]) - 1]
            total += _job_price(cfg, job)
        if invoice["status"] == "Part paid":
            payment["amount"] = round(total * 0.6, 2)
        else:
            payment["amount"] = round(total, 2)
        if payment["method"] == "Card":
            payment["fee"] = round(payment["amount"] * 0.0175, 2)

    # a deposit on the end of tenancy job, taken before the work
    if job_no("J_tenancy"):
        s.payments.append({
            "date": d(today, -12), "invoice": "", "amount": 100.00,
            "fee": 1.75, "method": "Card", "reference": "SUMUP-3290",
            "deposit": "Yes", "banked": "Yes",
            "notes": "Deposit to hold the date for the end of tenancy "
                     "clean.",
        })

    # ------------------------------------------------------------- quotes --
    quote_index = {tag: n for n, (tag, *_) in enumerate(QUOTES, 1)}
    for (tag, offset, client, prospect, prop, address, service,
         frequency, status, sent, followed, decided, discount, job_tag,
         decline, notes, lines) in QUOTES:
        s.quotes.append({
            "date": d(today, offset),
            "client": cid(client) if client else "",
            "prospect": prospect,
            "property": pid(prop) if prop else "",
            "address": address, "service": service,
            "frequency": frequency, "status": status,
            "sent_date": d(today, sent) if sent != "" else None,
            "followed_up": d(today, followed) if followed != ""
            else None,
            "decision_date": d(today, decided) if decided != ""
            else None,
            "discount": discount or None, "deposit_input": None,
            "valid_input": None, "job": job_no(job_tag) if job_tag
            else "", "decline_reason": decline, "notes": notes,
            "terms": "",
        })
        for (line_service, qty, price, line_discount, optional,
             wording) in lines:
            s.quote_lines.append({
                "quote": qid(quote_index[tag]), "service": line_service,
                "description": wording, "qty": qty,
                "price_input": price, "discount": line_discount or None,
                "optional": optional, "notes": "",
            })

    # -------------------------------------------------------------- leads --
    for (offset, name, phone, email, address, kind, source, service,
         frequency, stage, value, quote, visit, contacted, action,
         next_offset, client, lost, notes) in LEADS:
        s.leads.append({
            "date": d(today, offset), "name": name, "phone": phone,
            "email": email, "address": address, "type": kind,
            "source": source, "service": service,
            "frequency": frequency, "stage": stage, "value": value,
            "quote": quote, "visit_date": d(today, visit)
            if visit != "" else None, "contacted": contacted,
            "next_action": action,
            "next_date": d(today, next_offset) if next_offset != ""
            else None,
            "client": cid(client) if client else "",
            "lost_reason": lost, "notes": notes,
        })

    # --------------------------------------------------------- checklists --
    for service, tasks in CHECKLIST_TASKS.items():
        for order, (area, task, detail, minutes, frequency, required,
                    photo) in enumerate(tasks, 1):
            s.checklists.append({
                "template": service, "property": "", "area": area,
                "order": order, "task": task, "detail": detail,
                "minutes": minutes, "frequency": frequency,
                "required": required, "photo": photo,
                "products": "", "active": "Yes", "notes": "",
            })
    # one property specific extra
    s.checklists.append({
        "template": "Regular home clean", "property": pid(1),
        "area": "Whole property", "order": 17,
        "task": "Put the recycling box out on a Tuesday",
        "detail": "Box at the end of the drive, lid weighted down",
        "minutes": 3, "frequency": "Weekly", "required": "Yes",
        "photo": "No", "products": "", "active": "Yes",
        "notes": "Only at the Harper house - they asked specially.",
    })

    # --------------------------------------------------------- job tasks --
    recent_home = [n for n, row in enumerate(rows, 1)
                   if row["service"] == "Regular home clean"
                   and row["status"] in ("Completed", "Invoiced")
                   and (today - row["date"]).days <= 9]
    for n in recent_home[:2]:
        row = rows[n - 1]
        for order, (area, task, detail, minutes, frequency, required,
                    photo) in enumerate(
                        CHECKLIST_TASKS["Regular home clean"], 1):
            roll = rnd.random()
            status = "Done"
            quality = "Good"
            issue = ""
            reported = "No"
            if roll > 0.93:
                status = "Issue found"
                quality = "Acceptable"
                issue = "Limescale is building up again - needs a "\
                        "descale next visit."
                reported = "Yes"
            elif roll > 0.88:
                status = "Not needed"
                quality = ""
            s.job_tasks.append({
                "job": jid(n), "area": area, "task": task,
                "status": status, "employee": row["employee"],
                "minutes": minutes + rnd.choice([-2, 0, 0, 0, 3]),
                "quality": quality, "issue": issue,
                "reported": reported, "photo": "", "notes": "",
            })
    office = [n for n, row in enumerate(rows, 1)
              if row["service"] == "Office clean - small"
              and row["status"] in ("Completed", "Invoiced")
              and (today - row["date"]).days <= 6]
    for n in office[:1]:
        row = rows[n - 1]
        for area, task, detail, minutes, frequency, required, photo in \
                CHECKLIST_TASKS["Office clean - small"]:
            s.job_tasks.append({
                "job": jid(n), "area": area, "task": task,
                "status": "Done", "employee": row["employee"],
                "minutes": minutes, "quality": "Excellent",
                "issue": "", "reported": "No", "photo": "",
                "notes": "",
            })

    # ---------------------------------------------------------- supplies --
    for (name, category, supplier, unit, pack_size, pack_cost, stock,
         reorder, order_qty, eco, coshh, notes) in SUPPLIES:
        s.supplies.append({
            "name": name, "category": category, "supplier": supplier,
            "unit": unit, "pack_size": pack_size,
            "pack_cost": pack_cost, "stock": stock,
            "count_date": d(today, -14), "reorder": reorder,
            "order_qty": order_qty,
            "last_order": d(today, -rnd.choice([12, 19, 26, 33])),
            "eco": eco, "coshh": coshh, "notes": notes,
        })

    # ------------------------------------------------------ supply usage --
    usable = [n for n, row in enumerate(rows, 1)
              if row["status"] in ("Completed", "Invoiced")
              and (today - row["date"]).days <= 28]
    picks = ["Multi-surface cleaner", "Microfibre cloths - colour coded",
             "Bin bags - heavy duty", "Toilet cleaner", "Glass cleaner",
             "Nitrile gloves", "Bathroom descaler",
             "Toilet rolls - commercial", "Hand soap refill",
             "Kitchen degreaser"]
    for i, n in enumerate(usable[::3][:22]):
        row = rows[n - 1]
        item = picks[i % len(picks)]
        s.supply_usage.append({
            "date": row["date"], "supply": item,
            "qty": rnd.choice([0.25, 0.5, 1, 1, 2]),
            "job": jid(n), "employee": row["employee"],
            "reason": "Used on the job", "notes": "",
        })
    s.supply_usage.append({
        "date": d(today, -6), "supply": "Oven cleaning gel", "qty": 1,
        "job": job_no("J_tenancy"), "employee": "Priya Raman",
        "reason": "Used on the job",
        "notes": "Whole bottle on that oven. Charged for in the price.",
    })
    s.supply_usage.append({
        "date": d(today, -9), "supply": "Laundry detergent", "qty": 0.5,
        "job": "", "employee": "Jade Ellis", "reason": "Stock take",
        "notes": "Changeover linen for the week.",
    })

    # --------------------------------------------------------- equipment --
    for (name, kind, make, serial, ownership, status, assigned,
         location, bought, cost, life, warranty, last_service,
         interval, service_cost, safety_test, insured,
         notes) in EQUIPMENT:
        s.equipment.append({
            "name": name, "type": kind, "make": make, "serial": serial,
            "ownership": ownership, "status": status,
            "assigned": assigned, "location": location,
            "bought": d(today, bought), "cost": cost,
            "life_years": life,
            "warranty": d(today, warranty) if warranty != "" else None,
            "last_service": d(today, last_service)
            if last_service != "" else None,
            "interval_input": interval, "service_cost": service_cost,
            "safety_test": d(today, safety_test) if safety_test != ""
            else None, "insured": insured, "notes": notes,
        })

    # ---------------------------------------------------------- expenses --
    for back in (2, 1, 0):
        month_day = _dt.date(today.year, today.month, 1)
        for _ in range(back):
            month_day = (month_day - _dt.timedelta(days=1)).replace(day=1)
        for (category, kind, description, supplier, amount, tax,
             method, receipt, recurring, deductible,
             day_of_month) in MONTHLY_COSTS:
            try:
                when = month_day.replace(day=day_of_month)
            except ValueError:
                continue
            if when > today:
                continue
            wobble = 1.0
            if kind == "Variable":
                wobble = rnd.uniform(0.85, 1.2)
            s.expenses.append({
                "date": when, "category": category, "type": kind,
                "description": description, "supplier": supplier,
                "amount": round(amount * wobble, 2), "tax": tax or None,
                "method": method, "job": "", "employee": "",
                "receipt": receipt, "recurring": recurring,
                "deductible": deductible, "notes": "",
            })
    for (offset, category, kind, description, supplier, amount, tax,
         method, receipt, recurring, deductible) in ONE_OFF_COSTS:
        s.expenses.append({
            "date": d(today, offset), "category": category,
            "type": kind, "description": description,
            "supplier": supplier, "amount": amount, "tax": tax or None,
            "method": method, "job": "", "employee": "",
            "receipt": receipt, "recurring": recurring,
            "deductible": deductible, "notes": "",
        })
    # wages, the biggest line of all
    # Wages go out every Friday, and each run is sized from the hours
    # that were really worked that week - so the books and the
    # Timesheets tab tell the same story instead of two different
    # ones.
    cost_an_hour = cfg.pay_rate * (1 + cfg.employer_oncost)
    worked: dict = {}
    for sheet_row in s.timesheets:
        start, finish = sheet_row["start"], sheet_row["finish"]
        if start is None or finish is None:
            # a row clocked by hand, with the hours typed straight in
            typed = sheet_row.get("hours_input") or 0
            if typed:
                worked[sheet_row["date"]] = (
                    worked.get(sheet_row["date"], 0.0) + float(typed))
            continue
        minutes = ((finish.hour * 60 + finish.minute)
                   - (start.hour * 60 + start.minute)
                   - (sheet_row.get("break_min") or 0))
        if minutes > 0:
            worked[sheet_row["date"]] = (worked.get(sheet_row["date"], 0.0)
                                         + minutes / 60.0)
    if worked:
        # the Friday on or before today, then back a fortnight at a time
        payday = today - _dt.timedelta(days=(today.weekday() - 4) % 7)
        earliest = min(worked)
        while payday > earliest:
            opened = payday - _dt.timedelta(days=7)
            hours = sum(value for day, value in worked.items()
                        if opened <= day < payday)
            if hours > 0:
                s.expenses.append({
                    "date": payday, "category": "Wages & payroll",
                    "type": "Fixed monthly",
                    "description": "Weekly payroll run",
                    "supplier": "Payroll",
                    "amount": round(hours * cost_an_hour, 2),
                    "tax": None, "method": "Bank transfer", "job": "",
                    "employee": "", "receipt": "Yes",
                    "recurring": "Yes", "deductible": "Yes",
                    "notes": f"{hours:,.1f} cleaner hours from the "
                             f"Timesheets tab, plus employer "
                             f"on-costs.",
                })
            payday = opened

    # ----------------------------------------------------------- mileage --
    for i, n in enumerate(usable[::4][:16]):
        row = rows[n - 1]
        prop_no = int(row["property"][3:]) if row["property"] else 1
        distance = PROPERTY_EXTRA.get(prop_no, (5.0,))[0]
        s.mileage.append({
            "date": row["date"], "employee": row["employee"],
            "vehicle": "Van 1" if i % 2 == 0 else "Van 2",
            "purpose": "Travel to job", "from_place": "Base",
            "to_place": PROPERTIES[prop_no - 1][1], "job": jid(n),
            "distance": distance, "return_trip": "Yes",
            "rate_input": None, "parking": 0 if i % 5 else 4.50,
            "claimed": "Yes" if (today - row["date"]).days > 14
            else "No", "notes": "",
        })
    for offset, who, purpose, frm, to, distance, parking, note in [
            (-9, "Dana Whitfield", "Quote visit", "Base",
             "Marsh & Fletcher Dental", 7.4, 3.00,
             "Two quotes in one trip."),
            (-4, "Luke Bennett", "Supply run", "Base",
             "CleanCo Supplies", 11.2, 0, ""),
            (-16, "Dana Whitfield", "Bank & post", "Base",
             "Bank and post office", 3.8, 2.00, ""),
            (-6, "Marta Kovacs", "Training", "Base",
             "BICSc training centre", 14.6, 0, "COSHH refresher.")]:
        s.mileage.append({
            "date": d(today, offset), "employee": who,
            "vehicle": "Car", "purpose": purpose, "from_place": frm,
            "to_place": to, "job": "", "distance": distance,
            "return_trip": "Yes", "rate_input": None,
            "parking": parking or None, "claimed": "Yes",
            "notes": note,
        })

    # -------------------------------------------------------- complaints --
    for (offset, client, job_tag, employee, kind, severity,
         description, status, owner, action, resolution, reclean,
         cost, resolved, kept, rating_after, prevent) in COMPLAINTS:
        s.complaints.append({
            "date": d(today, offset), "client": cid(client),
            "job": job_no(job_tag) if job_tag else "",
            "employee": employee, "type": kind, "severity": severity,
            "description": description, "status": status,
            "owner": owner, "action": action, "resolution": resolution,
            "reclean_job": job_no(reclean) if reclean else "",
            "cost": cost or None,
            "resolved_date": d(today, resolved) if resolved != ""
            else None, "kept": kept, "rating_after": rating_after
            or None, "prevent": prevent, "notes": "",
        })

    # --------------------------------------------------------- incidents --
    for (offset, when, kind, employee, job_tag, prop, severity,
         description, injury, first_aid, medical, damage, cost,
         insurance, reported_to, action, prevent, status,
         closed) in INCIDENTS:
        s.incidents.append({
            "date": d(today, offset), "time": _dt.time(*when),
            "type": kind, "employee": employee,
            "job": job_no(job_tag) if job_tag else "",
            "property": pid(prop) if prop else "", "severity": severity,
            "description": description, "injury": injury,
            "first_aid": first_aid, "medical": medical,
            "damage": damage, "cost": cost or None,
            "insurance": insurance, "reported_to": reported_to,
            "reported_date": d(today, offset), "action": action,
            "prevent": prevent, "status": status,
            "closed_date": d(today, closed) if closed != "" else None,
            "notes": "",
        })

    # ------------------------------------------------------------ safety --
    for (document, applies, employee, reference, issuer, issued,
         expires, cost, holder, required, reviewed,
         notes) in SAFETY_DOCS:
        s.safety.append({
            "document": document, "applies_to": applies,
            "employee": employee, "reference": reference,
            "issuer": issuer, "issued": d(today, issued),
            "expires": d(today, expires) if expires != "" else None,
            "cost": cost or None, "holder": holder, "link": "",
            "required": required,
            "reviewed": d(today, reviewed) if reviewed != "" else None,
            "notes": notes,
        })

    return s


def _job_price(cfg: Config, job: Dict[str, Any]) -> float:
    """Mirror the Jobs price formula closely enough for the demo totals."""
    if job.get("price_input") is not None:
        return float(job["price_input"])
    service = job.get("service", "")
    prop = job.get("property", "")
    hours = job.get("hours_input")
    model = price = svc_hours = None
    for row in SERVICES:
        if row[0] == service:
            model, svc_hours, price = row[2], row[3], row[4]
            break
    if hours is None:
        if prop:
            hours = PROPERTIES[int(prop[3:]) - 1][14]
        else:
            hours = svc_hours or 2.0
        if model and model != "Per hour" and svc_hours:
            hours = svc_hours
    if model == "Flat price per visit" and price:
        base = float(price)
    else:
        client_no = int(job["client"][3:]) if job.get("client") else 0
        rate = CLIENTS[client_no - 1][13] if client_no else 45
        base = float(hours) * float(rate)
    factor = {
        "Weekly": 0.85, "Fortnightly": 0.90, "Every 3 weeks": 0.95,
        "Monthly": 1.00, "Quarterly": 1.05, "Twice a week": 0.80,
    }.get(job.get("frequency", "One-off"), 1.00)
    return round(max(cfg.minimum_charge, base * factor), 2)
