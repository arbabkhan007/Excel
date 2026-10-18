"""`Lists` - every drop down in the workbook, plus the pricing factor tables.

Each list becomes a workbook level defined name, so a column can point at
it with ``validate=list_from("JobStatusList")``. Edit a list here and
every drop down that uses it follows.

The two little tables at the bottom are the engine behind the quote
calculator: one turns a cleaning frequency into a price factor and a
number of days between visits, the other turns "how dirty is it" into a
time factor. Change a number there and every quote, every recurring
booking and every price on the workbook re-prices itself.
"""

from __future__ import annotations

from typing import List, Tuple

from ..config import (CONDITION_FACTORS, FREQUENCIES, FREQUENCY_FACTORS,
                      INVOICE_STATUSES, JOB_STATUSES, QUOTE_STATUSES,
                      Config, Palette)
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Lists"
LAST_COL = 9

#: name -> (header, [values]) - written one list per column block
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("YesNoNAList", ["Yes", "No", "N/A"]),
    ("ClientTypeList", ["Residential", "Commercial office", "Retail",
                        "Restaurant & cafe", "Holiday let / Airbnb",
                        "Landlord", "Letting agent", "Property manager",
                        "Builder / contractor", "Medical & dental",
                        "Gym & studio", "School & nursery", "Church & hall",
                        "Warehouse", "Other"]),
    ("ClientStatusList", ["Prospect", "New", "Active", "Regular",
                          "Occasional", "Paused", "Inactive", "Lost",
                          "Do not serve"]),
    ("PropertyTypeList", ["House", "Apartment", "Townhouse", "Bungalow",
                          "Studio", "Office suite", "Shop", "Restaurant",
                          "Clinic", "Salon", "Gym", "School", "Warehouse",
                          "Communal area", "Holiday let", "Other"]),
    ("AccessList", ["Client at home", "We hold a key", "Lockbox",
                    "Keypad code", "Smart lock", "Concierge / front desk",
                    "Reception sign-in", "Key collected each visit",
                    "Neighbour holds key", "Agent meets us", "Open access"]),
    ("KeyStatusList", ["Held by us", "With the cleaner", "In the key safe",
                       "Returned to client", "Not applicable", "Lost"]),
    ("ParkingList", ["Driveway", "Free street parking", "Metered street",
                     "Permit needed", "Visitor bay", "Car park",
                     "Loading bay", "Underground garage", "Difficult",
                     "Pay and display"]),
    ("ConditionList", list(CONDITION_FACTORS)),
    ("FrequencyList", list(FREQUENCIES)),
    ("ServiceCategoryList", ["Regular domestic clean", "Deep clean",
                             "End of tenancy", "Move in / move out",
                             "Commercial & office", "Spring clean",
                             "Post-construction", "Airbnb changeover",
                             "Carpet & upholstery", "Window cleaning",
                             "Oven cleaning", "Pressure washing",
                             "Add-on / extra", "One-off special"]),
    ("PricingModelList", ["Per hour", "Flat price per visit",
                          "Per area", "Per room", "Per bedroom & bathroom",
                          "Per item", "Monthly contract", "Quote only"]),
    ("JobStatusList", list(JOB_STATUSES)),
    ("JobTypeList", ["Recurring visit", "One-off clean", "First deep clean",
                     "Quote visit", "Re-clean (free)", "Emergency call out",
                     "Contract work"]),
    ("PriorityList", ["Normal", "High", "Urgent", "VIP client", "Flexible"]),
    ("TimeSlotList", ["Early (7-9)", "Morning (9-12)", "Midday (12-2)",
                      "Afternoon (2-5)", "Late (5-8)", "Evening",
                      "Flexible", "Client to confirm"]),
    ("QuoteStatusList", list(QUOTE_STATUSES)),
    ("ContractStatusList", ["No contract", "Draft", "Sent", "Signed",
                            "Active", "Due for renewal", "Renewed",
                            "Expired", "Terminated", "Declined"]),
    ("ContractTermList", ["Rolling monthly", "Trial period", "3 months",
                          "6 months", "12 months", "24 months",
                          "No fixed term"]),
    ("BillingCycleList", ["Pay per visit", "Weekly", "Fortnightly",
                          "Monthly in arrears", "Monthly in advance",
                          "Quarterly", "On completion",
                          "Deposit then balance"]),
    ("InvoiceStatusList", list(INVOICE_STATUSES)),
    ("PaymentMethodList", ["Bank transfer", "Card", "Cash", "Cheque",
                           "Direct debit", "Standing order", "PayPal",
                           "Online payment link", "Mobile wallet",
                           "Account credit", "Other"]),
    ("LeadSourceList", ["Google search", "Google Business Profile",
                        "Facebook", "Instagram", "Nextdoor", "Yelp",
                        "Directory listing", "Referral - client",
                        "Referral - staff", "Flyer / leaflet",
                        "Van signage", "Word of mouth", "Repeat client",
                        "Letting agent", "Walk-up", "Other"]),
    ("LeadStageList", ["New enquiry", "Contacted", "Visit booked",
                       "Quote sent", "Following up", "Negotiating", "Won",
                       "Lost", "On hold", "Not a fit"]),
    ("LostReasonList", ["Price too high", "Chose a competitor",
                        "No response", "Out of our area", "Wrong timing",
                        "Doing it themselves", "Needed a specialist",
                        "Not a fit", "Duplicate enquiry", "Other"]),
    ("TaskAreaList", ["Whole property", "Kitchen", "Bathroom",
                      "Second bathroom", "Bedroom", "Living room",
                      "Dining room", "Hallway & stairs", "Home office",
                      "Utility room", "Conservatory", "Reception",
                      "Open plan office", "Meeting room", "Toilets",
                      "Kitchenette", "Communal area", "Windows",
                      "Exterior", "Garage"]),
    ("TaskStatusList", ["To do", "In progress", "Done", "Skipped",
                        "Not needed", "Issue found", "Needs rework"]),
    ("QualityList", ["Excellent", "Good", "Acceptable", "Needs rework",
                     "Failed"]),
    ("RatingList", ["5 - Delighted", "4 - Happy", "3 - Fine",
                    "2 - Unhappy", "1 - Complaint"]),
    ("EmployeeRoleList", ["Owner", "Manager", "Team leader", "Supervisor",
                          "Cleaner", "Deep clean specialist",
                          "Window cleaner", "Carpet technician", "Driver",
                          "Trainee", "Subcontractor", "Office admin"]),
    ("EmploymentTypeList", ["Full time", "Part time", "Casual",
                            "Zero hours", "Subcontractor", "Agency",
                            "Apprentice", "Owner"]),
    ("EmployeeStatusList", ["Active", "Probation", "Training", "On leave",
                            "Sick", "Working notice", "Left", "Standby"]),
    ("ShiftList", ["Morning", "Afternoon", "Evening", "Night", "All day",
                   "Split shift", "On call", "Day off"]),
    ("AttendanceList", ["On time", "Late", "Left early", "Absent - told us",
                        "Absent - no show", "Holiday", "Sick", "Training",
                        "Day off", "Public holiday"]),
    ("LeaveTypeList", ["Holiday", "Sick", "Unpaid", "Parental",
                       "Bereavement", "Training", "Public holiday",
                       "Jury service"]),
    ("TrainingList", ["Induction", "COSHH & chemicals", "Manual handling",
                      "Working at height", "Infection control",
                      "Customer service", "Deep clean technique",
                      "Carpet machine", "Window technique", "First aid",
                      "Safeguarding", "Driving & vehicle", "Data privacy"]),
    ("SupplyCategoryList", ["Chemicals & detergents", "Cloths & sponges",
                            "Mops & buckets", "Paper & consumables",
                            "Bin bags & waste", "Gloves & PPE",
                            "Vacuum bags & filters", "Laundry",
                            "Specialist products", "Hardware & spares",
                            "Uniform", "Office"]),
    ("SupplyUnitList", ["Each", "Bottle", "Litre", "5L drum", "Gallon",
                        "Box", "Case", "Pack", "Roll", "Pair", "Kg",
                        "Refill", "Sachet"]),
    ("EquipmentTypeList", ["Upright vacuum", "Backpack vacuum",
                           "Wet & dry vacuum", "Carpet machine",
                           "Steam cleaner", "Floor polisher",
                           "Pressure washer", "Window pole system",
                           "Ladder & steps", "Trolley & caddy",
                           "Vehicle", "Power tool", "Other"]),
    ("EquipmentStatusList", ["In service", "In use", "Spare", "In repair",
                             "Awaiting parts", "Out of service", "Retired",
                             "Lost or stolen"]),
    ("OwnershipList", ["Owned outright", "Financed", "Leased", "Rented",
                       "Staff owned", "Client supplied"]),
    ("MaintenanceList", ["Routine service", "Filter change", "Belt change",
                         "Repair", "Deep clean", "Safety test",
                         "Calibration", "Inspection", "Replacement due"]),
    ("ExpenseCategoryList", ["Cleaning supplies", "Equipment purchase",
                             "Equipment repair", "Vehicle fuel",
                             "Vehicle maintenance", "Vehicle insurance",
                             "Business insurance", "Wages & payroll",
                             "Subcontractors", "Marketing & advertising",
                             "Website & hosting", "Software & apps",
                             "Phone & internet", "Office & rent",
                             "Uniforms", "Training & certification",
                             "Licences & memberships",
                             "Accounting & legal", "Bank & card fees",
                             "Laundry", "Waste disposal",
                             "Travel & parking", "Bad debt", "Other"]),
    ("ExpenseTypeList", ["Fixed monthly", "Variable", "One-off",
                         "Job related", "Annual"]),
    ("MileagePurposeList", ["Travel to job", "Between jobs", "Quote visit",
                            "Supply run", "Equipment repair",
                            "Staff pick-up", "Bank & post", "Training",
                            "Admin errand", "Marketing drop", "Other"]),
    ("VehicleList", ["Van 1", "Van 2", "Van 3", "Car", "Staff vehicle",
                     "Hire van", "Bike"]),
    ("ComplaintTypeList", ["Missed area", "Quality below standard",
                           "Arrived late", "Did not turn up",
                           "Damage to property", "Breakage",
                           "Item gone missing", "Staff conduct",
                           "Billing error", "Access problem", "Pet issue",
                           "Communication", "Wrong day", "Other"]),
    ("ComplaintStatusList", ["Logged", "Investigating", "Awaiting client",
                             "Re-clean booked", "Resolved", "Escalated",
                             "Refunded", "Closed - no action"]),
    ("SeverityList", ["Low", "Medium", "High", "Critical"]),
    ("ResolutionList", ["Free re-clean", "Partial refund", "Full refund",
                        "Discount next visit", "Apology", "Staff coaching",
                        "Item replaced", "Insurance claim",
                        "Process changed", "No action needed"]),
    ("IncidentTypeList", ["Slip, trip or fall", "Cut or laceration",
                          "Chemical splash", "Burn or scald",
                          "Manual handling strain", "Property damage",
                          "Equipment failure", "Near miss",
                          "Vehicle incident", "Aggression or abuse",
                          "Lone worker concern", "Illness at work",
                          "Other"]),
    ("SafetyDocList", ["COSHH data sheet", "Risk assessment",
                       "Method statement", "Public liability insurance",
                       "Employer liability insurance",
                       "Background check", "Training certificate",
                       "First aid certificate", "Electrical safety test",
                       "Vehicle insurance", "Driving licence",
                       "Right to work", "Client contract",
                       "Privacy notice", "Business licence"]),
    ("DocStatusList", ["Valid", "Due soon", "Expired", "Missing",
                       "Pending renewal", "Not required"]),
    ("RetentionList", ["Loyal", "Steady", "Slipping", "At risk", "Dormant",
                       "Won back", "Lost"]),
    ("MonthList", ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November",
                   "December"]),
    ("WeekdayList", ["Monday", "Tuesday", "Wednesday", "Thursday",
                     "Friday", "Saturday", "Sunday"]),
    ("WeekStartList", ["Monday", "Sunday"]),
    ("AreaUnitList", ["sq ft", "sq m"]),
    ("DistanceUnitList", ["miles", "km"]),
    ("RoundingList", ["Nearest 1", "Nearest 5", "Nearest 10", "Exact"]),
    ("GoalStatusList", ["Not started", "On track", "Behind", "At risk",
                        "Achieved", "Paused"]),
]

#: frequency -> (days between visits, visits a year, price factor)
#: The factor is why a weekly client pays less per visit than a one-off:
#: a place that is cleaned every week never gets badly dirty.
FREQUENCY_TABLE: List[Tuple[str, int, float, float]] = [
    ("One-off", 0, 1, FREQUENCY_FACTORS["One-off"]),
    ("Weekly", 7, 52, FREQUENCY_FACTORS["Weekly"]),
    ("Fortnightly", 14, 26, FREQUENCY_FACTORS["Fortnightly"]),
    ("Every 3 weeks", 21, 17, FREQUENCY_FACTORS["Every 3 weeks"]),
    ("Monthly", 28, 12, FREQUENCY_FACTORS["Monthly"]),
    ("Quarterly", 91, 4, FREQUENCY_FACTORS["Quarterly"]),
    ("Twice a week", 3, 104, FREQUENCY_FACTORS["Twice a week"]),
]

#: how dirty the place is -> how much longer it takes
CONDITION_TABLE: List[Tuple[str, float]] = list(CONDITION_FACTORS.items())


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.grey_text)

    s.banner("Drop Down Lists & Pricing Factors",
             "Every list the workbook offers you, and the two tables the "
             "quote calculator multiplies by. Edit an entry and the drop "
             "down that uses it changes everywhere.",
             LAST_COL)

    row = 4
    row = s.note(
        row,
        "Add new entries directly under the last one in a column - the "
        "lists are sized with room to grow, and blanks are ignored. Do "
        "not delete a whole column: a drop down pointing at a missing "
        "list stops working.",
        LAST_COL, height=30)
    row += 1

    # ------------------------------------------------------------ the lists
    row = s.section(row, "The drop downs", LAST_COL)
    per_row = 5
    spare = 8                       # blank rows kept inside every name
    top = row
    col = 0
    tallest = 0
    for index, (name, values) in enumerate(LISTS):
        if index and index % per_row == 0:
            top += tallest + 2
            col = 0
            tallest = 0
        s.write(top, col, name.replace("List", ""), fmts.chrome("header"))
        s.set_col(col, col, 24)
        for offset, value in enumerate(values):
            s.write(top + 1 + offset, col, value, fmts.input("text"))
        for offset in range(len(values), len(values) + spare):
            s.write(top + 1 + offset, col, None, fmts.input("text"))
        first = top + 2             # 1 based first value row
        last = top + 1 + len(values) + spare
        book.define_name(
            name,
            f"={quote_sheet(NAME)}!${chr(65 + col)}${first}:"
            f"${chr(65 + col)}${last}")
        tallest = max(tallest, len(values) + spare)
        col += 2
    row = top + tallest + 3

    # ------------------------------------------------------- frequency table
    row = s.section(row, "How often we clean, and what that does to the "
                         "price", LAST_COL)
    row = s.note(
        row,
        "Regular work is worth a discount - a weekly house never gets as "
        "dirty as a one-off, so it takes less time. The factor below is "
        "multiplied by the standard price on every quote and every "
        "recurring booking. 'Days between visits' is what the scheduler "
        "adds to today's date to work out the next clean.",
        LAST_COL, height=44)
    row += 1

    s.table_head(row, 0, ["Frequency", "Days between visits",
                          "Visits a year", "Price factor"],
                 [20, 20, 14, 14])
    first_freq = row + 1
    for index, (freq, days, per_year, factor) in enumerate(FREQUENCY_TABLE):
        r = first_freq + index
        s.write(r, 0, freq, fmts.input("text"))
        s.write(r, 1, days, fmts.input("int"))
        s.write(r, 2, per_year, fmts.input("int"))
        s.write(r, 3, factor, fmts.input("number"))
    spare_freq = 5
    for index in range(len(FREQUENCY_TABLE),
                       len(FREQUENCY_TABLE) + spare_freq):
        r = first_freq + index
        for c, style in enumerate(["text", "int", "int", "number"]):
            s.write(r, c, None, fmts.input(style))
    last_freq = first_freq + len(FREQUENCY_TABLE) + spare_freq

    sheet_ref = quote_sheet(NAME)
    book.define_name("FreqNames",
                     f"={sheet_ref}!$A${first_freq + 1}:$A${last_freq}")
    book.define_name("FreqDays",
                     f"={sheet_ref}!$B${first_freq + 1}:$B${last_freq}")
    book.define_name("FreqPerYear",
                     f"={sheet_ref}!$C${first_freq + 1}:$C${last_freq}")
    book.define_name("FreqFactors",
                     f"={sheet_ref}!$D${first_freq + 1}:$D${last_freq}")
    row = last_freq + 2

    # ------------------------------------------------------- condition table
    row = s.section(row, "How dirty is it, and what that does to the time",
                    LAST_COL)
    row = s.note(
        row,
        "The honest part of pricing a clean. The same three bedroom house "
        "can take three hours or six depending on what state it is in, so "
        "the calculator multiplies the estimated hours by the factor "
        "below. Walk the property, pick a line, and the quote follows.",
        LAST_COL, height=38)
    row += 1

    s.table_head(row, 0, ["Condition", "Time factor", "What it means"],
                 [26, 14, 52])
    meanings = [
        "Kept on top of between visits - a quick, tidy run through.",
        "A normal lived-in home or a tidy office.",
        "Behind on the housework - expect extra scrubbing.",
        "Neglected for months: grease, limescale, heavy build-up.",
        "Specialist job. Price it after a site visit, never over "
        "the phone.",
    ]
    first_cond = row + 1
    for index, (condition, factor) in enumerate(CONDITION_TABLE):
        r = first_cond + index
        s.write(r, 0, condition, fmts.input("text"))
        s.write(r, 1, factor, fmts.input("number"))
        s.write(r, 2,
                meanings[index] if index < len(meanings) else None,
                fmts.input("text"))
    spare_cond = 4
    for index in range(len(CONDITION_TABLE),
                       len(CONDITION_TABLE) + spare_cond):
        r = first_cond + index
        for c, style in enumerate(["text", "number", "text"]):
            s.write(r, c, None, fmts.input(style))
    last_cond = first_cond + len(CONDITION_TABLE) + spare_cond

    book.define_name("ConditionNames",
                     f"={sheet_ref}!$A${first_cond + 1}:$A${last_cond}")
    book.define_name("ConditionFactors",
                     f"={sheet_ref}!$B${first_cond + 1}:$B${last_cond}")
    row = last_cond + 2

    row = s.note(
        row,
        "A word on the factors: they are a starting point from real "
        "cleaning businesses, not a law. Track your actual hours on the "
        "Job Costing tab for a month, compare them with what you quoted, "
        "then nudge these numbers until the quotes match reality.",
        LAST_COL, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
