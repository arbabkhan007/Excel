"""`Lists` - every drop down in the workbook, plus the factor tables.

Each list becomes a workbook level defined name, so a column can point at
it with ``validate=list_from("CategoryList")``. Edit a list here and
every drop down that uses it follows.

The tables at the bottom are the engine behind the calculators: one
converts a buying unit into the unit you actually use (a 100 g skein
into grams of yarn), one turns "how fiddly is this" into extra minutes,
one turns a turnaround promise into a rush surcharge, and one holds the
target margin behind each pricing scenario. Change a number there and
every price in the workbook re-prices itself.
"""

from __future__ import annotations

from typing import List, Tuple

from ..config import (PRICING_SCENARIOS, PRODUCT_CATEGORIES,
                      PRODUCT_STATUSES, ROUNDING_RULES, SALES_CHANNELS,
                      UNIT_CONVERSIONS, Config, Palette)
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Lists"
LAST_COL = 9

#: name -> (header, [values]) - written one list per column block
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("YesNoNAList", ["Yes", "No", "N/A"]),
    ("CategoryList", list(PRODUCT_CATEGORIES)),
    ("StatusList", list(PRODUCT_STATUSES)),
    ("ChannelList", list(SALES_CHANNELS)),
    ("PricingMethodList", ["Cost plus margin", "Cost plus markup",
                           "Target price", "Keystone (x2)",
                           "Market matched", "Hourly rate only",
                           "Fixed price", "Made to order quote"]),
    ("UnitList", [u for u, _b, _f in UNIT_CONVERSIONS]),
    ("MaterialTypeList", ["Yarn & thread", "Fabric", "Beads & findings",
                          "Metal & wire", "Stones & crystals", "Wax",
                          "Fragrance & oils", "Wicks", "Vessels & jars",
                          "Timber & boards", "Finish & glue",
                          "Hardware & fixings", "Clay & glaze",
                          "Paper & card", "Ink & paint", "Leather",
                          "Elastic & notions", "Labels & tags",
                          "Packaging material", "Tools & consumables",
                          "Other supply"]),
    ("StockStatusList", ["In stock", "Low stock", "Out of stock",
                         "On order", "Discontinued", "Not tracked"]),
    ("SeasonList", ["All year", "Spring", "Summer", "Autumn", "Winter",
                    "Christmas", "Valentine's", "Mother's Day",
                    "Halloween", "Wedding season", "Back to school"]),
    ("DifficultyList", ["Very simple", "Simple", "Moderate", "Complex",
                        "Very complex"]),
    ("LaborStageList", ["Design", "Pattern & setup", "Cutting & prep",
                        "Making", "Assembly", "Sewing", "Finishing",
                        "Sanding & smoothing", "Painting & staining",
                        "Curing & drying", "Quality check", "Weaving in",
                        "Blocking", "Photography", "Listing & admin",
                        "Packing", "Post office run", "Customer messages"]),
    ("LaborWhoList", ["Me", "Partner", "Employee", "Contractor",
                      "Apprentice", "Studio assistant", "Outsourced"]),
    ("PerBasisList", ["Per item", "Per batch"]),
    ("PackProfileList", ["Budget", "Standard", "Gift", "Premium"]),
    ("PackTypeList", ["Mailer & box", "Tissue & wrap", "Filler",
                      "Tape & seal", "Label & sticker", "Thank you card",
                      "Care card", "Ribbon & twine", "Gift box",
                      "Polybag", "Insert & flyer", "Fragile protection"]),
    ("ShipServiceList", ["First class", "Priority", "Priority express",
                         "Ground", "Tracked 24", "Tracked 48",
                         "Signed for", "Courier", "International economy",
                         "International tracked", "Local delivery",
                         "Collection / pickup"]),
    ("ShipZoneList", ["Local", "Domestic", "Domestic remote", "Europe",
                      "North America", "Rest of world"]),
    ("ShipPayerList", ["Buyer pays", "Free shipping (I absorb it)",
                       "Free over threshold", "Flat rate",
                       "Calculated at checkout", "Collection only"]),
    ("OverheadTypeList", ["Studio rent", "Utilities", "Insurance",
                          "Shop subscription", "Website & domain",
                          "Software", "Accounting", "Bank & card fees",
                          "Marketing & ads", "Photography props",
                          "Tool replacement", "Equipment lease",
                          "Training & courses", "Memberships",
                          "Phone & internet", "Travel & fuel",
                          "Market stall fees", "Storage", "Other"]),
    ("FrequencyList", ["Monthly", "Quarterly", "Twice a year", "Yearly",
                       "One-off", "Per order"]),
    ("AllocationList", ["Per hour worked", "Per unit made", "Per order",
                        "Share of revenue", "Do not allocate"]),
    ("FeeTypeList", ["Percent of order", "Fixed per order",
                     "Fixed per item", "Listing fee",
                     "Percent + fixed", "Subscription", "None"]),
    ("PromoTypeList", ["Percent off", "Amount off", "Free shipping",
                       "Buy one get one", "Bundle price",
                       "Spend and save", "Sale price", "Coupon code",
                       "Early bird", "Loyalty reward"]),
    ("VariationTypeList", ["Size", "Colour", "Yarn choice", "Scent",
                           "Length", "Chain length", "Metal finish",
                           "Wood species", "Personalisation",
                           "Embroidery", "Gift wrap", "Rush order",
                           "Bundle add-on", "Custom request"]),
    ("BatchStatusList", ["Planned", "Materials ordered", "In progress",
                         "Curing", "Finishing", "Complete",
                         "Listed", "Sold out", "Abandoned"]),
    ("SaleStatusList", ["Paid", "Awaiting payment", "Made", "Shipped",
                        "Delivered", "Refunded", "Partly refunded",
                        "Cancelled", "Exchange", "Gift"]),
    ("WasteReasonList", ["Offcut", "Miscut", "Mistake / frogged",
                         "Failed pour", "Cracked", "Warped",
                         "Dye lot mismatch", "Spoiled in storage",
                         "Expired", "Sample / test piece",
                         "Customer return", "Breakage in post",
                         "Lost", "Normal trim waste"]),
    ("EventTypeList", ["Craft fair", "Farmers market", "Christmas market",
                       "Pop-up shop", "Trade show", "Wholesale market",
                       "Workshop", "Online market", "Charity event",
                       "School fete"]),
    ("CustomerTypeList", ["Independent shop", "Boutique", "Gallery",
                          "Gift shop", "Museum shop", "Online retailer",
                          "Subscription box", "Corporate gifting",
                          "Interior designer", "Market trader",
                          "Distributor"]),
    ("TermsList", ["Payment with order", "50% deposit, 50% on despatch",
                   "Net 7", "Net 14", "Net 30", "Net 60",
                   "Consignment", "Card on file"]),
    ("OrderStatusList", ["Enquiry", "Quoted", "Confirmed", "In production",
                         "Ready", "Despatched", "Invoiced", "Paid",
                         "Cancelled", "On hold"]),
    ("RoundingList", list(ROUNDING_RULES)),
    ("PriceReasonList", ["First price", "Material cost rose",
                         "Material cost fell", "Labour rate change",
                         "Fee change", "Competitor move", "Sale price",
                         "Sale ended", "Rebrand / repositioning",
                         "Rounding tidy-up", "Wholesale request",
                         "Testing a price"]),
    ("QualityList", ["Below mine", "Similar", "Above mine",
                     "Mass produced", "Designer / luxury", "Unknown"]),
    ("DevStageList", ["Idea", "Sketching", "Sampling", "Testing",
                      "Photography", "Listed", "Shelved",
                      "Recovered cost"]),
    ("RatingList", ["5 - Loved it", "4 - Happy", "3 - Fine",
                    "2 - Unhappy", "1 - Complaint"]),
    ("YarnWeightList", ["Lace", "Light fingering", "Fingering / 4 ply",
                        "Sport", "DK", "Worsted / Aran", "Bulky / chunky",
                        "Super bulky", "Jumbo", "Thread"]),
    ("HookSizeList", ["2.0 mm", "2.5 mm", "3.0 mm", "3.5 mm", "4.0 mm",
                      "4.5 mm", "5.0 mm", "5.5 mm", "6.0 mm", "6.5 mm",
                      "8.0 mm", "10.0 mm", "12.0 mm"]),
    ("MetalList", ["Sterling silver", "Silver plated", "Gold filled",
                   "9ct gold", "14ct gold", "Gold plated", "Rose gold",
                   "Brass", "Copper", "Stainless steel", "Titanium",
                   "Aluminium", "Mixed metal"]),
    ("StoneList", ["None", "Glass bead", "Crystal", "Freshwater pearl",
                   "Amethyst", "Rose quartz", "Turquoise", "Onyx",
                   "Moonstone", "Labradorite", "Cubic zirconia",
                   "Resin cabochon", "Sea glass", "Other gemstone"]),
    ("WaxList", ["Soy", "Coconut soy", "Coconut apricot", "Beeswax",
                 "Rapeseed", "Paraffin", "Palm", "Gel", "Wax melt blend"]),
    ("WickList", ["Cotton", "Wood", "Crackling wood", "Hemp", "Eco",
                  "Braided", "Multiple wicks"]),
    ("VesselList", ["Glass jar", "Amber jar", "Tin", "Ceramic vessel",
                    "Concrete vessel", "Teacup", "Tumbler", "Clamshell",
                    "Refill bag", "No vessel (melt)"]),
    ("FabricList", ["Cotton", "Organic cotton", "Linen", "Cotton jersey",
                    "Double gauze", "Denim", "Corduroy", "Wool",
                    "Fleece", "Canvas", "Viscose", "Silk", "Satin",
                    "Ripstop", "Oilcloth", "Deadstock"]),
    ("GarmentSizeList", ["Baby", "Toddler", "Child", "XS", "S", "M", "L",
                         "XL", "2XL", "3XL", "One size", "Made to measure"]),
    ("WoodList", ["Pine", "Oak", "Ash", "Beech", "Birch ply", "Walnut",
                  "Maple", "Cherry", "Sapele", "Reclaimed", "Plywood",
                  "MDF", "Live edge slab", "Offcut stock"]),
    ("FinishList", ["Oil", "Hard wax oil", "Danish oil", "Beeswax polish",
                    "Varnish", "Lacquer", "Stain", "Paint", "Milk paint",
                    "Resin", "Raw / unfinished"]),
    ("DigitalTypeList", ["Crochet pattern", "Knitting pattern",
                         "Sewing pattern", "SVG cut file", "Printable",
                         "Planner", "Clipart set", "Font",
                         "Photo pack", "Course / workshop",
                         "E-book", "Template"]),
    ("FileFormatList", ["PDF", "SVG", "PNG", "JPG", "DXF", "EPS", "AI",
                        "ZIP bundle", "Video", "Canva link"]),
    ("LicenceList", ["Personal use", "Small commercial",
                     "Extended commercial", "Limited run",
                     "No resale of files", "Credit required"]),
    ("BundleTypeList", ["Gift set", "Starter kit", "Collection",
                        "Seasonal set", "Multi-buy", "Mix and match",
                        "Subscription box", "Wholesale case"]),
    ("MonthList", ["January", "February", "March", "April", "May",
                   "June", "July", "August", "September", "October",
                   "November", "December"]),
]

#: buying unit -> the unit you actually count in, and the multiplier
UNIT_TABLE: List[Tuple[str, str, float]] = list(UNIT_CONVERSIONS)

#: how fiddly the piece is -> time factor, and the price factor some
#: makers like to add on top of the honest cost
DIFFICULTY_TABLE: List[Tuple[str, float, float, str]] = [
    ("Very simple", 0.85, 1.00,
     "Repeat work you could do watching television."),
    ("Simple", 1.00, 1.00,
     "A piece you have made many times - no surprises."),
    ("Moderate", 1.15, 1.05,
     "Some shaping, colour changes or fiddly assembly."),
    ("Complex", 1.35, 1.12,
     "Needs concentration, a pattern in front of you, few distractions."),
    ("Very complex", 1.60, 1.25,
     "Heirloom work, a skill few makers have - charge for it."),
]

#: how fast the customer wants it -> what that is worth
RUSH_TABLE: List[Tuple[str, float, str]] = [
    ("Standard", 0.00, "Your normal turnaround, no surcharge."),
    ("Priority", 0.15, "Jumps the queue by a few days."),
    ("Rush", 0.30, "Everything else waits - evenings and a weekend."),
    ("Same week", 0.50, "Dropping other work entirely. Price it properly."),
    ("Next day", 0.75, "Only if you genuinely can, and only at a premium."),
]

#: pricing scenario -> the margin it aims at
SCENARIO_TABLE: List[Tuple[str, float, str]] = [
    ("Budget market", PRICING_SCENARIOS["Budget market"],
     "Craft fair impulse buys and market stalls."),
    ("Conservative", PRICING_SCENARIOS["Conservative"],
     "Safe pricing while you build reviews."),
    ("Standard", PRICING_SCENARIOS["Standard"],
     "What a healthy handmade shop should be aiming for."),
    ("Premium", PRICING_SCENARIOS["Premium"],
     "Distinctive work, good photography, a following."),
    ("Luxury", PRICING_SCENARIOS["Luxury"],
     "Commissions, galleries and one-of-a-kind pieces."),
]

#: size -> how much more material and time it takes
SIZE_TABLE: List[Tuple[str, float, float]] = [
    ("Mini", 0.45, 0.50),
    ("Small", 0.75, 0.80),
    ("Standard", 1.00, 1.00),
    ("Large", 1.35, 1.25),
    ("Extra large", 1.75, 1.55),
    ("Custom size", 1.00, 1.00),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.grey_text)

    s.banner("Drop Down Lists & Pricing Factors",
             "Every list the workbook offers you, and the four tables the "
             "calculators multiply by. Edit an entry and the drop down "
             "that uses it changes everywhere.",
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

    sheet_ref = quote_sheet(NAME)

    # ----------------------------------------------------- unit conversions
    row = s.section(row, "Units: what you buy, what you use", LAST_COL)
    row = s.note(
        row,
        "Yarn arrives in 100 g skeins and gets used by the gram. Wax "
        "arrives by the kilo and gets poured by the gram. This table is "
        "how the Materials tab turns a purchase unit into a usable one, "
        "so a bill of materials can say '38 g' and still cost it "
        "correctly. 'Multiply by' is how many base units one of this "
        "unit contains.",
        LAST_COL, height=44)
    row += 1

    s.table_head(row, 0, ["Unit", "Base unit", "Multiply by"],
                 [18, 16, 16])
    first_unit = row + 1
    for index, (unit, base, factor) in enumerate(UNIT_TABLE):
        r = first_unit + index
        s.write(r, 0, unit, fmts.input("text"))
        s.write(r, 1, base, fmts.input("text"))
        s.write(r, 2, factor, fmts.input("number"))
    spare_unit = 6
    for index in range(len(UNIT_TABLE), len(UNIT_TABLE) + spare_unit):
        r = first_unit + index
        for c, style in enumerate(["text", "text", "number"]):
            s.write(r, c, None, fmts.input(style))
    last_unit = first_unit + len(UNIT_TABLE) + spare_unit

    book.define_name("UnitNames",
                     f"={sheet_ref}!$A${first_unit + 1}:$A${last_unit}")
    book.define_name("UnitBases",
                     f"={sheet_ref}!$B${first_unit + 1}:$B${last_unit}")
    book.define_name("UnitFactors",
                     f"={sheet_ref}!$C${first_unit + 1}:$C${last_unit}")
    row = last_unit + 2

    # ------------------------------------------------------ difficulty table
    row = s.section(row, "How fiddly is it, and what that does to the time",
                    LAST_COL)
    row = s.note(
        row,
        "The same shawl takes one maker four hours and another seven. "
        "Pick a line per product and the labour calculator multiplies "
        "your estimated minutes by the time factor. The price factor is "
        "separate and optional - it is the premium skilled work can "
        "carry on top of its honest cost.",
        LAST_COL, height=44)
    row += 1

    s.table_head(row, 0, ["Difficulty", "Time factor", "Price factor",
                          "What it means"], [20, 14, 14, 54])
    first_diff = row + 1
    for index, (label, time_f, price_f, meaning) in \
            enumerate(DIFFICULTY_TABLE):
        r = first_diff + index
        s.write(r, 0, label, fmts.input("text"))
        s.write(r, 1, time_f, fmts.input("number"))
        s.write(r, 2, price_f, fmts.input("number"))
        s.write(r, 3, meaning, fmts.input("text"))
    spare_diff = 4
    for index in range(len(DIFFICULTY_TABLE),
                       len(DIFFICULTY_TABLE) + spare_diff):
        r = first_diff + index
        for c, style in enumerate(["text", "number", "number", "text"]):
            s.write(r, c, None, fmts.input(style))
    last_diff = first_diff + len(DIFFICULTY_TABLE) + spare_diff

    book.define_name("DiffNames",
                     f"={sheet_ref}!$A${first_diff + 1}:$A${last_diff}")
    book.define_name("DiffTimeFactors",
                     f"={sheet_ref}!$B${first_diff + 1}:$B${last_diff}")
    book.define_name("DiffPriceFactors",
                     f"={sheet_ref}!$C${first_diff + 1}:$C${last_diff}")
    row = last_diff + 2

    # ------------------------------------------------------------ rush table
    row = s.section(row, "Turnaround: what a hurry is worth", LAST_COL)
    row = s.note(
        row,
        "A rush order costs you evenings, not materials. The custom "
        "order and variation tabs add the surcharge below to the price, "
        "so saying yes to 'can I have it by Friday?' stops being a "
        "favour you pay for yourself.",
        LAST_COL, height=38)
    row += 1

    s.table_head(row, 0, ["Turnaround", "Surcharge", "When you use it"],
                 [20, 14, 62])
    first_rush = row + 1
    for index, (label, pct, meaning) in enumerate(RUSH_TABLE):
        r = first_rush + index
        s.write(r, 0, label, fmts.input("text"))
        s.write(r, 1, pct, fmts.input("percent"))
        s.write(r, 2, meaning, fmts.input("text"))
    spare_rush = 3
    for index in range(len(RUSH_TABLE), len(RUSH_TABLE) + spare_rush):
        r = first_rush + index
        for c, style in enumerate(["text", "percent", "text"]):
            s.write(r, c, None, fmts.input(style))
    last_rush = first_rush + len(RUSH_TABLE) + spare_rush

    book.define_name("RushNames",
                     f"={sheet_ref}!$A${first_rush + 1}:$A${last_rush}")
    book.define_name("RushFactors",
                     f"={sheet_ref}!$B${first_rush + 1}:$B${last_rush}")
    row = last_rush + 2

    # -------------------------------------------------------- scenario table
    row = s.section(row, "Pricing scenarios and the margin each one aims at",
                    LAST_COL)
    row = s.note(
        row,
        "The Scenarios tab prices every product five ways at once so you "
        "can see what repositioning would be worth. These are the "
        "margins behind those five columns - change them to match the "
        "market you actually sell into.",
        LAST_COL, height=38)
    row += 1

    s.table_head(row, 0, ["Scenario", "Target margin", "Who prices this way"],
                 [20, 16, 62])
    first_scen = row + 1
    for index, (label, margin, meaning) in enumerate(SCENARIO_TABLE):
        r = first_scen + index
        s.write(r, 0, label, fmts.input("text"))
        s.write(r, 1, margin, fmts.input("percent"))
        s.write(r, 2, meaning, fmts.input("text"))
    last_scen = first_scen + len(SCENARIO_TABLE)

    book.define_name("ScenarioNames",
                     f"={sheet_ref}!$A${first_scen + 1}:$A${last_scen}")
    book.define_name("ScenarioMargins",
                     f"={sheet_ref}!$B${first_scen + 1}:$B${last_scen}")
    row = last_scen + 2

    # ------------------------------------------------------------ size table
    row = s.section(row, "Size: more material, more minutes", LAST_COL)
    row = s.note(
        row,
        "A baby blanket and a king size one are the same pattern and "
        "nowhere near the same price. The Variations tab multiplies the "
        "parent product's material cost and labour time by the factors "
        "below, so a size run prices itself.",
        LAST_COL, height=38)
    row += 1

    s.table_head(row, 0, ["Size", "Material factor", "Time factor"],
                 [20, 18, 16])
    first_size = row + 1
    for index, (label, mat_f, time_f) in enumerate(SIZE_TABLE):
        r = first_size + index
        s.write(r, 0, label, fmts.input("text"))
        s.write(r, 1, mat_f, fmts.input("number"))
        s.write(r, 2, time_f, fmts.input("number"))
    spare_size = 4
    for index in range(len(SIZE_TABLE), len(SIZE_TABLE) + spare_size):
        r = first_size + index
        for c, style in enumerate(["text", "number", "number"]):
            s.write(r, c, None, fmts.input(style))
    last_size = first_size + len(SIZE_TABLE) + spare_size

    book.define_name("SizeNames",
                     f"={sheet_ref}!$A${first_size + 1}:$A${last_size}")
    book.define_name("SizeMaterialFactors",
                     f"={sheet_ref}!$B${first_size + 1}:$B${last_size}")
    book.define_name("SizeTimeFactors",
                     f"={sheet_ref}!$C${first_size + 1}:$C${last_size}")
    row = last_size + 2

    row = s.note(
        row,
        "A word on the factors: they are a starting point from real "
        "handmade businesses, not a law. Track what a batch actually "
        "takes you on the Batch Production tab for a month, compare it "
        "with what you estimated, then nudge these numbers until the "
        "estimates match your bench.",
        LAST_COL, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
