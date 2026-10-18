"""Static checks on a generated Handmade Pricing workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_handmade_workbook.py output/pricing.xlsx

It never opens Excel - it parses the XML and reports anything that would
break for a buyer: unresolved template tokens, formulas pointing at the
wrong sheet, defined names that do not exist, missing sheet protection,
formula cells that are not locked, and input cells that are.
"""

from __future__ import annotations

import os
import re
import sys
import zipfile
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:                                   # pragma: no cover
    print("openpyxl is required: pip install openpyxl")
    raise SystemExit(2)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from handmade_pricing_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402

FUNCTION_RE = re.compile(r"\b([A-Z][A-Z0-9\.]{1,})\s*\(")
NAME_RE = re.compile(r"(?<![A-Za-z0-9_.!$\"])([A-Za-z_][A-Za-z0-9_.]{2,})"
                     r"(?![A-Za-z0-9_.(!])")
#: a column letter with no row number - almost always a missing {r}
BARE_COL_RE = re.compile(r"(?<![A-Za-z0-9_.!$\"])([A-Z]{1,3})"
                         r"(?![A-Za-z0-9_.(!$:])")

KNOWN_FUNCTIONS = {
    "IF", "IFERROR", "SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTA", "COUNTIF",
    "COUNTIFS", "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "MAX", "MIN", "ROUND",
    "ROUNDUP", "ROUNDDOWN", "INDEX", "MATCH", "OFFSET", "TEXT", "VALUE",
    "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "PROPER",
    "CONCATENATE", "TODAY", "NOW", "DATE", "DAY", "MONTH", "YEAR", "WEEKDAY",
    "WEEKNUM", "EOMONTH", "EDATE", "DATEDIF", "DAYS", "NETWORKDAYS", "HOUR",
    "MINUTE", "TIME", "SUMPRODUCT", "LARGE", "SMALL", "RANK", "ABS", "AND",
    "OR", "NOT", "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISNA", "NA",
    "CHOOSE", "VLOOKUP", "HLOOKUP", "LOOKUP", "SUBSTITUTE", "REPT",
    "MEDIAN", "STDEV", "FLOOR", "CEILING", "INT", "MOD", "SIGN", "RAND",
    "SEARCH", "FIND", "EXACT", "N", "T", "TRUE", "FALSE", "ROW", "ROWS",
    "COLUMN", "COLUMNS", "YEARFRAC", "INDIRECT", "ADDRESS", "CELL",
    "SUBTOTAL", "TRANSPOSE", "PRODUCT", "POWER", "SQRT", "COUNTBLANK",
    "ISERR", "ISLOGICAL", "ISREF", "TYPE", "NUMBERVALUE", "DATEVALUE",
    "TIMEVALUE", "WORKDAY", "EOMONTH", "FREQUENCY", "PERCENTILE",
}

#: functions that do not exist in Excel 2010 / Google Sheets
MODERN_ONLY = {
    "XLOOKUP", "IFS", "MAXIFS", "MINIFS", "TEXTJOIN", "CONCAT", "SWITCH",
    "LET", "LAMBDA", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE",
    "TEXTSPLIT", "AGGREGATE", "IFNA",
}


def _colnum(ref: str) -> tuple:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    letters, row = match.group(1), int(match.group(2))
    num = 0
    for char in letters:
        num = num * 26 + (ord(char) - 64)
    return num, row


def check_raw_xml(path: str) -> list:
    """Overlapping merges and stray hyperlinks - Excel calls these corrupt."""
    found = []
    with zipfile.ZipFile(path) as zf:
        book = zf.read("xl/workbook.xml").decode()
        sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="rId(\d+)"',
                            book)
        sheet_names = {name for name, _ in sheets}
        for name, rid in sheets:
            try:
                xml = zf.read(f"xl/worksheets/sheet{rid}.xml").decode()
            except KeyError:
                continue
            merges = re.findall(r'<mergeCell ref="([^"]+)"/>', xml)
            boxes = []
            for ref in merges:
                start, _, end = ref.partition(":")
                end = end or start
                c1, r1 = _colnum(start)
                c2, r2 = _colnum(end)
                boxes.append((c1, r1, c2, r2, ref))
            for i, a in enumerate(boxes):
                for b in boxes[i + 1:]:
                    if (a[0] <= b[2] and b[0] <= a[2]
                            and a[1] <= b[3] and b[1] <= a[3]):
                        found.append(f"{name}: merges overlap "
                                     f"{a[4]} / {b[4]}")
            for location in re.findall(
                    r'<hyperlink ref="[^"]+" location="([^"]+)"', xml):
                target = location.split("!")[0].strip("'").replace("''", "'")
                if target not in sheet_names:
                    found.append(f"{name}: hyperlink points at missing "
                                 f"sheet '{target}'")
            for link in re.findall(r'<hyperlink ref="([^"]+)"', xml):
                lc, lr = _colnum(link)
                for c1, r1, c2, r2, ref in boxes:
                    if (c1 <= lc <= c2 and r1 <= lr <= r2
                            and (lc, lr) != (c1, r1)):
                        found.append(f"{name}: hyperlink {link} is inside "
                                     f"merged range {ref}")
    return found


def paren_depth(formula: str) -> int:
    """Open brackets minus closing ones, ignoring anything inside text.

    Excel escapes a quote inside a string by doubling it, so `""` is not
    the end of the string. Deeply nested IF() chains built by string
    concatenation are the one bug this whole file exists to catch: Excel
    and openpyxl both accept a short formula silently, and the buyer
    sees the damage.
    """
    depth = 0
    in_text = False
    i = 0
    while i < len(formula):
        char = formula[i]
        if in_text:
            if char == '"':
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        i += 1
    return depth


#: divisions that really are meant to read `a + b/c`.  Each one is the
#: formula with its numbers blanked out, so one entry covers the whole
#: column.
PRECEDENCE_OK = {
    # variation labour = the base labour scaled, plus the extra minutes
    # over sixty at the labour rate - only the minutes are divided
    'IF(P#="",#,P#)*U#+IF(G#="",#,G#)/#*LaborCostHour',
    # jewelry consumables = solder, plating and hallmarking, plus bench
    # consumables charged by the hour
    'IF(Q#="",#,Q#)+IF(R#="",#,R#)+IF(S#="",#,S#)+AG#/#*SuppliesPerHour',
    # candle minutes = the minutes for one, plus the batch set-up spread
    # over the batch
    'IF(U#="",#,U#)+IF(T#="",#,T#)/MAX(#,IF(S#="",DefaultBatchSize,S#))',
    # woodworking consumables = glue and abrasives, plus shop
    # consumables by the hour
    'AG#+AD#/#*SuppliesPerHour',
    # a download: development spread over the sales it is expected to
    # make, plus the support minutes every sale really costs
    'V#/W#+IF(M#="",#,M#)/#*T#',
    'AL#+Z#-IF(M#="",#,M#)/#*T#',
    'AK#-AA#-(AK#*AF#+AG#)-IF(M#="",#,M#)/#*T#',
    'AK#-AA#-(AK#*AF#+AG#)-IF(M#="",#,M#)/#*T#<=#',
    # a wholesale line: carriage is per order, so only carriage is
    # divided by the quantity
    'C#-WcCost-WcCarriage/MAX(#,WcQty)',
    'WcPrice-WcCost-WcCarriage/MAX(#,WcQty)',
    'WcCost+WcCarriage/MAX(#,WcQty)',
}


def _top_level(text, wanted):
    """Positions of `wanted` characters outside brackets and strings."""
    out = []
    depth = 0
    in_text = False
    i = 0
    while i < len(text):
        char = text[i]
        if in_text:
            if char == '"':
                if i + 1 < len(text) and text[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif depth == 0 and char in wanted:
            before = text[i - 1] if i else ""
            if char != "/" and (before == "" or before in "(,*/+-^<>=&"):
                # a leading sign, not an operator between two values
                i += 1
                continue
            out.append(i)
        i += 1
    return out


def _arguments(formula):
    """Every comma separated argument, at every bracket depth."""
    starts = {0: 0}
    depth = 0
    in_text = False
    i = 0
    while i < len(formula):
        char = formula[i]
        if in_text:
            if char == '"':
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
            starts[depth] = i + 1
        elif char == ")":
            yield formula[starts.get(depth, 0):i]
            depth -= 1
        elif char == ",":
            yield formula[starts.get(depth, 0):i]
            starts[depth] = i + 1
        i += 1
    yield formula[starts.get(0, 0):]


def precedence_slip(formula):
    """`SUMIFS(..)+SUMIFS(..)/COUNTIFS(..)` - the sum is not divided.

    Excel divides before it adds, so a total built by adding two
    SUMIFS together has to be bracketed before it is divided by
    anything.  The file still opens, every bracket balances and the
    calculation engine returns a number - it is simply the wrong
    number, which is the worst kind of bug to ship.
    """
    for argument in _arguments(formula):
        if "/" not in argument:
            continue
        shape = re.sub(r"\d+", "#", argument)
        if shape in PRECEDENCE_OK:
            continue
        sums = _top_level(argument, "+-")
        if not sums:
            continue
        if any(divide > sums[0] for divide in _top_level(argument, "/")):
            return argument
    return ""


#: input cells that ship with a formula in them as a helpful default -
#: the buyer is meant to be able to type a fixed value straight over it
SEEDED_INPUTS = {
    "Retail Calculator!D9": "target margin starts at the Settings figure",
    "Wholesale Calculator!D7": "trade discount starts at the Settings "
                               "figure",
    "Wholesale Calculator!D10": "target wholesale margin starts at the "
                                "Settings figure",
    "Profit & Break-Even!D7": "fixed costs start at the overhead total",
    "Profit & Break-Even!D9": "the price starts at the catalog price",
    "Profit & Break-Even!D10": "the cost starts at the catalog cost",
    "Scenarios!D28": "the scenario target margin starts at the Settings figure",
}

#: the automation chain, link by link: (sheet, column key, what the
#: formula in that column has to still be pointing at)
CHAIN_CHECKS = [
    # ---- a material's cost per unit, and what it is used in -------------
    ("Materials", "net_pack", ("column", "Materials", "pack_cost")),
    ("Materials", "unit_cost", ("column", "Materials", "net_pack")),
    ("Materials", "base_cost", ("column", "Materials", "unit_cost")),
    ("Materials", "cost_used", ("sheet", "Material Usage")),
    ("Materials", "cover", ("column", "Materials", "stock")),
    # ---- bill of materials ----------------------------------------------
    ("Material Usage", "unit_cost", ("sheet", "Materials")),
    ("Material Usage", "qty_eff", ("column", "Material Usage",
                                   "waste_rate")),
    ("Material Usage", "line_cost", ("column", "Material Usage",
                                     "unit_cost")),
    # ---- labour: minutes over sixty, times a rate -----------------------
    ("Labor", "rate", ("text", "LaborCostHour")),
    ("Labor", "hours", ("column", "Labor", "minutes")),
    ("Labor", "stage_cost", ("column", "Labor", "rate")),
    ("Labor", "cost", ("column", "Labor", "stage_cost")),
    # ---- packaging, postage and overhead --------------------------------
    ("Packaging", "unit_cost", ("column", "Packaging", "pack_cost")),
    ("Packaging", "line_cost", ("column", "Packaging", "unit_cost")),
    ("Shipping", "total_cost", ("column", "Shipping", "cost")),
    ("Shipping", "absorbed", ("column", "Shipping", "charge")),
    ("Overhead", "monthly", ("column", "Overhead", "amount")),
    ("Overhead", "per_hour", ("text", "ProductiveHours")),
    ("Overhead", "per_unit", ("text", "UnitsPerMonth")),
    # ---- marketplace fees, which must stay editable ---------------------
    ("Channels", "pct_total", ("column", "Channels", "transaction")),
    ("Channels", "fixed_total", ("column", "Channels", "processing_fixed")),
    ("Channels", "take_rate", ("column", "Channels", "fee_20")),
    # ---- the catalog: C = M + L + P + S + O -----------------------------
    ("Products", "mat_cost", ("sheet", "Material Usage")),
    ("Products", "labor_cost", ("sheet", "Labor")),
    ("Products", "pack_cost", ("text", "PackStandard")),
    ("Products", "ship_pay", ("sheet", "Shipping")),
    ("Products", "oh_rate", ("text", "OverheadMethod")),
    ("Products", "cost_base", ("column", "Products", "mat_cost")),
    ("Products", "fee_pct", ("sheet", "Channels")),
    ("Products", "fee_fixed", ("sheet", "Channels")),
    ("Products", "denom", ("column", "Products", "fee_pct")),
    ("Products", "suggested_raw", ("column", "Products", "denom")),
    ("Products", "suggested", ("text", "RoundingRule")),
    ("Products", "price", ("column", "Products", "price_input")),
    ("Products", "total_cost", ("column", "Products", "overhead")),
    ("Products", "fees", ("column", "Products", "fee_fixed")),
    ("Products", "profit", ("column", "Products", "total_cost")),
    ("Products", "margin", ("column", "Products", "profit")),
    ("Products", "markup", ("column", "Products", "total_cost")),
    ("Products", "profit_hour", ("column", "Products", "hours")),
    ("Products", "wholesale", ("text", "WholesaleDiscount")),
    ("Products", "comp_price", ("sheet", "Competitors")),
    ("Products", "sold", ("sheet", "Sales")),
    ("Products", "revenue", ("sheet", "Sales")),
    # ---- the six specialist calculators each reach the catalog ----------
    ("Crochet", "yarn_cost", ("sheet", "Materials")),
    ("Crochet", "labor_cost", ("column", "Crochet", "rate")),
    ("Crochet", "suggested_raw", ("column", "Crochet", "denom")),
    ("Jewelry", "metal_cost", ("column", "Jewelry", "metal_rate")),
    ("Jewelry", "suggested_raw", ("column", "Jewelry", "denom")),
    ("Candles", "wax_cost", ("column", "Candles", "wax_total")),
    ("Candles", "fragrance_cost", ("column", "Candles", "fragrance_g")),
    ("Clothing", "fabric_cost", ("column", "Clothing", "fabric_total")),
    ("Woodworking", "wood_cost", ("sheet", "Materials")),
    ("Digital", "dev_cost", ("column", "Digital", "dev_hours")),
    ("Digital", "break_even", ("column", "Digital", "dev_cost")),
    # ---- what actually happened -----------------------------------------
    ("Sales", "catalog_price", ("sheet", "Products")),
    ("Sales", "price_used", ("column", "Sales", "catalog_price")),
    ("Sales", "cogs", ("column", "Sales", "unit_cost_total")),
    ("Sales", "fees_est", ("column", "Sales", "fee_pct")),
    ("Sales", "fee_gap", ("column", "Sales", "fees_est")),
    ("Sales", "profit", ("column", "Sales", "cogs")),
    ("Sales", "margin", ("column", "Sales", "profit")),
    ("Sales", "profit_hour", ("column", "Sales", "hours")),
    # ---- bundles, variations, batches, wholesale ------------------------
    ("Bundles", "parts_cost", ("sheet", "Bundle Lines")),
    ("Bundles", "profit", ("column", "Bundles", "fees")),
    ("Bundle Lines", "unit_cost", ("sheet", "Products")),
    ("Variations", "base_mat", ("sheet", "Products")),
    ("Variations", "total_cost", ("column", "Variations", "mat_cost")),
    ("Batches", "mat_total", ("column", "Batches", "est_mat_unit")),
    ("Batches", "cost_good", ("column", "Batches", "batch_cost")),
    ("Batches", "variance", ("column", "Batches", "est_unit")),
    ("Wholesale Orders", "unit_price", ("column", "Wholesale Orders",
                                        "discount")),
    ("Wholesale Orders", "unit_cost", ("sheet", "Products")),
    ("Wholesale Orders", "margin", ("column", "Wholesale Orders",
                                    "profit")),
    ("Wholesale Customers", "value", ("sheet", "Wholesale Orders")),
    ("Promotions", "margin_after", ("column", "Promotions",
                                    "discount_pct")),
    ("Events", "net", ("column", "Events", "total_cost")),
    ("Development", "recovery", ("column", "Development", "dev_cost")),
    ("Price History", "change_pct", ("column", "Price History",
                                     "old_price")),
    ("Competitors", "gap", ("column", "Competitors", "your_total")),
    ("Waste Log", "unit_cost", ("sheet", "Materials")),
    ("Waste Log", "cost", ("column", "Waste Log", "unit_cost")),
]

#: cells reached by a defined name that have to hold a live formula -
#: the calculator pages are where the value of the file lives, and a
#: name pointing at an empty cell means a page quietly stopped working.
PAGE_FORMULAS = [
    ("RcCost", "retail calculator: cost a unit"),
    ("RcDenom", "retail calculator: 1 - margin - fee rate"),
    ("RcPrice", "retail calculator: the price it recommends"),
    ("RcOverhead", "retail calculator: overhead per unit"),
    ("WcPrice", "wholesale calculator: the trade price"),
    ("WcMargin", "wholesale calculator: margin at that price"),
    ("WcRetail", "wholesale calculator: the retail it implies"),
    ("BeContribution", "break-even: contribution a unit"),
    ("BeFixedTotal", "break-even: fixed costs a month"),
    ("BeCm", "break-even: units needed"),
    ("DashKey", "dashboard: the month the tiles report on"),
]

#: names that must point at an unlocked cell the buyer types into
PAGE_INPUTS = [
    ("RcProduct", "retail calculator: which product"),
    ("RcTarget", "retail calculator: target margin"),
    ("RcChannel", "retail calculator: which channel"),
    ("WcProduct", "wholesale calculator: which product"),
    ("WcDiscount", "wholesale calculator: trade discount"),
    ("BeProduct", "break-even: which product"),
    ("BePrice", "break-even: price being tested"),
    ("ScProduct", "scenarios: which product"),
    ("PrintProduct", "print costing: which product"),
    ("QuoteCustomer", "print quote: who it is for"),
]


def _name_target(wb, name):
    """The one cell a defined name points at, or None."""
    defn = wb.defined_names.get(name)
    if defn is None:
        for ws in wb.worksheets:
            if name in ws.defined_names:
                defn = ws.defined_names[name]
                break
    if defn is None:
        return None
    try:
        sheet, ref = next(defn.destinations)
    except (StopIteration, Exception):          # noqa: B902
        return None
    if sheet not in wb.sheetnames or ":" in ref:
        return None
    return wb[sheet][ref.replace("$", "")]


def check_pages(wb, problems: list, notes: list) -> None:
    """The calculator pages, checked through their defined names."""
    live = 0
    for name, what in PAGE_FORMULAS:
        cell = _name_target(wb, name)
        if cell is None:
            problems.append(f"{name} is missing - {what}")
            continue
        value = cell.value
        if not (isinstance(value, str) and value.startswith("=")):
            problems.append(f"{name} no longer holds a formula - {what}")
            continue
        if cell.protection and not cell.protection.locked:
            problems.append(f"{name} is a formula but is unlocked - {what}")
            continue
        live += 1
    typed = 0
    for name, what in PAGE_INPUTS:
        cell = _name_target(wb, name)
        if cell is None:
            problems.append(f"{name} is missing - {what}")
            continue
        if cell.protection and cell.protection.locked:
            problems.append(f"{name} is locked, so nobody can change it "
                            f"- {what}")
            continue
        typed += 1
    notes.append(f"calculator pages: {live} live formulas, "
                 f"{typed} typeable inputs verified")


def _specs():
    """Every register specification, keyed by sheet name."""
    from handmade_pricing_spreadsheet.build import REGISTER_MODULES
    from handmade_pricing_spreadsheet.config import Config
    cfg = Config()
    return {module.NAME: module.spec(cfg) for module in REGISTER_MODULES}


def check_chain(wb, problems: list, notes: list) -> None:
    """The quote-to-job-to-invoice chain only holds while these do.

    Every one of them is a formula in row 5 of a register, so the check
    is cheap: find the column by its key in the specification, read the
    formula, and make sure it still mentions the thing it depends on.
    """
    specs = _specs()
    checked = 0
    for sheet_name, key, expect in CHAIN_CHECKS:
        spec = specs.get(sheet_name)
        if spec is None or sheet_name not in wb.sheetnames:
            problems.append(f"missing register {sheet_name}")
            continue
        try:
            column = spec.index(key) + 1
        except Exception:
            problems.append(f"{sheet_name}: column '{key}' has gone")
            continue
        formula = wb[sheet_name].cell(row=5, column=column).value
        if not isinstance(formula, str) or not formula.startswith("="):
            problems.append(f"{sheet_name}.{key} is no longer a formula")
            continue
        if expect[0] == "sheet":
            found = expect[1] in formula
            label = f"the {expect[1]} tab"
        elif expect[0] == "text":
            found = expect[1] in formula
            label = expect[1]
        else:
            letter = specs[expect[1]].letter(expect[2])
            label = f"{expect[1]}.{expect[2]} (column {letter})"
            if expect[1] == sheet_name:
                # a link inside the same register is a plain A1 reference
                found = bool(re.search(r"(?<![A-Z$])\$?" + letter +
                                       r"\$?5(?![0-9])", formula))
            else:
                found = f"${letter}$" in formula
        if not found:
            problems.append(
                f"{sheet_name}.{key} no longer points at {label} - "
                f"the automation chain is broken")
        checked += 1
    notes.append(f"automation chain: {checked} links verified")


def check(path: str) -> int:
    problems: list[str] = list(check_raw_xml(path))
    notes: list[str] = []
    if problems:
        print(f"--- {path}")
        for problem in problems[:40]:
            print(f"    - {problem}")
        return 1
    wb = openpyxl.load_workbook(path, data_only=False)

    # ---------------------------------------------------------- structure
    names = wb.sheetnames
    if names != EXPECTED_SHEETS:
        problems.append(f"sheet order/={names}")
    hidden = [ws.title for ws in wb.worksheets if ws.sheet_state != "visible"]
    if hidden:
        problems.append(f"no tab should be hidden, got {hidden}")

    check_chain(wb, problems, notes)
    check_pages(wb, problems, notes)

    # --------------------------------------------------------- properties
    props = wb.properties
    if props.creator != "Novality Store":
        problems.append(f"creator is {props.creator!r}")
    if props.lastModifiedBy != "Novality Store":
        problems.append(f"lastModifiedBy is {props.lastModifiedBy!r}")

    # --------------------------------------------------------- protection
    for ws in wb.worksheets:
        prot = ws.protection
        if not prot.sheet:
            problems.append(f"{ws.title}: sheet protection off")
        elif not prot.password:
            problems.append(f"{ws.title}: protected without a password")

    # -------------------------------------------------------- definednames
    defined = set(wb.defined_names)
    for ws in wb.worksheets:
        defined |= set(ws.defined_names)
    notes.append(f"{len(defined)} defined names")

    # ------------------------------------------------------------ formulas
    used_names: Counter = Counter()
    unknown_functions: Counter = Counter()
    bare_refs: dict = {}
    bad_functions: Counter = Counter()
    formula_cells = 0
    locked_formula_problem = 0
    unlocked_input_cells = 0
    per_sheet = defaultdict(int)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    formula_cells += 1
                    per_sheet[ws.title] += 1
                    if "[[" in value or "{" in value or "}" in value:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unresolved "
                            f"template token -> {value[:90]}")
                    if "#REF" in value:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: #REF -> "
                            f"{value[:90]}")
                    where = f"{ws.title}!{cell.coordinate}"
                    if (cell.protection and not cell.protection.locked
                            and where not in SEEDED_INPUTS):
                        locked_formula_problem += 1
                        if locked_formula_problem < 25:
                            problems.append(
                                f"{where}: formula cell is NOT locked")
                    stripped = re.sub(r'"[^"]*"', '""', value)
                    stripped = re.sub(r"'[^']*'!", "SHEET!", stripped)
                    if value.count('"') % 2:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: odd number of "
                            f"quotes -> {value[:80]}")
                    depth = paren_depth(value)
                    if depth:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unbalanced "
                            f"brackets ({depth:+d}) -> {value[:80]}")
                    slip = precedence_slip(value)
                    if slip:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: a+b/c where "
                            f"(a+b)/c was meant -> {slip[:90]}")
                    if ",)" in stripped or "(," in stripped or \
                            ",," in stripped:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: empty argument "
                            f"-> {value[:80]}")
                    for func in FUNCTION_RE.findall(stripped):
                        if func in MODERN_ONLY:
                            bad_functions[func] += 1
                        elif func not in KNOWN_FUNCTIONS:
                            unknown_functions[func] += 1
                    for token in NAME_RE.findall(stripped):
                        if token in KNOWN_FUNCTIONS:
                            continue
                        used_names[token] += 1
                    for token in BARE_COL_RE.findall(stripped):
                        if token in KNOWN_FUNCTIONS:
                            continue
                        bare_refs[f"{ws.title}!{cell.coordinate} {token}"] = \
                            value[:90]
                elif value is not None and cell.protection and \
                        not cell.protection.locked:
                    unlocked_input_cells += 1

    # sheet names used inside formulas must exist
    sheet_refs = Counter()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    body = re.sub(r'"[^"]*"', '""', value)
                    for ref in re.findall(r"'([^']+)'!", body):
                        sheet_refs[ref] += 1
                    for ref in re.findall(r"(?<![A-Za-z0-9_'])"
                                          r"([A-Za-z][A-Za-z0-9_]*)!", body):
                        sheet_refs[ref] += 1
    for ref, count in sheet_refs.items():
        if ref not in names:
            problems.append(f"formula points at missing sheet "
                            f"'{ref}' ({count}x)")
    notes.append("cross sheet references: " + ", ".join(
        f"{k} {v}" for k, v in sorted(sheet_refs.items(),
                                      key=lambda kv: -kv[1])[:6]))

    for where, formula in list(bare_refs.items())[:15]:
        if where.rsplit(" ", 1)[-1] not in defined:
            problems.append(f"{where}: column letter with no row "
                            f"number -> {formula}")
    for func, count in unknown_functions.items():
        problems.append(f"unknown function {func}() used {count}x")
    for func, count in bad_functions.items():
        problems.append(f"{func}() needs a modern Excel - used {count}x")

    missing = {}
    for token, count in used_names.items():
        if token in defined:
            continue
        if token in KNOWN_FUNCTIONS or token.upper() in KNOWN_FUNCTIONS:
            continue
        if re.match(r"^[A-Z]{1,3}\d+$", token):
            continue
        missing[token] = count
    for token, count in sorted(missing.items(), key=lambda kv: -kv[1])[:25]:
        problems.append(f"undefined name used {count}x: {token}")

    # ----------------------------------------------------- data validation
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            formula = (dv.formula1 or "").lstrip("=")
            if dv.type == "list" and formula and \
                    re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", formula):
                if formula not in defined:
                    problems.append(
                        f"{ws.title}: drop down uses missing name "
                        f"{formula}")

    # ------------------------------------------------------------- report
    with zipfile.ZipFile(path) as zf:
        size = sum(i.file_size for i in zf.infolist())
    notes.append(f"{formula_cells:,} formula cells, "
                 f"{unlocked_input_cells:,} unlocked value cells")
    notes.append(f"uncompressed xml {size / 1_000_000:.1f} MB")

    print(f"--- {path}")
    for note in notes:
        print(f"  . {note}")
    print("  . formulas per sheet: " + ", ".join(
        f"{k} {v}" for k, v in sorted(per_sheet.items(),
                                      key=lambda kv: -kv[1])[:10]))
    if problems:
        print(f"  ! {len(problems)} problem(s):")
        for problem in problems[:60]:
            print(f"    - {problem}")
        return 1
    print("  OK - no problems found")
    return 0


if __name__ == "__main__":
    target = (sys.argv[1] if len(sys.argv) > 1
              else "output/Handmade Product Pricing & Profit "
                   "Calculator.xlsx")
    raise SystemExit(check(target))
