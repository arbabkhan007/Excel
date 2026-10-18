"""The three printable forms.

* `Print Costing` - a one page costing card for a single product:
  every material, every labour stage, the packaging, the fees and the
  price, in a form you can put in a ring binder or hand to an
  accountant.
* `Print Quote` - a custom order quote for a buyer, with the lines, the
  deposit and an acceptance box.
* `Print Order` - a wholesale line sheet and order form: your range at
  retail and trade prices, with a column the shop fills in.

Pick an ID at the top of each page and the form fills itself in.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from, number_check

COSTING_NAME = "Print Costing"
QUOTE_NAME = "Print Quote"
ORDER_NAME = "Print Order"

BOM_LINES = 12
LABOUR_LINES = 7
QUOTE_LINES = 8
ORDER_LINES = 16


def _letterhead(s: SheetWriter, fmts, row: int, last_col: int,
                heading: str) -> int:
    """Shop name, contact details and the form's title."""
    ws = s.ws
    s.merge(row, 0, 3, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0, "=ShopName", fmts.chrome("print_title"))
    s.merge(row, 4, last_col, heading, fmts.chrome("print_title"))
    row += 1
    s.merge(row, 0, 3, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0, "=ShopAddress", fmts.chrome("print_sub"))
    s.merge(row, 4, last_col, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 4,
                     '=ShopEmail&IF(ShopPhone="",""," - "&ShopPhone)&'
                     'IF(ShopWebsite="",""," - "&ShopWebsite)',
                     fmts.chrome("print_sub"))
    return row + 2


def _picker(s: SheetWriter, fmts, book, row: int, label: str,
            default: str, list_name: str, define: str,
            help_text: str, last_col: int = 7) -> int:
    """The 'which one are we printing?' box."""
    ws = s.ws
    s.write(row, 0, label, fmts.chrome("label_bold"))
    s.write(row, 1, default, fmts.input("id"))
    ws.data_validation(row, 1, row, 1, list_from(list_name))
    book.define_name(define, f"={quote_sheet(s.name)}!$B${row + 1}")
    s.merge(row, 2, last_col, help_text, fmts.chrome("note"))
    return row + 2


# ---------------------------------------------------------------------------
# Print Costing
# ---------------------------------------------------------------------------

def build_costing(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = COSTING_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Costing Sheet", "", LAST_COL)
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 28)
    s.set_col(2, 2, 14)
    s.set_col(3, 3, 11)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 13)
    s.set_col(7, 7, 15)

    row = 4
    row = _picker(s, fmts, book, row, "Product", "P-0001", "ProductIDs",
                  "PrintProduct",
                  "Pick a product. Every line below fills itself in "
                  "from the Materials, Material Usage and Labor tabs.",
                  LAST_COL)
    row = _letterhead(s, fmts, row, LAST_COL, "COSTING SHEET")

    def p(key: str, style: str = "text") -> str:
        return s.reg.render(f'=IFERROR(INDEX([[Products.{key}]],'
                            f'MATCH(PrintProduct,[[Products.id]],0)),"")')

    # ------------------------------------------------------- the header --
    pairs = [
        ("Product", "name", "text", "SKU", "sku", "text"),
        ("Category", "category", "text", "Status", "status", "text"),
        ("Channel", "channel", "text", "Difficulty", "difficulty", "text"),
        ("Packaging", "pack_profile", "text", "Postage", "ship_option",
         "text"),
    ]
    for label_a, key_a, style_a, label_b, key_b, style_b in pairs:
        s.write(row, 0, label_a, fmts.chrome("label_bold"))
        s.merge(row, 1, 2, None, fmts.cell(style_a, "plain"))
        ws.write_formula(row, 1, p(key_a), fmts.cell(style_a, "plain"))
        s.write(row, 4, label_b, fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.cell(style_b, "plain"))
        ws.write_formula(row, 5, p(key_b), fmts.cell(style_b, "plain"))
        row += 1
    row += 1

    # ---------------------------------------------------- the materials --
    row = s.section(row, "Materials", LAST_COL)
    row = s.table_head(row, 0, ["Material", "Detail", "Unit", "Quantity",
                                "Waste", "Unit cost", "Line cost"],
                       [18, 28, 14, 11, 10, 12, 13])
    bom_first = row
    for n in range(1, BOM_LINES + 1):
        key = f'PrintProduct&"|{n}"'
        fields = [("material_name", "text"), ("notes", "text"),
                  ("unit", "text"), ("qty", "number1"),
                  ("waste_rate", "percent"), ("unit_cost", "money4"),
                  ("line_cost", "money")]
        for col, (field, style) in enumerate(fields):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Material Usage.{field}]],'
                             f'MATCH({key},[[Material Usage.key]],0)),"")'),
                fmts.cell(style, "plain"))
        row += 1
    s.merge(row, 0, 5, "Materials total", fmts.chrome("total_label"))
    ws.write_formula(row, 6,
                     f"=ROUND(SUM(G{bom_first + 1}:G{row}),2)",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    # ------------------------------------------------------- the labour --
    row = s.section(row, "Labour", LAST_COL)
    row = s.table_head(row, 0, ["Stage", "Who", "Basis", "Minutes",
                                "Per item", "Rate", "Cost"],
                       [18, 28, 14, 11, 10, 12, 13])
    lab_first = row
    for n in range(1, LABOUR_LINES + 1):
        key = f'PrintProduct&"|{n}"'
        fields = [("stage", "text"), ("who", "text"), ("basis", "text"),
                  ("minutes", "number1"), ("mins_item", "number1"),
                  ("rate", "money"), ("cost", "money")]
        for col, (field, style) in enumerate(fields):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Labor.{field}]],'
                             f'MATCH({key},[[Labor.key]],0)),"")'),
                fmts.cell(style, "plain"))
        row += 1
    s.merge(row, 0, 5, "Labour total", fmts.chrome("total_label"))
    ws.write_formula(row, 6, f"=ROUND(SUM(G{lab_first + 1}:G{row}),2)",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    # ------------------------------------------------- everything else --
    row = s.section(row, "The rest of the cost", LAST_COL)
    row = s.table_head(row, 0, ["Cost", "What it covers", "", "", "", "",
                                "Amount"],
                       [18, 28, 14, 11, 10, 12, 13])
    others = [
        ("Materials", "Everything in the piece, waste included",
         "mat_cost"),
        ("Labour", "Your hours at your rate, difficulty applied",
         "labour_adj"),
        ("Consumables", "Thread, glue, gas, abrasives", "consumables"),
        ("Packaging", "Box, tissue, card, label", "pack_cost"),
        ("Postage absorbed", "What the label costs less what the buyer "
                             "pays", "ship_cost"),
        ("Overhead", "This product's share of the studio", "overhead"),
    ]
    other_first = row
    for label, what, key in others:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 5, what, fmts.chrome("note"))
        ws.write_formula(row, 6, p(key), fmts.cell("money", "plain"))
        row += 1
    s.merge(row, 0, 5, "Total cost", fmts.chrome("total_label"))
    ws.write_formula(row, 6,
                     f"=ROUND(SUM(G{other_first + 1}:G{row}),2)",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    # -------------------------------------------------------- the price --
    row = s.section(row, "The price", LAST_COL)
    price_rows = [
        ("Price", "price", "money", "What it sells for."),
        ("Channel fees", "fees", "money",
         "What the marketplace takes on that price."),
        ("Profit", "profit", "money", "What is left for you."),
        ("Margin", "margin", "percent", "Profit over the price."),
        ("Markup", "markup", "percent", "Profit over the cost."),
        ("Minutes of work", "minutes", "number1",
         "Including the difficulty factor."),
        ("You earn an hour", "wage_hour", "money",
         "Your wage plus the profit, over the hours."),
        ("Wholesale price", "wholesale", "money",
         "At the standard trade discount."),
        ("Suggested price", "suggested", "money",
         "What the target margin says it should be."),
    ]
    for label, key, style, what in price_rows:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 5, what, fmts.chrome("note"))
        ws.write_formula(row, 6, p(key), fmts.cell(style, "plain"))
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(
        row, 0,
        '="Costed by "&ShopName&" on "&TEXT(TodayDate,"d mmmm yyyy")&'
        '". Labour at "&TEXT(LaborCostHour,"0.00")&" an hour. '
        'Marketplace fee rates last reviewed "&'
        'TEXT(FeesReviewed,"d mmm yyyy")&"."',
        fmts.chrome("note"))
    row += 1

    s.finish(freeze_row=None, print_area=f"A1:H{row + 1}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Quote
# ---------------------------------------------------------------------------

def build_quote(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = QUOTE_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Quote", "", LAST_COL)
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 16)
    s.set_col(3, 3, 9)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 13)
    s.set_col(7, 7, 14)

    row = 4
    s.merge(row, 0, LAST_COL,
            "For custom and commission work. Type the customer's "
            "details and pick the products - the prices come from "
            "your catalog and the totals work themselves out.",
            fmts.chrome("note_box"))
    row += 2
    row = _letterhead(s, fmts, row, LAST_COL, "QUOTE")

    # -------------------------------------------------- the two blocks --
    s.write(row, 0, "Quote for", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.input("text"))
    s.write(row, 4, "Quote number", fmts.chrome("label_bold"))
    s.write(row, 5, None, fmts.input("text"))
    book.define_name("QuoteCustomer", f"={quote_sheet(NAME)}!$B${row + 1}")
    row += 1
    s.write(row, 0, "Email", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.input("text"))
    s.write(row, 4, "Date", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, "=TodayDate", fmts.cell("date", "plain"))
    row += 1
    s.write(row, 0, "What they want", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.input("wrap"))
    s.write(row, 4, "Valid for", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, "30 days", fmts.input("text"))
    s.row_height(row, 32)
    row += 2

    # ------------------------------------------------------- the lines --
    row = s.table_head(row, 0, ["Product", "Description", "Variation",
                                "Qty", "Price each", "Line total"],
                       [18, 30, 16, 9, 12, 13])
    first_line = row
    for _ in range(QUOTE_LINES):
        s.write(row, 0, None, fmts.input("text"))
        ws.data_validation(row, 0, row, 0, list_from("ProductNames"))
        s.merge(row, 1, 2, None, fmts.input("text"))
        s.write(row, 3, None, fmts.input("text"))
        s.write(row, 4, None, fmts.input("int"))
        ws.data_validation(row, 4, row, 4,
                           number_check(0, 10000, decimal=False))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=IF(A{row + 1}="","",IFERROR(INDEX('
                         f'[[Products.price]],MATCH(A{row + 1},'
                         f'[[Products.name]],0)),""))'),
            fmts.cell("money", "plain"))
        ws.write_formula(
            row, 6,
            f'=IF(OR(E{row + 1}="",F{row + 1}=""),"",ROUND(E{row + 1}*'
            f'F{row + 1},2))', fmts.cell("money", "plain"))
        row += 1
    last_line = row - 1
    row += 1

    # ------------------------------------------------------ the totals --
    totals = [
        ("Subtotal",
         f"=ROUND(SUM(G{first_line + 1}:G{last_line + 1}),2)", "money",
         "The lines above."),
        ("Rush or custom surcharge", None, "money",
         "Type it in if the turnaround is short."),
        ("Postage", None, "money", "What you will charge for postage."),
        ("Discount", None, "money", "Anything you are knocking off."),
    ]
    total_cells = []
    for label, formula, style, note in totals:
        s.merge(row, 0, 2, note, fmts.chrome("note"))
        s.merge(row, 3, 5, label, fmts.chrome("label_right"))
        if formula:
            ws.write_formula(row, 6, formula, fmts.cell(style, "plain"))
        else:
            s.write(row, 6, None, fmts.input(style))
        total_cells.append(row)
        row += 1
    sub, rush, post, disc = (f"G{n + 1}" for n in total_cells)
    s.merge(row, 3, 5, "Total", fmts.chrome("total_label"))
    ws.write_formula(
        row, 6,
        f'=ROUND({sub}+IF({rush}="",0,{rush})+IF({post}="",0,{post})-'
        f'IF({disc}="",0,{disc}),2)',
        fmts.chrome("total_value", num_format=cfg.money_format))
    total_row = row
    row += 1
    s.merge(row, 3, 5, "Deposit to start", fmts.chrome("label_right"))
    ws.write_formula(row, 6, f"=ROUND(G{total_row + 1}*0.5,2)",
                     fmts.cell("money", "plain"))
    s.merge(row, 0, 2, "Half up front is normal for custom work.",
            fmts.chrome("note"))
    row += 2

    # --------------------------------------------------------- the terms
    row = s.section(row, "What you are agreeing to", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        '="Custom work begins once the deposit is received. The '
        'balance is due before despatch. Handmade pieces vary '
        'slightly - colour, grain and texture will not be identical '
        'to the photographs. Custom and personalised orders cannot '
        'be returned unless faulty. Prices hold for 30 days from '
        'the date above. "&ShopName&" - "&ShopEmail',
        fmts.cell("wrap", "plain"))
    s.row_height(row, 56)
    row += 2

    row = s.section(row, "Acceptance", LAST_COL)
    for label in ("Signed", "Print name"):
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 3, None, fmts.chrome("form_line"))
        s.write(row, 4, "Date", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 26)
        row += 1
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(row, 0,
                     '="Quote prepared by "&ShopName&" on "&'
                     'TEXT(TodayDate,"d mmmm yyyy")',
                     fmts.chrome("note"))
    row += 1

    s.finish(freeze_row=None, print_area=f"A1:H{row + 1}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Order
# ---------------------------------------------------------------------------

def build_order(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = ORDER_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Wholesale Line Sheet & Order Form", "", LAST_COL)
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 16)
    s.set_col(3, 3, 11)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 10)
    s.set_col(6, 6, 13)
    s.set_col(7, 7, 14)

    row = 4
    row = _picker(s, fmts, book, row, "Customer", "WC-0001",
                  "CustomerIDs", "PrintCustomer",
                  "Pick a stockist and the sheet prices itself at "
                  "their discount. Leave the quantity column blank - "
                  "that is the bit they fill in.", LAST_COL)
    row = _letterhead(s, fmts, row, LAST_COL, "LINE SHEET")

    def c(key: str) -> str:
        return s.reg.render(
            f'=IFERROR(INDEX([[Wholesale Customers.{key}]],'
            f'MATCH(PrintCustomer,[[Wholesale Customers.id]],0)),"")')

    s.write(row, 0, "Shop", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 1, c("name"), fmts.cell("text", "plain"))
    s.write(row, 4, "Discount", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, c("discount_used"),
                     fmts.cell("percent", "plain"))
    row += 1
    s.write(row, 0, "Contact", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 1, c("contact"), fmts.cell("text", "plain"))
    s.write(row, 4, "Terms", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 5, c("terms"), fmts.cell("text", "plain"))
    row += 1
    s.write(row, 0, "Date", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=TodayDate", fmts.cell("date", "plain"))
    s.write(row, 4, "Order ref", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, None, fmts.input("text"))
    row += 2

    # -------------------------------------------------------- the range --
    row = s.table_head(row, 0, ["SKU", "Product", "Category", "Retail",
                                "Trade price", "Qty", "Line total"],
                       [18, 30, 16, 11, 12, 10, 13])
    first = row
    discount = 'IFERROR(INDEX([[Wholesale Customers.discount_used]],' \
               'MATCH(PrintCustomer,[[Wholesale Customers.id]],0)),' \
               'WholesaleDiscount)'
    for n in range(1, ORDER_LINES + 1):
        for col, field, style in ((0, "sku", "text"),
                                  (1, "name", "text"),
                                  (2, "category", "text")):
            # &"" keeps an empty catalog row blank on the printed sheet
            # instead of printing a zero
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Products.{field}]],{n})'
                             f'&"","")'), fmts.cell(style, "plain"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF(B{row + 1}="","",IFERROR('
                         f'INDEX([[Products.price]],{n}),""))'),
            fmts.cell("money", "plain"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IF(D{row + 1}="","",ROUND(D{row + 1}*(1-'
                         f'{discount}),2))'),
            fmts.cell("money", "plain"))
        s.write(row, 5, None, fmts.input("int"))
        ws.data_validation(row, 5, row, 5,
                           number_check(0, 100000, decimal=False))
        ws.write_formula(
            row, 6,
            f'=IF(OR(E{row + 1}="",F{row + 1}=""),"",ROUND(E{row + 1}*'
            f'F{row + 1},2))', fmts.cell("money", "plain"))
        row += 1
    last = row - 1
    row += 1

    s.merge(row, 0, 3, "Items ordered", fmts.chrome("label_right"))
    ws.write_formula(row, 4, f"=SUM(F{first + 1}:F{last + 1})",
                     fmts.cell("int", "plain"))
    s.merge(row, 5, 5, "Total", fmts.chrome("total_label"))
    ws.write_formula(row, 6,
                     f"=ROUND(SUM(G{first + 1}:G{last + 1}),2)",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    total_row = row
    row += 1
    s.merge(row, 0, 5, None, fmts.cell("text", "plain"))
    ws.write_formula(
        row, 0,
        f'=IF(G{total_row + 1}=0,"Add quantities in the Qty column.",'
        f'IF(G{total_row + 1}<MinOrderValue,"This order is under the '
        f'minimum of "&TEXT(MinOrderValue,"0.00")&".",'
        f'"Order accepted - thank you."))',
        fmts.cell("text", "plain"))
    ws.conditional_format(row, 0, row, 5,
                          {"type": "text", "criteria": "containing",
                           "value": "under the minimum",
                           "format": fmts.cf("warn")})
    row += 2

    row = s.section(row, "Trade terms", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        '="Minimum order "&TEXT(MinOrderValue,"0.00")&" or "&'
        'TEXT(MinOrderQty,"0")&" pieces. Lead time 2-4 weeks '
        'depending on the season. Carriage charged at cost unless '
        'agreed otherwise. Everything is made by hand, so small '
        'variations in colour and texture are part of the product '
        'rather than a fault. Payment terms as agreed above. "&'
        'ShopName&" - "&ShopEmail',
        fmts.cell("wrap", "plain"))
    s.row_height(row, 56)
    row += 2

    for label in ("Signed for the shop", "Print name"):
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 3, None, fmts.chrome("form_line"))
        s.write(row, 4, "Date", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 26)
        row += 1
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(row, 0,
                     '="Line sheet produced by "&ShopName&" on "&'
                     'TEXT(TodayDate,"d mmmm yyyy")&". Prices hold '
                     'for 60 days."', fmts.chrome("note"))
    row += 1

    s.finish(freeze_row=None, print_area=f"A1:H{row + 1}",
             gridlines=False, fit_width=True)
    return s
