"""`Wholesale Calculator` - can you afford to sell at half price?

Wholesale is simple arithmetic that catches a lot of makers out. A
shop expects around 50% off your retail price, sells it at your
retail price, and keeps the difference. That only works if your
retail price was built on your real costs in the first place - which
is exactly what the rest of this workbook is for.

This page works it both ways: what a given retail price leaves you at
wholesale, and what retail price you would need for wholesale to work
at the margin you want.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check
from .common import round_price

NAME = "Wholesale Calculator"
LAST_COL = 11

#: discounts a shop might ask for
DISCOUNTS = [0.30, 0.35, 0.40, 0.45, 0.50, 0.55, 0.60]


def _pick(field: str, default: str = "0") -> str:
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH(WcProduct,'
            f'[[Products.name]],0)),{default})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Wholesale Price Calculator",
             "What a stockist costs you, what you need to charge, and "
             "whether the order is worth the fortnight.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 32, fmts.chrome("label"))
    s.set_col(2, 2, 11, fmts.chrome("label"))
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 15)
    s.set_col(5, 5, 15)
    s.set_col(6, LAST_COL, 14, fmts.chrome("note"))

    r = 4

    # ------------------------------------------------------- the inputs
    r = s.section(r, "1. The product and the deal", LAST_COL)
    r = s.field(r, "Product", "Willow cable beanie", name="WcProduct",
                validate=list_from("ProductNames"), value_end=5,
                help_text="From the Product Catalog.")
    r = s.field(r, "Discount the shop wants", None, style="percent",
                name="WcDiscount", validate=number_check(0, 0.9),
                formula="=WholesaleDiscount",
                help_text="50% is the usual ask. 40% is common for "
                          "small independents, 60% for distributors.")
    r = s.field(r, "Units in the order", 24, style="int", name="WcQty",
                validate=number_check(1, 100000, decimal=False),
                help_text="Used for the order totals and the hours "
                          "further down.")
    r = s.field(r, "Carriage you pay", 0, style="money",
                name="WcCarriage", validate=number_check(0, 100000),
                help_text="One box to the shop, not one box per item "
                          "- which is why wholesale postage is so "
                          "much cheaper per piece.")
    r = s.field(r, "Target wholesale margin", None, style="percent",
                name="WcTarget", validate=number_check(0, 0.95),
                formula="=TargetWholesaleMargin",
                help_text="Thinner than retail on purpose: no "
                          "marketplace fees and the shop does the "
                          "selling.")
    r += 1

    # -------------------------------------------------------- the costs
    r = s.section(r, "2. What the piece costs you", LAST_COL)
    r = s.note(r, "Wholesale skips the gift packaging and the retail "
                  "postage, so the cost base is lower than it is for "
                  "a retail sale. Override anything that is "
                  "different for a trade order.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Cost", "From the catalog", "Your override",
                            "Used", "Note"],
                     [32, 16, 15, 15, 42])
    rows = [
        ("Materials", _pick("mat_cost"), "WcMat", "money",
         "The same whoever buys it."),
        ("Labour", _pick("labour_adj"), "WcLab", "money",
         "Often lower per piece on a batch - override it if so."),
        ("Consumables", _pick("consumables"), "WcCons", "money",
         "Glue, thread, gas, sandpaper."),
        ("Trade packaging", "0.35", "WcPack", "money",
         "A poly bag and a swing tag, not a gift box."),
        ("Overhead", _pick("overhead"), "WcOh", "money",
         "The same share of the studio as any other piece."),
    ]
    first = r
    for label, formula, name, style, note in rows:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render("=" + formula),
                         fmts.calc(style))
        s.write(r, 3, None, fmts.input(style))
        ws.data_validation(r, 3, r, 3, number_check(0, 1000000))
        book.define_name(name, f"='{NAME}'!$D${r + 1}")
        ws.write_formula(r, 4, f'=IF(D{r + 1}<>"",D{r + 1},C{r + 1})',
                         fmts.calc(style))
        book.define_name(name + "Used", f"='{NAME}'!$E${r + 1}")
        s.merge(r, 5, LAST_COL, note, fmts.chrome("note"))
        r += 1
    s.write(r, 1, "Cost a unit", fmts.chrome("total_label"))
    ws.write_formula(r, 4, f'=ROUND(SUM(E{first + 1}:E{r}),2)',
                     fmts.chrome("total_value",
                                 num_format=s.num_format("money")))
    book.define_name("WcCost", f"='{NAME}'!$E${r + 1}")
    r += 2

    # ------------------------------------------------- the two directions
    r = s.section(r, "3. Working it both ways", LAST_COL)
    r = s.subsection(r, "A. From your retail price down", LAST_COL,
                     first_col=1)
    rows_a = [
        ("Your retail price", "=" + _pick("price", '""'), "money",
         "WcRetail", "From the catalog."),
        ("Wholesale price", '=IF(WcRetail="","",ROUND(WcRetail*'
         '(1-WcDiscount),2))', "money", "WcPrice",
         "Retail less the discount the shop is asking for."),
        ("Profit a unit", '=IF(WcPrice="","",ROUND(WcPrice-WcCost-'
         'WcCarriage/MAX(1,WcQty),2))', "money", "WcProfit",
         "After the cost above and a share of the carriage."),
        ("Wholesale margin", '=IF(OR(WcPrice="",WcPrice=0),"",'
         'WcProfit/WcPrice)', "percent", "WcMargin",
         "Compare it with your target above."),
        ("You earn an hour", '=IF(OR(WcPrice="",' + _pick("hours", "0") +
         '=0),"",ROUND((WcProfit+WcLabUsed)/' + _pick("hours", "1") +
         ',2))', "money", "WcWage",
         "Profit plus your own wage, over the hours it takes."),
    ]
    for label, formula, style, name, note in rows_a:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 3, s.reg.render(formula), fmts.calc(style))
        book.define_name(name, f"='{NAME}'!$D${r + 1}")
        s.merge(r, 4, LAST_COL, note, fmts.chrome("note"))
        r += 1
    ws.conditional_format(r - 2, 3, r - 2, 3,
                          {"type": "cell", "criteria": "<",
                           "value": "=WcTarget",
                           "format": fmts.cf("warn")})
    ws.conditional_format(r - 2, 3, r - 2, 3,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    r += 1

    r = s.subsection(r, "B. From the margin you need, back up", LAST_COL,
                     first_col=1)
    rows_b = [
        ("Wholesale price you need",
         '=IF(WcTarget>=1,"",ROUND((WcCost+WcCarriage/MAX(1,WcQty))/'
         '(1-WcTarget),2))', "money", "WcNeedWholesale",
         "Cost divided by one minus your target margin."),
        ("Retail price that supports it",
         '=IF(OR(WcNeedWholesale="",WcDiscount>=1),"",ROUND(' +
         round_price("(WcNeedWholesale/(1-WcDiscount))") + ',2))',
         "money", "WcNeedRetail",
         "The retail price the shop will put on the shelf. If it is "
         "above what the market pays, the costs have to come down."),
        ("Versus your current retail",
         '=IF(OR(WcRetail="",WcNeedRetail=""),"",ROUND(WcNeedRetail-'
         'WcRetail,2))', "money", "WcGap",
         "Positive means your retail price is too low to wholesale "
         "from."),
        ("Keystone check (cost x 4)",
         '=ROUND(WcCost*4,2)', "money", "WcKeystone",
         "The old trade rule: double for wholesale, double again for "
         "retail. Crude, but a useful second opinion."),
    ]
    for label, formula, style, name, note in rows_b:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 3, s.reg.render(formula), fmts.calc(style))
        book.define_name(name, f"='{NAME}'!$D${r + 1}")
        s.merge(r, 4, LAST_COL, note, fmts.chrome("note"))
        r += 1
    r += 1

    # ------------------------------------------------------- the order --
    r = s.section(r, "4. The order itself", LAST_COL)
    r = s.table_head(r, 1, ["Units", "Order value", "Cost to make",
                            "Carriage", "Profit", "Margin",
                            "Hours of work", "Profit an hour"],
                     [10, 14, 14, 11, 12, 10, 14, 14])
    ws.write_formula(r, 1, "=WcQty", fmts.calc("int"))
    ws.write_formula(r, 2, '=IF(WcPrice="","",ROUND(WcQty*WcPrice,2))',
                     fmts.calc("money"))
    ws.write_formula(r, 3, "=ROUND(WcQty*WcCost,2)", fmts.calc("money"))
    ws.write_formula(r, 4, "=WcCarriage", fmts.calc("money"))
    ws.write_formula(r, 5, f'=IF(C{r + 1}="","",ROUND(C{r + 1}-'
                           f'D{r + 1}-E{r + 1},2))', fmts.calc("money"))
    ws.write_formula(r, 6, f'=IF(OR(C{r + 1}="",C{r + 1}=0),"",'
                           f'F{r + 1}/C{r + 1})', fmts.calc("percent"))
    ws.write_formula(r, 7, "=ROUND(WcQty*" + s.reg.render(_pick("hours")) +
                     ",1)", fmts.calc("number1"))
    ws.write_formula(r, 8, f'=IF(OR(H{r + 1}=0,F{r + 1}=""),"",'
                           f'ROUND(F{r + 1}/H{r + 1},2))',
                     fmts.calc("money"))
    order_row = r
    r += 1
    s.merge(r, 1, LAST_COL,
            "The hours column is the one to look at before you say "
            "yes. A profitable order that fills six weeks can still "
            "be the wrong order.", fmts.chrome("note"))
    ws.conditional_format(order_row, 5, order_row, 5,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(order_row, 8, order_row, 8,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    r += 2

    # --------------------------------------------------- discount ladder
    r = s.section(r, "5. What each discount does", LAST_COL)
    r = s.note(r, "The same product at every discount a shop might "
                  "ask for. The last column is the honest answer to "
                  "'can you do a bit better on the price?'.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Discount", "Wholesale price", "Profit a unit",
                            "Margin", "You earn an hour",
                            "On a full order", "Verdict"],
                     [11, 16, 14, 10, 16, 15, 26])
    ladder = r
    for disc in DISCOUNTS:
        price = f'IF(WcRetail="","",ROUND(WcRetail*{1 - disc},2))'
        s.write(r, 1, disc, fmts.calc("percent0"))
        ws.write_formula(r, 2, f"={price}", fmts.calc("money"))
        ws.write_formula(r, 3, f'=IF(C{r + 1}="","",ROUND(C{r + 1}-'
                               f'WcCost-WcCarriage/MAX(1,WcQty),2))',
                         fmts.calc("money"))
        ws.write_formula(r, 4, f'=IF(OR(C{r + 1}="",C{r + 1}=0),"",'
                               f'D{r + 1}/C{r + 1})',
                         fmts.calc("percent"))
        ws.write_formula(
            r, 5, '=IF(OR(C{0}="",{1}=0),"",ROUND((D{0}+WcLabUsed)/'
                  '{1},2))'.format(r + 1, s.reg.render(_pick("hours"))),
            fmts.calc("money"))
        ws.write_formula(r, 6, f'=IF(D{r + 1}="","",ROUND(D{r + 1}*'
                               f'WcQty,2))', fmts.calc("money"))
        ws.write_formula(
            r, 7,
            f'=IF(C{r + 1}="","",IF(D{r + 1}<0,"No - below cost",'
            f'IF(E{r + 1}<0.1,"No - nothing left in it",'
            f'IF(E{r + 1}<WcTarget,"Thin - only for a big order",'
            f'"Workable"))))', fmts.calc("text"))
        r += 1
    ws.conditional_format(ladder, 3, r - 1, 3,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(ladder, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "No -", "format": fmts.cf("bad")})
    ws.conditional_format(ladder, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "Workable", "format": fmts.cf("good")})
    r += 1

    # ------------------------------------------------------- the verdict
    r = s.section(r, "6. Should you take it?", LAST_COL)
    checks = [
        ("The order makes money",
         '=IF(WcPrice="","Pick a product first",'
         'IF(WcProfit>0,"Yes - there is profit in every piece",'
         '"No - every piece loses money at this discount"))'),
        ("It pays you properly",
         '=IF(OR(WcWage="",WcWage=0),"Add the labour minutes first",'
         'IF(WcWage>=MinimumWage,"Yes - above minimum wage",'
         '"No - you would be paying to make them"))'),
        ("Your retail price supports it",
         '=IF(WcGap="","-",IF(WcGap<=0,'
         '"Yes - your retail price has room in it",'
         '"No - retail needs to be "&TEXT(WcNeedRetail,"0.00")&'
         '" for this to work"))'),
        ("It meets your minimum",
         '=IF(WcPrice="","-",IF(AND(WcQty*WcPrice>=MinOrderValue,'
         'WcQty>=MinOrderQty),"Yes - above your minimum order",'
         '"Check - below the minimum on your Settings tab"))'),
        ("Can you actually make them?",
         '=IF(' + _pick("hours", "0") + '=0,'
         '"Add the labour minutes to find out",'
         '"It is "&TEXT(WcQty*' + _pick("hours", "0") +
         ',"0.0")&" hours of making - about "&'
         'TEXT(ROUNDUP(WcQty*' + _pick("hours", "0") +
         '/8,0),"0")&" full days")'),
    ]
    first_check = r
    for label, formula in checks:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc("text"))
        r += 1
    for word, kind in (("No -", "bad"), ("Yes -", "good"),
                       ("Check", "warn")):
        ws.conditional_format(first_check, 2, r - 1, 2,
                              {"type": "text", "criteria": "containing",
                               "value": word, "format": fmts.cf(kind)})
    r += 1

    s.note(r,
           "Taking the order? Add the shop to the Wholesale Customers "
           "tab with its discount and terms, then enter the lines on "
           "Wholesale Orders - the margins there follow whatever you "
           "agree here.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
