"""`Retail Calculator` - the page the whole workbook exists for.

Pick a product, or type the costs in by hand. The page adds materials,
labour, packaging, postage and overhead, works out what the
marketplace takes, and solves for the price that leaves the margin you
asked for:

    Price = (Cost + fixed fees) / (1 - target margin - fee rate)

Then it shows you the same product at a dozen other prices, so the
number you choose is a decision rather than a hope.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check
from .common import round_price

NAME = "Retail Calculator"
LAST_COL = 11

#: margins offered in the ladder
LADDER = [0.25, 0.30, 0.35, 0.40, 0.45, 0.50, 0.55, 0.60, 0.65, 0.70]
#: price points shown around the suggested price
STEPS = [-0.30, -0.20, -0.10, -0.05, 0.0, 0.05, 0.10, 0.20, 0.30, 0.50]


def _pick(field: str, default: str = "0") -> str:
    """Look a field up on the catalog by the product name chosen above."""
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH(RcProduct,'
            f'[[Products.name]],0)),{default})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    hours = s.reg.render(_pick("hours", "0"))

    s.banner("Retail Price Calculator",
             "Everything that goes into the price, and the number it "
             "comes out at. Fill in the white cells.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 30, fmts.chrome("label"))
    s.set_col(2, 2, 11, fmts.chrome("label"))
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 15)
    s.set_col(5, 5, 15)
    s.set_col(6, LAST_COL, 14, fmts.chrome("note"))

    r = 4

    # ----------------------------------------------- 1. which product --
    r = s.section(r, "1. What are you pricing?", LAST_COL)
    r = s.note(r, "Choose a product from your catalog and every cost "
                  "below fills itself in. Or leave it blank and type "
                  "the costs by hand - nothing on this page changes "
                  "your saved data.",
               LAST_COL, first_col=1, height=26)
    r = s.field(r, "Product", "Willow cable beanie", name="RcProduct",
                validate=list_from("ProductNames"), value_end=5,
                help_text="From the Product Catalog. Clear the cell to "
                          "price something that is not listed yet.")
    r = s.field(r, "Where are you selling it?", "Etsy",
                name="RcChannel", validate=list_from("ChannelList"),
                value_end=5,
                help_text="The fees come from the Channels tab, which "
                          "you can edit.")
    r = s.field(r, "Target margin", None, style="percent",
                name="RcTarget", validate=number_check(0, 0.95),
                formula="=TargetRetailMargin",
                help_text="Profit as a share of the price. Change it "
                          "here to test, or on Settings to change it "
                          "everywhere.")
    r = s.field(r, "Quantity you make at a time", 1, style="int",
                name="RcBatch", validate=number_check(1, 10000,
                                                      decimal=False),
                help_text="Used for the per unit setup time below.")
    r += 1

    # ------------------------------------------------- 2. the costs ----
    r = s.section(r, "2. What it costs you to make", LAST_COL)
    r = s.note(r, "The middle column is what the catalog says. Type "
                  "anything in the override column to try a different "
                  "number - an override always wins.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Cost", "From the catalog", "Your override",
                            "Used", "What it is"],
                     [30, 16, 15, 15, 44])
    cost_rows = [
        ("Materials", _pick("mat_cost"), "RcMat", "money",
         "Everything that ends up in the piece, waste included."),
        ("Labour", _pick("labour_adj"), "RcLab", "money",
         "Your hours at your rate - the cost most makers leave out."),
        ("Consumables", _pick("consumables"), "RcCons", "money",
         "Glue, thread, sandpaper, gas. Small, real, forgotten."),
        ("Packaging", _pick("pack_cost"), "RcPack", "money",
         "Box, tissue, card, label, tape."),
        ("Postage you absorb", _pick("ship_cost"), "RcShip", "money",
         "What the label costs you less what the buyer pays."),
        ("Other costs", "0", "RcOther", "money",
         "Anything else this piece carries - a pattern licence, a "
         "stamp, a fee."),
    ]
    first_cost = r
    for label, formula, name, style, what in cost_rows:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render("=" + formula),
                         fmts.calc(style))
        s.write(r, 3, None, fmts.input(style))
        ws.data_validation(r, 3, r, 3, number_check(0, 1000000))
        book.define_name(name, f"='{NAME}'!$D${r + 1}")
        ws.write_formula(r, 4, f'=IF(D{r + 1}<>"",D{r + 1},C{r + 1})',
                         fmts.calc(style))
        book.define_name(name + "Used", f"='{NAME}'!$E${r + 1}")
        s.merge(r, 5, LAST_COL, what, fmts.chrome("note"))
        r += 1
    direct_first, direct_last = first_cost, r - 1
    s.write(r, 1, "Direct cost", fmts.chrome("total_label"))
    ws.write_formula(r, 4,
                     f'=ROUND(SUM(E{direct_first + 1}:E{direct_last + 1}),2)',
                     fmts.chrome("total_value",
                                 num_format=s.num_format("money")))
    book.define_name("RcDirect", f"='{NAME}'!$E${r + 1}")
    s.merge(r, 5, LAST_COL, "Materials + labour + packaging + postage - "
                            "everything except overhead.",
            fmts.chrome("note"))
    r += 2

    # ----------------------------------------------- 3. the overhead ---
    r = s.section(r, "3. Overhead - the costs that are not in the piece",
                  LAST_COL)
    r = s.note(r, "Studio rent, insurance, your website, the shop "
                  "subscription. The Overhead tab works out how it is "
                  "shared; this is that share applied to this product.",
               LAST_COL, first_col=1, height=26)
    r = s.field(r, "How overhead is shared", None, style="text",
                mode="calc", formula="=OverheadMethod", value_end=5,
                help_text="Change the method on the Settings tab.")
    r = s.field(r, "Overhead on this product", None, style="money",
                mode="calc", name="RcOverhead", value_end=5,
                formula='=IF(OverheadMethod="Per hour worked",'
                        'ROUND(OverheadPerHour*' + _pick("hours") + ',2),'
                        'IF(OR(OverheadMethod="Per unit made",'
                        'OverheadMethod="Per order"),OverheadPerUnit,0))',
                help_text="With 'Share of revenue' the overhead goes "
                          "into the price as a percentage instead - "
                          "see the next section.")
    r = s.field(r, "Total cost", None, style="money", mode="calc",
                name="RcCost", value_end=5,
                formula="=ROUND(RcDirect+RcOverhead,2)",
                help_text="This is C in the pricing formula.")
    r += 1

    # ------------------------------------------- 4. what the shop takes
    r = s.section(r, "4. What the marketplace takes", LAST_COL)
    r = s.note(r, "Percentage fees and fixed fees behave completely "
                  "differently, which is why adding a flat 'Etsy 20%' "
                  "to your cost gets the price wrong. Percentages come "
                  "out of the price you are solving for; fixed fees "
                  "are added to the cost.",
               LAST_COL, first_col=1, height=30)
    r = s.field(r, "Percentage fees", None, style="percent", mode="calc",
                name="RcFeePct", value_end=5,
                formula='=IFERROR(INDEX([[Channels.pct_total]],'
                        'MATCH(RcChannel,[[Channels.name]],0)),0)',
                help_text="Listing share, transaction, payment "
                          "processing, ads, regulatory - all editable "
                          "on the Channels tab.")
    r = s.field(r, "Fixed fees", None, style="money", mode="calc",
                name="RcFeeFixed", value_end=5,
                formula='=IFERROR(INDEX([[Channels.fixed_total]],'
                        'MATCH(RcChannel,[[Channels.name]],0)),0)',
                help_text="The listing fee and the fixed part of the "
                          "payment processing charge.")
    r = s.field(r, "Overhead as a share of revenue", None, style="percent",
                mode="calc", name="RcOhRate", value_end=5,
                formula='=IF(OverheadMethod="Share of revenue",'
                        'IF(RevenueTargetMonth<=0,0,OverheadMonthly/'
                        'RevenueTargetMonth),0)',
                help_text="Only used with the 'Share of revenue' "
                          "method - it comes out of the price, like a "
                          "fee.")
    r += 1

    # ------------------------------------------------ 5. the answer ----
    r = s.section(r, "5. The price", LAST_COL)
    denom_row = r
    s.write(r, 1, "Denominator", fmts.chrome("label_bold"))
    ws.write_formula(r, 3, "=1-RcTarget-RcFeePct-RcOhRate",
                     fmts.calc("number"))
    book.define_name("RcDenom", f"='{NAME}'!$D${denom_row + 1}")
    s.merge(r, 4, LAST_COL,
            "1 minus the margin you want, minus the percentage fees, "
            "minus the overhead share. If this is zero or less, no "
            "price works and the maths says so.",
            fmts.chrome("note"))
    r += 1
    s.write(r, 1, "Suggested price", fmts.chrome("label_bold"))
    ws.write_formula(
        r, 3,
        '=IF(RcDenom<=0,"No price works",ROUND((RcCost+RcFeeFixed)/'
        'RcDenom,2))',
        fmts.calc("money"))
    book.define_name("RcRaw", f"='{NAME}'!$D${r + 1}")
    s.merge(r, 4, LAST_COL,
            "(Cost + fixed fees) divided by the denominator. This is "
            "the price that leaves exactly the margin you asked for.",
            fmts.chrome("note"))
    r += 1
    s.write(r, 1, "Rounded for the listing", fmts.chrome("label_bold"))
    ws.write_formula(
        r, 3,
        '=IF(ISNUMBER(RcRaw),ROUND(' + round_price("RcRaw") + ',2),"")',
        fmts.calc("money"))
    book.define_name("RcPrice", f"='{NAME}'!$D${r + 1}")
    s.merge(r, 4, LAST_COL,
            "Using the rounding rule on Settings. Rounding up is free "
            "margin; rounding down is a gift.",
            fmts.chrome("note"))
    r += 2

    # ---------------------------------------------------- the KPI tiles
    tiles = [
        ("Price", "=IF(RcPrice=\"\",\"\",RcPrice)", "money",
         "what to charge"),
        ("Fees on that price",
         '=IF(RcPrice="","",ROUND(RcPrice*RcFeePct+RcFeeFixed,2))',
         "money", "what the channel takes"),
        ("Profit",
         '=IF(RcPrice="","",ROUND(RcPrice-RcCost-(RcPrice*RcFeePct+'
         'RcFeeFixed),2))', "money", "what you keep"),
        ("Margin",
         '=IF(OR(RcPrice="",RcPrice=0),"",(RcPrice-RcCost-'
         '(RcPrice*RcFeePct+RcFeeFixed))/RcPrice)', "percent",
         "profit over price"),
    ]
    for i, (label, formula, style, note) in enumerate(tiles):
        s.kpi(r, 1 + i * 2, label, formula, style=style, span=2, note=note)
    r += 3
    tiles_2 = [
        ("Markup",
         '=IF(OR(RcPrice="",RcCost=0),"",(RcPrice-RcCost-'
         '(RcPrice*RcFeePct+RcFeeFixed))/RcCost)', "percent",
         "profit over cost"),
        ("Hours in it", "=" + _pick("hours", '0'), "number",
         "from the catalog"),
        ("You earn an hour",
         '=IF(OR(' + _pick("hours", "0") + '=0,RcPrice=""),"",'
         'ROUND((RcPrice-RcCost-(RcPrice*RcFeePct+RcFeeFixed)+'
         'RcLabUsed)/' + _pick("hours", "1") + ',2))', "money",
         "profit plus your wage"),
        ("Wholesale at this price",
         '=IF(RcPrice="","",ROUND(RcPrice*(1-WholesaleDiscount),2))',
         "money", "half retail, by default"),
    ]
    for i, (label, formula, style, note) in enumerate(tiles_2):
        s.kpi(r, 1 + i * 2, label, formula, style=style, span=2, note=note)
    r += 4

    # --------------------------------------------------- the workings --
    r = s.section(r, "6. How that number was worked out", LAST_COL)
    r = s.table_head(r, 1, ["Step", "Value", "Where it comes from"],
                     [30, 15, 60])
    steps = [
        ("Materials, labour and the rest", "=RcDirect", "money",
         "Section 2 above"),
        ("Overhead", "=RcOverhead", "money",
         "The Overhead tab, shared by your chosen method"),
        ("C - total cost", "=RcCost", "money", "The two lines above"),
        ("F - fixed fees", "=RcFeeFixed", "money",
         "Channels tab: listing fee plus the fixed processing charge"),
        ("r - target margin", "=RcTarget", "percent",
         "Your target, from Settings or the box above"),
        ("f - percentage fees", "=RcFeePct", "percent",
         "Channels tab: everything charged as a percentage"),
        ("1 - r - f", "=RcDenom", "number",
         "The share of the price that is left to cover the cost"),
        ("(C + F) / (1 - r - f)", "=RcRaw", "money",
         "The price that leaves exactly your target margin"),
        ("Rounded", "=RcPrice", "money",
         "Your rounding rule from the Settings tab"),
    ]
    for label, formula, style, where in steps:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, formula, fmts.calc(style))
        s.merge(r, 3, LAST_COL, where, fmts.chrome("note"))
        r += 1
    r += 1

    # ------------------------------------------------- the margin ladder
    r = s.section(r, "7. The same product at other margins", LAST_COL)
    r = s.note(r, "What you would have to charge for each level of "
                  "margin, and what you would earn an hour at it. The "
                  "row matching your target is the suggestion above.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Target margin", "Price", "Fees", "Profit",
                            "You earn an hour", "Markup"],
                     [16, 14, 12, 13, 17, 12])
    ladder_first = r
    for margin in LADDER:
        denom = f"(1-{margin}-RcFeePct-RcOhRate)"
        price = f"(RcCost+RcFeeFixed)/{denom}"
        s.write(r, 1, margin, fmts.calc("percent"))
        ws.write_formula(r, 2, f'=IF({denom}<=0,"",ROUND({price},2))',
                         fmts.calc("money"))
        ws.write_formula(
            r, 3, f'=IF(C{r + 1}="","",ROUND(C{r + 1}*RcFeePct+'
                  f'RcFeeFixed,2))', fmts.calc("money"))
        ws.write_formula(
            r, 4, f'=IF(C{r + 1}="","",ROUND(C{r + 1}-RcCost-'
                  f'D{r + 1},2))', fmts.calc("money"))
        ws.write_formula(
            r, 5, '=IF(OR(C{0}="",{1}=0),"",ROUND((E{0}+RcLabUsed)/'
                  '{1},2))'.format(r + 1, hours),
            fmts.calc("money"))
        ws.write_formula(
            r, 6, f'=IF(OR(C{r + 1}="",RcCost=0),"",E{r + 1}/RcCost)',
            fmts.calc("percent"))
        r += 1
    ws.conditional_format(ladder_first, 2, r - 1, 2,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    ws.conditional_format(ladder_first, 5, r - 1, 5,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    ws.conditional_format(ladder_first, 5, r - 1, 5,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetHourly",
                           "format": fmts.cf("good")})
    r += 1

    # ----------------------------------------------- what if you charged
    r = s.section(r, "8. What if you charged something else?", LAST_COL)
    r = s.note(r, "The suggested price, moved up and down. This is the "
                  "table to look at before you knock a few dollars off "
                  "because a price feels high.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Price", "Versus suggested", "Fees", "Profit",
                            "Margin", "Markup", "You earn an hour",
                            "Verdict"],
                     [14, 14, 12, 12, 10, 10, 16, 26])
    step_first = r
    for step in STEPS:
        price = f'IF(ISNUMBER(RcRaw),ROUND(RcRaw*{1 + step},2),"")'
        ws.write_formula(r, 1, f"={price}", fmts.calc("money"))
        s.write(r, 2, step, fmts.calc("percent0"))
        ws.write_formula(
            r, 3, f'=IF(B{r + 1}="","",ROUND(B{r + 1}*RcFeePct+'
                  f'RcFeeFixed,2))', fmts.calc("money"))
        ws.write_formula(
            r, 4, f'=IF(B{r + 1}="","",ROUND(B{r + 1}-RcCost-'
                  f'D{r + 1},2))', fmts.calc("money"))
        ws.write_formula(
            r, 5, f'=IF(OR(B{r + 1}="",B{r + 1}=0),"",E{r + 1}/'
                  f'B{r + 1})', fmts.calc("percent"))
        ws.write_formula(
            r, 6, f'=IF(OR(B{r + 1}="",RcCost=0),"",E{r + 1}/RcCost)',
            fmts.calc("percent"))
        ws.write_formula(
            r, 7, '=IF(OR(B{0}="",{1}=0),"",ROUND((E{0}+RcLabUsed)/'
                  '{1},2))'.format(r + 1, hours),
            fmts.calc("money"))
        ws.write_formula(
            r, 8,
            f'=IF(B{r + 1}="","",IF(E{r + 1}<0,"Loss - below cost",'
            f'IF(F{r + 1}<MarginWarning,"Too thin",'
            f'IF(F{r + 1}<RcTarget,"Under your target",'
            f'IF(F{r + 1}>RcTarget+0.1,"Above target - test it",'
            f'"On target")))))', fmts.calc("text"))
        r += 1
    ws.conditional_format(step_first, 4, r - 1, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(step_first, 8, r - 1, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "Loss", "format": fmts.cf("bad")})
    ws.conditional_format(step_first, 8, r - 1, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "On target", "format": fmts.cf("good")})
    r += 1

    # -------------------------------------------------- the sanity check
    r = s.section(r, "9. Does this price make sense?", LAST_COL)
    checks = [
        ("Covers your costs",
         '=IF(RcPrice="","Pick a product or type the costs",'
         'IF(RcPrice>RcCost+RcPrice*RcFeePct+RcFeeFixed,'
         '"Yes - there is profit in it","No - this price loses money"))'),
        ("Pays you properly",
         '=IF(OR(RcPrice="",' + _pick("hours", "0") + '=0),'
         '"Add the labour minutes on the Labor tab",'
         'IF((RcPrice-RcCost-(RcPrice*RcFeePct+RcFeeFixed)+RcLabUsed)/'
         + _pick("hours", "1") + '>=MinimumWage,'
         '"Yes - above minimum wage",'
         '"No - you are paying to work"))'),
        ("Works for wholesale",
         '=IF(RcPrice="","-",IF(RcPrice*(1-WholesaleDiscount)>RcCost,'
         '"Yes - there is room for a stockist",'
         '"No - wholesale would sell below cost"))'),
        ("Against the market",
         '=IF(OR(RcPrice="",' + _pick("comp_price", '""') + '=""),'
         '"No competitor research for this product yet",'
         '"The shops you logged average "&TEXT(' +
         _pick("comp_price", "0") + ',"0.00")&" - you would be "&'
         'TEXT(ABS(RcPrice-' + _pick("comp_price", "0") +
         '),"0.00")&IF(RcPrice>' + _pick("comp_price", "0") +
         '," dearer"," cheaper"))'),
        ("Fee settings",
         '=IF(FeesStale="Yes","Check - your fee rates were last '
         'reviewed "&TEXT(FeesReviewed,"d mmm yyyy")&". Marketplace '
         'fees change; verify the current official fee schedule.",'
         '"Fee rates reviewed "&TEXT(FeesReviewed,"d mmm yyyy"))'),
    ]
    check_first = r
    for label, formula in checks:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc("text"))
        r += 1
    for word, kind in (("No -", "bad"), ("Yes -", "good"),
                       ("Check", "warn")):
        ws.conditional_format(check_first, 2, r - 1, 2,
                              {"type": "text", "criteria": "containing",
                               "value": word, "format": fmts.cf(kind)})
    r += 1

    s.note(r,
           "Happy with the number? Put it in the 'Your price' column "
           "on the Product Catalog and log the change on the Price "
           "History tab. Nothing on this page is saved - it is a "
           "calculator, not a record.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
