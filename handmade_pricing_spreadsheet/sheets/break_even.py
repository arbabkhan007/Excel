"""`Profit & Break-Even` - how many you have to sell before you earn anything.

Break-even is the least glamorous number in the workbook and the most
useful. It answers one question: given what the studio costs every
month and what each piece contributes, how many do you have to sell
before you are working for yourself rather than for the landlord?

    Contribution a unit = price - fees - the costs that vary
    Break-even units    = fixed costs / contribution a unit

If the contribution is zero or negative, no number of sales works and
the page says so rather than printing something enormous.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

NAME = "Profit & Break-Even"
LAST_COL = 11
PRODUCT_ROWS = 15
PRICE_STEPS = [-0.20, -0.10, 0.0, 0.10, 0.20, 0.30]


def _pick(field: str, default: str = "0") -> str:
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH(BeProduct,'
            f'[[Products.name]],0)),{default})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Profit, Margin & Break-Even",
             "What each piece contributes, how many pay for the month, "
             "and the difference between margin and markup.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 32, fmts.chrome("label"))
    s.set_col(2, 2, 12, fmts.chrome("label"))
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 15)
    s.set_col(5, 5, 15)
    s.set_col(6, LAST_COL, 14, fmts.chrome("note"))

    r = 4

    # -------------------------------------------------------- the inputs
    r = s.section(r, "1. The numbers it works from", LAST_COL)
    r = s.field(r, "Product", "Willow cable beanie", name="BeProduct",
                validate=list_from("ProductNames"), value_end=5,
                help_text="Or leave it and use the override boxes "
                          "below to test something hypothetical.")
    r = s.field(r, "Fixed costs a month", None, style="money",
                name="BeFixed", validate=number_check(0, 10000000),
                formula="=ROUND(OverheadMonthly,2)",
                help_text="From the Overhead tab. Type over it to "
                          "include a wage you want to draw.")
    r = s.field(r, "Wage you want to draw a month", 0, style="money",
                name="BeWage", validate=number_check(0, 10000000),
                help_text="Treat your own pay as a fixed cost and "
                          "break-even becomes a living, not a "
                          "survival.")
    r = s.field(r, "Price", None, style="money", name="BePrice",
                validate=number_check(0, 1000000),
                formula="=" + _pick("price", '""'),
                help_text="The catalog price. Override it to test a "
                          "different one.")
    r = s.field(r, "Variable cost a unit", None, style="money",
                name="BeVariable", validate=number_check(0, 1000000),
                formula="=" + _pick("cost_base", "0"),
                help_text="Materials, labour, packaging and postage - "
                          "everything that only happens when you make "
                          "one. Overhead is the fixed part and is "
                          "counted above.")
    r = s.field(r, "Percentage fees", None, style="percent",
                mode="calc", name="BeFeePct",
                formula="=" + _pick("fee_pct", "0"),
                help_text="From the channel this product sells on.")
    r = s.field(r, "Fixed fees a sale", None, style="money", mode="calc",
                name="BeFeeFixed",
                formula="=" + _pick("fee_fixed", "0"))
    r += 1

    # --------------------------------------------------- the calculation
    r = s.section(r, "2. The calculation, step by step", LAST_COL)
    r = s.table_head(r, 1, ["Step", "Value", "What it means"],
                     [32, 15, 58])
    steps = [
        ("Fixed costs a month", "=ROUND(BeFixed+BeWage,2)", "money",
         "BeFixedTotal",
         "Overhead plus whatever you want to pay yourself."),
        ("Price", '=IF(BePrice="",0,BePrice)', "money", "BeP",
         "What the buyer pays for the item."),
        ("Fees on that price",
         "=ROUND(BeP*BeFeePct+BeFeeFixed,2)", "money", "BeFees",
         "Percentage fees plus the fixed ones."),
        ("Variable cost", '=IF(BeVariable="",0,BeVariable)', "money",
         "BeVar", "Everything that only happens when you make one."),
        ("Contribution a unit",
         "=ROUND(BeP-BeFees-BeVar,2)", "money", "BeContribution",
         "What each sale puts towards the fixed costs. This is the "
         "number that matters."),
        ("Contribution margin",
         '=IF(BeP=0,"",BeContribution/BeP)', "percent", "BeCm",
         "Contribution as a share of the price."),
    ]
    for label, formula, style, name, note in steps:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc(style))
        book.define_name(name, f"='{NAME}'!$C${r + 1}")
        s.merge(r, 3, LAST_COL, note, fmts.chrome("note"))
        r += 1
    ws.conditional_format(r - 2, 2, r - 2, 2,
                          {"type": "cell", "criteria": "<=", "value": 0,
                           "format": fmts.cf("bad")})
    r += 1

    # ----------------------------------------------------- the KPI tiles
    tiles = [
        ("Break-even units a month",
         '=IF(BeContribution<=0,"",ROUNDUP(BeFixedTotal/'
         'BeContribution,0))', "int", "fixed costs / contribution"),
        ("Break-even revenue",
         '=IF(BeContribution<=0,"",ROUNDUP(BeFixedTotal/'
         'BeContribution,0)*BeP)', "money0", "at this price"),
        ("Units a week",
         '=IF(BeContribution<=0,"",ROUNDUP(BeFixedTotal/'
         'BeContribution/4.33,0))', "int", "to keep pace"),
        ("Hours a month to make them",
         '=IF(BeContribution<=0,"",ROUND(ROUNDUP(BeFixedTotal/'
         'BeContribution,0)*' + _pick("hours", "0") + ',1))',
         "number1", "is that even possible?"),
    ]
    for i, (label, formula, style, note) in enumerate(tiles):
        s.kpi(r, 1 + i * 2, label, formula, style=style, span=2,
              note=note)
    r += 4

    warn_row = r
    s.merge(r, 1, LAST_COL, None, fmts.calc("text"))
    ws.write_formula(
        r, 1,
        '=IF(BeContribution>0,"Every sale puts "&TEXT(BeContribution,'
        '"0.00")&" towards the "&TEXT(BeFixedTotal,"0")&" of fixed '
        'costs - "&TEXT(ROUNDUP(BeFixedTotal/BeContribution,0),"0")&'
        '" sales a month and the studio pays for itself.",'
        '"Warning: this product contributes nothing. At this price '
        'the fees and the variable costs are more than the price, so '
        'no number of sales will ever break even. Raise the price or '
        'cut the cost.")', fmts.calc("text"))
    s.row_height(r, 30)
    ws.conditional_format(warn_row, 1, warn_row, 1,
                          {"type": "text", "criteria": "containing",
                           "value": "Warning", "format": fmts.cf("bad")})
    r += 2

    # ----------------------------------------------- margin versus markup
    r = s.section(r, "3. Margin and markup are not the same thing",
                  LAST_COL)
    r = s.note(r, "Margin is profit over the price; markup is profit "
                  "over the cost. A 50% markup is a 33% margin, and "
                  "mixing them up is how a shop ends up pricing at "
                  "cost plus a bit.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["This product", "Value", "Worked out as"],
                     [32, 15, 58])
    pair = [
        ("Profit a unit", "=ROUND(BeContribution,2)", "money",
         "Price - fees - variable cost"),
        ("Margin", '=IF(BeP=0,"",BeContribution/BeP)', "percent",
         "Profit / price x 100"),
        ("Markup", '=IF(BeVar=0,"",BeContribution/BeVar)', "percent",
         "Profit / cost x 100"),
        ("Multiplier on cost", '=IF(BeVar=0,"",BeP/BeVar)', "number",
         "How many times your cost the price is"),
    ]
    for label, formula, style, how in pair:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, formula, fmts.calc(style))
        s.merge(r, 3, LAST_COL, how, fmts.chrome("note"))
        r += 1
    r += 1

    r = s.table_head(r, 1, ["If you want this margin", "The markup is",
                            "Multiply your cost by", "Example on $10 of cost"],
                     [24, 16, 22, 24])
    conv_first = r
    for margin in (0.20, 0.30, 0.35, 0.40, 0.45, 0.50, 0.55, 0.60, 0.65,
                   0.70):
        s.write(r, 1, margin, fmts.calc("percent0"))
        ws.write_formula(r, 2, f'={margin}/(1-{margin})',
                         fmts.calc("percent"))
        ws.write_formula(r, 3, f'=1/(1-{margin})', fmts.calc("number"))
        ws.write_formula(r, 4, f'=ROUND(10/(1-{margin}),2)',
                         fmts.calc("money"))
        r += 1
    ws.conditional_format(conv_first, 3, r - 1, 3,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    r += 1

    # -------------------------------------------------- price sensitivity
    r = s.section(r, "4. Break-even at other prices", LAST_COL)
    r = s.table_head(r, 1, ["Price", "Change", "Contribution",
                            "Margin", "Break-even units",
                            "Break-even revenue", "Verdict"],
                     [13, 10, 14, 10, 17, 18, 24])
    sens = r
    for step in PRICE_STEPS:
        price = f"ROUND(BeP*{1 + step},2)"
        contrib = f"ROUND({price}-({price}*BeFeePct+BeFeeFixed)-BeVar,2)"
        ws.write_formula(r, 1, f"={price}", fmts.calc("money"))
        s.write(r, 2, step, fmts.calc("percent0"))
        ws.write_formula(r, 3, f"={contrib}", fmts.calc("money"))
        ws.write_formula(r, 4, f'=IF(B{r + 1}=0,"",D{r + 1}/B{r + 1})',
                         fmts.calc("percent"))
        ws.write_formula(r, 5, f'=IF(D{r + 1}<=0,"Never",'
                               f'ROUNDUP(BeFixedTotal/D{r + 1},0))',
                         fmts.calc("int"))
        ws.write_formula(r, 6, f'=IF(D{r + 1}<=0,"",ROUNDUP('
                               f'BeFixedTotal/D{r + 1},0)*B{r + 1})',
                         fmts.calc("money0"))
        ws.write_formula(
            r, 7,
            f'=IF(D{r + 1}<=0,"Loses money on every sale",'
            f'IF(E{r + 1}<MarginWarning,"Too thin to live on",'
            f'IF(F{r + 1}<=UnitsPerMonth,"Reachable at your volume",'
            f'"More than you currently make")))', fmts.calc("text"))
        r += 1
    ws.conditional_format(sens, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "Reachable", "format": fmts.cf("good")})
    ws.conditional_format(sens, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "Loses", "format": fmts.cf("bad")})
    ws.conditional_format(sens, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "More than", "format": fmts.cf("warn")})
    r += 1

    # ----------------------------------------------- every product's view
    r = s.section(r, "5. Every product against the monthly overhead",
                  LAST_COL)
    r = s.note(r, "How many of each you would have to sell in a month "
                  "to pay the fixed costs on that product alone. It is "
                  "a blunt test and a very clear one.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Product", "Price", "Cost", "Fees",
                            "Contribution", "Margin",
                            "Units to cover the month", "Hours that takes"],
                     [30, 11, 11, 11, 13, 10, 20, 15])
    table_first = r
    for i in range(1, PRODUCT_ROWS + 1):
        name = f'IFERROR(INDEX([[Products.name]],{i})&"","")'
        price = f'IFERROR(INDEX([[Products.price]],{i}),"")'
        cost = f'IFERROR(INDEX([[Products.cost_base]],{i}),"")'
        fees = f'IFERROR(INDEX([[Products.fees]],{i}),"")'
        hours = f'IFERROR(INDEX([[Products.hours]],{i}),0)'
        ws.write_formula(r, 1, s.reg.render(f"={name}"),
                         fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(f'=IF({name}="","",{price})'),
                         fmts.calc("money"))
        ws.write_formula(r, 3, s.reg.render(f'=IF({name}="","",{cost})'),
                         fmts.calc("money"))
        ws.write_formula(r, 4, s.reg.render(f'=IF({name}="","",{fees})'),
                         fmts.calc("money"))
        ws.write_formula(
            r, 5,
            f'=IF(C{r + 1}="","",ROUND(C{r + 1}-D{r + 1}-E{r + 1},2))',
            fmts.calc("money"))
        ws.write_formula(
            r, 6,
            f'=IF(OR(C{r + 1}="",C{r + 1}=0),"",F{r + 1}/C{r + 1})',
            fmts.calc("percent"))
        ws.write_formula(
            r, 7,
            f'=IF(C{r + 1}="","",IF(F{r + 1}<=0,"Never",'
            f'ROUNDUP(BeFixedTotal/F{r + 1},0)))', fmts.calc("int"))
        ws.write_formula(
            r, 8,
            s.reg.render(f'=IF(OR(C{r + 1}="",F{r + 1}<=0),"",'
                         f'ROUND(H{r + 1}*{hours},0))'),
            fmts.calc("int"))
        r += 1
    ws.conditional_format(table_first, 6, r - 1, 6,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("warn")})
    ws.conditional_format(table_first, 7, r - 1, 7,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    r += 1

    # ------------------------------------------------- where you are now
    r = s.section(r, "6. Where you actually are", LAST_COL)
    r = s.table_head(r, 1, ["This month", "Value", "What it tells you"],
                     [30, 15, 58])
    now = [
        ("Revenue so far",
         '=ROUND(SUMIFS([[Sales.revenue]],[[Sales.date]],">="&MonthStart,'
         '[[Sales.date]],"<="&MonthEnd),2)', "money",
         "Everything logged on the Sales tab this month."),
        ("Contribution so far",
         '=ROUND(SUMIFS([[Sales.profit]],[[Sales.date]],">="&MonthStart,'
         '[[Sales.date]],"<="&MonthEnd),2)', "money",
         "Profit on the orders logged, before the studio overhead."),
        ("Fixed costs to cover", "=ROUND(BeFixed+BeWage,2)", "money",
         "What the month has to pay for."),
        ("Still to find",
         '=ROUND(MAX(0,BeFixed+BeWage-SUMIFS([[Sales.profit]],'
         '[[Sales.date]],">="&MonthStart,[[Sales.date]],"<="&MonthEnd)),2)',
         "money", "The gap between here and break-even."),
        ("Days left in the month",
         "=MAX(0,MonthEnd-TodayDate)", "int",
         "How long you have to find it."),
    ]
    for label, formula, style, what in now:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc(style))
        s.merge(r, 3, LAST_COL, what, fmts.chrome("note"))
        r += 1
    r += 1

    s.note(r,
           "Break-even is not a target, it is a floor. Pay yourself a "
           "wage in the box at the top and the number becomes the "
           "honest answer to 'how much do I have to sell to make this "
           "a job?'.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
