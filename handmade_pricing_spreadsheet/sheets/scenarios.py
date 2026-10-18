"""`Scenarios` - the same product priced five ways, on every channel, with
the materials going up.

Three questions this page answers that no single price can:

* What would my whole range look like if I repositioned it?
* What do I have to charge on each channel to keep the same money?
* What happens to my margins when materials go up 15%?

Nothing here changes your catalog. It is a thinking tool.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check
from .common import round_price

NAME = "Scenarios"
LAST_COL = 11
PRODUCT_ROWS = 15
SCENARIOS = 5
CHANNEL_ROWS = 10
SHOCKS = [0.0, 0.05, 0.10, 0.15, 0.25, 0.40]


def _pick(field: str, default: str = "0") -> str:
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH(ScProduct,'
            f'[[Products.name]],0)),{default})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Advanced Pricing Scenarios",
             "Reposition the whole range, compare every channel, and "
             "see what a price rise in your materials really does.",
             LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 30, fmts.chrome("label"))
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 14)

    r = 4

    # ------------------------------------------- 1. the whole range, 5 ways
    r = s.section(r, "1. Your range at five different positions", LAST_COL)
    r = s.note(r, "Each column prices every product at that scenario's "
                  "target margin, using its own costs and its own "
                  "channel fees. The margins behind the five columns "
                  "are on the Lists tab and you can change them.",
               LAST_COL, first_col=1, height=30)

    r = s.table_head(r, 1, ["Product", "Cost", "Your price", "", "", "",
                            "", "", "Where your price sits"],
                     [30, 11, 12, 12, 12, 12, 12, 12, 22])
    header_row = r - 1
    for i in range(1, SCENARIOS + 1):
        ws.write_formula(header_row, 3 + i,
                         f'=IFERROR(INDEX(ScenarioNames,{i}),"")',
                         fmts.chrome("header"))

    first = r
    for n in range(1, PRODUCT_ROWS + 1):
        # &"" so an empty catalog row reads as blank: INDEX on an
        # empty cell hands back a zero, and a zero is not ""
        name = f'IFERROR(INDEX([[Products.name]],{n})&"","")'
        cost = f'IFERROR(INDEX([[Products.total_cost]],{n}),"")'
        price = f'IFERROR(INDEX([[Products.price]],{n}),"")'
        # *1 turns the blank an empty catalog row returns into an
        # error, which IFERROR then reads as zero - subtracting "" from
        # 1 is a #VALUE! and it would spread across the whole block
        fee_pct = f'IFERROR(INDEX([[Products.fee_pct]],{n})*1,0)'
        fee_fix = f'IFERROR(INDEX([[Products.fee_fixed]],{n})*1,0)'
        ws.write_formula(r, 1, s.reg.render(f"={name}"), fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(f'=IF({name}="","",{cost})'),
                         fmts.calc("money"))
        ws.write_formula(r, 3, s.reg.render(f'=IF({name}="","",{price})'),
                         fmts.calc("money"))
        for i in range(1, SCENARIOS + 1):
            margin = f'IFERROR(INDEX(ScenarioMargins,{i}),0)'
            denom = f"(1-{margin}-{fee_pct})"
            body = f"(C{r + 1}+{fee_fix})/{denom}"
            ws.write_formula(
                r, 3 + i,
                s.reg.render(f'=IF(C{r + 1}="","",IF({denom}<=0,"",'
                             f'ROUND(' + round_price(f"({body})") + ',2)))'),
                fmts.calc("money"))
        ws.write_formula(
            r, 9,
            f'=IF(D{r + 1}="","",IF(D{r + 1}<E{r + 1},"Below the lowest",'
            f'IF(D{r + 1}<F{r + 1},"Budget",IF(D{r + 1}<G{r + 1},'
            f'"Conservative",IF(D{r + 1}<H{r + 1},"Standard",'
            f'IF(D{r + 1}<I{r + 1},"Premium","Luxury"))))))',
            fmts.calc("text"))
        r += 1
    ws.conditional_format(first, 9, r - 1, 9,
                          {"type": "text", "criteria": "containing",
                           "value": "Below the lowest",
                           "format": fmts.cf("bad")})
    ws.conditional_format(first, 9, r - 1, 9,
                          {"type": "text", "criteria": "containing",
                           "value": "Budget", "format": fmts.cf("warn")})
    r += 1
    s.merge(r, 1, LAST_COL,
            "If most of your range sits in the 'Budget' column, that "
            "is a positioning decision you probably never made on "
            "purpose.", fmts.chrome("note"))
    r += 2

    # ------------------------------------------- 2. one product, each channel
    r = s.section(r, "2. One product on every channel", LAST_COL)
    r = s.field(r, "Product", "Willow cable beanie", name="ScProduct",
                validate=list_from("ProductNames"), value_end=5,
                help_text="The channel comparison below follows this "
                          "choice.")
    r = s.field(r, "Target margin", None, style="percent", name="ScTarget",
                validate=number_check(0, 0.95),
                formula="=TargetRetailMargin")
    r += 1
    r = s.note(r, "The same piece, the same costs, different fees. The "
                  "'price to keep the same margin' column is what you "
                  "should actually list it at on each site - charging "
                  "one price everywhere quietly gives the dearest "
                  "channel your profit.",
               LAST_COL, first_col=1, height=30)
    r = s.table_head(r, 1, ["Channel", "Percentage fees", "Fixed fees",
                            "Price for the same margin", "Fees you pay",
                            "You keep", "At your current price",
                            "Margin there"],
                     [22, 15, 12, 20, 13, 12, 18, 13])
    ch_first = r
    for n in range(1, CHANNEL_ROWS + 1):
        name = f'IFERROR(INDEX([[Channels.name]],{n})&"","")'
        pct = f'IFERROR(INDEX([[Channels.pct_total]],{n})*1,0)'
        fix = f'IFERROR(INDEX([[Channels.fixed_total]],{n})*1,0)'
        ws.write_formula(r, 1, s.reg.render(f"={name}"), fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(f'=IF({name}="","",{pct})'),
                         fmts.calc("percent"))
        ws.write_formula(r, 3, s.reg.render(f'=IF({name}="","",{fix})'),
                         fmts.calc("money"))
        cost = s.reg.render(_pick("total_cost"))
        ws.write_formula(
            r, 4,
            f'=IF(B{r + 1}="","",IF(1-ScTarget-C{r + 1}<=0,"",'
            f'ROUND(({cost}+D{r + 1})/(1-ScTarget-C{r + 1}),2)))',
            fmts.calc("money"))
        ws.write_formula(
            r, 5,
            f'=IF(E{r + 1}="","",ROUND(E{r + 1}*C{r + 1}+D{r + 1},2))',
            fmts.calc("money"))
        ws.write_formula(
            r, 6,
            f'=IF(E{r + 1}="","",ROUND(E{r + 1}-F{r + 1}-{cost},2))',
            fmts.calc("money"))
        price = s.reg.render(_pick("price", '""'))
        ws.write_formula(
            r, 7,
            f'=IF(OR(B{r + 1}="",{price}=""),"",ROUND({price}-'
            f'({price}*C{r + 1}+D{r + 1})-{cost},2))',
            fmts.calc("money"))
        ws.write_formula(
            r, 8,
            f'=IF(OR(H{r + 1}="",{price}=""),"",H{r + 1}/{price})',
            fmts.calc("percent"))
        r += 1
    ws.conditional_format(ch_first, 7, r - 1, 7,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(ch_first, 8, r - 1, 8,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("warn")})
    ws.conditional_format(ch_first, 4, r - 1, 4,
                          {"type": "data_bar", "bar_color": "#8C4A32"})
    r += 2

    # --------------------------------------------- 3. the cost shock table
    r = s.section(r, "3. What if your materials go up?", LAST_COL)
    r = s.note(r, "Yarn, silver, wax and timber all moved a long way "
                  "in the last few years. This is what each rise does "
                  "to the product above if you hold your price - and "
                  "what you would have to charge instead.",
               LAST_COL, first_col=1, height=30)
    r = s.table_head(r, 1, ["Materials go up by", "New material cost",
                            "New total cost", "Margin if you hold your price",
                            "Price to keep your margin", "Rise needed",
                            "Verdict"],
                     [18, 17, 16, 24, 20, 13, 24])
    shock_first = r
    mat = s.reg.render(_pick("mat_cost"))
    cost = s.reg.render(_pick("total_cost"))
    price = s.reg.render(_pick("price", '""'))
    fee_pct = s.reg.render(_pick("fee_pct"))
    fee_fix = s.reg.render(_pick("fee_fixed"))
    for shock in SHOCKS:
        s.write(r, 1, shock, fmts.calc("percent0"))
        ws.write_formula(r, 2, f"=ROUND({mat}*{1 + shock},2)",
                         fmts.calc("money"))
        ws.write_formula(r, 3, f"=ROUND({cost}+{mat}*{shock},2)",
                         fmts.calc("money"))
        ws.write_formula(
            r, 4,
            f'=IF(OR({price}="",{price}=0),"",({price}-D{r + 1}-'
            f'({price}*{fee_pct}+{fee_fix}))/{price})',
            fmts.calc("percent"))
        ws.write_formula(
            r, 5,
            f'=IF(1-ScTarget-{fee_pct}<=0,"",ROUND((D{r + 1}+{fee_fix})/'
            f'(1-ScTarget-{fee_pct}),2))', fmts.calc("money"))
        ws.write_formula(
            r, 6,
            f'=IF(OR({price}="",{price}=0,F{r + 1}=""),"",F{r + 1}/'
            f'{price}-1)', fmts.calc("percent"))
        ws.write_formula(
            r, 7,
            f'=IF(E{r + 1}="","",IF(E{r + 1}<0,"You would be losing money",'
            f'IF(E{r + 1}<MarginWarning,"Margin gone - must re-price",'
            f'IF(G{r + 1}>0.1,"Needs a real price rise",'
            f'"Absorbable"))))', fmts.calc("text"))
        r += 1
    ws.conditional_format(shock_first, 4, r - 1, 4,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("bad")})
    ws.conditional_format(shock_first, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "Absorbable", "format": fmts.cf("good")})
    r += 2

    # ------------------------------------------------ 4. volume and time
    r = s.section(r, "4. Quantity, time and what your hour is worth",
                  LAST_COL)
    r = s.note(r, "Making more of the same thing is usually faster per "
                  "piece. This is what that saving is worth - and what "
                  "it lets you offer a shop without losing anything.",
               LAST_COL, first_col=1, height=28)
    r = s.table_head(r, 1, ["If a batch of", "Minutes each",
                            "Labour each", "Cost each", "At your price",
                            "Margin", "You earn an hour"],
                     [14, 13, 13, 13, 14, 10, 16])
    minutes = s.reg.render(_pick("minutes"))
    labour = s.reg.render(_pick("labour_adj"))
    vol_first = r
    for qty, factor in ((1, 1.0), (3, 0.88), (6, 0.80), (12, 0.72),
                        (24, 0.66), (48, 0.60)):
        s.write(r, 1, qty, fmts.calc("int"))
        ws.write_formula(r, 2, f"=ROUND({minutes}*{factor},1)",
                         fmts.calc("number1"))
        ws.write_formula(r, 3, f"=ROUND({labour}*{factor},2)",
                         fmts.calc("money"))
        ws.write_formula(r, 4, f"=ROUND({cost}-{labour}+D{r + 1},2)",
                         fmts.calc("money"))
        ws.write_formula(
            r, 5, f'=IF({price}="","",ROUND({price}-E{r + 1}-'
                  f'({price}*{fee_pct}+{fee_fix}),2))',
            fmts.calc("money"))
        ws.write_formula(
            r, 6, f'=IF(OR({price}="",{price}=0),"",F{r + 1}/{price})',
            fmts.calc("percent"))
        ws.write_formula(
            r, 7, f'=IF(OR(C{r + 1}=0,F{r + 1}=""),"",ROUND((F{r + 1}+'
                  f'D{r + 1})/(C{r + 1}/60),2))', fmts.calc("money"))
        r += 1
    ws.conditional_format(vol_first, 7, r - 1, 7,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    r += 1
    s.merge(r, 1, LAST_COL,
            "The batch factors here are a reasonable rule of thumb. "
            "Log a few runs on the Batch Production tab and you will "
            "have your own.", fmts.chrome("note"))
    r += 2

    s.note(r,
           "Nothing on this page writes to your catalog. When a "
           "scenario convinces you, change the price on the Product "
           "Catalog and log it on Price History.",
           LAST_COL, first_col=1, height=28)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
