"""`Reports` - the month in one page, and the year in one table.

Pick a month at the top and every table underneath follows it: what
you spent, what you cooked, what you wasted, how the week-by-week
pattern moved. Underneath that, a twelve row year table you can copy
straight into a message to whoever else pays for the food.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from
#: the Budget tab owns the list; the two pages must agree
from .budget import SPEND_TYPES

NAME = "Reports"
LAST_COL = 9


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Reporting & Analytics",
             "Pick a month; every table on this page follows it.",
             LAST_COL)
    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 14)

    # ------------------------------------------------------ the picker --
    row = 4
    s.write(row, 0, "Month", fmts.chrome("label_bold"))
    s.write(row, 1, "", fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("MonthList"))
    book.define_name("ReportMonth", f"={quote_sheet(NAME)}!$B${row + 1}")
    ws.write_formula(row, 2,
                     '=IFERROR(DATE(VALUE(LEFT(ReportMonth,4)),'
                     'VALUE(RIGHT(ReportMonth,2)),1),'
                     'DATE(YEAR(TODAY()),MONTH(TODAY()),1))',
                     fmts.calc("date"))
    book.define_name("ReportStart", f"={quote_sheet(NAME)}!$C${row + 1}")
    s.merge(row, 3, LAST_COL,
            "Leave it blank and the page shows the month you are in. "
            "The list comes from the Lists tab.", fmts.chrome("note"))
    ws.write_formula(row, 11, s.reg.render('=MAX(1,(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,1,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,2,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,3,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,4,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,5,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,6,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,7,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,8,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,9,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,10,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,11,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,12,1),"yyyy-mm"))>0)*1)'),
                     fmts.calc("int"))
    book.define_name("ReportMonths", f"={quote_sheet(NAME)}!$L${row + 1}")
    ws.set_column(11, 11, None, None, {"hidden": 1})
    row += 2

    M = 'TEXT(ReportStart,"yyyy-mm")'

    kpis = [
        ("Total food spend",
         f'=SUMIFS([[Grocery Spending.amount]],'
         f'[[Grocery Spending.month]],{M})', "money", "every receipt"),
        ("Cost per head per day",
         f'=IFERROR(SUMIFS([[Grocery Spending.amount]],'
         f'[[Grocery Spending.month]],{M})/HouseMembers/'
         f'DAY(EOMONTH(ReportStart,0)),"")', "money",
         "the comparable number"),
        ("Meals cooked at home",
         f'=SUMIFS([[Meal Plan.home_cooked]],[[Meal Plan.month]],{M})',
         "int", "from the meal plan"),
        ("Food binned",
         f'=SUMIFS([[Food Waste.value]],[[Food Waste.month]],{M})',
         "money", "the avoidable part hurts"),
        ("Against budget",
         f'=IFERROR(SUMIFS([[Grocery Spending.amount]],'
         f'[[Grocery Spending.month]],{M})/MonthlyBudget,"")',
         "percent0", "under 100% is the win"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ------------------------------------------------- spend breakdown --
    row = s.section(row, "Where the money went", LAST_COL)
    row = s.table_head(row, 0, ["Type of spend", "This month", "Share",
                                "Last month", "Change",
                                "Monthly average"],
                       [26, 14, 12, 14, 14, 16])
    first = row
    for kind in SPEND_TYPES:
        s.write(row, 0, kind, fmts.chrome("label"))
        this = (f'SUMIFS([[Grocery Spending.amount]],'
                f'[[Grocery Spending.month]],{M},'
                f'[[Grocery Spending.type]],"{kind}")')
        prev = (f'SUMIFS([[Grocery Spending.amount]],'
                f'[[Grocery Spending.month]],'
                f'TEXT(EDATE(ReportStart,-1),"yyyy-mm"),'
                f'[[Grocery Spending.type]],"{kind}")')
        s.calc_row(row, 1, [
            f'={this}',
            f'=IFERROR({this}/SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{M}),"")',
            f'={prev}',
            f'=IF({prev}=0,"",{this}/{prev}-1)',
            f'=IFERROR(SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}")/ReportMonths,"")',
        ], ["money", "percent0", "money", "percent", "money"])
        row += 1
    s.write(row, 0, "Total", fmts.chrome("total_label"))
    for col in (1, 3):
        letter = chr(ord("A") + col)
        ws.write_formula(row, col,
                         f"=SUM({letter}{first + 1}:{letter}{row})",
                         fmts.chrome("total_value",
                                     num_format=cfg.money_format))
    for col in (2, 4, 5):
        s.write(row, col, None, fmts.chrome("total_value"))
    ws.conditional_format(first, 4, row - 1, 4,
                          {"type": "cell", "criteria": ">",
                           "value": 0.1, "format": fmts.cf("bad")})
    ws.conditional_format(first, 4, row - 1, 4,
                          {"type": "cell", "criteria": "<",
                           "value": -0.1, "format": fmts.cf("good")})
    row += 2

    # ------------------------------------------------ cooking the month --
    row = s.section(row, "What the month looked like in the kitchen",
                    LAST_COL)
    row = s.table_head(row, 0, ["Measure", "This month", "Target",
                                "Verdict", "Notes"],
                       [30, 14, 14, 16, 40])
    lines = [
        ("Meals planned",
         f'=COUNTIFS([[Meal Plan.month]],{M})', '=""', "int",
         "every slot with something in it"),
        ("Dinners cooked at home",
         f'=COUNTIFS([[Meal Plan.month]],{M},[[Meal Plan.slot]],'
         f'"Dinner",[[Meal Plan.home_cooked]],1)',
         '=HomeCookedTarget*DAY(EOMONTH(ReportStart,0))/7', "int",
         "the headline habit"),
        ("Takeaways and meals out",
         f'=COUNTIFS([[Meal Plan.month]],{M},[[Meal Plan.home_cooked]],0)',
         '=TakeawayLimit*DAY(EOMONTH(ReportStart,0))/7', "int",
         "no judgement, just the number"),
        ("Different recipes used",
         f'=SUMPRODUCT((COUNTIFS([[Meal Plan.month]],{M},'
         f'[[Meal Plan.recipe]],[[Recipes.id]])>0)*1)',
         '=""', "int", "variety, not volume"),
        ("New recipes added",
         '=COUNTIFS([[Recipes.added]],">="&ReportStart,'
         '[[Recipes.added]],"<="&EOMONTH(ReportStart,0))',
         '=NewRecipeTarget', "int", "one a month is plenty"),
        ("Packed lunches",
         f'=COUNTIFS([[Meal Plan.month]],{M},[[Meal Plan.slot]],'
         f'"Lunch",[[Meal Plan.home_cooked]],1)',
         '=PackedLunchTarget*DAY(EOMONTH(ReportStart,0))/7', "int",
         "where the easy money is"),
        ("Batch cook sessions",
         '=COUNTIFS([[Meal Prep.date]],">="&ReportStart,'
         '[[Meal Prep.date]],"<="&EOMONTH(ReportStart,0))',
         '=""', "int", "from the Meal Prep tab"),
        ("Portions frozen",
         '=SUMIFS([[Meal Prep.portions]],[[Meal Prep.date]],">="&'
         'ReportStart,[[Meal Prep.date]],"<="&EOMONTH(ReportStart,0))',
         '=""', "number1", "future you says thanks"),
    ]
    for label, formula, target, style, note in lines:
        s.write(row, 0, label, fmts.chrome("label"))
        s.calc_row(row, 1, [
            formula, target,
            f'=IF(C{row + 1}="","",IF(B{row + 1}>=C{row + 1},'
            f'"On target","Short"))',
        ], [style, style, "status"])
        s.write(row, 4, note, fmts.chrome("note"))
        row += 1
    row += 2

    # ------------------------------------------------------ week by week --
    row = s.section(row, "Week by week inside the month", LAST_COL)
    row = s.table_head(row, 0, ["Week beginning", "Meals planned",
                                "Planned cost", "Spent", "Wasted",
                                "Cost per head"],
                       [18, 14, 14, 14, 14, 16])
    for n in range(5):
        start = f"A{row + 1}"
        ws.write_formula(
            row, 0,
            f'=IF(ReportStart+{n * 7}>EOMONTH(ReportStart,0),"",'
            f'ReportStart+{n * 7})', fmts.calc("date"))
        s.calc_row(row, 1, [
            f'=IF({start}="","",COUNTIFS([[Meal Plan.date]],">="&{start},'
            f'[[Meal Plan.date]],"<="&{start}+6))',
            f'=IF({start}="","",SUMIFS([[Meal Plan.cost]],'
            f'[[Meal Plan.date]],">="&{start},[[Meal Plan.date]],'
            f'"<="&{start}+6))',
            f'=IF({start}="","",SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.date]],">="&{start},'
            f'[[Grocery Spending.date]],"<="&{start}+6))',
            f'=IF({start}="","",SUMIFS([[Food Waste.value]],'
            f'[[Food Waste.date]],">="&{start},[[Food Waste.date]],'
            f'"<="&{start}+6))',
            f'=IF({start}="","",IFERROR(SUMIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.date]],'
            f'">="&{start},[[Grocery Spending.date]],"<="&{start}+6)/'
            f'HouseMembers,""))',
        ], ["int", "money", "money", "money", "money"])
        row += 1
    row += 2

    # ------------------------------------------------------- year table --
    row = s.section(row, "The year in one table", LAST_COL)
    row = s.table_head(row, 0, ["Month", "Spend", "Budget", "Over/under",
                                "Meals at home", "Waste",
                                "Cost per head per day", "Takeaways",
                                "Shops"],
                       [14, 13, 13, 13, 14, 12, 18, 12, 10])
    y_first = row
    for m in range(12):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        ws.write_formula(row, 0, f"=DATE(PlanYear,{m + 1},1)",
                         fmts.calc("text", num_format="mmmm"))
        spend = (f'SUMIFS([[Grocery Spending.amount]],'
                 f'[[Grocery Spending.month]],{key})')
        s.calc_row(row, 1, [
            f'={spend}',
            '=MonthlyBudget',
            f'=MonthlyBudget-{spend}',
            f'=SUMIFS([[Meal Plan.home_cooked]],[[Meal Plan.month]],'
            f'{key})',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.month]],{key})',
            f'=IFERROR({spend}/HouseMembers/DAY(DATE(PlanYear,'
            f'{m + 2},0)),"")',
            f'=COUNTIFS([[Meal Plan.month]],{key},'
            f'[[Meal Plan.home_cooked]],0)',
            f'=COUNTIFS([[Grocery Spending.month]],{key})',
        ], ["money", "money", "money", "int", "money", "money", "int",
            "int"])
        row += 1
    s.write(row, 0, "Year", fmts.chrome("total_label"))
    for col in (1, 2, 3, 4, 5, 7, 8):
        letter = chr(ord("A") + col)
        style = "int" if col in (4, 7, 8) else "money"
        ws.write_formula(
            row, col, f"=SUM({letter}{y_first + 1}:{letter}{row})",
            fmts.chrome("total_value",
                        num_format=s.num_format(style)))
    ws.write_formula(
        row, 6, f"=IFERROR(AVERAGE(G{y_first + 1}:G{row}),0)",
        fmts.chrome("total_value", num_format=cfg.money_format))
    ws.conditional_format(y_first, 1, row - 1, 1,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    ws.conditional_format(y_first, 3, row - 1, 3,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row += 2

    # ---------------------------------------------------- store summary --
    row = s.section(row, "Shops, ranked by what you spend there",
                    LAST_COL)
    row = s.table_head(row, 0, ["Store", "Visits", "Total spent",
                                "Average basket", "Share of spend",
                                "Price level"],
                       [24, 12, 14, 16, 14, 14])
    for n in range(8):
        pick = f'MATCH({n + 1},[[Stores.spend_rank]],0)'
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX([[Stores.name]],{pick}),"")'),
            fmts.calc("text"))
        store = f"A{row + 1}"
        s.calc_row(row, 1, [
            f'=IF({store}="","",COUNTIFS([[Grocery Spending.store]],'
            f'{store}))',
            f'=IF({store}="","",SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.store]],{store}))',
            f'=IF({store}="","",IFERROR(SUMIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.store]],'
            f'{store})/COUNTIFS([[Grocery Spending.store]],{store}),""))',
            f'=IF({store}="","",IFERROR(SUMIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.store]],'
            f'{store})/SUM([[Grocery Spending.amount]]),""))',
            f'=IF({store}="","",IFERROR(INDEX([[Stores.price_level]],'
            f'MATCH({store},[[Stores.name]],0)),""))',
        ], ["int", "money", "money", "percent0", "text"])
        row += 1
    row += 1

    s.note(row,
           "Copy any of these tables straight out - they are plain "
           "cells, not a pivot table, so paste works properly. If a "
           "month looks empty, check the Grocery Spending tab has "
           "dates in it: every table on this page keys off the month "
           "column, which is worked out from the date you type.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
