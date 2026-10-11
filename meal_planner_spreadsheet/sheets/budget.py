"""`Budget` - where the food money actually goes.

Twelve months down the page: what you budgeted, what you spent, what
you spent it on and what was left. Underneath, the split by type of
spend and by store, the cost of eating out, and the weekly rhythm that
tells you whether the budget is realistic or just optimistic.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Budget"
LAST_COL = 10
MONTHS = 12

SPEND_TYPES = ["Weekly shop", "Top-up shop", "Bulk shop", "Takeaway",
               "Eating out", "Work lunch", "Coffee & treats",
               "Delivery fee", "Special occasion", "Household"]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Food Budget",
             "Planned against actual, month by month, with the split "
             "that shows where it really goes.", LAST_COL)
    s.set_col(0, 0, 16)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 14)

    # ------------------------------------------------------- the targets
    row = 4
    s.merge(row, 0, 0, "Weekly budget", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=WeeklyBudget", fmts.calc("money"))
    s.merge(row, 2, 2, "Monthly", fmts.chrome("label_right"))
    ws.write_formula(row, 3, "=MonthlyBudget", fmts.calc("money"))
    s.merge(row, 4, 4, "A year", fmts.chrome("label_right"))
    ws.write_formula(row, 5, "=AnnualBudget", fmts.calc("money"))
    s.merge(row, 6, 6, "Per person a week", fmts.chrome("label_right"))
    ws.write_formula(row, 7,
                     "=IFERROR(WeeklyBudget/MAX(1,HouseMembers),\"\")",
                     fmts.calc("money"))
    s.merge(row, 8, 8, "Per meal", fmts.chrome("label_right"))
    ws.write_formula(row, 9,
                     "=IFERROR(WeeklyBudget/MAX(1,HouseMembers*21),\"\")",
                     fmts.calc("money"))
    row += 1
    row = s.note(row, "Change these on the Settings tab. Per meal "
                      "assumes three meals a day for everyone - a "
                      "useful reality check when the weekly figure "
                      "feels mean.", LAST_COL, height=26)

    # ------------------------------------------------- month by month --
    row = s.section(row, "Month by month", LAST_COL)
    row = s.table_head(row, 0, ["Month", "Budget", "Groceries",
                                "Eating out & takeaway", "Other",
                                "Total spent", "Left", "% of budget",
                                "Food waste", "Spend a head"],
                       [12, 12, 13, 19, 11, 13, 12, 12, 12, 13])
    first = row
    for m in range(MONTHS):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        groceries = (f'SUMIFS([[Grocery Spending.amount]],'
                     f'[[Grocery Spending.month]],{key},'
                     f'[[Grocery Spending.type]],"*shop*")')
        out = (f'SUMIFS([[Grocery Spending.amount]],'
               f'[[Grocery Spending.month]],{key},'
               f'[[Grocery Spending.type]],"Takeaway")+'
               f'SUMIFS([[Grocery Spending.amount]],'
               f'[[Grocery Spending.month]],{key},'
               f'[[Grocery Spending.type]],"Eating out")')
        total = (f'SUMIFS([[Grocery Spending.amount]],'
                 f'[[Grocery Spending.month]],{key})')
        ws.write_formula(row, 0,
                         f'=TEXT(DATE(PlanYear,{m + 1},1),"mmmm")',
                         fmts.calc("text"))
        s.calc_row(row, 1, [
            "=MonthlyBudget",
            f"={groceries}",
            f"={out}",
            f"={total}-{groceries}-({out})",
            f"={total}",
            f"=MonthlyBudget-{total}",
            f'=IFERROR(({total})/MonthlyBudget,"")',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.month]],{key})',
            f'=IFERROR(({total})/MAX(1,HouseMembers),"")',
        ], ["money", "money", "money", "money", "money", "money",
            "percent0", "money", "money"])
        row += 1
    last = row - 1
    ws.conditional_format(first, 6, last, 6,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first, 6, last, 6,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(first, 7, last, 7,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    ws.conditional_format(first, 5, last, 5,
                          {"type": "data_bar", "bar_color": "#2F5E43"})

    # year total line
    s.merge(row, 0, 0, "Year", fmts.chrome("total_label"))
    totals = [
        "=AnnualBudget",
        f'=SUM(C{first + 1}:C{last + 1})',
        f'=SUM(D{first + 1}:D{last + 1})',
        f'=SUM(E{first + 1}:E{last + 1})',
        f'=SUM(F{first + 1}:F{last + 1})',
        f'=AnnualBudget-SUM(F{first + 1}:F{last + 1})',
        f'=IFERROR(SUM(F{first + 1}:F{last + 1})/AnnualBudget,"")',
        f'=SUM(I{first + 1}:I{last + 1})',
        f'=IFERROR(SUM(F{first + 1}:F{last + 1})/MAX(1,HouseMembers),"")',
    ]
    for i, formula in enumerate(totals):
        num = "0%" if i == 6 else cfg.money_format
        ws.write_formula(row, i + 1, formula,
                         fmts.chrome("total_value", num_format=num))
    row += 2

    # ------------------------------------------------- split by type ---
    row = s.section(row, "What the money goes on", LAST_COL)
    row = s.table_head(row, 0, ["Type of spend", "This month",
                                "Year to date", "Share of the year",
                                "Times logged", "Average"],
                       [20, 14, 14, 16, 13, 13])
    first_type = row
    for kind in SPEND_TYPES:
        s.write(row, 0, kind, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}",'
            f'[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
            f'=SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}")',
            f'=IFERROR(SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}")/'
            f'SUM([[Grocery Spending.amount]]),"")',
            f'=COUNTIFS([[Grocery Spending.type]],"{kind}")',
            f'=IFERROR(AVERAGEIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}"),"")',
        ], ["money", "money", "percent0", "int", "money"])
        row += 1
    ws.conditional_format(first_type, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#D9703C"})
    row += 1

    # ------------------------------------------------- split by store --
    row = s.section(row, "Store by store", LAST_COL)
    row = s.table_head(row, 0, ["Store", "Shops", "Spent",
                                "Average shop", "Share", "Price level"],
                       [20, 10, 13, 14, 12, 14])
    for i in range(8):
        cell = f"A{row + 1}"
        ws.write_formula(row, 0,
                         f'=IFERROR(IF(INDEX(StoreNames,{i + 1})=0,"",'
                         f'INDEX(StoreNames,{i + 1})),"")',
                         fmts.calc("text"))
        s.calc_row(row, 1, [
            f'=IF({cell}="","",COUNTIFS([[Grocery Spending.store]],'
            f'{cell}))',
            f'=IF({cell}="","",SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.store]],{cell}))',
            f'=IF({cell}="","",IFERROR(AVERAGEIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.store]],'
            f'{cell}),""))',
            f'=IF({cell}="","",IFERROR(SUMIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.store]],'
            f'{cell})/SUM([[Grocery Spending.amount]]),""))',
            f'=IF({cell}="","",IFERROR(INDEX([[Stores.price_level]],'
            f'MATCH({cell},[[Stores.name]],0)),""))',
        ], ["int", "money", "money", "percent0", "text"])
        row += 1
    row += 1

    # ---------------------------------------------- planned vs actual --
    row = s.section(row, "Planned against actual", LAST_COL)
    row = s.note(row, "The planner prices your meals from the "
                      "Ingredients tab. The till tells you what really "
                      "happened. The gap between them is the honest "
                      "measure of how good your price book is - and of "
                      "how much gets bought that was never planned.",
                 LAST_COL, height=32)
    row = s.table_head(row, 0, ["", "Planned (recipes)",
                                "Actual (receipts)", "Difference",
                                "Unplanned buys", "Impulse spend"],
                       [16, 18, 18, 14, 14, 14])
    gap_row = row
    s.write(row, 0, "This month", fmts.chrome("label_bold"))
    s.calc_row(row, 1, [
        '=SUMIFS([[Meal Plan.cost]],[[Meal Plan.month]],'
        'TEXT(TODAY(),"yyyy-mm"))',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))-'
        'SUMIFS([[Meal Plan.cost]],[[Meal Plan.month]],'
        'TEXT(TODAY(),"yyyy-mm"))',
        '=COUNTIFS([[Grocery Spending.planned]],"No",'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.planned]],"No",'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
    ], ["money", "money", "money", "int", "money"])
    row += 1
    s.write(row, 0, "The year", fmts.chrome("label_bold"))
    s.calc_row(row, 1, [
        '=SUM([[Meal Plan.cost]])',
        '=SUM([[Grocery Spending.amount]])',
        '=SUM([[Grocery Spending.amount]])-SUM([[Meal Plan.cost]])',
        '=COUNTIF([[Grocery Spending.planned]],"No")',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.planned]],"No")',
    ], ["money", "money", "money", "int", "money"])
    ws.conditional_format(gap_row, 3, row, 3,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 2

    # ------------------------------------------------- savings tracker --
    row = s.section(row, "What the workbook has saved you", LAST_COL)
    row = s.table_head(row, 0, ["Leftovers eaten",
                                "Cheaper shops available",
                                "Waste avoided against target",
                                "Batch cooking against takeaway"],
                       [18, 22, 26, 28])
    s.calc_row(row, 0, [
        '=SUMIFS([[Leftovers.value]],[[Leftovers.status]],"Eaten")',
        '=SUM([[Ingredients.saving]])',
        '=MAX(0,WasteCostTarget-SUMIFS([[Food Waste.value]],'
        '[[Food Waste.month]],TEXT(TODAY(),"yyyy-mm")))',
        '=IFERROR(SUM([[Meal Prep.portions]])*'
        '(EatingOutBudget/MAX(1,TakeawayLimit*2))-'
        'SUM([[Meal Prep.cost]]),"")',
    ], ["money", "money", "money", "money"])
    row += 2
    s.note(row, "These are estimates, not accounting. 'Leftovers "
                "eaten' values the portions you did not have to cook "
                "again; 'cheaper shops available' is the ceiling if "
                "you always bought where it is cheapest; the batch "
                "cooking figure compares your prepped portions with "
                "what a takeaway portion costs you.", LAST_COL,
           height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
