"""`Financial Plan` - targets, break-even and what if you grew.

Three honest questions: how much work do you have to do to cover the
bills, are you on track against your own targets, and what happens if
you put the rate up or take on another cleaner. The white cells are
yours to play with; nothing here touches your records.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, number_check

NAME = "Financial Plan"
LAST_COL = 11
MONTHS = 12

_REV = ('SUMIFS([[Jobs.price]],[[Jobs.month]],{key},[[Jobs.status]],'
        '"Completed")+SUMIFS([[Jobs.price]],[[Jobs.month]],{key},'
        '[[Jobs.status]],"Invoiced")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Financial Planning & Forecasting",
             "What you must earn to break even, how the year is "
             "tracking and what growing would actually be worth.",
             LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 32, fmts.chrome("label"))
    s.set_col(2, 2, 10, fmts.chrome("label"))
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 14)
    s.set_col(5, LAST_COL, 14)

    r = 4

    # ----------------------------------------------------- break-even ---
    r = s.section(r, "1. Break-even: the number you must hit", LAST_COL)
    r = s.note(r, "Fixed costs are the ones that arrive whether you "
                  "clean anything or not - insurance, the van, "
                  "software, your phone. Fill those in and the "
                  "workbook tells you how many jobs a month keep the "
                  "lights on.", LAST_COL, first_col=1, height=30)
    r = s.field(r, "Fixed costs a month", 1800, style="money",
                name="FpFixed", validate=number_check(0, 1000000),
                help_text="Insurance, vehicle, software, phone, "
                          "accountant, rent, your own basic draw.")
    r = s.field(r, "Average job value", None, style="money", mode="calc",
                name="FpAvgJob",
                formula=s.reg.render(
                    '=IFERROR(AVERAGEIFS([[Jobs.price]],[[Jobs.price]],'
                    '">0",[[Jobs.status]],"Completed"),MinimumCharge)'),
                help_text="Taken from your completed jobs. If you have "
                          "none yet it falls back to the minimum "
                          "charge.")
    r = s.field(r, "Average hours a job", None, style="number1",
                mode="calc", name="FpAvgHours",
                formula=s.reg.render(
                    '=IFERROR(AVERAGEIFS([[Jobs.est_hours]],'
                    '[[Jobs.est_hours]],">0"),2)'))
    r = s.field(r, "Variable cost a job", None, style="money",
                mode="calc", name="FpVarCost",
                formula="=ROUND(FpAvgHours*(CostPerHour+"
                        "SuppliesPerHour),2)",
                help_text="Wages and supplies for an average job. "
                          "Travel sits in fixed costs unless you drive "
                          "a lot.")
    r = s.field(r, "Contribution a job", None, style="money",
                mode="calc", name="FpContribution",
                formula="=FpAvgJob-FpVarCost",
                help_text="What each job leaves towards the fixed "
                          "costs. If this is near zero, no amount of "
                          "work will save the business.")
    r = s.field(r, "Jobs a month to break even", None, style="int",
                mode="calc", name="FpBreakEvenJobs",
                formula='=IF(FpContribution<=0,"Never - the price is '
                        'too low",ROUNDUP(FpFixed/FpContribution,0))')
    r = s.field(r, "Revenue a month to break even", None, style="money0",
                mode="calc", name="FpBreakEvenRevenue",
                formula='=IF(FpContribution<=0,"",ROUNDUP(FpFixed/'
                        'FpContribution,0)*FpAvgJob)')
    r = s.field(r, "Hours a month to break even", None, style="number1",
                mode="calc", name="FpBreakEvenHours",
                formula='=IF(FpContribution<=0,"",ROUNDUP(FpFixed/'
                        'FpContribution,0)*FpAvgHours)')
    r = s.field(r, "Jobs a month you are doing now", None, style="int",
                mode="calc", name="FpJobsNow",
                formula=s.reg.render(
                    '=ROUND((COUNTIF([[Jobs.status]],"Completed")+'
                    'COUNTIF([[Jobs.status]],"Invoiced"))/12,0)'),
                help_text="A twelve month average of completed work.")
    verdict = r
    r = s.field(r, "Where that leaves you", None, mode="calc",
                value_end=7, name="FpBreakEvenVerdict",
                formula='=IF(FpContribution<=0,"Put the price up - '
                        'every job loses money",IF(FpJobsNow>='
                        'FpBreakEvenJobs,"Covered - you are "&'
                        '(FpJobsNow-FpBreakEvenJobs)&" jobs a month '
                        'past break-even","Short by "&'
                        '(FpBreakEvenJobs-FpJobsNow)&" jobs a month"))')
    ws.conditional_format(verdict, 3, verdict, 7, {
        "type": "text", "criteria": "containing", "value": "Covered",
        "format": fmts.cf("good")})
    ws.conditional_format(verdict, 3, verdict, 7, {
        "type": "text", "criteria": "containing", "value": "Short",
        "format": fmts.cf("warn")})
    ws.conditional_format(verdict, 3, verdict, 7, {
        "type": "text", "criteria": "containing", "value": "Put the",
        "format": fmts.cf("bad")})
    r += 2

    # -------------------------------------------------- the year plan ---
    r = s.section(r, "2. The year against your targets", LAST_COL)
    r = s.table_head(r, 1, ["Month", "Target", "Revenue", "Difference",
                            "Costs", "Profit", "Running profit",
                            "Hit target?"],
                     [10, 13, 13, 13, 12, 12, 14, 12])
    first_month = r
    for m in range(MONTHS):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        revenue = _REV.format(key=key)
        ws.write_formula(r, 1, f'=TEXT(DATE(PlanYear,{m + 1},1),"mmm")',
                         fmts.calc("text"))
        ws.write_formula(r, 2, "=MonthlyRevenueTarget",
                         fmts.calc("money0"))
        ws.write_formula(r, 3, s.reg.render(f"={revenue}"),
                         fmts.calc("money0"))
        ws.write_formula(r, 4, f"=D{r + 1}-C{r + 1}", fmts.calc("money0"))
        ws.write_formula(
            r, 5,
            s.reg.render(f'=SUMIFS([[Expenses.total]],'
                         f'[[Expenses.month]],{key})'),
            fmts.calc("money0"))
        ws.write_formula(r, 6, f"=D{r + 1}-F{r + 1}", fmts.calc("money0"))
        if m == 0:
            ws.write_formula(r, 7, f"=G{r + 1}", fmts.calc("money0"))
        else:
            ws.write_formula(r, 7, f"=H{r}+G{r + 1}", fmts.calc("money0"))
        ws.write_formula(
            r, 8,
            f'=IF(D{r + 1}=0,"",IF(D{r + 1}>=C{r + 1},"Yes","No"))',
            fmts.calc("status"))
        r += 1
    last_month = r - 1
    ws.conditional_format(first_month, 3, last_month, 3,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    ws.conditional_format(first_month, 4, last_month, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_month, 6, last_month, 6,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    for text, fmt in [("Yes", "good"), ("No", "warn")]:
        ws.conditional_format(first_month, 8, last_month, 8, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(fmt)})
    # totals
    ws.write(r, 1, "Year", fmts.chrome("label_bold"))
    for col, letter in [(2, "C"), (3, "D"), (4, "E"), (5, "F"),
                        (6, "G")]:
        ws.write_formula(
            r, col,
            f"=SUM({letter}{first_month + 1}:{letter}{last_month + 1})",
            fmts.chrome("total_value", num_format=cfg.money_format_0))
    ws.write_formula(r, 7, f"=H{last_month + 1}",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format_0))
    ws.write_formula(
        r, 8,
        f'=IF(D{r + 1}=0,"",TEXT(G{r + 1}/D{r + 1},"0%")&" margin")',
        fmts.chrome("total_value"))
    r += 2

    # ------------------------------------------------- what if ----------
    r = s.section(r, "3. What if the business grew?", LAST_COL)
    r = s.note(r, "Change the three white cells and watch the table "
                  "below. It is the quickest way to see whether you "
                  "need more clients, better prices or another pair "
                  "of hands.", LAST_COL, first_col=1, height=26)
    r = s.field(r, "Put the hourly rate up by", 0.05, style="percent0",
                name="FpRateRise", validate=number_check(-0.5, 2))
    r = s.field(r, "Extra regular clients a month", 2, style="int",
                name="FpNewClients", validate=number_check(0, 500,
                                                           decimal=False))
    r = s.field(r, "Extra cleaners", 1, style="int", name="FpNewStaff",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "Jobs a new client books a month", 4, style="number1",
                name="FpClientJobs", validate=number_check(0, 60))
    r += 1

    r = s.table_head(r, 1, ["Scenario", "Revenue a month",
                            "Costs a month", "Profit a month",
                            "Profit a year", "Change"],
                     [30, 15, 15, 15, 15, 14])
    base_rev = "FpJobsNow*FpAvgJob"
    base_cost = "FpFixed+FpJobsNow*FpVarCost"
    scenarios = [
        ("As you are now", base_rev, base_cost),
        ("With the rate rise",
         f"({base_rev})*(1+FpRateRise)", base_cost),
        ("With the extra clients",
         f"({base_rev})+FpNewClients*FpClientJobs*FpAvgJob",
         f"({base_cost})+FpNewClients*FpClientJobs*FpVarCost"),
        ("With another cleaner",
         f"({base_rev})+FpNewStaff*ProductiveHours*21*HourlyRate*"
         f"UtilisationTarget",
         f"({base_cost})+FpNewStaff*ProductiveHours*21*CostPerHour"),
        ("All three together",
         f"(({base_rev})+FpNewClients*FpClientJobs*FpAvgJob+"
         f"FpNewStaff*ProductiveHours*21*HourlyRate*UtilisationTarget)"
         f"*(1+FpRateRise)",
         f"({base_cost})+FpNewClients*FpClientJobs*FpVarCost+"
         f"FpNewStaff*ProductiveHours*21*CostPerHour"),
    ]
    first_s = r
    for label, rev, cost in scenarios:
        ws.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, f"={rev}", fmts.calc("money0"))
        ws.write_formula(r, 3, f"={cost}", fmts.calc("money0"))
        ws.write_formula(r, 4, f"=C{r + 1}-D{r + 1}", fmts.calc("money0"))
        ws.write_formula(r, 5, f"=E{r + 1}*12", fmts.calc("money0"))
        if r == first_s:
            ws.write(r, 6, "baseline", fmts.chrome("note"))
        else:
            ws.write_formula(
                r, 6, f"=E{r + 1}-E${first_s + 1}", fmts.calc("money0"))
        r += 1
    last_s = r - 1
    ws.conditional_format(first_s, 4, last_s, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_s + 1, 6, last_s, 6,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    r += 1
    s.merge(r, 1, LAST_COL,
            "The cleaner scenario assumes 21 working days a month at "
            "your utilisation target - be careful, a new cleaner "
            "costs money from day one and earns it from week three.",
            fmts.chrome("note"))
    r += 2

    # ------------------------------------------------- cash position ----
    r = s.section(r, "4. Cash position today", LAST_COL)
    r = s.table_head(r, 1, ["Measure", "Amount", "What it means"],
                     [28, 15, 54])
    cash = [
        ("Invoiced, not yet paid", '=SUM([[Invoices.outstanding]])',
         '="Money you have earned and not received"'),
        ("Overdue right now",
         '=SUMIFS([[Invoices.outstanding]],[[Invoices.age]],'
         '"1-30 days")+SUMIFS([[Invoices.outstanding]],'
         '[[Invoices.age]],"31-60 days")+SUMIFS('
         '[[Invoices.outstanding]],[[Invoices.age]],"61-90 days")+'
         'SUMIFS([[Invoices.outstanding]],[[Invoices.age]],"90+ days")',
         '="Past the due date. Chase today"'),
        ("Work done, not invoiced",
         '=SUMIFS([[Jobs.price]],[[Jobs.status]],"Completed",'
         '[[Jobs.invoice]],"")',
         '="Raise these invoices and the cash follows"'),
        ("Deposits taken",
         '=SUMIFS([[Payments.amount]],[[Payments.deposit]],"Yes")',
         '="Money in advance - lovely, but it is not profit yet"'),
        ("Booked work, not yet done",
         '=SUMIFS([[Jobs.price]],[[Jobs.date]],">"&TodayDate,'
         '[[Jobs.status]],"<>Cancelled")',
         '="The diary, in money"'),
        ("Quotes awaiting an answer",
         '=SUMIFS([[Quotes.total]],[[Quotes.status]],"Sent")+'
         'SUMIFS([[Quotes.total]],[[Quotes.status]],"Followed up")',
         '="Possible, not certain"'),
        ("Expenses this month",
         '=SUMIFS([[Expenses.total]],[[Expenses.month]],DashKey)',
         '="Against a budget of "&TEXT(MonthlyExpenseBudget,"#,##0")'),
        ("Stock on the shelf",
         '=SUM([[Supplies.stock_value]])',
         '="Cash you have already spent, sitting in the cupboard"'),
    ]
    first_cash = r
    for label, formula, meaning in cash:
        ws.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc("money0"))
        s.merge(r, 3, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(r, 3, s.reg.render(meaning), fmts.calc("text"))
        r += 1
    ws.conditional_format(first_cash + 1, 2, first_cash + 1, 2,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_cash + 2, 2, first_cash + 2, 2,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    r += 1

    # ------------------------------------------------- the tax reminder -
    r = s.section(r, "5. Money you should not spend", LAST_COL)
    r = s.field(r, "Set aside for tax", 0.25, style="percent0",
                name="FpTaxSet", validate=number_check(0, 1),
                help_text="A guess, not advice. Ask an accountant "
                          "what your rate really is.")
    r = s.field(r, "Profit so far this year", None, style="money0",
                mode="calc", name="FpYearProfit",
                formula=s.reg.render(
                    '=SUMIFS([[Jobs.price]],[[Jobs.status]],"Completed",'
                    '[[Jobs.date]],">="&DATE(PlanYear,1,1))+SUMIFS('
                    '[[Jobs.price]],[[Jobs.status]],"Invoiced",'
                    '[[Jobs.date]],">="&DATE(PlanYear,1,1))-SUMIFS('
                    '[[Expenses.total]],[[Expenses.date]],">="&'
                    'DATE(PlanYear,1,1))'))
    r = s.field(r, "Put this much aside", None, style="money0",
                mode="calc", name="FpTaxPot",
                formula="=MAX(0,ROUND(FpYearProfit*FpTaxSet,0))",
                help_text="Move it to a separate account the day the "
                          "money lands. It is the single habit that "
                          "keeps small businesses out of trouble.")
    r += 1

    s.note(r,
           "This page is a planning tool, not an accounts package. It "
           "does not know about loans, drawings, depreciation or your "
           "local tax rules, and it is not financial advice. Use it "
           "to decide what to charge and when to hire - then show the "
           "real numbers to a real accountant.",
           LAST_COL, first_col=1, height=40)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
