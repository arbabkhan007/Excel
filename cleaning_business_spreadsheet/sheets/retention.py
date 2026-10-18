"""`Client Retention` - keeping the clients you already have.

Winning a new cleaning client costs several times what it costs to
keep one. This page shows how much of your income is recurring, which
clients are quietly drifting away, what each one is worth over a year
and who to ring this week.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Client Retention"
LAST_COL = 11
LIST_ROWS = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Client Retention & Recurring Services",
             "How much of your work comes back every week, who is "
             "slipping away and what it is worth to keep them.",
             LAST_COL)

    s.set_col(0, 0, 24)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.note(row,
                 "Everything on this page reads the Clients and Jobs "
                 "tabs. Nothing here needs filling in - but the phone "
                 "calls it suggests are the most profitable half hour "
                 "in your week.", LAST_COL, height=26)
    row += 1

    # ------------------------------------------------------------ KPIs --
    active = ('COUNTIF([[Clients.status]],"Active")+'
              'COUNTIF([[Clients.status]],"Regular")')
    tiles_1 = [
        ("Active clients", f"={active}", "int", "Active or Regular"),
        ("Recurring clients",
         '=COUNTIFS([[Clients.frequency]],"<>One-off",'
         '[[Clients.frequency]],"<>",[[Clients.next_job]],"<>")',
         "int", "with a booking in the diary"),
        ("Recurring value a year",
         '=SUMIFS([[Clients.year_value]],[[Clients.frequency]],'
         '"<>One-off",[[Clients.frequency]],"<>")', "money0",
         "at today's frequencies"),
        ("Average client value a year",
         '=IFERROR(AVERAGEIFS([[Clients.year_value]],'
         '[[Clients.year_value]],">0"),"")', "money0",
         "across everyone on the books"),
    ]
    tiles_2 = [
        ("Clients at risk",
         '=COUNTIF([[Clients.retention]],"At risk")', "int",
         f"nothing for {cfg.inactive_client_days}+ days"),
        ("Slipping away",
         '=COUNTIF([[Clients.retention]],"Slipping")', "int",
         "no next visit booked"),
        ("Dormant or lost",
         '=COUNTIF([[Clients.retention]],"Dormant")+'
         'COUNTIF([[Clients.retention]],"Lost")', "int",
         "gone, for now"),
        ("Retention rate",
         f'=IFERROR(({active})/({active}+COUNTIF('
         f'[[Clients.retention]],"Dormant")+COUNTIF('
         f'[[Clients.retention]],"Lost")),"")', "percent0",
         f"target {cfg.retention_target:.0%}"),
    ]
    for tiles in (tiles_1, tiles_2):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, s.reg.render(formula), style=style,
                  span=3, note=note)
        row += 4

    # ------------------------------------------------- recurring mix ----
    row = s.section(row, "How much of the work comes back", LAST_COL)
    row = s.table_head(row, 0, ["Frequency", "Clients", "Value a year",
                                "Share of the year", "Jobs booked"],
                       [22, 11, 15, 15, 13])
    first_freq = row
    for i in range(7):
        n = i + 1
        name = f'IFERROR(INDEX(FreqNames,{n}),"")'
        cell = f"$A${row + 1}"
        ws.write_formula(row, 0, f'={name}', fmts.calc("text"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IF({cell}="","",COUNTIFS('
                         f'[[Clients.frequency]],{cell}))'),
            fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=IF({cell}="","",SUMIFS('
                         f'[[Clients.year_value]],[[Clients.frequency]],'
                         f'{cell}))'), fmts.calc("money0"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF({cell}="","",IFERROR(C{row + 1}/'
                         f'SUM([[Clients.year_value]]),""))'),
            fmts.calc("percent0"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IF({cell}="","",COUNTIFS('
                         f'[[Jobs.frequency]],{cell}))'),
            fmts.calc("int"))
        row += 1
    ws.conditional_format(first_freq, 2, row - 1, 2,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    row += 1
    s.merge(row, 0, LAST_COL,
            "A business built on weekly and fortnightly work survives "
            "a quiet January. One built on one-off deep cleans starts "
            "every month from zero.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------------- ring these ----
    row = s.section(row, "Ring these clients this week", LAST_COL)
    row = s.table_head(row, 0, ["Client", "Phone", "Standing",
                                "Last clean", "Days", "Next booked",
                                "Worth a year", "Jobs done",
                                "What to say"],
                       [22, 15, 12, 12, 8, 12, 13, 10, 30])
    first_risk = row
    for n in range(1, LIST_ROWS + 1):
        pick = f'MATCH({n},[[Clients.risk_rank]],0)'
        for col, (key, style) in enumerate([
                ("name", "text"), ("phone", "text"),
                ("retention", "status"), ("last_job", "date"),
                ("days_since", "int"), ("next_job", "date"),
                ("year_value", "money"), ("jobs_done", "int")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Clients.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 8,
            f'=IF($A${row + 1}="","",IF($F${row + 1}="",'
            f'"Nothing booked - offer them a slot",'
            f'"Check they are happy with the service"))',
            fmts.calc("text"))
        row += 1
    ws.conditional_format(first_risk, 4, row - 1, 4,
                          {"type": "cell", "criteria": ">",
                           "value": "=InactiveClientDays",
                           "format": fmts.cf("bad")})
    row += 1

    # ------------------------------------------------- most valuable ----
    row = s.section(row, "The clients worth keeping most", LAST_COL)
    row = s.table_head(row, 0, ["Client", "Type", "Frequency",
                                "Worth a year", "Profit to date",
                                "Margin", "Rating", "Complaints",
                                "Standing"],
                       [22, 15, 14, 13, 13, 10, 10, 11, 12])
    first_top = row
    for n in range(1, LIST_ROWS + 1):
        pick = f'MATCH({n},[[Clients.year_value_rank]],0)'
        for col, (key, style) in enumerate([
                ("name", "text"), ("type", "text"),
                ("frequency", "text"), ("year_value", "money"),
                ("profit", "money"), ("margin", "percent0"),
                ("rating", "number1"), ("complaints", "int"),
                ("retention", "status")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Clients.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        row += 1
    ws.conditional_format(first_top, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    ws.conditional_format(first_top, 5, row - 1, 5,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginFloor",
                           "format": fmts.cf("bad")})
    row += 1

    # ------------------------------------------------------ contracts ---
    row = s.section(row, "Contracts and renewals", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "How many", "What it means"],
                       [28, 12, 60])
    rows = [
        ("Clients on a contract",
         '=COUNTIFS([[Clients.contract]],"<>No contract",'
         '[[Clients.contract]],"<>")',
         '="Rolling or fixed term - the work you can count on"'),
        ("Renewing soon",
         '=COUNTIFS([[Clients.contract_end]],">="&TodayDate,'
         '[[Clients.contract_end]],"<="&TodayDate+RenewalWarningDays)',
         '="Within "&RenewalWarningDays&" days. Ask early and you '
         'usually keep them"'),
        ("Contracts already expired",
         '=COUNTIFS([[Clients.contract_end]],"<"&TodayDate,'
         '[[Clients.contract_end]],"<>")',
         '="Still cleaning, no agreement in place"'),
        ("Clients with no next visit",
         '=COUNTIFS([[Clients.next_job]],"",[[Clients.status]],'
         '"Active")+COUNTIFS([[Clients.next_job]],"",'
         '[[Clients.status]],"Regular")',
         '="Active clients with an empty diary - the quickest money '
         'in the workbook"'),
        ("Won back this year",
         '=COUNTIF([[Clients.retention]],"Won back")',
         '="Mark a returning client as Won back on the Clients tab"'),
    ]
    first_c = row
    for label, formula, meaning in rows:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula), fmts.calc("int"))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(meaning), fmts.calc("text"))
        row += 1
    ws.conditional_format(first_c + 1, 1, first_c + 3, 1,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    s.note(row,
           "Three numbers worth remembering: a cleaning client who "
           "stays two years is worth roughly double one who stays "
           "one; most clients who leave never complained first; and "
           "the single best predictor of leaving is an empty diary. "
           "That is exactly what the first table on this page "
           "lists.", LAST_COL, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
