"""`Analytics` - the patterns you cannot see from behind the desk.

Which categories earn, which days pay, which hours are dead, where
the clients come from, and whether the ones you win stay. Every
figure is a formula over the registers, so it moves the moment the
diary does.
"""

from __future__ import annotations

from ..config import (MARKETING_CHANNELS, SERVICE_CATEGORIES, Config,
                      Palette)
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Analytics"
LAST_COL = 10
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
        "Saturday", "Sunday"]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws

    s.banner("Business Analytics",
             "Categories, days, hours, sources and retention - the "
             "patterns behind the takings.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 28)
    s.set_col(2, LAST_COL, 14)

    row = 4
    row = s.note(
        row,
        "Nothing on this tab is typed. It is the same data as the rest "
        "of the workbook, cut four different ways, because the answer "
        "to 'why was last month quiet?' is usually in one of these "
        "tables.",
        LAST_COL, height=28)
    row += 1

    # ---------------------------------------------- service categories --
    row = s.section(row, "Where the money comes from, by category",
                    LAST_COL)
    s.table_head(row, 1, [
        "Category", "Services on the menu", "Booked", "Revenue",
        "Share of revenue", "Average price", "Chair hours",
        "Revenue an hour"],
        [26, 16, 11, 14, 14, 13, 13, 14])
    row += 1
    cat_first = row
    for category in SERVICE_CATEGORIES:
        s.write(row, 1, category, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Services.category]],"{category}")', "int"),
            (3, f'=SUMIFS([[Services.done]],[[Services.category]],'
                f'"{category}")', "int"),
            (4, f'=ROUND(SUMIFS([[Services.revenue]],'
                f'[[Services.category]],"{category}"),2)', "money"),
            (5, f'=IFERROR(SUMIFS([[Services.revenue]],'
                f'[[Services.category]],"{category}")/'
                f'MAX(0.01,ServiceRevenueAll),"")', "percent"),
            (6, f'=IFERROR(ROUND(SUMIFS([[Services.revenue]],'
                f'[[Services.category]],"{category}")/'
                f'MAX(1,SUMIFS([[Services.units]],'
                f'[[Services.category]],"{category}")),2),"")', "money"),
            (7, f'=ROUND(SUMIFS([[Services.hours_total]],'
                f'[[Services.category]],"{category}"),1)', "number1"),
            (8, f'=IFERROR(ROUND(SUMIFS([[Services.revenue]],'
                f'[[Services.category]],"{category}")/'
                f'MAX(0.1,SUMIFS([[Services.hours_total]],'
                f'[[Services.category]],"{category}")),2),"")', "money"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    cat_last = row - 1
    ws.conditional_format(cat_first, 4, cat_last, 4,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(cat_first, 8, cat_last, 8,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetChairRate",
                           "format": fmts.cf("good")})
    ws.conditional_format(cat_first, 8, cat_last, 8,
                          {"type": "cell", "criteria": "<",
                           "value": "=TargetChairRate",
                           "format": fmts.cf("warn")})
    row += 1

    # ----------------------------------------------------- by weekday ---
    row = s.section(row, "Which days pay", LAST_COL)
    s.table_head(row, 1, [
        "Day", "Appointments", "Completed", "Cancelled", "No shows",
        "Revenue", "Average ticket", "Chair hours", "Revenue an hour"],
        [16, 14, 12, 12, 11, 14, 14, 13, 14])
    row += 1
    day_first = row
    for index, day in enumerate(DAYS):
        short = day[:3]
        s.write(row, 1, day, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Appointments.weekday]],"{short}")', "int"),
            (3, f'=COUNTIFS([[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")', "int"),
            (4, f'=COUNTIFS([[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Cancelled")', "int"),
            (5, f'=COUNTIFS([[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"No show")', "int"),
            (6, f'=ROUND(SUMIFS([[Appointments.billed]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed"),2)', "money"),
            (7, f'=IFERROR(ROUND(SUMIFS([[Appointments.billed]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")/'
                f'MAX(1,COUNTIFS([[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")),2),"")',
             "money"),
            (8, f'=ROUND(SUMIFS([[Appointments.hours]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed"),1)', "number1"),
            (9, f'=IFERROR(ROUND(SUMIFS([[Appointments.billed]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")/'
                f'MAX(0.1,SUMIFS([[Appointments.hours]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")),2),"")',
             "money"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    day_last = row - 1
    ws.conditional_format(day_first, 6, day_last, 6,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(day_first, 5, day_last, 5,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # -------------------------------------------------------- by hour ---
    row = s.section(row, "Which hours are dead", LAST_COL)
    row = s.note(
        row,
        "Bookings by start time, across the whole diary. A column of "
        "empty mornings is a column you can fill with an offer, a "
        "training slot, or a later opening time.",
        LAST_COL, height=26)
    row += 1
    s.table_head(row, 1, [
        "Hour", "Appointments", "Completed", "Revenue",
        "Share of bookings", "Cancelled or no show"],
        [12, 14, 12, 14, 16, 20])
    row += 1
    hour_first = row
    for hour in range(7, 21):
        label = _hour_label(hour)
        s.write(row, 1, label, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Appointments.hour]],{hour})', "int"),
            (3, f'=COUNTIFS([[Appointments.hour]],{hour},'
                f'[[Appointments.status]],"Completed")', "int"),
            (4, f'=ROUND(SUMIFS([[Appointments.billed]],'
                f'[[Appointments.hour]],{hour},'
                f'[[Appointments.status]],"Completed"),2)', "money"),
            (5, f'=IFERROR(COUNTIFS([[Appointments.hour]],{hour})/'
                f'MAX(1,COUNTIF([[Appointments.date]],"<>")),"")',
             "percent"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        ws.write_formula(
            row, 6,
            s.reg.render(
                f'=COUNTIFS([[Appointments.hour]],{hour},'
                f'[[Appointments.status]],"Cancelled")+'
                f'COUNTIFS([[Appointments.hour]],{hour},'
                f'[[Appointments.status]],"No show")'),
            fmts.calc("int"))
        row += 1
    hour_last = row - 1
    ws.conditional_format(hour_first, 2, hour_last, 2,
                          {"type": "data_bar",
                           "bar_color": Palette.accent})
    row += 1

    # ------------------------------------------------- where they come --
    row = s.section(row, "Where the clients come from", LAST_COL)
    s.table_head(row, 1, [
        "Channel", "Clients", "New in 12 months", "Bookings",
        "Revenue from them", "Average spend", "Still active"],
        [24, 12, 16, 12, 16, 14, 13])
    row += 1
    src_first = row
    for channel in MARKETING_CHANNELS[:12] + ["Walk-in", "Referral",
                                              "Repeat client"]:
        s.write(row, 1, channel, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Clients.source]],"{channel}")', "int"),
            (3, f'=COUNTIFS([[Clients.source]],"{channel}",'
                f'[[Clients.first_visit]],">="&TodayDate-365)', "int"),
            (4, f'=COUNTIFS([[Appointments.source]],"{channel}")',
             "int"),
            (5, f'=ROUND(SUMIFS([[Clients.spend]],[[Clients.source]],'
                f'"{channel}"),2)', "money"),
            (6, f'=IFERROR(ROUND(SUMIFS([[Clients.spend]],'
                f'[[Clients.source]],"{channel}")/'
                f'MAX(1,COUNTIFS([[Clients.source]],"{channel}")),2),'
                f'"")', "money"),
            (7, f'=COUNTIFS([[Clients.source]],"{channel}",'
                f'[[Clients.segment]],"Active")+'
                f'COUNTIFS([[Clients.source]],"{channel}",'
                f'[[Clients.segment]],"Regular")+'
                f'COUNTIFS([[Clients.source]],"{channel}",'
                f'[[Clients.segment]],"VIP")', "int"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    src_last = row - 1
    ws.conditional_format(src_first, 5, src_last, 5,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    row += 1

    # ------------------------------------------------------- retention --
    row = s.section(row, "Do they come back?", LAST_COL)
    s.table_head(row, 1, [
        "Month", "Visits", "New clients", "Visits by returning clients",
        "Repeat rate", "Rebooked at the desk", "Rebook rate",
        "Gone quiet since"],
        [14, 12, 13, 18, 13, 18, 13, 15])
    row += 1
    ret_first = row
    for index in range(11, -1, -1):
        key = (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
               f'"yyyy-mm")')
        s.write(row, 1, None, fmts.calc("text"))
        ws.write_formula(
            row, 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"mmm yy")', fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Appointments.month]],{key},'
                f'[[Appointments.status]],"Completed")', "int"),
            (3, f'=COUNTIFS([[Clients.first_visit]],">="&'
                f'DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
                f'[[Clients.first_visit]],"<"&'
                f'DATE(YEAR(TodayDate),MONTH(TodayDate)-{index}+1,1))',
             "int"),
            (6, f'=COUNTIFS([[Appointments.month]],{key},'
                f'[[Appointments.rebooked]],"Yes")', "int"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        ws.write_formula(row, 4, f"=MAX(0,C{row + 1}-D{row + 1})",
                         fmts.calc("int"))
        ws.write_formula(row, 5,
                         f'=IFERROR(E{row + 1}/MAX(1,C{row + 1}),"")',
                         fmts.calc("percent"))
        ws.write_formula(row, 7,
                         f'=IFERROR(G{row + 1}/MAX(1,C{row + 1}),"")',
                         fmts.calc("percent"))
        ws.write_formula(
            row, 8,
            s.reg.render(
                f'=COUNTIFS([[Clients.last_visit]],">="&'
                f'DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
                f'[[Clients.last_visit]],"<"&'
                f'DATE(YEAR(TodayDate),MONTH(TodayDate)-{index}+1,1),'
                f'[[Clients.segment]],"Inactive")'),
            fmts.calc("int"))
        row += 1
    ret_last = row - 1
    for column in (5, 7):
        ws.conditional_format(ret_first, column, ret_last, column,
                              {"type": "cell", "criteria": ">=",
                               "value": "=TargetRebook",
                               "format": fmts.cf("good")})
        ws.conditional_format(ret_first, column, ret_last, column,
                              {"type": "cell", "criteria": "<",
                               "value": "=TargetRebook",
                               "format": fmts.cf("warn")})
    ws.add_sparkline(row, 2, {
        "range": (f"{quote_sheet(NAME)}!$C${ret_first + 1}:"
                  f"$C${ret_last + 1}"),
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 3, 5, "Visits, twelve months", fmts.chrome("note"))
    row += 2

    # ---------------------------------------------- the client pyramid --
    row = s.section(row, "The client book, by value", LAST_COL)
    s.table_head(row, 1, [
        "Segment", "Clients", "Share", "Lifetime spend",
        "Average spend", "Average visits", "Average ticket"],
        [18, 12, 12, 16, 14, 14, 14])
    row += 1
    for segment in ("VIP", "Regular", "Active", "New", "Lapsing",
                    "Inactive", "Not been yet"):
        s.write(row, 1, segment, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Clients.segment]],"{segment}")', "int"),
            (3, f'=IFERROR(COUNTIFS([[Clients.segment]],"{segment}")/'
                f'MAX(1,ClientCount),"")', "percent"),
            (4, f'=ROUND(SUMIFS([[Clients.spend]],[[Clients.segment]],'
                f'"{segment}"),2)', "money"),
            (5, f'=IFERROR(ROUND(SUMIFS([[Clients.spend]],'
                f'[[Clients.segment]],"{segment}")/'
                f'MAX(1,COUNTIFS([[Clients.segment]],"{segment}")),2),'
                f'"")', "money"),
            (6, f'=IFERROR(ROUND(SUMIFS([[Clients.visits]],'
                f'[[Clients.segment]],"{segment}")/'
                f'MAX(1,COUNTIFS([[Clients.segment]],"{segment}")),1),'
                f'"")', "number1"),
            (7, f'=IFERROR(ROUND(SUMIFS([[Clients.spend]],'
                f'[[Clients.segment]],"{segment}")/'
                f'MAX(1,SUMIFS([[Clients.visits]],'
                f'[[Clients.segment]],"{segment}")),2),"")', "money"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    row += 1

    row = s.note(
        row,
        "Two tables worth acting on: the hours table tells you when to "
        "run an offer, and the category table tells you which service "
        "to run it on. Discounting the thing that already earns the "
        "most an hour is the most common mistake in the trade.",
        LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s


def _hour_label(hour: int) -> str:
    suffix = "am" if hour < 12 else "pm"
    display = hour if hour <= 12 else hour - 12
    if hour == 0:
        display = 12
    return f"{display}{suffix}"
