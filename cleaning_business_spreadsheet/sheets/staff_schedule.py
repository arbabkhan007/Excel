"""`Staff Schedule` - the week's rota, and who has room for one more.

Pick a week and every cleaner's day-by-day load appears: hours booked,
jobs, how full their week is against their contracted hours and
whether anything clashes. Underneath it, the attendance for the week
and the jobs still waiting for a name.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Staff Schedule"
LAST_COL = 12
STAFF_ROWS = 12
UNASSIGNED = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Employee Schedule & Attendance",
             "Who is working when, how full their week is and who has "
             "capacity for one more job.", LAST_COL)

    s.set_col(0, 0, 20)
    for c in range(1, 8):
        s.set_col(c, c, 11)
    for c in range(8, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Week starting", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=WeekStartDate", fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("RotaStart", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 3, None, fmts.calc("text"))
    ws.write_formula(row, 2,
                     '="Week "&TEXT(RotaStart,"d mmm")&" to "&'
                     'TEXT(RotaStart+6,"d mmm yyyy")', fmts.calc("text"))
    s.merge(row, 4, 4, "Jobs", fmts.chrome("label_right"))
    ws.write_formula(
        row, 5,
        s.reg.render('=COUNTIFS([[Jobs.date]],">="&RotaStart,'
                     '[[Jobs.date]],"<="&RotaStart+6,[[Jobs.status]],'
                     '"<>Cancelled")'), fmts.calc("int"))
    s.merge(row, 6, 6, "Hours", fmts.chrome("label_right"))
    ws.write_formula(
        row, 7,
        s.reg.render('=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],'
                     '">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6,'
                     '[[Jobs.status]],"<>Cancelled")'),
        fmts.calc("number1"))
    s.merge(row, 8, 9, "Value of the week", fmts.chrome("label_right"))
    ws.write_formula(
        row, 10,
        s.reg.render('=SUMIFS([[Jobs.price]],[[Jobs.date]],'
                     '">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6,'
                     '[[Jobs.status]],"<>Cancelled")'),
        fmts.calc("money0"))
    row += 2

    # ---------------------------------------------------------- the rota --
    row = s.section(row, "Hours booked, by cleaner", LAST_COL)
    s.write(row, 0, "Cleaner", fmts.chrome("header_left"))
    for i in range(7):
        ws.write_formula(row, i + 1, f"=RotaStart+{i}",
                         fmts.chrome("header", num_format="ddd d"))
    for i, head in enumerate(["Week hours", "Contracted", "How full",
                              "Jobs", "Check"]):
        s.write(row, 8 + i, head, fmts.chrome("header"))
    day_row = row
    row += 1

    first_staff = row
    for n in range(1, STAFF_ROWS + 1):
        pick = f'MATCH({n},[[Employees.name_rank]],0)'
        name_cell = f"$A${row + 1}"
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX([[Employees.name]],{pick}),"")'),
            fmts.chrome("label_bold"))
        for i in range(7):
            cell = f"{chr(66 + i)}${day_row + 1}"
            ws.write_formula(
                row, i + 1,
                s.reg.render(
                    f'=IF({name_cell}="","",IF(SUMIFS('
                    f'[[Jobs.hours_each]],[[Jobs.employee]],{name_cell},'
                    f'[[Jobs.date]],{cell},[[Jobs.status]],'
                    f'"<>Cancelled")=0,"",SUMIFS([[Jobs.hours_each]],'
                    f'[[Jobs.employee]],{name_cell},[[Jobs.date]],'
                    f'{cell},[[Jobs.status]],"<>Cancelled")))'),
                fmts.calc("number1"))
        # week totals
        ws.write_formula(
            row, 8,
            s.reg.render(f'=IF({name_cell}="","",SUMIFS('
                         f'[[Jobs.hours_each]],[[Jobs.employee]],'
                         f'{name_cell},[[Jobs.date]],">="&RotaStart,'
                         f'[[Jobs.date]],"<="&RotaStart+6,'
                         f'[[Jobs.status]],"<>Cancelled"))'),
            fmts.calc("number1"))
        ws.write_formula(
            row, 9,
            s.reg.render(f'=IF({name_cell}="","",IFERROR(INDEX('
                         f'[[Employees.hours]],MATCH({name_cell},'
                         f'[[Employees.name]],0)),""))'),
            fmts.calc("number1"))
        ws.write_formula(
            row, 10,
            f'=IF(OR({name_cell}="",J{row + 1}="",J{row + 1}=0),"",'
            f'I{row + 1}/J{row + 1})', fmts.calc("percent0"))
        ws.write_formula(
            row, 11,
            s.reg.render(f'=IF({name_cell}="","",COUNTIFS('
                         f'[[Jobs.employee]],{name_cell},[[Jobs.date]],'
                         f'">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6,'
                         f'[[Jobs.status]],"<>Cancelled"))'),
            fmts.calc("int"))
        ws.write_formula(
            row, 12,
            s.reg.render(
                f'=IF({name_cell}="","",IF(COUNTIFS([[Jobs.clash]],'
                f'"Clash",[[Jobs.employee]],{name_cell},[[Jobs.date]],'
                f'">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6)>0,'
                f'"Clash",IF(K{row + 1}="","",IF(K{row + 1}>1,'
                f'"Overloaded",IF(K{row + 1}<0.6,"Has room","OK")))))'),
            fmts.calc("status"))
        row += 1
    last_staff = row - 1
    ws.conditional_format(first_staff, 1, last_staff, 7,
                          {"type": "3_color_scale",
                           "min_color": "#FFFFFF",
                           "mid_color": "#DCEEF0",
                           "max_color": "#2F8F9D"})
    ws.conditional_format(first_staff, 10, last_staff, 10,
                          {"type": "cell", "criteria": ">", "value": 1,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_staff, 10, last_staff, 10,
                          {"type": "cell", "criteria": "<", "value": 0.6,
                           "format": fmts.cf("warn")})
    for text, fmt in [("Clash", "bad"), ("Overloaded", "bad"),
                      ("Has room", "info"), ("OK", "good")]:
        ws.conditional_format(first_staff, 12, last_staff, 12, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(fmt)})
    row += 1

    # --------------------------------------------------- the day totals --
    s.write(row, 0, "Team total", fmts.chrome("label_bold"))
    for i in range(7):
        cell = f"{chr(66 + i)}${day_row + 1}"
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],'
                         f'{cell},[[Jobs.status]],"<>Cancelled")'),
            fmts.chrome("total_value", num_format="0.0"))
    ws.write_formula(
        row, 8,
        s.reg.render('=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],'
                     '">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6,'
                     '[[Jobs.status]],"<>Cancelled")'),
        fmts.chrome("total_value", num_format="0.0"))
    s.merge(row, 9, LAST_COL,
            "Hours on site, not hours paid - add travel between jobs.",
            fmts.chrome("note"))
    row += 2

    # ----------------------------------------------------- capacity check
    row = s.section(row, "Can you take more work this week?", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "This week", "What it means"],
                       [26, 14, 60])
    checks = [
        ("Cleaners available",
         '=COUNTIF([[Employees.status]],"Active")',
         '="Everyone not on leave or off sick"'),
        ("Hours you could sell",
         '=COUNTIF([[Employees.status]],"Active")*ProductiveHours*5',
         '="Active cleaners times "&TEXT(ProductiveHours,"0.0")&'
         '" billable hours over five days"'),
        ("Hours booked",
         '=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],">="&RotaStart,'
         '[[Jobs.date]],"<="&RotaStart+6,[[Jobs.status]],'
         '"<>Cancelled")', '="What is in the diary"'),
        ("Spare capacity",
         '=COUNTIF([[Employees.status]],"Active")*ProductiveHours*5-'
         'SUMIFS([[Jobs.hours_each]],[[Jobs.date]],">="&RotaStart,'
         '[[Jobs.date]],"<="&RotaStart+6,[[Jobs.status]],'
         '"<>Cancelled")',
         '="Hours you could still sell this week. Negative means you '
         'need help"'),
        ("Jobs with nobody on them",
         '=COUNTIFS([[Jobs.employee]],"",[[Jobs.client]],"<>",'
         '[[Jobs.date]],">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6)',
         '="Assign these before the week starts"'),
        ("Double bookings",
         '=COUNTIFS([[Jobs.clash]],"Clash",[[Jobs.date]],'
         '">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6)',
         '="Same cleaner, same time, two places"'),
    ]
    first_check = row
    for label, formula, meaning in checks:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula),
                         fmts.calc("number1"))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(meaning), fmts.calc("text"))
        row += 1
    ws.conditional_format(first_check + 3, 1, first_check + 3, 1,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_check + 4, 1, first_check + 5, 1,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # ------------------------------------------------- unassigned jobs --
    row = s.section(row, "Jobs still waiting for a cleaner", LAST_COL)
    row = s.table_head(row, 0, ["Date", "Client", "Property", "Service",
                                "Hours", "Price", "Priority"],
                       [12, 22, 24, 20, 9, 11, 12])
    for n in range(1, UNASSIGNED + 1):
        pick = f'MATCH({n},[[Jobs.unassigned_rank]],0)'
        for col, (key, style) in enumerate([
                ("date", "date"), ("client_name", "text"),
                ("property_name", "text"), ("service", "text"),
                ("hours_each", "number1"), ("price", "money"),
                ("priority", "text")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Jobs.{key}]],{pick}),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    # ------------------------------------------------ attendance summary
    row = s.section(row, "Attendance this week", LAST_COL)
    row = s.table_head(row, 0, ["Cleaner", "Hours paid", "Hours billed",
                                "Utilisation", "Late", "Absent",
                                "Cost", "Approved?"],
                       [20, 12, 12, 12, 9, 9, 12, 11])
    first_att = row
    for n in range(1, STAFF_ROWS + 1):
        pick = f'MATCH({n},[[Employees.name_rank]],0)'
        name_cell = f"$A${row + 1}"
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX([[Employees.name]],{pick}),"")'),
            fmts.chrome("label_bold"))
        week = (f'[[Timesheets.employee]],{name_cell},'
                f'[[Timesheets.date]],">="&RotaStart,'
                f'[[Timesheets.date]],"<="&RotaStart+6')
        cells = [
            (f'=IF({name_cell}="","",SUMIFS([[Timesheets.hours]],{week}))',
             "number1"),
            (f'=IF({name_cell}="","",SUMIFS([[Timesheets.hours]],{week},'
             f'[[Timesheets.billable]],"Yes"))', "number1"),
            (f'=IF(OR({name_cell}="",B{row + 1}=0),"",C{row + 1}/'
             f'B{row + 1})', "percent0"),
            (f'=IF({name_cell}="","",COUNTIFS([[Timesheets.attendance]],'
             f'"Late",{week}))', "int"),
            (f'=IF({name_cell}="","",COUNTIFS([[Timesheets.attendance]],'
             f'"Absent - no show",{week})+COUNTIFS('
             f'[[Timesheets.attendance]],"Absent - told us",{week}))',
             "int"),
            (f'=IF({name_cell}="","",SUMIFS([[Timesheets.cost]],{week}))',
             "money"),
            (f'=IF({name_cell}="","",COUNTIFS([[Timesheets.approved]],'
             f'"Yes",{week})&" of "&COUNTIFS([[Timesheets.employee]],'
             f'{name_cell},[[Timesheets.date]],">="&RotaStart,'
             f'[[Timesheets.date]],"<="&RotaStart+6))', "text"),
        ]
        for col, (formula, style) in enumerate(cells, start=1):
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    last_att = row - 1
    ws.conditional_format(first_att, 3, last_att, 3,
                          {"type": "cell", "criteria": ">=",
                           "value": "=UtilisationTarget",
                           "format": fmts.cf("good")})
    ws.conditional_format(first_att, 3, last_att, 3,
                          {"type": "cell", "criteria": "<",
                           "value": "=UtilisationTarget",
                           "format": fmts.cf("warn")})
    ws.conditional_format(first_att, 4, last_att, 5,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    s.note(row,
           "The rota reads the Jobs tab - it does not change it. To "
           "move a job, change the cleaner or the date on the Jobs "
           "tab and this page follows. Holidays and sickness come "
           "from the Timesheets tab.",
           LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
