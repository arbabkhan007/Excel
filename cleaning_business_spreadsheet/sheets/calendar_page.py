"""`Calendar` - the month at a glance, and the next jobs in order.

A real diary built from the Jobs tab: a month grid showing how busy
each day is, and a running list of the next thirty visits with the
client, the address, the cleaner and the price. Change the month in
one cell and the whole page moves.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Calendar"
LAST_COL = 10
WEEKS = 6
AGENDA = 30


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Job Calendar",
             "Every booked visit, laid out as a month and as a list. "
             "It is the Jobs tab seen sideways - change a job there "
             "and it moves here.", LAST_COL)

    s.set_col(0, 0, 16)
    for c in range(1, 8):
        s.set_col(c, c, 15)
    s.set_col(8, LAST_COL, 14)

    # ------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Month starting", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=MonthStart", fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("CalStart", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "Showing", fmts.chrome("label_right"))
    s.merge(row, 3, 4, None, fmts.calc("text"))
    ws.write_formula(row, 3, '=TEXT(CalStart,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 5, 5, "Jobs", fmts.chrome("label_right"))
    ws.write_formula(
        row, 6,
        s.reg.render('=COUNTIFS([[Jobs.date]],">="&CalStart,'
                     '[[Jobs.date]],"<"&DATE(YEAR(CalStart),'
                     'MONTH(CalStart)+1,1),[[Jobs.status]],'
                     '"<>Cancelled")'), fmts.calc("int"))
    s.merge(row, 7, 7, "Value", fmts.chrome("label_right"))
    ws.write_formula(
        row, 8,
        s.reg.render('=SUMIFS([[Jobs.price]],[[Jobs.date]],">="&CalStart,'
                     '[[Jobs.date]],"<"&DATE(YEAR(CalStart),'
                     'MONTH(CalStart)+1,1),[[Jobs.status]],'
                     '"<>Cancelled")'), fmts.calc("money0"))
    row += 1
    row = s.note(row,
                 "Type any date in the white cell - the grid always "
                 "starts on the week containing it and runs for six "
                 "weeks.", LAST_COL, height=22)
    row += 1

    # --------------------------------------------------- the month grid --
    row = s.section(row, "The month", LAST_COL)
    first_cell = ('CalStart-WEEKDAY(CalStart,IF(WeekStart="Sunday",1,2))'
                  '+1')
    for i in range(7):
        ws.write_formula(
            row, i + 1,
            f'=TEXT({first_cell}+{i},"dddd")',
            fmts.chrome("header"))
    s.write(row, 0, "Week", fmts.chrome("header_left"))
    row += 1

    grid_top = row
    for w in range(WEEKS):
        offset = w * 7
        # week label
        s.merge(row, 0, 0, None, fmts.calc("date"))
        ws.write_formula(row, 0, f'={first_cell}+{offset}',
                         fmts.chrome("label_bold", num_format="d mmm"))
        for i in range(7):
            ws.write_formula(
                row, i + 1, f'={first_cell}+{offset + i}',
                fmts.chrome("kpi_label", num_format="ddd d"))
        row += 1
        s.write(row, 0, "Jobs", fmts.chrome("label"))
        for i in range(7):
            cell = f"{chr(66 + i)}${row}"
            ws.write_formula(
                row, i + 1,
                s.reg.render(f'=IF(COUNTIFS([[Jobs.date]],{cell},'
                             f'[[Jobs.status]],"<>Cancelled")=0,"",'
                             f'COUNTIFS([[Jobs.date]],{cell},'
                             f'[[Jobs.status]],"<>Cancelled")&" jobs, "&'
                             f'TEXT(SUMIFS([[Jobs.hours_each]],'
                             f'[[Jobs.date]],{cell},[[Jobs.status]],'
                             f'"<>Cancelled"),"0.0")&"h")'),
                fmts.cell("text", "calc"))
        row += 1
        s.write(row, 0, "Value", fmts.chrome("label"))
        for i in range(7):
            cell = f"{chr(66 + i)}${row - 1}"
            ws.write_formula(
                row, i + 1,
                s.reg.render(f'=IF(SUMIFS([[Jobs.price]],[[Jobs.date]],'
                             f'{cell},[[Jobs.status]],"<>Cancelled")=0,'
                             f'"",SUMIFS([[Jobs.price]],[[Jobs.date]],'
                             f'{cell},[[Jobs.status]],"<>Cancelled"))'),
                fmts.calc("money0"))
            ws.conditional_format(row, i + 1, row, i + 1, {
                "type": "cell", "criteria": ">", "value": 0,
                "format": fmts.cf("good")})
        s.row_height(row, 16)
        row += 1
    grid_bottom = row - 1
    ws.conditional_format(grid_top, 1, grid_bottom, 7, {
        "type": "text", "criteria": "containing", "value": "jobs",
        "format": fmts.cf("info")})
    row += 1

    # ------------------------------------------------------- the agenda --
    row = s.section(row, "The next thirty visits", LAST_COL)
    row = s.table_head(row, 0, ["Date", "Day", "Start", "Client",
                                "Property", "Service", "Cleaner",
                                "Hours", "Price", "Status", "Action"],
                       [11, 8, 8, 20, 22, 20, 16, 8, 11, 13, 26])
    first_row = row
    for n in range(1, AGENDA + 1):
        pick = f'MATCH({n},[[Jobs.future_rank]],0)'
        for col, (key, style) in enumerate([
                ("date", "date"), ("weekday", "text"), ("start", "time"),
                ("client_name", "text"), ("property_name", "text"),
                ("service", "text"), ("employee", "text"),
                ("hours_each", "number1"), ("price", "money"),
                ("status", "status"), ("alert", "status")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Jobs.{key}]],{pick}),"")'),
                fmts.calc(style))
        row += 1
    last_row = row - 1
    for col, text, fmt in [(9, "Completed", "good"), (9, "Confirmed", "info"),
                           (9, "Booked", "info"), (9, "Enquiry", "warn"),
                           (10, "Assign", "warn"), (10, "Confirm", "warn"),
                           (10, "Overdue", "bad"), (10, "Late", "bad"),
                           (10, "OK", "good")]:
        ws.conditional_format(first_row, col, last_row, col, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(fmt)})
    ws.conditional_format(first_row, 8, last_row, 8,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    row += 1

    s.note(row,
           "The list is in date order, nearest first, and it only "
           "shows visits from today onwards. An empty row means you "
           "have run out of bookings - which is either a quiet week "
           "or a sales problem.",
           LAST_COL, height=30)
    row += 2

    s.note(row,
           "A note on calendar sync: a spreadsheet cannot push these "
           "visits into Google Calendar, Outlook or a phone by "
           "itself. What it can do is print - set the print area to "
           "the agenda and the week goes on the van dashboard. If you "
           "need live sync, export the Jobs tab to CSV and import it "
           "into your calendar app.",
           LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
