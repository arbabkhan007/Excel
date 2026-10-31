"""`Turnover Checklist` - one changeover, on one sheet of paper.

Type a cleaning job ID at the top and this page becomes that job:
the property, the address, the wifi, the access code, when the next
guest arrives and how long there is to do it.  Underneath, five
groups of tick boxes - bedroom, bathroom, kitchen, living space and
a final inspection - with somewhere to sign.

Print it, hand it over, get it back.  The nine summary flags at the
bottom are the ones you transfer to the Cleaning Schedule afterwards,
which is what makes the completion percentage on that tab mean
something.
"""

from __future__ import annotations

from ..config import (CLEANING_TASKS, Config, Palette, TURNOVER_CHECKLIST)
from ..layout import Registry, col_letter
from ..writer import SheetWriter

NAME = "Turnover Checklist"
LAST_COL = 7


def _job(field: str, cast: str = '&""') -> str:
    """A field from the chosen cleaning job, blank when nothing chosen."""
    return (f'=IF(JobId="","",IFERROR(INDEX('
            f'[[Cleaning Schedule.{field}]],MATCH(JobId,'
            f'[[Cleaning Schedule.id]],0)){cast},""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Turnover Checklist",
             "One job, one page. Type the cleaning job ID and print "
             "it.", LAST_COL)
    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 10)
    s.set_col(3, 3, 26)
    s.set_col(4, 4, 10)
    s.set_col(5, 5, 24)
    s.set_col(6, LAST_COL, 16)

    row = 4
    s.merge(row, 1, 1, "Cleaning job ID", fmts.chrome("label_bold"))
    s.write(row, 2, None, fmts.chrome("form_box", locked=False))
    book.define_name("JobId", f"='{NAME}'!$C${row + 1}")
    s.merge(row, 3, LAST_COL,
            "From the Cleaning Schedule - for example CLN-0001. "
            "Everything below fills itself in.",
            fmts.chrome("note"))
    s.row_height(row, 24)
    row += 2

    # ----------------------------------------------------- the job card --
    row = s.section(row, "The job", LAST_COL)
    card = [
        ("Property", _job("property_name"), "Address",
         '=IF(JobId="","",IFERROR(INDEX('
         '[[Property Database.address]],MATCH(INDEX('
         '[[Cleaning Schedule.property]],MATCH(JobId,'
         '[[Cleaning Schedule.id]],0)),'
         '[[Property Database.id]],0))&"",""))'),
        ("Clean on", _job("clean_date", "*1"), "Cleaner",
         _job("cleaner_name")),
        ("Guest leaving", _job("guest_name"), "Checked out at",
         '=IF(JobId="","",IFERROR(INDEX('
         '[[Property Database.check_out_time]],MATCH(INDEX('
         '[[Cleaning Schedule.property]],MATCH(JobId,'
         '[[Cleaning Schedule.id]],0)),'
         '[[Property Database.id]],0))&"",""))'),
        ("Next guest arrives", _job("next_arrival", "*1"),
         "Hours to do it", _job("window_hours", "*1")),
        ("Same-day turnover", _job("same_day"), "Access code",
         '=IF(JobId="","",IFERROR(INDEX('
         '[[Property Database.access]],MATCH(INDEX('
         '[[Cleaning Schedule.property]],MATCH(JobId,'
         '[[Cleaning Schedule.id]],0)),'
         '[[Property Database.id]],0))&"",""))'),
        ("Wifi network",
         '=IF(JobId="","",IFERROR(INDEX('
         '[[Property Database.wifi_name]],MATCH(INDEX('
         '[[Cleaning Schedule.property]],MATCH(JobId,'
         '[[Cleaning Schedule.id]],0)),'
         '[[Property Database.id]],0))&"",""))',
         "Who to ring", "=HostName&\" - \"&OwnerName"),
    ]
    styles = {"Clean on": "date", "Next guest arrives": "date",
              "Hours to do it": "number1"}
    for left_label, left_formula, right_label, right_formula in card:
        s.write(row, 1, left_label, fmts.chrome("label_bold"))
        s.merge(row, 2, 2, None,
                fmts.calc(styles.get(left_label, "text")))
        s.formula(row, 2, s.reg.render(left_formula),
                         fmts.calc(styles.get(left_label, "text")))
        s.write(row, 3, right_label, fmts.chrome("label_bold"))
        s.merge(row, 4, LAST_COL, None,
                fmts.calc(styles.get(right_label, "text")))
        s.formula(row, 4, s.reg.render(right_formula),
                         fmts.calc(styles.get(right_label, "text")))
        s.row_height(row, 20)
        row += 1

    s.write(row, 1, "Warning", fmts.chrome("label_bold"))
    s.merge(row, 2, LAST_COL, None, fmts.chrome("note_box"))
    s.formula(
        row, 2,
        '=IF(JobId="","Type a cleaning job ID above.",'
        'IF(COUNTIF([[Cleaning Schedule.id]],JobId)=0,'
        '"No job with that ID - check the Cleaning Schedule.",'
        'IF(IFERROR(INDEX([[Cleaning Schedule.same_day]],MATCH(JobId,'
        '[[Cleaning Schedule.id]],0))&"","")="Yes",'
        '"SAME-DAY TURNOVER. A guest arrives today - finish and '
        'report before check-in time.",'
        '"No arrival today. Normal turnaround.")))',
        fmts.chrome("note_box"))
    s.row_height(row, 28)
    ws.conditional_format(row, 2, row, LAST_COL, {
        "type": "text", "criteria": "containing", "value": "SAME-DAY",
        "format": fmts.cf("bold_bad")})
    row += 2

    # ----------------------------------------------- the tick boxes -----
    row = s.section(row, "The checklist", LAST_COL)
    row = s.note(
        row,
        "Tick as you go. Anything you cannot finish goes in the notes "
        "at the bottom - a cleaner who writes 'kettle broken' saves "
        "the next guest's review.",
        LAST_COL, height=24)
    row += 1

    for key, group, items in TURNOVER_CHECKLIST:
        s.merge(row, 0, 3, group, fmts.chrome("subsection"))
        s.merge(row, 4, 5, "Tick", fmts.chrome("subsection"))
        s.merge(row, 6, LAST_COL, None, fmts.chrome("subsection"))
        s.formula(
            row, 6,
            s.reg.render(
                f'=IF(JobId="","",IFERROR(IF(INDEX('
                f'[[Cleaning Schedule.{key}]],MATCH(JobId,'
                f'[[Cleaning Schedule.id]],0))="Yes","Signed off",'
                f'"Outstanding"),""))'),
            fmts.chrome("subsection"))
        s.row_height(row, 20)
        row += 1
        for item in items:
            s.write(row, 0, None, fmts.chrome("form_line", locked=False))
            s.merge(row, 1, 3, item, fmts.chrome("label"))
            s.merge(row, 4, LAST_COL, None,
                    fmts.chrome("form_line", locked=False))
            s.row_height(row, 19)
            row += 1
        row += 1

    # --------------------------------------------- the nine flags -------
    row = s.section(row, "Report back", LAST_COL)
    row = s.note(
        row,
        "These nine are the ones that go onto the Cleaning Schedule. "
        "Everything above is detail for the person doing the work; "
        "these are what the workbook counts.",
        LAST_COL, height=24)
    row = s.table_head(row, 0, ["", "Task", "Done?", "Notes"],
                       [4, 30, 10, 66])
    for key, label in CLEANING_TASKS:
        s.write(row, 0, None, fmts.chrome("form_line", locked=False))
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, 2, None, fmts.calc("flag"))
        s.formula(
            row, 2,
            s.reg.render(f'=IF(JobId="","",IFERROR(INDEX('
                         f'[[Cleaning Schedule.{key}]],MATCH(JobId,'
                         f'[[Cleaning Schedule.id]],0))&"",""))'),
            fmts.calc("flag"))
        s.merge(row, 3, LAST_COL, None,
                fmts.chrome("form_line", locked=False))
        s.row_height(row, 20)
        row += 1

    s.write(row, 1, "Completed", fmts.chrome("total_label"))
    s.merge(row, 2, 2, None,
            fmts.chrome("total_value", num_format="0.0%"))
    s.formula(
        row, 2,
        s.reg.render('=IF(JobId="","",IFERROR(INDEX('
                     '[[Cleaning Schedule.completion]],MATCH(JobId,'
                     '[[Cleaning Schedule.id]],0))*1,""))'),
        fmts.chrome("total_value", num_format="0.0%"))
    s.merge(row, 3, LAST_COL,
            "Red until every one of the nine is ticked on the "
            "Cleaning Schedule.", fmts.chrome("note"))
    ws.conditional_format(row, 2, row, 2, {
        "type": "cell", "criteria": ">=", "value": 1,
        "format": fmts.cf("good")})
    ws.conditional_format(row, 2, row, 2, {
        "type": "cell", "criteria": "<", "value": 1,
        "format": fmts.cf("bad")})
    row += 2

    # ---------------------------------------------------- sign off ------
    row = s.section(row, "Sign off", LAST_COL)
    for label in ("Anything broken or missing",
                  "Anything left behind by the guest",
                  "Supplies that need restocking"):
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, None,
                fmts.chrome("form_box", locked=False))
        s.row_height(row, 26)
        row += 1
    row += 1
    for left, right in (("Started at", "Finished at"),
                        ("Cleaner's signature", "Checked by")):
        s.write(row, 1, left, fmts.chrome("label_bold"))
        s.merge(row, 2, 3, None, fmts.chrome("form_line", locked=False))
        s.write(row, 4, right, fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None,
                fmts.chrome("form_line", locked=False))
        s.row_height(row, 26)
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_sub"))
    s.formula(
        row, 0,
        '=HostName&"  -  turnover checklist  -  printed "&'
        'TEXT(TodayDate,"d mmm yyyy")', fmts.chrome("print_sub"))
    row += 1

    row = s.note(
        row,
        "Blank is fine. With no job ID typed this prints as an empty "
        "checklist you can keep a stack of in the cupboard - the "
        "property, the date and the cleaner just get written in by "
        "hand.",
        LAST_COL, height=26)

    s.finish(freeze_row=None, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
