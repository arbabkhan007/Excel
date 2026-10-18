"""`Mileage & Travel` - the miles between the jobs.

Cleaning is a driving business. Logged properly, travel is both a real
cost to put against a job and, in most countries, a deduction. Logged
nowhere, it is the quietest hole in a cleaning company's margin.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, month_col

NAME = "Mileage"


def spec(cfg: Config) -> SheetSpec:
    unit = cfg.distance_unit
    columns = [
        id_col("date", '"MIL-"', digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("employee", "Driver", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("vehicle", "Vehicle", 13, "text",
            validate=list_from("VehicleList")),
        Col("purpose", "Purpose", 18, "text",
            validate=list_from("MileagePurposeList")),
        Col("from_place", "From", 22, "text"),
        Col("to_place", "To", 22, "text"),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Link it to a job and the travel cost lands on that "
                 "job's profit."),
        Col("distance", f"Distance ({unit})", 11, "number1",
            validate=number_check(0, 10000)),
        Col("return_trip", "Return trip?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Yes doubles the distance. Most job trips are "
                 "returns."),
        Col("rate_input", "Rate (override)", 10, "money",
            validate=number_check(0, 100),
            note=f"Blank uses the rate per {unit[:-1]} from the "
                 f"Settings tab."),
        Col("parking", "Parking & tolls", 10, "money",
            validate=number_check(0, 10000)),
        Col("claimed", "Claimed?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("client_name", "Client", 18, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))',
            hidden=True),
        Col("total_distance", f"Total {unit}", 10, "number1",
            formula=blank("date",
                          'IF({distance}{r}="",0,{distance}{r})*'
                          'IF({return_trip}{r}="Yes",2,1)'),
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("rate_used", "Rate", 9, "money",
            formula=blank("date",
                          'IF({rate_input}{r}<>"",{rate_input}{r},'
                          'MileageRate)'),
            hidden=True),
        Col("amount", "Travel cost", 11, "money",
            formula=blank("date",
                          'ROUND({total_distance}{r}*{rate_used}{r}+'
                          'IF({parking}{r}="",0,{parking}{r}),2)'),
            note="Distance times the rate, plus parking. This is "
                 "what you claim and what the job carries."),
        Col("alert", "Check", 22, "status",
            formula=blank(
                "date",
                'IF({distance}{r}="","Add - how far?",'
                'IF({employee}{r}="","Add - who drove?",'
                'IF({claimed}{r}="No","Due - not claimed yet",'
                'IF({total_distance}{r}>200,"Check - a long day",'
                '"OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Mileage & Travel",
        subtitle="Every trip, what it cost and which job should "
                 "carry it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).mileage,
        help_text="Log the trip on the day - a week later nobody "
                  "remembers. Link it to a job and the travel cost "
                  "appears in that job's profit.",
        extra_notes=[
            "Travel between two jobs is not billable time, but it "
            "is paid time. Twenty minutes between every visit is "
            "most of an hour a day you are paying for and nobody is "
            "buying.",
            "The rate per mile is on the Settings tab. Check what "
            "your tax authority allows - it changes, and it is "
            "rarely what you guess.",
            "A mileage log is only evidence if it is contemporary. "
            "From, to, purpose and distance, written the same week.",
        ],
    )
