"""`Equipment` - vacuums, machines, vans and ladders.

What you own, what it cost, who has it and when it was last serviced.
A vacuum that dies on a Monday morning costs you a day's work, so the
next service date on this tab is worth more than it looks.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Equipment"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "EquipmentPrefix", digits=3, width=10),
        Col("name", "Equipment", 26, "text",
            note="Something you would recognise on a van check: "
                 "'Numatic Henry #2', not 'vacuum'."),
        Col("type", "Type", 18, "text",
            validate=list_from("EquipmentTypeList")),
        Col("make", "Make & model", 20, "text"),
        Col("serial", "Serial / registration", 16, "text",
            note="Worth recording. Stolen equipment without a serial "
                 "number is gone."),
        Col("ownership", "Ownership", 14, "text",
            validate=list_from("OwnershipList")),
        Col("status", "Status", 14, "status",
            validate=list_from("EquipmentStatusList"), cf=status_cf()),
        Col("assigned", "Assigned to", 16, "text",
            validate=list_from("EmployeeNames"),
            note="Who has it. The commonest cause of a 'missing' "
                 "machine is that somebody took it home."),
        Col("location", "Kept at", 16, "text",
            validate=list_from("VehicleList")),
        Col("bought", "Bought on", 11, "date", validate=date_check()),
        Col("cost", "Cost", 11, "money", validate=number_check(0, 1000000)),
        Col("life_years", "Expected life (years)", 10, "number1",
            validate=number_check(0, 50),
            note="Used to work out what it costs you a year, and "
                 "when to start saving for its replacement."),
        Col("warranty", "Warranty to", 11, "date", validate=date_check()),
        Col("last_service", "Last serviced", 11, "date",
            validate=date_check()),
        Col("interval_input", "Service every (days)", 10, "int",
            validate=number_check(0, 3650, decimal=False),
            note="Blank uses the standard interval from the Settings "
                 "tab."),
        Col("service_cost", "Spent on repairs", 11, "money",
            validate=number_check(0, 1000000),
            note="Running total of what you have spent keeping it "
                 "going. When it passes the replacement cost, stop."),
        Col("safety_test", "Safety tested on", 11, "date",
            validate=date_check(),
            note="Electrical safety testing, where your country "
                 "requires it."),
        Col("insured", "Insured?", 9, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf()),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("age_years", "Age (years)", 10, "number1",
            formula='=IF(OR({name}{r}="",{bought}{r}=""),"",'
                    'ROUND((TodayDate-{bought}{r})/365.25,1))'),
        Col("annual_cost", "Costs a year", 11, "money",
            formula='=IF(OR({name}{r}="",{cost}{r}="",'
                    '{life_years}{r}="",{life_years}{r}=0),"",'
                    'ROUND({cost}{r}/{life_years}{r}+IF('
                    '{service_cost}{r}="",0,{service_cost}{r})/'
                    'MAX(1,{age_years}{r}),2))',
            note="Purchase spread over its life, plus what you have "
                 "spent repairing it. The real cost of owning it."),
        Col("interval", "Interval used", 9, "int",
            formula=blank("name",
                          'IF({interval_input}{r}="",MaintenanceDays,'
                          '{interval_input}{r})'),
            hidden=True),
        Col("next_service", "Next service due", 12, "date",
            formula='=IF(OR({name}{r}="",{last_service}{r}=""),"",'
                    '{last_service}{r}+{interval}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=TodayDate",
                 "fmt": "bad"}]),
        Col("days_to_service", "Days to service", 10, "int",
            formula='=IF({next_service}{r}="","",'
                    '{next_service}{r}-TodayDate)',
            hidden=True),
        Col("replace_year", "Replace by", 10, "int",
            formula='=IF(OR({name}{r}="",{bought}{r}="",'
                    '{life_years}{r}=""),"",YEAR({bought}{r})+'
                    '{life_years}{r})',
            hidden=True),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Retired","Archived",'
                'IF({status}{r}="Lost or stolen","Action - report it",'
                'IF({status}{r}="In repair","Check - in for repair",'
                'IF(AND({days_to_service}{r}<>"",'
                '{days_to_service}{r}<0),"Overdue - service it",'
                'IF(AND({days_to_service}{r}<>"",'
                '{days_to_service}{r}<=14),"Due - service coming up",'
                'IF(AND({last_service}{r}="",{bought}{r}<>""),'
                '"Add - never serviced",'
                'IF(AND({age_years}{r}<>"",{life_years}{r}<>"",'
                '{age_years}{r}>{life_years}{r}),'
                '"Check - past its expected life","OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Equipment & Assets",
        subtitle="What you own, what it costs to keep and when it "
                 "next needs looking at.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).equipment,
        help_text="Add every machine worth more than a week's "
                  "takings, plus the vans. Record the service date "
                  "each time and the next one works itself out.",
        extra_notes=[
            "'Costs a year' spreads the purchase over the expected "
            "life and adds the repairs. It is the number to compare "
            "with the price of a new one.",
            "A machine that has cost more in repairs than it did to "
            "buy is telling you something. Replace it before it "
            "breaks down in a client's hallway.",
            "Electrical safety testing is a legal requirement in "
            "some countries and not in others. Check what applies "
            "where you are - the column is here either way.",
        ],
    )
