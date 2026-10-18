"""`Quote Calculator` - price a clean in thirty seconds, honestly.

Answer eight questions about the property and the workbook prices the
job four different ways, shows you what it will cost you to deliver,
what margin each price leaves and what the same client is worth over a
year at each frequency. Nothing here is hidden: every number is a
formula you can read.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import (SheetWriter, list_from, number_check)

NAME = "Quote Calculator"
LAST_COL = 11
EXTRAS = 4

#: how much floor area one cleaner gets through in an hour
_AREA_PER_HOUR = 'IF(AreaUnit="sq m",60,650)'


def _round(expr: str) -> str:
    """Apply the rounding rule from the Settings tab."""
    return (f'IF(PriceRounding="Nearest 10",ROUND(({expr})/10,0)*10,'
            f'IF(PriceRounding="Nearest 5",ROUND(({expr})/5,0)*5,'
            f'IF(PriceRounding="Nearest 1",ROUND({expr},0),'
            f'ROUND({expr},2))))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Pricing & Quote Calculator",
             "Fill in the white cells. The price, the cost, the margin "
             "and the yearly value all work themselves out.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 34, fmts.chrome("label"))
    s.set_col(2, 2, 10, fmts.chrome("label"))
    s.set_col(3, 3, 16)
    s.set_col(4, 4, 14)
    s.set_col(5, LAST_COL, 15, fmts.chrome("note"))

    r = 4

    # ------------------------------------------------------- 1. the job --
    r = s.section(r, "1. What are we cleaning?", LAST_COL)
    r = s.note(r, "You can start from a property you already have on "
                  "file, or just type the numbers in. Nothing here "
                  "changes any of your saved data.",
               LAST_COL, first_col=1, height=24)
    r = s.field(r, "Service", "Regular home clean", name="QcService",
                validate=list_from("ServiceNames"), value_end=4,
                help_text="From the Services tab. Its hours, price and "
                          "uplift are used below.")
    r = s.field(r, "Client type", "Residential", name="QcType",
                validate=list_from("ClientTypeList"), value_end=4)
    r = s.field(r, f"Floor area ({cfg.area_unit})", 1450, style="number",
                name="QcArea", validate=number_check(0, 1000000),
                help_text="Leave at 0 if you do not know it - the room "
                          "count will carry the estimate.")
    r = s.field(r, "Bedrooms", 3, style="int", name="QcBeds",
                validate=number_check(0, 50, decimal=False))
    r = s.field(r, "Bathrooms", 2, style="int", name="QcBaths",
                validate=number_check(0, 50, decimal=False),
                help_text="Bathrooms and kitchens are where the time "
                          "really goes.")
    r = s.field(r, "Other rooms", 4, style="int", name="QcRooms",
                validate=number_check(0, 200, decimal=False))
    r = s.field(r, "Condition", "Normal", name="QcCondition",
                validate=list_from("ConditionList"), value_end=4,
                help_text="Be honest. This multiplies the hours - the "
                          "factors are on the Lists tab.")
    r = s.field(r, "How often", "Weekly", name="QcFrequency",
                validate=list_from("FrequencyList"), value_end=4,
                help_text="Regular work takes less time per visit, so "
                          "it is priced lower.")
    r = s.field(r, "Pets at the property?", "Yes", name="QcPets",
                validate=list_from("YesNoList"))
    r = s.field(r, f"Distance one way ({cfg.distance_unit})", 6,
                style="number1", name="QcDistance",
                validate=number_check(0, 1000))
    r = s.field(r, "Add the call-out fee?", "No", name="QcCallout",
                validate=list_from("YesNoList"),
                help_text=f"Adds the {cfg.currency_symbol}"
                          f"{cfg.callout_fee:,.0f} call-out fee from "
                          f"Settings.")
    r = s.field(r, "Cleaners on the job", cfg.cleaners_per_job,
                style="int", name="QcCleaners",
                validate=number_check(1, 20, decimal=False))
    r = s.field(r, "Hours (override the estimate)", 0, style="number1",
                name="QcHoursOverride", validate=number_check(0, 200),
                help_text="Leave at 0 to let the calculator estimate. "
                          "Type a number if you have walked the job.")
    r = s.field(r, "Hourly rate (override)", 0, style="money",
                name="QcRateOverride", validate=number_check(0, 10000),
                help_text="0 uses the standard rate from Settings.")
    r = s.field(r, "Discount", 0, style="percent0", name="QcDiscount",
                validate=number_check(0, 1),
                help_text="A first-clean discount is the cheapest way "
                          "to win recurring work - but watch the margin "
                          "below.")

    # ----------------------------------------------------- 2. the extras --
    r += 1
    r = s.section(r, "2. Extras", LAST_COL)
    r = s.note(r, "Add-ons priced straight from the Services tab. "
                  "Ovens, interior windows, inside the fridge - the "
                  "jobs that turn an average clean into a good one.",
               LAST_COL, first_col=1, height=24)
    s.write(r, 1, "Extra", fmts.chrome("header_left"))
    s.write(r, 2, "Qty", fmts.chrome("header"))
    s.write(r, 3, "Each", fmts.chrome("header"))
    s.write(r, 4, "Total", fmts.chrome("header"))
    s.merge(r, 5, LAST_COL, "Hours it adds", fmts.chrome("header"))
    r += 1
    extras_top = r
    for i in range(EXTRAS):
        s.write(r, 1, None if i else "Oven clean", fmts.input("text"))
        ws.data_validation(r, 1, r, 1, list_from("ServiceNames"))
        s.write(r, 2, 1 if not i else None, fmts.input("int"))
        ws.data_validation(r, 2, r, 2, number_check(0, 1000))
        ws.write_formula(
            r, 3,
            s.reg.render(f'=IF(B{r + 1}="","",IFERROR(INDEX('
                         f'[[Services.standard_price]],MATCH(B{r + 1},'
                         f'[[Services.name]],0)),0))'),
            fmts.calc("money"))
        ws.write_formula(
            r, 4,
            f'=IF(B{r + 1}="","",IF(C{r + 1}="",0,C{r + 1})*D{r + 1})',
            fmts.calc("money"))
        ws.write_formula(
            r, 5,
            s.reg.render(f'=IF(B{r + 1}="","",IFERROR(INDEX('
                         f'[[Services.hours]],MATCH(B{r + 1},'
                         f'[[Services.name]],0)),0)*IF(C{r + 1}="",0,'
                         f'C{r + 1}))'),
            fmts.calc("number1"))
        s.merge(r, 6, LAST_COL, None, fmts.chrome("note"))
        r += 1
    extras_bottom = r - 1
    extras_sum = f"SUM($E${extras_top + 1}:$E${extras_bottom + 1})"
    extras_hours = f"SUM($F${extras_top + 1}:$F${extras_bottom + 1})"
    s.write(r, 1, "Extras total", fmts.chrome("label_bold"))
    ws.write_formula(r, 4, f"={extras_sum}",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    ws.write_formula(r, 5, f"={extras_hours}",
                     fmts.chrome("total_value", num_format="0.0"))
    book.define_name("QcExtras", f"={quote_sheet(NAME)}!$E${r + 1}")
    book.define_name("QcExtraHours", f"={quote_sheet(NAME)}!$F${r + 1}")
    r += 2

    # ------------------------------------------------- 3. how long it takes
    r = s.section(r, "3. How long it should take", LAST_COL)
    r = s.field(r, "Hours from the room count", None, style="number1",
                mode="calc", name="QcRoomHours",
                formula="=QcBeds*0.45+QcBaths*0.6+QcRooms*0.3",
                help_text="Bedrooms at 27 minutes, bathrooms at 36, "
                          "everything else at 18. Change the weights "
                          "here if your own figures differ.")
    r = s.field(r, "Hours from the floor area", None, style="number1",
                mode="calc", name="QcAreaHours",
                formula=f"=IF(QcArea=0,0,QcArea/{_AREA_PER_HOUR})",
                help_text="One cleaner covers about 650 sq ft "
                          "(60 sq m) an hour on a standard clean.")
    r = s.field(r, "Hours the service allows", None, style="number1",
                mode="calc", name="QcServiceHours",
                formula=s.reg.render(
                    '=IFERROR(INDEX([[Services.hours]],MATCH(QcService,'
                    '[[Services.name]],0)),0)'))
    r = s.field(r, "Base hours used", None, style="number1", mode="calc",
                name="QcBaseHours",
                formula=s.reg.render(
                    '=IF(QcHoursOverride>0,QcHoursOverride,'
                    'IF(AND(IFERROR(INDEX([[Services.model]],'
                    'MATCH(QcService,[[Services.name]],0)),"")'
                    '<>"Per hour",QcServiceHours>0),QcServiceHours,'
                    'MAX(QcRoomHours,QcAreaHours)))'),
                help_text="Your override if you gave one, the service "
                          "time for a fixed-price job, otherwise the "
                          "larger of the room and area estimates.")
    r = s.field(r, "Condition factor", None, style="number1", mode="calc",
                name="QcConditionFactor",
                formula='=IFERROR(INDEX(ConditionFactors,'
                        'MATCH(QcCondition,ConditionNames,0)),1)')
    r = s.field(r, "Hours for this job", None, style="number1",
                mode="calc", name="QcHours",
                formula="=ROUND(QcBaseHours*QcConditionFactor+"
                        "QcExtraHours,2)",
                help_text="Total cleaner-hours, including the extras.")
    r = s.field(r, "Time on site", None, style="number1", mode="calc",
                name="QcVisitHours",
                formula="=ROUND(QcHours/MAX(1,QcCleaners),2)",
                help_text="How long the visit lasts with your crew on "
                          "it. This is what you tell the client.")

    # ---------------------------------------------- 4. the four prices --
    r += 1
    r = s.section(r, "4. Four ways to price it", LAST_COL)
    r = s.note(r, "The same job, priced four different ways. They "
                  "rarely agree - that is the point. Pick the method "
                  "that suits the client and sanity-check it against "
                  "the others.",
               LAST_COL, first_col=1, height=26)
    r = s.field(r, "Rate used", None, style="money", mode="calc",
                name="QcRate",
                formula="=IF(QcRateOverride>0,QcRateOverride,HourlyRate)")
    r = s.field(r, "Service uplift", None, style="percent0", mode="calc",
                name="QcUplift",
                formula=s.reg.render(
                    '=IFERROR(INDEX([[Services.uplift]],MATCH(QcService,'
                    '[[Services.name]],0)),0)'),
                help_text="Deep cleans and end of tenancy work carry "
                          "an uplift - it is set on the Services tab.")
    r = s.field(r, "A. By the hour", None, style="money", mode="calc",
                name="QcPriceHour",
                formula="=ROUND(QcHours*QcRate*(1+QcUplift),2)")
    r = s.field(r, "B. By floor area", None, style="money", mode="calc",
                name="QcPriceArea",
                formula="=ROUND(QcArea*RatePerArea*(1+QcUplift),2)",
                help_text="The usual way to price offices and end of "
                          "tenancy work.")
    r = s.field(r, "C. By the room", None, style="money", mode="calc",
                name="QcPriceRoom",
                formula="=ROUND((QcBeds*RatePerBedroom+QcBaths*"
                        "RatePerBathroom+QcRooms*RatePerRoom)*"
                        "(1+QcUplift),2)")
    r = s.field(r, "D. The catalogue price", None, style="money",
                mode="calc", name="QcPriceService",
                formula=s.reg.render(
                    '=IFERROR(INDEX([[Services.standard_price]],'
                    'MATCH(QcService,[[Services.name]],0)),0)'))
    r = s.field(r, "Price for your target margin", None, style="money",
                mode="calc", name="QcPriceTarget",
                formula="=IF(TargetMargin>=1,0,ROUND(QcCost/"
                        "(1-TargetMargin),2))",
                help_text="What you would have to charge to hit the "
                          "margin on your Settings tab.")
    r = s.field(r, "Price using", "A - by the hour", name="QcMethod",
                value_end=4,
                validate={"validate": "list",
                          "source": ["A - by the hour",
                                     "B - by floor area",
                                     "C - by the room",
                                     "D - catalogue price",
                                     "E - target margin",
                                     "F - my own price"]},
                help_text="Choose the method this quote uses.")
    r = s.field(r, "My own price", 0, style="money", name="QcManual",
                validate=number_check(0, 1000000),
                help_text="Only used when the method above is F.")

    # ------------------------------------------------- 5. what it costs --
    r += 1
    r = s.section(r, "5. What the job costs you", LAST_COL)
    r = s.field(r, "Wages", None, style="money", mode="calc",
                name="QcLabour",
                formula="=ROUND(QcHours*CostPerHour,2)",
                help_text="Hours times the true cost of a cleaner "
                          "hour - pay plus your on-costs.")
    r = s.field(r, "Supplies", None, style="money", mode="calc",
                name="QcSupplies",
                formula="=ROUND(QcHours*SuppliesPerHour,2)")
    r = s.field(r, "Travel", None, style="money", mode="calc",
                name="QcTravel",
                formula="=ROUND(QcDistance*2*MileageRate,2)",
                help_text="There and back, at the mileage rate on the "
                          "Settings tab.")
    r = s.field(r, "Total cost to deliver", None, style="money",
                mode="calc", name="QcCost",
                formula="=QcLabour+QcSupplies+QcTravel")

    # ---------------------------------------------------- 6. the answer --
    r += 1
    r = s.section(r, "6. The quote", LAST_COL)
    r = s.field(r, "Chosen price", None, style="money", mode="calc",
                name="QcChosen",
                formula='=IF(LEFT(QcMethod,1)="A",QcPriceHour,'
                        'IF(LEFT(QcMethod,1)="B",QcPriceArea,'
                        'IF(LEFT(QcMethod,1)="C",QcPriceRoom,'
                        'IF(LEFT(QcMethod,1)="D",QcPriceService,'
                        'IF(LEFT(QcMethod,1)="E",QcPriceTarget,'
                        'QcManual)))))')
    r = s.field(r, "Frequency factor", None, style="number1", mode="calc",
                name="QcFreqFactor",
                formula='=IFERROR(INDEX(FreqFactors,MATCH(QcFrequency,'
                        'FreqNames,0)),1)',
                help_text="Weekly work is priced below a one-off "
                          "because it takes less time. The factors are "
                          "on the Lists tab.")
    r = s.field(r, "After the frequency discount", None, style="money",
                mode="calc", name="QcAfterFreq",
                formula="=ROUND(QcChosen*QcFreqFactor,2)")
    r = s.field(r, "Pet surcharge", None, style="money", mode="calc",
                name="QcPetFee",
                formula='=IF(QcPets="Yes",PetSurcharge,0)')
    r = s.field(r, "Call-out fee", None, style="money", mode="calc",
                name="QcCalloutFee",
                formula='=IF(QcCallout="Yes",CalloutFee,0)')
    r = s.field(r, "Extras", None, style="money", mode="calc",
                name="QcExtrasTotal", formula="=QcExtras")
    r = s.field(r, "Discount", None, style="money", mode="calc",
                name="QcDiscountValue",
                formula="=ROUND((QcAfterFreq+QcPetFee+QcCalloutFee+"
                        "QcExtras)*QcDiscount,2)")
    r = s.field(r, "Price before rounding", None, style="money",
                mode="calc", name="QcRaw",
                formula="=MAX(MinimumCharge,QcAfterFreq+QcPetFee+"
                        "QcCalloutFee+QcExtras-QcDiscountValue)",
                help_text=f"The minimum charge of "
                          f"{cfg.currency_symbol}"
                          f"{cfg.minimum_charge:,.0f} applies to every "
                          f"job.")
    r = s.field(r, "Quote this price", None, style="money", mode="calc",
                name="QcPrice", formula="=" + _round("QcRaw"),
                help_text="Rounded the way you asked on the Settings "
                          "tab. This is the number you give the "
                          "client.")
    r = s.field(r, f"{cfg.tax_label}", None, style="money", mode="calc",
                name="QcTax", formula="=ROUND(QcPrice*TaxRate,2)")
    r = s.field(r, "Total including tax", None, style="money",
                mode="calc", name="QcTotal", formula="=QcPrice+QcTax")
    r = s.field(r, "Deposit to ask for", None, style="money", mode="calc",
                name="QcDeposit",
                formula="=ROUND(QcTotal*DepositPercent,2)")

    # ---------------------------------------------------- 7. the verdict --
    r += 1
    r = s.section(r, "7. Is it worth doing?", LAST_COL)
    r = s.field(r, "Profit on the visit", None, style="money", mode="calc",
                name="QcProfit", formula="=QcPrice-QcCost")
    r = s.field(r, "Margin", None, style="percent0", mode="calc",
                name="QcMargin",
                formula='=IF(QcPrice=0,"",QcProfit/QcPrice)')
    r = s.field(r, "You earn an hour", None, style="money", mode="calc",
                name="QcPerHour",
                formula='=IF(QcHours=0,"",QcProfit/QcHours)',
                help_text="Profit per cleaner-hour. Compare it with "
                          "your other work before you say yes.")
    r = s.field(r, "Price per sq unit", None, style="money", mode="calc",
                name="QcPerArea",
                formula='=IF(QcArea=0,"",QcPrice/QcArea)')
    verdict_row = r
    r = s.field(r, "Verdict", None, mode="calc", value_end=6,
                name="QcVerdict",
                formula='=IF(QcPrice=0,"Fill in the job above",'
                        'IF(QcProfit<0,"No - this job loses money",'
                        'IF(QcMargin<MarginFloor,'
                        '"Careful - the margin is under your floor",'
                        'IF(QcMargin<TargetMargin,'
                        '"Acceptable - but under your target margin",'
                        '"Good - this one is worth having"))))')
    ws.conditional_format(verdict_row, 3, verdict_row, 6, {
        "type": "text", "criteria": "containing", "value": "Good",
        "format": fmts.cf("good")})
    ws.conditional_format(verdict_row, 3, verdict_row, 6, {
        "type": "text", "criteria": "containing", "value": "Careful",
        "format": fmts.cf("warn")})
    ws.conditional_format(verdict_row, 3, verdict_row, 6, {
        "type": "text", "criteria": "containing", "value": "No -",
        "format": fmts.cf("bad")})
    r += 1

    # ------------------------------------------- 8. the year, at each gap
    r = s.section(r, "8. The same client at every frequency", LAST_COL)
    r = s.note(r, "What this job is worth over twelve months if they "
                  "book weekly, fortnightly or monthly. It is the "
                  "argument for offering the regular discount - and "
                  "the reason a one-off deep clean is worth less than "
                  "it looks.",
               LAST_COL, first_col=1, height=30)
    r = s.table_head(r, 1, ["Frequency", "Factor", "Price a visit",
                            "Visits a year", "Value a year",
                            "Profit a year", "Margin"],
                     [22, 9, 14, 12, 14, 14, 10])
    first_freq = r
    for i in range(7):
        n = i + 1
        name = f'IFERROR(INDEX(FreqNames,{n}),"")'
        factor = f'IFERROR(INDEX(FreqFactors,{n}),"")'
        visits = f'IFERROR(INDEX(FreqPerYear,{n}),"")'
        body = _round(f"MAX(MinimumCharge,(QcChosen*{factor}+QcPetFee+"
                      f"QcCalloutFee+QcExtras)*(1-QcDiscount))")
        ws.write_formula(r, 1, f'={name}', fmts.calc("text"))
        ws.write_formula(r, 2, f'={factor}', fmts.calc("number1"))
        ws.write_formula(r, 3, f'=IF({name}="","",{body})',
                         fmts.calc("money"))
        ws.write_formula(r, 4, f'={visits}', fmts.calc("int"))
        ws.write_formula(r, 5, f'=IF(D{r + 1}="","",D{r + 1}*E{r + 1})',
                         fmts.calc("money0"))
        ws.write_formula(r, 6,
                         f'=IF(D{r + 1}="","",(D{r + 1}-QcCost)*'
                         f'E{r + 1})', fmts.calc("money0"))
        ws.write_formula(r, 7,
                         f'=IF(OR(D{r + 1}="",D{r + 1}=0),"",'
                         f'(D{r + 1}-QcCost)/D{r + 1})',
                         fmts.calc("percent0"))
        r += 1
    ws.conditional_format(first_freq, 5, r - 1, 5,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    ws.conditional_format(first_freq, 7, r - 1, 7,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginFloor",
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_freq, 7, r - 1, 7,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetMargin",
                           "format": fmts.cf("good")})
    r += 1

    s.note(r,
           "Happy with the number? Raise it on the Quotes tab, add the "
           "lines on Quote Lines, and the follow-up reminders start by "
           "themselves. Nothing you do on this page is saved - it is a "
           "calculator, not a record.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
