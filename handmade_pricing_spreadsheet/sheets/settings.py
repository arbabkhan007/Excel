"""
The `Settings` tab - shop profile, labour rates, fees, margins, targets.

Everything the rest of the workbook needs to know about your craft
business is entered once here and published as a defined name
(ShopName, HourlyRate, TargetRetailMargin, TransactionFee ...) so the
formulas stay readable and there is exactly one place to change a
number. Nothing in this workbook hard codes a fee or a rate.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

SHEET = "Settings"
LAST_COL = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, SHEET, tab_color=Palette.accent)
    b = cfg.shop

    s.banner("Business Settings",
             "Fill in the white cells once. Shaded cells are locked "
             "formulas. Every price, every margin and every warning in "
             "the workbook reads these values.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 38, fmts.chrome("label"))
    s.set_col(2, 2, 14, fmts.chrome("label"))
    s.set_col(3, 3, 24)
    s.set_col(4, 4, 18)
    s.set_col(5, 8, 17, fmts.chrome("note"))

    r = 4  # 0 based row index -> Excel row 5

    # ------------------------------------------------------------ 1. the shop
    r = s.section(r, "1. Your shop", LAST_COL)
    r = s.field(r, "Business name", b.name, name="ShopName", value_end=4,
                help_text="Shown on every quote, order form and "
                          "printable page.")
    r = s.field(r, "Maker / owner", b.owner, name="ShopOwner", value_end=4)
    r = s.field(r, "Brand name", b.brand, name="ShopBrand", value_end=4,
                help_text="If you trade under a shorter name than the "
                          "business one.")
    r = s.field(r, "Email", b.email, name="ShopEmail", value_end=4)
    r = s.field(r, "Phone", b.phone, name="ShopPhone", value_end=4)
    r = s.field(r, "Address", b.address, name="ShopAddress", value_end=4)
    r = s.field(r, "Website", b.website, name="ShopWebsite", value_end=4)
    r = s.field(r, "Marketplace shop name", b.etsy_shop, name="ShopEtsy",
                value_end=4,
                help_text="Your Etsy / Amazon Handmade / Folksy shop "
                          "name, for the order forms.")
    r = s.field(r, "Social handle", b.instagram, name="ShopSocial",
                value_end=4)
    r = s.field(r, "Tagline", b.tagline, name="ShopTagline", value_end=4)
    r = s.field(r, "Business / tax number", b.tax_id, name="ShopTaxId",
                value_end=4)

    # ------------------------------------------------------- 2. the calendar
    r += 1
    r = s.section(r, "2. Dates", LAST_COL)
    r = s.field(r, "Year", cfg.year, style="int", name="PlanYear",
                validate=number_check(2000, 2100, decimal=False),
                help_text="The year the reports and the sales summaries "
                          "open on.")
    r = s.field(r, "Today", None, style="date", mode="calc",
                name="TodayDate", formula="=TODAY()",
                help_text="Every 'due', 'stale price' and 'review the "
                          "fees' flag counts from this cell, so they all "
                          "move on by themselves overnight.")
    r = s.field(r, "This month starts", None, style="date", mode="calc",
                name="MonthStart",
                formula="=DATE(YEAR(TodayDate),MONTH(TodayDate),1)")
    r = s.field(r, "This month ends", None, style="date", mode="calc",
                name="MonthEnd",
                formula="=DATE(YEAR(TodayDate),MONTH(TodayDate)+1,1)-1")
    r = s.field(r, "This year starts", None, style="date", mode="calc",
                name="YearStart", formula="=DATE(PlanYear,1,1)")

    # --------------------------------------------------- 3. your time is cost
    r += 1
    r = s.section(r, "3. What your time is worth", LAST_COL)
    r = s.note(r, "The single number most handmade sellers get wrong. If "
                  "your products do not pay you for the hours you spend "
                  "making them, you do not have a business, you have an "
                  "expensive hobby. Pay yourself a real wage here and let "
                  "the prices follow.",
               LAST_COL, first_col=1, height=38)
    r = s.field(r, "What you pay yourself an hour", cfg.hourly_rate,
                style="money", name="HourlyRate",
                validate=number_check(0, 1000),
                help_text="What one hour of your making time costs. Not "
                          "what is left over - what you are worth.")
    r = s.field(r, "Local minimum wage (for the warning)", cfg.minimum_wage,
                style="money", name="MinimumWage",
                validate=number_check(0, 1000),
                help_text="The workbook flags any product whose price "
                          "pays you less than this an hour.")
    r = s.field(r, "Employee hourly rate", cfg.employee_rate, style="money",
                name="EmployeeRate", validate=number_check(0, 1000))
    r = s.field(r, "Contractor / outsourced hourly rate",
                cfg.contractor_rate, style="money", name="ContractorRate",
                validate=number_check(0, 1000),
                help_text="Used when a labour line says somebody else "
                          "made it.")
    r = s.field(r, "On-costs on paid labour", cfg.payroll_burden,
                style="percent", name="PayrollBurden",
                validate=number_check(0, 1),
                help_text="Employer tax, insurance and holiday pay. "
                          "Added to employee and contractor hours.")
    r = s.field(r, "Your hour, with on-costs", None, style="money",
                mode="calc", name="LaborCostHour",
                formula="=ROUND(HourlyRate*(1+PayrollBurden),2)",
                help_text="This is the rate the labour calculator "
                          "multiplies your minutes by.")
    r = s.field(r, "Hours a month you actually make things",
                cfg.productive_hours, style="number1",
                name="ProductiveHours", validate=number_check(1, 744),
                help_text="Bench hours only - not admin, not photos, not "
                          "scrolling. Overhead per hour is divided by "
                          "this.")
    r = s.field(r, "Items a month you usually finish", cfg.units_per_month,
                style="int", name="UnitsPerMonth",
                validate=number_check(1, 100000, decimal=False),
                help_text="Used when overhead is spread per item instead "
                          "of per hour.")
    r = s.field(r, "Default batch size", cfg.default_batch_size,
                style="int", name="DefaultBatchSize",
                validate=number_check(1, 10000, decimal=False))
    r = s.field(r, "Consumables an hour", cfg.supplies_per_hour,
                style="money", name="SuppliesPerHour",
                validate=number_check(0, 1000),
                help_text="Glue, needles, sandpaper, thread - the small "
                          "things nobody itemises. Added per labour hour.")

    # ---------------------------------------------------- 4. materials, waste
    r += 1
    r = s.section(r, "4. Materials and waste", LAST_COL)
    r = s.field(r, "Measurement system", cfg.measure_system,
                name="MeasureSystem",
                validate={"validate": "list",
                          "source": ["Metric", "Imperial"]},
                help_text="Metric weighs in grams, imperial in ounces. "
                          "The unit table on Lists converts either way.")
    r = s.field(r, "Default waste allowance", cfg.default_waste,
                style="percent", name="DefaultWaste",
                validate=number_check(0, 1),
                help_text="Offcuts, frogged rows, the last inch of a "
                          "board. Added to every material line that does "
                          "not set its own.")
    r = s.field(r, "Default defect / seconds rate", cfg.default_defect_rate,
                style="percent", name="DefaultDefect",
                validate=number_check(0, 1),
                help_text="The share of a batch that never makes it to "
                          "the shop. The batch tab spreads its cost over "
                          "the pieces that do.")
    r = s.field(r, "Days of stock that counts as low", cfg.low_stock_days,
                style="int", name="LowStockDays",
                validate=number_check(0, 365, decimal=False))

    # ------------------------------------------------ 5. packaging, shipping
    r += 1
    r = s.section(r, "5. Packaging and shipping", LAST_COL)
    r = s.field(r, "Fallback packaging cost", cfg.default_packaging,
                style="money", name="DefaultPackaging",
                validate=number_check(0, 1000),
                help_text="Used for a product that has not picked a "
                          "packaging profile on the Packaging tab.")
    r = s.field(r, "Shipping you absorb", cfg.default_shipping,
                style="money", name="DefaultShipping",
                validate=number_check(0, 1000),
                help_text="Only what you pay and do not charge on. "
                          "Postage the buyer pays is not your cost - the "
                          "workbook keeps the two apart.")
    r = s.field(r, "Free shipping over", cfg.minimum_order_value / 3.0,
                style="money", name="FreeShipThreshold",
                validate=number_check(0, 100000),
                help_text="Orders above this are shipped free, so the "
                          "postage becomes your cost.")

    # ------------------------------------------------------------ 6. overhead
    r += 1
    r = s.section(r, "6. Overhead", LAST_COL)
    r = s.note(r, "Rent, insurance, the shop subscription, the website, "
                  "the accountant. Type the bills on the Overhead tab - "
                  "this is only how they get spread across your "
                  "products.",
               LAST_COL, first_col=1, height=30)
    r = s.field(r, "Spread overhead", "Per hour worked",
                name="OverheadMethod", validate=list_from("AllocationList"),
                help_text="Per hour is the fairest for slow, careful "
                          "work. Per item suits a maker who turns out "
                          "many small, similar pieces.")
    r = s.field(r, "Overhead a month (from the Overhead tab)", None,
                style="money", mode="calc", name="OverheadMonthly",
                formula="=ROUND(SUMPRODUCT(IFERROR("
                        "[[Overhead.monthly]]*1,0)),2)")
    r = s.field(r, "Overhead per hour worked", None, style="money",
                mode="calc", name="OverheadPerHour",
                formula="=IF(ProductiveHours<=0,0,"
                        "ROUND(OverheadMonthly/ProductiveHours,2))")
    r = s.field(r, "Overhead per item made", None, style="money",
                mode="calc", name="OverheadPerUnit",
                formula="=IF(UnitsPerMonth<=0,0,"
                        "ROUND(OverheadMonthly/UnitsPerMonth,2))")

    # ------------------------------------------------- 7. marketplace fees
    r += 1
    r = s.section(r, "7. Marketplace fees", LAST_COL)
    r = s.note(r, "IMPORTANT: marketplace fees change, and they differ by "
                  "country. The figures below are the defaults this "
                  "workbook shipped with - verify them against the "
                  "current official fee schedule for your shop and your "
                  "region before you trust a price. Every one of them is "
                  "editable here, and every channel can override them on "
                  "the Channels tab.",
               LAST_COL, first_col=1, height=52)
    r = s.field(r, "Listing fee (per listing)", cfg.listing_fee,
                style="money", name="ListingFee",
                validate=number_check(0, 100),
                help_text="Charged when you list and again when the "
                          "listing renews.")
    r = s.field(r, "Transaction fee", cfg.transaction_fee, style="percent",
                name="TransactionFee", validate=number_check(0, 1),
                help_text="A percentage of what the buyer pays.")
    r = s.field(r, "Payment processing %", cfg.processing_percent,
                style="percent", name="ProcessingPercent",
                validate=number_check(0, 1))
    r = s.field(r, "Payment processing fixed", cfg.processing_fixed,
                style="money", name="ProcessingFixed",
                validate=number_check(0, 100),
                help_text="The flat part - this is what quietly eats "
                          "small orders alive.")
    r = s.field(r, "Offsite ads fee", cfg.offsite_ads_fee, style="percent",
                name="OffsiteAdsFee", validate=number_check(0, 1),
                help_text="Only charged on orders that came from the "
                          "marketplace's own advertising.")
    r = s.field(r, "Share of orders from offsite ads",
                cfg.offsite_ads_share, style="percent",
                name="OffsiteAdsShare", validate=number_check(0, 1),
                help_text="Used to work out the average ads cost per "
                          "order. Check your own statements.")
    r = s.field(r, "Regulatory operating fee", cfg.regulatory_fee,
                style="percent", name="RegulatoryFee",
                validate=number_check(0, 1),
                help_text="Charged in some countries only. Set it to 0 "
                          "if your shop is not in one.")
    r = s.field(r, "Currency conversion fee", cfg.currency_fee,
                style="percent", name="CurrencyFee",
                validate=number_check(0, 1),
                help_text="Applies when you are paid in a currency other "
                          "than your bank's.")
    r = s.field(r, "Fees charged on shipping too?", cfg.fees_include_shipping,
                name="FeesOnShipping", validate=list_from("YesNoList"),
                help_text="On most marketplaces the transaction fee "
                          "applies to the postage the buyer paid as well.")
    r = s.field(r, "Fees charged on sales tax?", cfg.fees_include_tax,
                name="FeesOnTax", validate=list_from("YesNoList"))
    r = s.field(r, "Fee schedule last reviewed", cfg.today, style="date",
                name="FeesReviewed",
                help_text="Put today's date here each time you check the "
                          "official fee page. The dashboard nags you "
                          "when it goes stale.")
    r = s.field(r, "Review the fee schedule every (days)",
                cfg.fee_review_days, style="int", name="FeeReviewDays",
                validate=number_check(0, 3650, decimal=False))
    r = s.field(r, "Fees are stale", None, mode="calc", name="FeesStale",
                formula='=IF(FeesReviewed="","Never reviewed",'
                        'IF(TodayDate-FeesReviewed>FeeReviewDays,'
                        '"Check the fee schedule","OK"))')
    r = s.field(r, "Default sales channel", "Etsy", name="DefaultChannel",
                validate=list_from("ChannelList"),
                help_text="The channel a new product is priced for "
                          "until you pick another.")

    # ------------------------------------------------------ 8. pricing rules
    r += 1
    r = s.section(r, "8. Pricing rules", LAST_COL)
    r = s.field(r, "Target retail margin", cfg.target_retail_margin,
                style="percent", name="TargetRetailMargin",
                validate=number_check(0, 0.95),
                help_text="Margin, not markup. 55% margin means 55 cents "
                          "of every dollar you take is left after costs "
                          "and fees.")
    r = s.field(r, "Target wholesale margin", cfg.target_wholesale_margin,
                style="percent", name="TargetWholesaleMargin",
                validate=number_check(0, 0.95),
                help_text="Lower than retail, because the shop does the "
                          "selling. It must still be positive.")
    r = s.field(r, "Default markup on cost", cfg.default_markup,
                style="percent", name="DefaultMarkup",
                validate=number_check(0, 20),
                help_text="Markup is a percentage OF YOUR COST. 100% "
                          "markup is a 50% margin - the workbook shows "
                          "both so you never confuse them again.")
    r = s.field(r, "Wholesale is this much off retail",
                cfg.wholesale_discount, style="percent",
                name="WholesaleDiscount", validate=number_check(0, 0.9),
                help_text="The traditional handmade rule is half the "
                          "retail price. Only sustainable if the retail "
                          "price was built properly in the first place.")
    r = s.field(r, "Warn below this margin", cfg.margin_warning,
                style="percent", name="MarginWarning",
                validate=number_check(0, 0.95),
                help_text="Any product under this turns amber on the "
                          "catalog and appears on the dashboard.")
    r = s.field(r, "Round prices to", cfg.rounding_rule,
                name="RoundingRule", validate=list_from("RoundingList"),
                help_text="Charm pricing always rounds UP, so a rounded "
                          "price can never fall below the price the "
                          "maths asked for.")
    r = s.field(r, "Usual discount", cfg.default_discount, style="percent",
                name="DefaultDiscount", validate=number_check(0, 1),
                help_text="What a sale usually knocks off. The promotion "
                          "tab shows what each one costs you.")
    r = s.field(r, "Sales tax / VAT on your prices", cfg.tax_rate,
                style="percent", name="TaxRate",
                validate=number_check(0, 1),
                help_text="Leave at 0 if you are not registered or the "
                          "marketplace handles it for you.")
    r = s.field(r, "Minimum wholesale order (items)", cfg.minimum_order_qty,
                style="int", name="MinOrderQty",
                validate=number_check(0, 10000, decimal=False))
    r = s.field(r, "Minimum wholesale order (value)",
                cfg.minimum_order_value, style="money",
                name="MinOrderValue", validate=number_check(0, 1000000))
    r = s.field(r, "Review a price after (days)", cfg.price_review_days,
                style="int", name="PriceReviewDays",
                validate=number_check(0, 3650, decimal=False),
                help_text="Material costs drift upwards quietly. This is "
                          "how long a price can go unchecked.")

    # ----------------------------------------------------------- 9. targets
    r += 1
    r = s.section(r, "9. Targets", LAST_COL)
    r = s.field(r, "Revenue target a month", cfg.revenue_target_month,
                style="money", name="RevenueTargetMonth",
                validate=number_check(0, 10000000))
    r = s.field(r, "Revenue target a year", cfg.revenue_target_year,
                style="money", name="RevenueTargetYear",
                validate=number_check(0, 100000000))
    r = s.field(r, "Profit target a month", cfg.profit_target_month,
                style="money", name="ProfitTargetMonth",
                validate=number_check(0, 10000000))
    r = s.field(r, "Hourly pay you are aiming for", cfg.hourly_rate,
                style="money", name="TargetHourly",
                validate=number_check(0, 1000),
                help_text="The dashboard works out what you are actually "
                          "earning an hour and compares it with this.")

    # ------------------------------------------------------ 10. ID prefixes
    r += 1
    r = s.section(r, "10. ID prefixes", LAST_COL)
    r = s.note(r, "The letters in front of the automatic numbers. Change "
                  "them before you start entering data, not after.",
               LAST_COL, first_col=1, height=24)
    for label, value, name in [
            ("Product IDs", "P-", "ProductPrefix"),
            ("Material IDs", "M-", "MaterialPrefix"),
            ("Bundle IDs", "BN-", "BundlePrefix"),
            ("Variation IDs", "VR-", "VariationPrefix"),
            ("Batch IDs", "BT-", "BatchPrefix"),
            ("Sale IDs", "SL-", "SalePrefix"),
            ("Wholesale customer IDs", "WC-", "CustomerPrefix"),
            ("Wholesale order IDs", "WO-", "OrderPrefix"),
            ("Development IDs", "DV-", "DevPrefix"),
            ("Event IDs", "EV-", "EventPrefix"),
            ("Waste log IDs", "WL-", "WastePrefix"),
            ("Crochet sheet IDs", "CRO-", "CrochetPrefix"),
            ("Jewelry sheet IDs", "JWL-", "JewelryPrefix"),
            ("Candle sheet IDs", "CND-", "CandlePrefix"),
            ("Clothing sheet IDs", "CLO-", "ClothingPrefix"),
            ("Woodworking sheet IDs", "WDW-", "WoodPrefix"),
            ("Digital sheet IDs", "DIG-", "DigitalPrefix")]:
        r = s.field(r, label, value, name=name)

    # --------------------------------------------------------- 11. about it
    r += 1
    r = s.section(r, "11. About this workbook", LAST_COL)
    r = s.field(r, "Product", cfg.product_name, mode="calc", value_end=5)
    r = s.field(r, "Version", cfg.version, mode="calc")
    r = s.field(r, "Created by", cfg.author, mode="calc")
    r = s.field(r, "Sheet protection password", cfg.password, mode="calc",
                help_text="Review > Unprotect Sheet if you need to change "
                          "a formula.")
    r = s.field(r, "Currency", f"{cfg.currency_symbol} "
                               f"({cfg.currency_code})", mode="calc")
    r += 1
    s.note(r,
           "Changed a rate here? Everything updates straight away - every "
           "product in the catalog, both price calculators, the "
           "break-even tab and the dashboard all read these cells rather "
           "than keeping their own copy. Sales you have already recorded "
           "keep the price you actually charged.",
           LAST_COL, first_col=1, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
