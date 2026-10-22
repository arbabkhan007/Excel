"""Command line front end.

    python -m salon_business_spreadsheet
    python -m salon_business_spreadsheet -o "output/Salon.xlsx" --blank
    python -m salon_business_spreadsheet --currency "£" --tax 20 --rows 400
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import (AUTHOR, COMMISSION_METHODS, PRODUCT_LISTING_TITLE,
                     PRODUCT_NAME, ROUNDING_RULES, Config)

DEFAULT_OUTPUT = os.path.join(
    "output", "Salon Business Management Spreadsheet.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="salon_business_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty (no demo salon)")
    parser.add_argument("--no-protect", action="store_true",
                        help="do not protect the sheets (for development)")
    parser.add_argument("--password", default=None,
                        help="sheet protection password")
    parser.add_argument("--author", default=AUTHOR,
                        help="document author / creator")
    parser.add_argument("--salon", default=None,
                        help="pre-fill the salon name on Settings")
    parser.add_argument("--owner", default=None,
                        help="pre-fill the owner's name")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="the year the reports open on")
    parser.add_argument("--tax", type=float, default=None,
                        help="sales tax / VAT rate as a percentage, "
                             "e.g. 20 for 20%%")
    parser.add_argument("--tax-on", choices=["both", "retail",
                                             "services", "neither"],
                        default=None,
                        help="what the tax rate applies to")
    parser.add_argument("--commission", type=float, default=None,
                        help="default service commission as a "
                             "percentage, e.g. 40")
    parser.add_argument("--commission-method", choices=COMMISSION_METHODS,
                        default=None,
                        help="the method new staff rows start on")
    parser.add_argument("--hourly", type=float, default=None,
                        help="default hourly rate")
    parser.add_argument("--burden", type=float, default=None,
                        help="employer costs on wages as a percentage")
    parser.add_argument("--deposit", type=float, default=None,
                        help="deposit as a percentage of the price")
    parser.add_argument("--rounding", choices=ROUNDING_RULES,
                        default=None,
                        help="how calculated prices are rounded")
    parser.add_argument("--target", type=float, default=None,
                        help="monthly revenue target")
    parser.add_argument("--chairs", type=int, default=None,
                        help="how many chairs or stations the salon has")
    parser.add_argument("--open-days", type=int, default=None,
                        help="trading days a week")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the default row "
                             "counts are scaled to fit this many "
                             "appointment rows)")
    parser.add_argument("--today", default=None,
                        help="treat this ISO date as 'today' when "
                             "generating the example data")
    parser.add_argument("--quiet", action="store_true",
                        help="only print the output path")
    return parser


def config_from_args(args: argparse.Namespace) -> Config:
    kwargs = {}
    if args.blank:
        kwargs["include_sample_data"] = False
    if args.no_protect:
        kwargs["protect_sheets"] = False
    if args.password:
        kwargs["password"] = args.password
    if args.author:
        kwargs["author"] = args.author
    if args.currency:
        kwargs["currency_symbol"] = args.currency
    if args.currency_code:
        kwargs["currency_code"] = args.currency_code
    if args.year:
        kwargs["plan_year"] = args.year
    if args.tax is not None:
        kwargs["tax_rate"] = round(args.tax / 100.0, 5)
    if args.tax_on:
        kwargs["tax_on_retail"] = ("Yes" if args.tax_on in
                                   ("both", "retail") else "No")
        kwargs["tax_on_services"] = ("Yes" if args.tax_on in
                                     ("both", "services") else "No")
    if args.commission is not None:
        kwargs["service_commission"] = round(args.commission / 100.0, 4)
    if args.commission_method:
        kwargs["default_commission_method"] = args.commission_method
    if args.hourly is not None:
        kwargs["hourly_rate"] = args.hourly
    if args.burden is not None:
        kwargs["payroll_burden"] = round(args.burden / 100.0, 4)
    if args.deposit is not None:
        kwargs["deposit_percent"] = round(args.deposit / 100.0, 4)
    if args.rounding:
        kwargs["rounding_rule"] = args.rounding
    if args.target is not None:
        kwargs["monthly_revenue_target"] = args.target
        kwargs["sample_monthly_revenue_target"] = args.target
    if args.chairs is not None:
        kwargs["chairs"] = args.chairs
    if args.open_days is not None:
        kwargs["open_days"] = args.open_days
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.salon:
        cfg.salon.name = args.salon
    if args.owner:
        cfg.salon.owner = args.owner
    if args.rows:
        cfg.scale_rows(args.rows)
    return cfg


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        cfg = config_from_args(args)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    path = build_workbook(args.output, cfg)
    size = os.path.getsize(path) / 1024.0

    if args.quiet:
        print(path)
        return 0

    tax_on = []
    if cfg.tax_on_services == "Yes":
        tax_on.append("services")
    if cfg.tax_on_retail == "Yes":
        tax_on.append("retail")

    print(f"{cfg.product_name}")
    print(f"  listing title : {PRODUCT_LISTING_TITLE}")
    print(f"  author        : {cfg.author}")
    print(f"  sheets        : {len(SHEET_ORDER)}")
    print(f"  year          : {cfg.year}")
    print(f"  salon         : {cfg.salon.name}")
    print(f"  chairs        : {cfg.chairs} over {cfg.open_days} days "
          f"({cfg.chair_hours_week:,.0f} sellable hours a week)")
    print(f"  tax           : {cfg.tax_rate:.2%} on "
          f"{' and '.join(tax_on) if tax_on else 'nothing'}")
    print(f"  commission    : {cfg.service_commission:.0%} default, "
          f"method '{cfg.default_commission_method}'")
    print(f"  target        : {cfg.currency_symbol}"
          f"{cfg.revenue_target_month:,.0f} a month")
    print(f"  example data  : "
          f"{'no' if not cfg.include_sample_data else 'yes'}")
    print(f"  protection    : "
          f"{'off' if not cfg.protect_sheets else 'on, password ' + cfg.password}")
    print(f"  written to    : {os.path.abspath(path)}  ({size:,.0f} KB)")
    return 0


if __name__ == "__main__":        # pragma: no cover
    raise SystemExit(main())
