"""Command line front end.

    python -m airbnb_rental_spreadsheet
    python -m airbnb_rental_spreadsheet -o "output/Airbnb.xlsx" --blank
    python -m airbnb_rental_spreadsheet --currency "GBP" --rows 1200
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import (AUTHOR, LISTING_TITLE, MAX_PROPERTIES,
                     PRODUCT_NAME, Config)

DEFAULT_OUTPUT = os.path.join(
    "output", "Airbnb & Short-Term Rental Management Spreadsheet.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="airbnb_rental_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty (no example "
                             "portfolio)")
    parser.add_argument("--no-protect", action="store_true",
                        help="do not protect the sheets (for "
                             "development)")
    parser.add_argument("--password", default=None,
                        help="sheet protection password")
    parser.add_argument("--author", default=AUTHOR,
                        help="document author / creator")
    parser.add_argument("--business", default=None,
                        help="pre-fill the business name on Settings")
    parser.add_argument("--owner", default=None,
                        help="pre-fill the host or owner's name")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="the year the reports open on")
    parser.add_argument("--nightly", type=float, default=None,
                        help="default nightly rate for a new property")
    parser.add_argument("--cleaning-fee", type=float, default=None,
                        help="default cleaning fee charged to guests")
    parser.add_argument("--cleaning-cost", type=float, default=None,
                        help="default cost of a turnover to you")
    parser.add_argument("--deposit", type=float, default=None,
                        help="default security deposit")
    parser.add_argument("--min-nights", type=int, default=None,
                        help="default minimum stay")
    parser.add_argument("--tax", type=float, default=None,
                        help="lodging tax rate as a percentage, e.g. "
                             "6 for 6%% (zero by default: most "
                             "platforms collect and remit it)")
    parser.add_argument("--tax-self", action="store_true",
                        help="you collect and remit lodging tax "
                             "yourself rather than the platform")
    parser.add_argument("--occupancy-target", type=float, default=None,
                        help="occupancy target as a percentage, e.g. 72")
    parser.add_argument("--adr-target", type=float, default=None,
                        help="average daily rate target")
    parser.add_argument("--revenue-target", type=float, default=None,
                        help="monthly revenue target")
    parser.add_argument("--properties", type=int, default=None,
                        help="how many properties the example "
                             "portfolio has")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the default row "
                             "counts are scaled to fit this many "
                             "booking rows)")
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
    if args.nightly is not None:
        kwargs["default_nightly_rate"] = args.nightly
    if args.cleaning_fee is not None:
        kwargs["default_cleaning_fee"] = args.cleaning_fee
    if args.cleaning_cost is not None:
        kwargs["default_cleaning_cost"] = args.cleaning_cost
    if args.deposit is not None:
        kwargs["default_security_deposit"] = args.deposit
    if args.min_nights is not None:
        kwargs["default_min_nights"] = args.min_nights
    if args.tax is not None:
        kwargs["tax_rate"] = round(args.tax / 100.0, 5)
    if args.tax_self:
        kwargs["tax_collected_by_platform"] = "No"
    if args.occupancy_target is not None:
        kwargs["occupancy_target"] = round(
            args.occupancy_target / 100.0, 4)
    if args.adr_target is not None:
        kwargs["adr_target"] = args.adr_target
    if args.revenue_target is not None:
        kwargs["revenue_target_month"] = args.revenue_target
        kwargs["revenue_target_year"] = args.revenue_target * 12
    if args.properties is not None:
        kwargs["property_count"] = args.properties
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.business:
        cfg.host.name = args.business
    if args.owner:
        cfg.host.owner = args.owner
    if args.rows:
        # every register scales against the Booking Tracker, which is
        # what `row_count()` and `scale_rows()` both read
        cfg.rows_override = args.rows
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

    platforms = ", ".join(name for name, _c, _p, counts, _n
                          in cfg.platform_table if counts)
    who = ("the platform" if cfg.tax_collected_by_platform == "Yes"
           else "you")

    print(f"{cfg.product_name}")
    print(f"  listing title : {LISTING_TITLE}")
    print(f"  author        : {cfg.author}")
    print(f"  sheets        : {len(SHEET_ORDER)}")
    print(f"  year          : {cfg.year}")
    print(f"  business      : {cfg.host.name}")
    print(f"  properties    : {cfg.property_count} in the example "
          f"portfolio (room for {MAX_PROPERTIES} on the "
          f"dashboard)")
    print(f"  platforms     : {platforms}")
    print(f"  nightly rate  : {cfg.currency_symbol}"
          f"{cfg.default_nightly_rate:,.2f}, cleaning fee "
          f"{cfg.currency_symbol}{cfg.default_cleaning_fee:,.2f}, "
          f"cost {cfg.currency_symbol}"
          f"{cfg.default_cleaning_cost:,.2f}")
    print(f"  lodging tax   : {cfg.tax_rate:.2%}, remitted by {who}")
    print(f"  targets       : {cfg.occupancy_target:.0%} occupancy, "
          f"ADR {cfg.currency_symbol}{cfg.adr_target:,.0f}, RevPAR "
          f"{cfg.currency_symbol}{cfg.revpar_target:,.0f}")
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
