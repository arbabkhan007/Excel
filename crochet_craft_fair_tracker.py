#!/usr/bin/env python3
"""
Command-line builder for the Ultimate Crochet Craft Fair Tracker.

Examples
--------
    python3 crochet_craft_fair_tracker.py --all --outdir products/
    python3 crochet_craft_fair_tracker.py --edition premium --theme berry \
        --mode demo --out products/preview.xlsx
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from crochet_tracker import workbook as W          # noqa: E402
from crochet_tracker import config as C            # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--edition", choices=sorted(C.EDITIONS), default="premium")
    p.add_argument("--theme", choices=["berry", "mint"], default="berry")
    p.add_argument("--mode", choices=["blank", "demo"], default="blank")
    p.add_argument("--all", action="store_true",
                   help="build the curated six-file product set")
    p.add_argument("--out", default="crochet_tracker.xlsx")
    p.add_argument("--outdir", default="products")
    p.add_argument("--no-protect", dest="protect",
                   action="store_false", default=None)
    p.add_argument("--no-images", dest="images", action="store_false")
    args = p.parse_args(argv)

    if args.all:
        rows = W.build_all(args.outdir, protect=args.protect,
                           images=args.images)
    else:
        stats = W.build_workbook(args.out, args.edition, args.theme,
                                 args.mode, protect=args.protect,
                                 images=args.images)
        rows = [(os.path.basename(args.out), stats)]

    print("=" * 78)
    print("BUILD COMPLETE")
    print("=" * 78)
    for name, st in rows:
        print("  %-58s %7.1f KB" % (name, st["size"] / 1024.0))
        print("      formulas %5d   validations %3d   cond. formats %3d   "
              "charts %2d   links %3d"
              % (st["formulas"], st["validations"], st["cond_formats"],
                 st["charts"], st["links"]))
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
