#!/usr/bin/env python3
"""Build the Ultimate Airbnb & Short-Term Rental Management Spreadsheet.

    python3 build_airbnb_workbook.py
    python3 build_airbnb_workbook.py -o "output/Airbnb.xlsx" --blank
    python3 build_airbnb_workbook.py --currency "GBP" --rows 1200
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from airbnb_rental_spreadsheet.cli import main   # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
