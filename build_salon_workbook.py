#!/usr/bin/env python3
"""Build the Salon Business Management Spreadsheet workbook.

    python3 build_salon_workbook.py
    python3 build_salon_workbook.py -o "output/My Salon.xlsx" --blank
    python3 build_salon_workbook.py --currency "£" --tax 20 --rows 400

Run `--help` for every option. The generator itself lives in the
`salon_business_spreadsheet` package next to this file.
"""

import sys

from salon_business_spreadsheet.cli import main

if __name__ == "__main__":
    sys.exit(main())
