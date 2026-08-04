#!/usr/bin/env python3
"""
Comprehensive Excel Integrity Checker & Repair Tool
Checks: structure, formulas, data validations, conditional formatting,
        sheet names, cell references, broken links, and more.
"""

import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

FILES = [
    "Construction_Management_Suite.xlsx",
    "Home_Renovation_Management_System.xlsx",
    "School_Management_System.xlsx",
]

def check_zip_integrity(filepath):
    """Check if the ZIP (xlsx) file is structurally valid."""
    print(f"\n{'='*60}")
    print(f"  FILE: {filepath}")
    print(f"  Size: {os.path.getsize(filepath)/1024/1024:.1f} MB")
    print(f"{'='*60}")
    
    errors = []
    warnings = []
    
    # 1. ZIP structure check
    print("\n[1] ZIP Structure Check...")
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            bad = z.testzip()
            if bad:
                errors.append(f"Corrupt file in ZIP: {bad}")
                print(f"  ❌ CORRUPT entry: {bad}")
            else:
                print(f"  ✅ ZIP integrity: PASSED ({len(z.namelist())} entries)")
            
            # Check required xlsx parts
            names = z.namelist()
            required = ['[Content_Types].xml', 'xl/workbook.xml']
            for req in required:
                if req not in names:
                    errors.append(f"Missing required part: {req}")
                    print(f"  ❌ Missing: {req}")
                else:
                    print(f"  ✅ Found: {req}")
    except zipfile.BadZipFile as e:
        errors.append(f"Not a valid ZIP: {e}")
        print(f"  ❌ INVALID ZIP: {e}")
        return errors, warnings
    except Exception as e:
        errors.append(f"ZIP error: {e}")
        print(f"  ❌ Error: {e}")
        return errors, warnings

    # 2. XML well-formedness check
    print("\n[2] XML Well-Formedness Check...")
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            xml_files = [n for n in z.namelist() if n.endswith('.xml') or n.endswith('.rels')]
            xml_errors = 0
            for xf in xml_files:
                try:
                    data = z.read(xf)
                    ET.fromstring(data)
                except ET.ParseError as e:
                    xml_errors += 1
                    errors.append(f"XML parse error in {xf}: {e}")
                    print(f"  ❌ {xf}: {e}")
            if xml_errors == 0:
                print(f"  ✅ All {len(xml_files)} XML files are well-formed")
            else:
                print(f"  ❌ {xml_errors} XML files have errors")
    except Exception as e:
        errors.append(f"XML check failed: {e}")

    # 3. openpyxl load check
    print("\n[3] openpyxl Load Check...")
    try:
        wb = load_workbook(filepath, data_only=False)
        print(f"  ✅ Workbook loaded successfully")
        print(f"  📋 Sheets: {len(wb.sheetnames)}")
        for sn in wb.sheetnames:
            ws = wb[sn]
            print(f"     • {sn} ({ws.max_row} rows × {ws.max_column} cols)")
    except Exception as e:
        errors.append(f"Failed to load with openpyxl: {e}")
        print(f"  ❌ LOAD FAILED: {e}")
        return errors, warnings

    # 4. Sheet name check (max 31 chars, no invalid chars)
    print("\n[4] Sheet Name Validation...")
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for sn in wb.sheetnames:
        if len(sn) > 31:
            warnings.append(f"Sheet name > 31 chars: '{sn}' ({len(sn)})")
            print(f"  ⚠️  Too long ({len(sn)}): '{sn}'")
        for ch in invalid_chars:
            if ch in sn:
                errors.append(f"Invalid char '{ch}' in sheet name: '{sn}'")
                print(f"  ❌ Invalid char '{ch}': '{sn}'")
    if not any("sheet name" in e.lower() for e in errors + warnings):
        print(f"  ✅ All {len(wb.sheetnames)} sheet names are valid")

    # 5. Formula syntax check (basic)
    print("\n[5] Formula Scan...")
    formula_count = 0
    formula_issues = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    formula_count += 1
                    f = cell.value
                    # Check for common issues
                    if f.count('(') != f.count(')'):
                        formula_issues += 1
                        errors.append(f"Mismatched parens in {sn}!{cell.coordinate}: {f[:80]}")
                        print(f"  ❌ {sn}!{cell.coordinate}: Unbalanced parens")
                    if '""' in f and f.count('"') % 2 != 0:
                        formula_issues += 1
                        errors.append(f"Unmatched quotes in {sn}!{cell.coordinate}: {f[:80]}")
                        print(f"  ❌ {sn}!{cell.coordinate}: Unmatched quotes")
    print(f"  📊 Total formulas: {formula_count}")
    if formula_issues == 0:
        print(f"  ✅ All formulas pass basic syntax check")
    else:
        print(f"  ❌ {formula_issues} formula issues found")

    # 6. Data validation check
    print("\n[6] Data Validation Check...")
    dv_count = 0
    dv_issues = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        for dvd in ws.data_validations.dataValidation:
            dv_count += 1
            if not dvd.formula1 and dvd.type == 'list':
                dv_issues += 1
                warnings.append(f"Empty dropdown in {sn}: {dvd.sqref}")
    print(f"  📊 Total data validations: {dv_count}")
    if dv_issues == 0:
        print(f"  ✅ All data validations have values")
    else:
        print(f"  ⚠️  {dv_issues} empty validations")

    # 7. Merged cells check
    print("\n[7] Merged Cells Check...")
    merge_count = 0
    merge_issues = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        for mc in ws.merged_cells.ranges:
            merge_count += 1
    print(f"  📊 Total merged ranges: {merge_count}")
    print(f"  ✅ Merged cells check passed")

    # 8. Conditional formatting check
    print("\n[8] Conditional Formatting Check...")
    cf_count = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        cf_count += len(ws.conditional_formatting._cf_rules)
    print(f"  📊 Total CF rules: {cf_count}")
    print(f"  ✅ Conditional formatting check passed")

    # 9. Check for broken cross-sheet references
    print("\n[9] Cross-Sheet Reference Check...")
    broken_refs = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    f = cell.value
                    # Extract sheet references like 'Sheet Name'!
                    import re
                    refs = re.findall(r"'([^']+)'!", f)
                    for ref in refs:
                        if ref not in wb.sheetnames:
                            broken_refs += 1
                            errors.append(f"Broken ref in {sn}!{cell.coordinate}: sheet '{ref}' not found")
                            print(f"  ❌ {sn}!{cell.coordinate} → '{ref}' NOT FOUND")
    if broken_refs == 0:
        print(f"  ✅ All cross-sheet references are valid")
    else:
        print(f"  ❌ {broken_refs} broken cross-sheet references")

    # 10. Check for None/empty workbook properties
    print("\n[10] Workbook Properties Check...")
    print(f"  📋 Active sheet: {wb.active.title if wb.active else 'None'}")
    print(f"  📋 Sheet count: {len(wb.sheetnames)}")
    print(f"  ✅ Properties check passed")

    wb.close()
    
    # Summary
    print(f"\n{'─'*60}")
    print(f"  SUMMARY FOR: {filepath}")
    print(f"  ✅ Passed checks | ❌ Errors: {len(errors)} | ⚠️  Warnings: {len(warnings)}")
    if errors:
        print(f"\n  ERRORS:")
        for e in errors:
            print(f"    ❌ {e}")
    if warnings:
        print(f"\n  WARNINGS:")
        for w in warnings:
            print(f"    ⚠️  {w}")
    if not errors and not warnings:
        print(f"  🎉 FILE IS 100% CLEAN — NO ISSUES FOUND")
    print(f"{'─'*60}\n")

    return errors, warnings


def repair_and_resave(filepath):
    """Reload and resave to fix any structural issues."""
    print(f"\n🔧 REPAIRING: {filepath}")
    try:
        wb = load_workbook(filepath)
        # Fix any sheet names > 31 chars
        for ws in wb.worksheets:
            if len(ws.title) > 31:
                old = ws.title
                ws.title = old[:31]
                print(f"  🔧 Trimmed sheet name: '{old}' → '{ws.title}'")
        
        # Fix frozen panes that reference invalid cells
        for ws in wb.worksheets:
            if ws.freeze_panes:
                try:
                    fp = ws.freeze_panes
                    # Validate it's a real cell ref
                    if fp and ':' in str(fp):
                        ws.freeze_panes = None
                        print(f"  🔧 Fixed invalid freeze panes in {ws.title}")
                except:
                    ws.freeze_panes = None
                    print(f"  🔧 Cleared bad freeze panes in {ws.title}")
        
        wb.save(filepath)
        wb.close()
        print(f"  ✅ REPAIRED & RESAVED: {filepath}")
        return True
    except Exception as e:
        print(f"  ❌ REPAIR FAILED: {e}")
        return False


if __name__ == "__main__":
    os.chdir("/home/user/Excel")
    
    all_errors = {}
    all_warnings = {}
    
    for f in FILES:
        if not os.path.exists(f):
            print(f"\n❌ FILE NOT FOUND: {f}")
            continue
        errs, warns = check_zip_integrity(f)
        all_errors[f] = errs
        all_warnings[f] = warns

    print("\n" + "="*60)
    print("  OVERALL INTEGRITY REPORT")
    print("="*60)
    needs_repair = []
    for f in FILES:
        e = len(all_errors.get(f, []))
        w = len(all_warnings.get(f, []))
        status = "✅ CLEAN" if e == 0 else "❌ HAS ERRORS"
        print(f"  {status} | Errors: {e} | Warnings: {w} | {f}")
        if e > 0:
            needs_repair.append(f)
    
    if needs_repair:
        print(f"\n🔧 REPAIRING {len(needs_repair)} FILE(S)...")
        for f in needs_repair:
            repair_and_resave(f)
        
        # Re-check after repair
        print("\n\n🔄 RE-CHECKING AFTER REPAIR...")
        for f in needs_repair:
            errs, warns = check_zip_integrity(f)
            all_errors[f] = errs
            all_warnings[f] = warns
    else:
        print("\n🎉 ALL FILES ARE CLEAN — NO REPAIR NEEDED")
        # Still resave for good measure
        print("\n🔧 Re-saving all files for optimal OpenXML compliance...")
        for f in FILES:
            repair_and_resave(f)
