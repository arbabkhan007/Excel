#!/usr/bin/env python3
"""
FINAL REPAIR SCRIPT
Fixes:
1. CellIsRule with 'containsText' operator → FormulaRule (ISNUMBER/SEARCH)
   This is the cause of "Removed Records: Conditional formatting" in ALL 3 files
2. Any remaining leaked Python variables in formulas
3. Circular references
4. Re-saves with clean OpenXML
"""
import re, os
from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import PatternFill

LIGHT_GREEN = "C6EFCE"
LIGHT_RED   = "FFC7CE"
YELLOW_FILL = "FFEB9C"
BDD7EE      = "BDD7EE"
GREY_BG     = "F2F2F2"

FILES = [
    "Construction_Management_Suite.xlsx",
    "Home_Renovation_Management_System.xlsx",
    "School_Management_System.xlsx",
]

def fix_contains_text_rules(ws, sheet_name):
    """Replace CellIsRule(containsText) with FormulaRule using ISNUMBER(SEARCH())"""
    fixed = 0
    rules_to_remove = []
    rules_to_add = []
    
    for cf_idx, cf_rule in enumerate(ws.conditional_formatting._cf_rules):
        rule = cf_rule.rule
        sqref = str(cf_rule.sqref)
        
        # Check if it's a CellIsRule with containsText
        if isinstance(rule, CellIsRule) and rule.operator == 'containsText':
            # Get the formula text
            formula_text = rule.formula[0] if rule.formula else ""
            # Extract the search string (remove quotes)
            search_str = formula_text.strip('"')
            
            # Get fill color from the rule
            fill = None
            if rule.dxf and rule.dxf.fill:
                fill = rule.dxf.fill
            
            # Create a FormulaRule that does the same thing
            # ISNUMBER(SEARCH("text", cell_ref)) 
            # We need the top-left cell of the range
            range_cells = sqref.split()
            if range_cells:
                first_cell = range_cells[0].split(':')[0]
                # Make it relative to the first cell in the range
                formula_str = f'ISNUMBER(SEARCH("{search_str}",{first_cell}))'
                
                new_rule = FormulaRule(
                    formula=[formula_str],
                    fill=fill
                )
                rules_to_remove.append(cf_idx)
                rules_to_add.append((sqref, new_rule))
                fixed += 1
                print(f"    🔧 {sheet_name}: containsText '{search_str}' on {sqref} → FormulaRule")
    
    # Remove old rules (in reverse order to preserve indices)
    for idx in sorted(rules_to_remove, reverse=True):
        del ws.conditional_formatting._cf_rules[idx]
    
    # Add new rules
    for sqref, rule in rules_to_add:
        ws.conditional_formatting.add(sqref, rule)
    
    return fixed


def fix_leaked_variables(ws, sheet_name):
    """Fix any remaining leaked Python variables (r, c, i, etc.) in formulas"""
    fixed = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                f = cell.value
                original = f
                # Fix standalone 'r' that's a Python variable leak (not part of a cell ref)
                # Pattern: after ( or , followed by 'r' followed by > < = + -
                f = re.sub(r'(?<=[(,])r(?=[><=+\-*/])', 'ROW()', f)
                # Fix r-1, r+1 patterns
                f = re.sub(r'(?<=[(,])r(?=\d)', 'ROW()', f)
                if f != original:
                    cell.value = f
                    fixed += 1
                    print(f"    🔧 {sheet_name}!{cell.coordinate}: {original[:60]} → {f[:60]}")
    return fixed


def fix_circular_refs(ws, sheet_name):
    """Detect and fix circular references"""
    fixed = 0
    from openpyxl.utils import get_column_letter, column_index_from_string
    
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                f = cell.value
                cell_ref = cell.coordinate
                col_letter = get_column_letter(cell.column)
                row_num = cell.row
                
                # Check if formula references its own cell
                # Pattern: exact match of cell coordinate in formula
                # e.g., cell W23 has formula referencing W23
                self_refs = [
                    f'{col_letter}{row_num}',  # Direct self-ref
                ]
                
                for self_ref in self_refs:
                    # Check if the formula contains a reference to itself
                    # But exclude cases where it's part of a range like SUM(A1:A100)
                    pattern = rf'(?<![A-Z]){re.escape(self_ref)}(?![A-Z0-9])'
                    if re.search(pattern, f[1:]):  # Skip the = sign
                        # Check it's not in a range context
                        range_pattern = rf'[A-Z]+\d+:{re.escape(self_ref)}'
                        if not re.search(range_pattern, f):
                            print(f"    ⚠️  CIRCULAR REF: {sheet_name}!{cell_ref}: {f[:80]}")
                            # Try to fix common patterns
                            if 'SUM(' in f and self_ref in f:
                                # If it's SUM(range:current_cell), exclude current cell
                                f = f.replace(f'+{self_ref}', '').replace(f',{self_ref}', '')
                                cell.value = f
                                fixed += 1
    return fixed


def remove_invalid_cf(ws, sheet_name):
    """Remove any conditional formatting rules that are structurally invalid"""
    removed = 0
    rules_to_remove = []
    
    for cf_idx, cf_rule in enumerate(ws.conditional_formatting._cf_rules):
        rule = cf_rule.rule
        
        # Check for rules with empty formulas
        if hasattr(rule, 'formula') and (not rule.formula or all(f == '' for f in rule.formula)):
            rules_to_remove.append(cf_idx)
            removed += 1
            print(f"    🔧 {sheet_name}: Removed empty CF rule on {cf_rule.sqref}")
        
        # Check for CellIsRule with invalid operators
        if isinstance(rule, CellIsRule):
            valid_ops = {'equal', 'notEqual', 'greaterThan', 'lessThan', 
                        'between', 'notBetween', 'greaterThanOrEqual', 'lessThanOrEqual'}
            if rule.operator not in valid_ops:
                rules_to_remove.append(cf_idx)
                removed += 1
                print(f"    🔧 {sheet_name}: Removed invalid CF operator '{rule.operator}' on {cf_rule.sqref}")
    
    for idx in sorted(rules_to_remove, reverse=True):
        del ws.conditional_formatting._cf_rules[idx]
    
    return removed


def process_file(filepath):
    """Process a single file: fix all issues"""
    print(f"\n{'='*70}")
    print(f"  REPAIRING: {filepath}")
    print(f"{'='*70}")
    
    wb = load_workbook(filepath)
    total_fixed = 0
    total_cf_fixed = 0
    total_cf_removed = 0
    
    for sn in wb.sheetnames:
        ws = wb[sn]
        
        # 1. Fix containsText rules
        cf_fixed = fix_contains_text_rules(ws, sn)
        total_cf_fixed += cf_fixed
        
        # 2. Fix leaked variables
        var_fixed = fix_leaked_variables(ws, sn)
        total_fixed += var_fixed
        
        # 3. Fix circular references
        circ_fixed = fix_circular_refs(ws, sn)
        total_fixed += circ_fixed
        
        # 4. Remove invalid CF
        cf_removed = remove_invalid_cf(ws, sn)
        total_cf_removed += cf_removed
    
    print(f"\n  SUMMARY: {filepath}")
    print(f"    CF containsText→FormulaRule: {total_cf_fixed}")
    print(f"    Leaked variables fixed: {total_fixed}")
    print(f"    Invalid CF removed: {total_cf_removed}")
    
    wb.save(filepath)
    wb.close()
    print(f"  ✅ SAVED: {filepath}")
    return total_cf_fixed, total_fixed, total_cf_removed


def verify_file(filepath):
    """Verify the file after repair"""
    print(f"\n  VERIFYING: {filepath}")
    wb = load_workbook(filepath)
    
    issues = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        for cf_rule in ws.conditional_formatting._cf_rules:
            rule = cf_rule.rule
            if isinstance(rule, CellIsRule) and rule.operator == 'containsText':
                issues += 1
                print(f"    ❌ Still has containsText: {sn} on {cf_rule.sqref}")
    
    if issues == 0:
        print(f"  ✅ VERIFIED CLEAN - No containsText rules remain")
    else:
        print(f"  ❌ {issues} containsText rules still remain")
    
    wb.close()
    return issues


if __name__ == "__main__":
    os.chdir("/home/user/Excel")
    
    grand_cf = 0
    grand_var = 0
    grand_rem = 0
    
    for f in FILES:
        if os.path.exists(f):
            cf, var, rem = process_file(f)
            grand_cf += cf
            grand_var += var
            grand_rem += rem
    
    print(f"\n{'='*70}")
    print(f"  TOTAL REPAIRS ACROSS ALL FILES")
    print(f"{'='*70}")
    print(f"  CF containsText→FormulaRule: {grand_cf}")
    print(f"  Leaked variables fixed: {grand_var}")
    print(f"  Invalid CF removed: {grand_rem}")
    
    # Verify
    print(f"\n{'='*70}")
    print(f"  POST-REPAIR VERIFICATION")
    print(f"{'='*70}")
    for f in FILES:
        if os.path.exists(f):
            verify_file(f)
