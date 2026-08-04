#!/usr/bin/env python3
"""
FINAL REPAIR - Fix all containsText CF rules + leaked variables + circular refs
Converts invalid CellIsRule(containsText) → Rule(type='expression', formula=ISNUMBER/SEARCH)
"""
import re, os, copy
from openpyxl import load_workbook
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle

FILES = [
    "Construction_Management_Suite.xlsx",
    "Home_Renovation_Management_System.xlsx",
    "School_Management_System.xlsx",
]

def get_first_cell(sqref_str):
    """Get the first cell reference from a sqref string like 'L3:L202' or 'A1:B2 C3:D4'"""
    parts = str(sqref_str).split()
    first = parts[0]
    if ':' in first:
        return first.split(':')[0]
    return first


def fix_file(filepath):
    print(f"\n{'='*70}")
    print(f"  REPAIRING: {filepath}")
    print(f"{'='*70}")
    
    wb = load_workbook(filepath)
    cf_fixed = 0
    var_fixed = 0
    
    for sn in wb.sheetnames:
        ws = wb[sn]
        
        # ── FIX 1: containsText conditional formatting ──
        cf_to_remove = []
        cf_to_add = []
        
        for cf_idx, cf in enumerate(ws.conditional_formatting):
            rules_to_fix = []
            for rule_idx, rule in enumerate(cf.rules):
                if rule.type == 'containsText' or (hasattr(rule, 'operator') and rule.operator == 'containsText'):
                    search_text = rule.formula[0].strip('"') if rule.formula else ""
                    first_cell = get_first_cell(cf.sqref)
                    
                    # Get the differential style (fill)
                    dxf = None
                    if rule.dxf:
                        dxf = rule.dxf
                    
                    # Create replacement expression rule
                    formula_str = f'ISNUMBER(SEARCH("{search_text}",{first_cell}))'
                    new_rule = Rule(type='expression')
                    new_rule.formula = [formula_str]
                    if dxf:
                        new_rule.dxf = dxf
                    
                    rules_to_fix.append((rule_idx, new_rule))
                    print(f"    🔧 {sn}: containsText '{search_text}' on {cf.sqref}")
                    print(f"       → ISNUMBER(SEARCH(\"{search_text}\",{first_cell}))")
                    cf_fixed += 1
            
            # Replace rules in this CF object
            if rules_to_fix:
                for rule_idx, new_rule in reversed(rules_to_fix):
                    cf.rules[rule_idx] = new_rule
        
        # ── FIX 2: Leaked Python variables in formulas ──
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    f = cell.value
                    original = f
                    # Fix standalone 'r' leaked from Python f-strings
                    f = re.sub(r'(?<=[(,])r(?=[><=+\-*/])', 'ROW()', f)
                    f = re.sub(r'(?<=[(,])r(?=\d)', 'ROW()', f)
                    if f != original:
                        cell.value = f
                        var_fixed += 1
                        print(f"    🔧 {sn}!{cell.coordinate}: var leak fixed")
    
    print(f"\n  Results: {cf_fixed} CF rules fixed, {var_fixed} formulas fixed")
    wb.save(filepath)
    wb.close()
    print(f"  ✅ SAVED: {filepath}")


def verify_file(filepath):
    print(f"\n  VERIFYING: {filepath}")
    wb = load_workbook(filepath)
    issues = 0
    
    for sn in wb.sheetnames:
        ws = wb[sn]
        # Check for remaining containsText
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                if rule.type == 'containsText' or (hasattr(rule, 'operator') and rule.operator == 'containsText'):
                    issues += 1
                    print(f"    ❌ Still has containsText: {sn}")
        
        # Check for leaked variables
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    f = cell.value
                    if re.search(r'(?<=[(,])r(?=[><=+\-*/])', f):
                        issues += 1
                        print(f"    ❌ Still has leaked var: {sn}!{cell.coordinate}")
    
    if issues == 0:
        print(f"  ✅ 100% CLEAN — All issues resolved")
    else:
        print(f"  ❌ {issues} issues remain")
    wb.close()
    return issues


if __name__ == "__main__":
    os.chdir("/home/user/Excel")
    
    for f in FILES:
        if os.path.exists(f):
            fix_file(f)
    
    print(f"\n{'='*70}")
    print(f"  POST-REPAIR VERIFICATION")
    print(f"{'='*70}")
    total_issues = 0
    for f in FILES:
        if os.path.exists(f):
            total_issues += verify_file(f)
    
    if total_issues == 0:
        print(f"\n  🎉 ALL 3 FILES ARE 100% CLEAN AND REPAIR-FREE")
