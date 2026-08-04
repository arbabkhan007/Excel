#!/usr/bin/env python3
"""
NUCLEAR FIX - Completely rebuild conditional formatting
Removes all containsText rules and recreates them as expression rules
"""
import re, os
from openpyxl import load_workbook
from openpyxl.formatting.rule import Rule
from openpyxl.styles import PatternFill

FILES = [
    "Construction_Management_Suite.xlsx",
    "Home_Renovation_Management_System.xlsx",
    "School_Management_System.xlsx",
]

# Map of (sheet_name, sqref, search_text) → fill color
# These are the EXACT rules we need to fix
CF_FIXES = {
    "Construction_Management_Suite.xlsx": [
        ("TL Resource Loading", "L3:L202", "YES", "FFC7CE"),  # red fill
    ],
    "Home_Renovation_Management_System.xlsx": [
        ("Material Inventory", "L3:L5002", "REORDER", "FFC7CE"),  # red
        ("Material Inventory", "L3:L5002", "OK", "C6EFCE"),  # green
        ("Cash Flow Tracker", "H3:H202", "Positive", "C6EFCE"),  # green
        ("Cash Flow Tracker", "H3:H202", "Negative", "FFC7CE"),  # red
    ],
    "School_Management_System.xlsx": [
        ("Timetable", "I3:I2000", "CONFLICT", "FFC7CE"),  # red
        ("Fee Defaulters", "K3:K5002", "CRITICAL", "FFC7CE"),  # red
        ("Fee Defaulters", "K3:K5002", "HIGH", "FFEB9C"),  # yellow
        ("Vehicle Master", "N3:N200", "EXPIRING", "FFEB9C"),  # yellow
        ("Stationery Stock", "K3:K2002", "REORDER", "FFC7CE"),  # red
        ("Stationery Stock", "K3:K2002", "OK", "C6EFCE"),  # green
        ("Medical Inventory", "I3:I1002", "REORDER", "FFC7CE"),  # red
        ("Medical Inventory", "I3:I1002", "EXPIRED", "FFC7CE"),  # red
        ("Budget Planning", "Q3:Q103", "OVER", "FFC7CE"),  # red
        ("Budget Planning", "Q3:Q103", "Within", "C6EFCE"),  # green
    ],
}

def make_fill(hex_color):
    """Create a PatternFill with the given hex color."""
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')

def fix_file(filepath):
    print(f"\n{'='*70}")
    print(f"  REPAIRING: {filepath}")
    print(f"{'='*70}")
    
    wb = load_workbook(filepath)
    fixes = CF_FIXES.get(os.path.basename(filepath), [])
    fixed_count = 0
    
    for sheet_name, sqref, search_text, fill_color in fixes:
        ws = wb[sheet_name]
        
        # Step 1: Find and remove the old containsText CF rule for this sqref+text
        cf_list = ws.conditional_formatting
        indices_to_remove = []
        
        for cf_idx, cf in enumerate(cf_list):
            if str(cf.sqref) == sqref:
                # Check if any rule in this CF is a containsText
                contains_text_rules = []
                other_rules = []
                for rule in cf.rules:
                    if rule.type == 'containsText' or (hasattr(rule, 'operator') and rule.operator == 'containsText'):
                        contains_text_rules.append(rule)
                    else:
                        other_rules.append(rule)
                
                if contains_text_rules:
                    # We need to rebuild this CF entry
                    # Remove the entire CF and recreate with valid rules
                    indices_to_remove.append(cf_idx)
                    
                    # Create new expression-based rule
                    first_cell = sqref.split(':')[0]
                    formula_str = f'ISNUMBER(SEARCH("{search_text}",{first_cell}))'
                    fill = make_fill(fill_color)
                    
                    new_rule = Rule(
                        type='expression',
                        formula=[formula_str],
                    )
                    # Create differential style for the fill
                    from openpyxl.styles.differential import DifferentialStyle
                    dxf = DifferentialStyle(fill=fill)
                    new_rule.dxf = dxf
                    
                    # Add the new CF
                    ws.conditional_formatting.add(sqref, new_rule)
                    
                    fixed_count += 1
                    print(f"    ✅ {sheet_name}: '{search_text}' on {sqref}")
                    print(f"       New rule: ISNUMBER(SEARCH(\"{search_text}\",{first_cell}))")
        
        # Remove old CF entries (reverse order)
        for idx in sorted(indices_to_remove, reverse=True):
            del cf_list._cf_rules[idx]
    
    # Also fix any leaked variables (from previous repair attempt)
    var_fixed = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    f = cell.value
                    original = f
                    f = re.sub(r'(?<=[(,])r(?=[><=+\-*/])', 'ROW()', f)
                    f = re.sub(r'(?<=[(,])r(?=\d)', 'ROW()', f)
                    if f != original:
                        cell.value = f
                        var_fixed += 1
    
    print(f"\n  Fixed: {fixed_count} CF rules, {var_fixed} formulas")
    wb.save(filepath)
    wb.close()
    print(f"  ✅ SAVED: {filepath}")


def verify_file(filepath):
    print(f"\n  VERIFYING: {filepath}")
    wb = load_workbook(filepath)
    issues = 0
    
    for sn in wb.sheetnames:
        ws = wb[sn]
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                if rule.type == 'containsText' or (hasattr(rule, 'operator') and rule.operator == 'containsText'):
                    issues += 1
                    print(f"    ❌ Still has containsText: {sn} sqref={cf.sqref}")
    
    if issues == 0:
        print(f"  ✅ 100% CLEAN")
    else:
        print(f"  ❌ {issues} issues remain")
    wb.close()
    return issues


if __name__ == "__main__":
    os.chdir("/home/user/Excel")
    
    for f in FILES:
        fix_file(f)
    
    print(f"\n{'='*70}")
    print(f"  VERIFICATION")
    print(f"{'='*70}")
    total = 0
    for f in FILES:
        total += verify_file(f)
    
    if total == 0:
        print(f"\n  🎉 ALL 3 FILES ARE 100% CLEAN — ZERO ISSUES")
    else:
        print(f"\n  ❌ {total} issues remain — need different approach")
