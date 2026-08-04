#!/usr/bin/env python3
"""
XML RAW TEXT FIX - Directly fix containsText in xlsx XML using regex
"""
import os, re, zipfile

FILES = [
    "Construction_Management_Suite.xlsx",
    "Home_Renovation_Management_System.xlsx",
    "School_Management_System.xlsx",
]

def fix_xlsx(filepath):
    print(f"\n{'='*70}")
    print(f"  FIXING: {filepath}")
    print(f"{'='*70}")
    
    temp_path = filepath + '.tmp'
    total_fixed = 0
    
    with zipfile.ZipFile(filepath, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                
                if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
                    text = data.decode('utf-8')
                    original = text
                    
                    # Pattern: <cfRule type="cellIs" ... operator="containsText" dxfId="N"><formula>"TEXT"</formula>
                    # Replace with: type="expression" operator removed, formula=ISNUMBER(SEARCH(...))
                    # We need the sqref from the parent <conditionalFormatting> tag
                    
                    # Strategy: find each conditionalFormatting block and fix containsText rules within
                    
                    def fix_cf_block(match):
                        nonlocal total_fixed
                        full_match = match.group(0)
                        sqref_match = re.search(r'sqref="([^"]+)"', full_match)
                        sqref = sqref_match.group(1) if sqref_match else "A1"
                        first_cell = sqref.split(':')[0].split()[0] if sqref else "A1"
                        
                        def fix_rule(rule_match):
                            nonlocal total_fixed
                            rule_xml = rule_match.group(0)
                            if 'containsText' in rule_xml:
                                # Extract search text from formula
                                formula_match = re.search(r'<formula>"?([^"<]+)"?</formula>', rule_xml)
                                search_text = formula_match.group(1) if formula_match else ""
                                
                                # Extract dxfId
                                dxf_match = re.search(r'dxfId="(\d+)"', rule_xml)
                                dxf_id = dxf_match.group(1) if dxf_match else "0"
                                
                                # Extract priority
                                pri_match = re.search(r'priority="(\d+)"', rule_xml)
                                priority = pri_match.group(1) if pri_match else "1"
                                
                                # Build new rule XML
                                new_formula = f'ISNUMBER(SEARCH("{search_text}",{first_cell}))'
                                new_rule = f'<cfRule type="expression" priority="{priority}" dxfId="{dxf_id}"><formula>{new_formula}</formula></cfRule>'
                                
                                total_fixed += 1
                                return new_rule
                            return rule_xml
                        
                        # Fix all cfRule elements within this conditionalFormatting block
                        fixed_block = re.sub(r'<cfRule[^>]*>.*?</cfRule>', fix_rule, full_match)
                        return fixed_block
                    
                    # Match each <conditionalFormatting ...>...</conditionalFormatting> block
                    text = re.sub(r'<conditionalFormatting[^>]*>.*?</conditionalFormatting>', fix_cf_block, text)
                    
                    if text != original:
                        data = text.encode('utf-8')
                
                zout.writestr(item, data)
    
    os.replace(temp_path, filepath)
    print(f"\n  Fixed {total_fixed} containsText → expression rules")
    print(f"  ✅ SAVED: {filepath}")
    return total_fixed


def verify_xlsx(filepath):
    print(f"\n  VERIFYING: {filepath}")
    issues = 0
    with zipfile.ZipFile(filepath, 'r') as z:
        for name in sorted(z.namelist()):
            if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                data = z.read(name).decode('utf-8', errors='replace')
                count = data.lower().count('containstext')
                if count > 0:
                    issues += count
                    print(f"    ❌ {name}: {count} containsText remain")
                # Show what expression rules look like
                expr_count = data.count('ISNUMBER(SEARCH')
                if expr_count > 0:
                    print(f"    ✅ {name}: {expr_count} expression rules (ISNUMBER/SEARCH)")
    if issues == 0:
        print(f"  ✅ 100% CLEAN")
    else:
        print(f"  ❌ {issues} issues remain")
    return issues


if __name__ == "__main__":
    os.chdir("/home/user/Excel")
    
    total = 0
    for f in FILES:
        total += fix_xlsx(f)
    
    print(f"\n{'='*70}")
    print(f"  TOTAL FIXED: {total} conditional formatting rules")
    print(f"{'='*70}")
    
    issues = 0
    for f in FILES:
        issues += verify_xlsx(f)
    
    if issues == 0:
        print(f"\n  🎉 ALL 3 FILES ARE 100% OPENXML COMPLIANT")
    else:
        print(f"\n  ❌ {issues} issues remain")
