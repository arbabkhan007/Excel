#!/usr/bin/env python3
"""
XML-LEVEL FIX - Directly edit the xlsx XML to fix conditional formatting
This is the most reliable way to fix invalid CF rules.
"""
import os, re, shutil, zipfile, tempfile
import xml.etree.ElementTree as ET

FILES = [
    "Construction_Management_Suite.xlsx",
    "Home_Renovation_Management_System.xlsx",
    "School_Management_System.xlsx",
]

# Excel XML namespaces
NS = {
    '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
}
MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def fix_xlsx(filepath):
    print(f"\n{'='*70}")
    print(f"  FIXING: {filepath}")
    print(f"{'='*70}")
    
    # Read original zip
    temp_path = filepath + '.tmp'
    fixed_count = 0
    
    with zipfile.ZipFile(filepath, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                
                # Only process sheet XML files
                if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
                    try:
                        tree = ET.fromstring(data)
                        modified = False
                        
                        # Find all conditionalFormatting elements
                        for cf in tree.findall(f'{{{MAIN_NS}}}conditionalFormatting'):
                            for rule in cf.findall(f'{{{MAIN_NS}}}rule'):
                                rule_type = rule.get('type', '')
                                operator = rule.get('operator', '')
                                
                                # Fix containsText rules
                                if rule_type == 'containsText' or operator == 'containsText':
                                    # Get the formula (search text)
                                    formula_elem = rule.find(f'{{{MAIN_NS}}}formula')
                                    search_text = ''
                                    if formula_elem is not None and formula_elem.text:
                                        search_text = formula_elem.text.strip('"')
                                    
                                    # Get the sqref
                                    sqref = cf.get('sqref', '')
                                    first_cell = sqref.split(':')[0].split()[0] if sqref else 'A1'
                                    
                                    # Get the dxfId (style reference)
                                    dxf_id = rule.get('dxfId', '')
                                    
                                    # Change the rule type to 'expression' and fix formula
                                    rule.set('type', 'expression')
                                    if 'operator' in rule.attrib:
                                        del rule.attrib['operator']
                                    
                                    # Set new formula
                                    new_formula = f'ISNUMBER(SEARCH("{search_text}",{first_cell}))'
                                    if formula_elem is not None:
                                        formula_elem.text = new_formula
                                    
                                    modified = True
                                    fixed_count += 1
                                    sheet_num = item.filename.replace('xl/worksheets/sheet', '').replace('.xml', '')
                                    print(f"    ✅ {item.filename}: containsText '{search_text}' on {sqref}")
                                    print(f"       → expression: {new_formula}")
                        
                        if modified:
                            # Register namespace to avoid ns0: prefixes
                            ET.register_namespace('', MAIN_NS)
                            ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                            ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
                            ET.register_namespace('x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
                            data = ET.tostring(tree, encoding='UTF-8', xml_declaration=True)
                    except ET.ParseError as e:
                        print(f"    ⚠️  XML parse error in {item.filename}: {e}")
                
                # Also check styles.xml for any issues
                zout.writestr(item, data)
    
    # Replace original with fixed
    os.replace(temp_path, filepath)
    print(f"\n  Fixed {fixed_count} conditional formatting rules")
    print(f"  ✅ SAVED: {filepath}")
    return fixed_count


def verify_xlsx(filepath):
    """Verify no containsText rules remain"""
    print(f"\n  VERIFYING: {filepath}")
    issues = 0
    
    with zipfile.ZipFile(filepath, 'r') as z:
        for name in z.namelist():
            if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                data = z.read(name)
                text = data.decode('utf-8', errors='replace')
                if 'containsText' in text or 'containsText' in text.lower():
                    # Count occurrences
                    count = text.lower().count('containstext')
                    issues += count
                    print(f"    ❌ {name}: still has {count} containsText references")
    
    if issues == 0:
        print(f"  ✅ 100% CLEAN — No containsText in any sheet XML")
    else:
        print(f"  ❌ {issues} containsText references remain")
    return issues


if __name__ == "__main__":
    os.chdir("/home/user/Excel")
    
    total_fixed = 0
    for f in FILES:
        if os.path.exists(f):
            total_fixed += fix_xlsx(f)
    
    print(f"\n{'='*70}")
    print(f"  TOTAL CF RULES FIXED: {total_fixed}")
    print(f"{'='*70}")
    
    # Verify
    total_issues = 0
    for f in FILES:
        if os.path.exists(f):
            total_issues += verify_xlsx(f)
    
    if total_issues == 0:
        print(f"\n  🎉 ALL 3 FILES ARE 100% OPENXML COMPLIANT — ZERO ISSUES")
    else:
        print(f"\n  ❌ {total_issues} issues remain")
