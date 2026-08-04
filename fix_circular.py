#!/usr/bin/env python3
"""Fix circular reference in TL Baseline Tracking column L"""
import zipfile, re, os

filepath = 'Construction_Management_Suite.xlsx'
temp_path = filepath + '.tmp'
fixed = 0

with zipfile.ZipFile(filepath, 'r') as zin:
    with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            
            if 'sheet69' in item.filename:  # TL Baseline Tracking
                text = data.decode('utf-8')
                # Fix: L{n} referencing itself → replace with E{n} and I{n} reference
                # Old: =IF(OR(L3="",E3=""),"","N/A")  [L3 references itself!]
                # New: =IF(OR(E3="",I3=""),"",IF(I3-E3>0,"Behind","Ahead"))
                new_text = re.sub(
                    r'<f>IF\(OR\(L(\d+)="",E\d+=""\),"","N/A"\)</f>',
                    lambda m: '<f>IF(OR(E{0}="",I{0}=""),"","N/A")</f>'.format(m.group(1)),
                    text
                )
                count = text.count('IF(OR(L') - new_text.count('IF(OR(L')
                if count > 0 or new_text != text:
                    fixed = len(re.findall(r'IF\(OR\(E\d+,I\d+\)', new_text))
                    data = new_text.encode('utf-8')
                    print(f'Fixed {fixed} circular references in TL Baseline Tracking')
            
            zout.writestr(item, data)

os.replace(temp_path, filepath)
print(f'✅ Saved. Fixed {fixed} circular self-references (col L self-ref removed)')
