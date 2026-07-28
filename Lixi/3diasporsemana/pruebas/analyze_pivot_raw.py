import zipfile

z = zipfile.ZipFile(r'excels\Iphone (1).xlsx')

for i in range(1, 5):
    pt = z.read(f'xl/pivotTables/pivotTable{i}.xml').decode('utf-8')
    name_start = pt.find('name="')
    name_end = pt.find('"', name_start + 6)
    name = pt[name_start+6:name_end] if name_start != -1 else "?"
    loc_start = pt.find('location="')
    loc_end = pt.find('"', loc_start + 10)
    loc = pt[loc_start+10:loc_end] if loc_start != -1 else "?"
    
    print(f'\n=== PIVOT TABLE {i}: {name} ===')
    print(f'  Location: {loc}')
    
    # Print pageFields, rowFields, colFields sections
    for section in ['pageFields', 'rowFields', 'colFields', 'dataFields']:
        idx = pt.find(f'<{section}')
        if idx != -1:
            end = pt.find(f'</{section}>', idx)
            segment = pt[idx:end+len(section)+3]
            print(f'  {segment[:500]}')
        else:
            print(f'  {section}: NOT FOUND')

z.close()
