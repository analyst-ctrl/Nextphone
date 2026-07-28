import zipfile, re

z = zipfile.ZipFile(r'excels\Iphone (1).xlsx')

for i in range(1, 5):
    pt = z.read(f'xl/pivotTables/pivotTable{i}.xml').decode('utf-8')
    name_m = re.search(r'name="([^"]+)"', pt)
    loc = re.search(r'location="([^"]+)"', pt)
    
    print(f'\n=== PIVOT TABLE {i}: {name_m.group(1) if name_m else "?"} ===')
    if loc:
        print(f'  Location: {loc.group(1)}')
    
    # Find page/row/col fields
    pf = re.findall(r'<pageField x="(\d+)"', pt)
    rf = re.findall(r'<rowField x="(\d+)"', pt)
    cf = re.findall(r'<colField x="(\d+)"', pt)
    print(f'  PageField indices: {pf}')
    print(f'  RowField indices: {rf}')
    print(f'  ColField indices: {cf}')

# Get cache field names
cache = z.read('xl/pivotCache/pivotCacheDefinition1.xml').decode('utf-8')
field_names = re.findall(r'cacheField name="([^"]+)"', cache)
print(f'\n=== CACHE FIELDS ({len(field_names)}) ===')
for i, name in enumerate(field_names):
    print(f'  [{i}] {name}')

z.close()
