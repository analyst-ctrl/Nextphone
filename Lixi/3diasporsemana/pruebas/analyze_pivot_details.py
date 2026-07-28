import zipfile, re

z = zipfile.ZipFile(r'excels\Iphone (1).xlsx')

# Read full pivot table definitions
for i in range(1, 5):
    pt = z.read(f'xl/pivotTables/pivotTable{i}.xml').decode('utf-8')
    tag = re.search(r'<pivotTableDefinition[^>]*>', pt)
    loc = re.search(r'location="([^"]+)"', pt)
    
    # Extract row/col/filter fields
    pageFields = re.findall(r'<pageField x="(.*?)"', pt)
    rowFields = re.findall(r'<rowField x="(.*?)"', pt)
    colFields = re.findall(r'<colField x="(.*?)"', pt)
    
    # Extract pivotFields to understand axis
    pf_count = re.search(r'count="(\d+)"', pt.split('pivotFields')[0] if 'pivotFields' in pt else '')
    
    # Get the cacheField names to map indices
    cache = z.read('xl/pivotCache/pivotCacheDefinition1.xml').decode('utf-8')
    field_names = re.findall(r'cacheField name="([^"]+)"', cache)
    
    print(f'\n=== PIVOT TABLE {i} ===')
    if tag:
        attrs = tag.group(0)
        name_m = re.search(r'name="([^"]+)"', attrs)
        print(f'  Name: {name_m.group(1) if name_m else "N/A"}')
    if loc:
        print(f'  Location: {loc.group(1)}')
    
    print(f'  Page fields (filters): {[(int(x), field_names[int(x)]) for x in pageFields]}')
    print(f'  Row fields: {[(int(x), field_names[int(x)]) for x in rowFields]}')
    print(f'  Col fields: {[(int(x), field_names[int(x)]) for x in colFields]}')

z.close()
