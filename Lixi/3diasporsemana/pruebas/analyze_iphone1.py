import zipfile, re

z = zipfile.ZipFile(r'excels\Iphone (1).xlsx')

# Workbook structure
wb = z.read('xl/workbook.xml').decode('utf-8')
sheets = re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb)
print('=== HOJAS ===')
for name, rid in sheets:
    print(f'  {name} -> {rid}')

# Workbook rels
wb_rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
print('\n=== WORKBOOK RELS ===')
rels = re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', wb_rels)
for rid, target in rels:
    print(f'  {rid} -> {target}')

# Table definition
table = z.read('xl/tables/table1.xml').decode('utf-8')
tmatch = re.search(r'<table[^>]*>', table)
print('\n=== TABLE1 ===')
print(tmatch.group(0)[:600])

# Pivot cache
cache = z.read('xl/pivotCache/pivotCacheDefinition1.xml').decode('utf-8')
cmatch = re.search(r'<pivotCacheDefinition[^>]*>', cache)
wss = re.search(r'<worksheetSource[^>]*/>', cache)
print('\n=== PIVOT CACHE ===')
print(cmatch.group(0)[:500])
print('WorksheetSource:', wss.group(0) if wss else 'N/A')

# Pivot tables details
for i in range(1, 5):
    pt = z.read(f'xl/pivotTables/pivotTable{i}.xml').decode('utf-8')
    tag = re.search(r'<pivotTableDefinition[^>]*>', pt)
    loc = re.search(r'location="([^"]+)"', pt)
    print(f'\n=== PIVOT TABLE {i} ===')
    if tag:
        print(f'  {tag.group(0)[:400]}')
    if loc:
        print(f'  Location: {loc.group(1)}')

# Sheet rels
print('\n=== SHEET RELS ===')
for s in ['sheet1.xml.rels', 'sheet2.xml.rels']:
    try:
        content = z.read(f'xl/worksheets/_rels/{s}').decode('utf-8')
        rels = re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', content)
        print(f'{s}:')
        for rid, target in rels:
            print(f'  {rid} -> {target}')
    except:
        print(f'{s}: not found')

z.close()
