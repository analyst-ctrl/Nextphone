import zipfile, re

z = zipfile.ZipFile(r'excels\hoy.xlsx')

# Check workbook
wb = z.read('xl/workbook.xml').decode('utf-8')
print('=== WORKBOOK ===')
sheets = re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb)
print('Sheets:', sheets)

# Check table definition
table = z.read('xl/tables/table1.xml').decode('utf-8')
print('\n=== TABLE1 ===')
tmatch = re.search(r'<table[^>]*>', table)
if tmatch:
    print(tmatch.group(0)[:500])

# Check pivot cache
cache = z.read('xl/pivotCache/pivotCacheDefinition1.xml').decode('utf-8')
print('\n=== PIVOT CACHE ===')
cmatch = re.search(r'<pivotCacheDefinition[^>]*>', cache)
if cmatch:
    print(cmatch.group(0)[:500])

# Check worksheet source in cache
wss = re.search(r'<worksheetSource[^>]*/>', cache)
if wss:
    print('WorksheetSource:', wss.group(0))

# Check sheet relationships to find which sheet is which
print('\n=== SHEET RELS ===')
for s in ['sheet1.xml.rels', 'sheet2.xml.rels']:
    try:
        content = z.read(f'xl/worksheets/_rels/{s}').decode('utf-8')
        print(f'{s}:')
        rels = re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', content)
        for rid, target in rels:
            print(f'  {rid} -> {target}')
    except:
        print(f'{s}: not found')

# Check workbook rels
print('\n=== WORKBOOK RELS ===')
wb_rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
rels = re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', wb_rels)
for rid, target in rels:
    print(f'  {rid} -> {target}')

z.close()
