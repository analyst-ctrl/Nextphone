import zipfile, re

z = zipfile.ZipFile(r'excels\hoy.xlsx')

# Full cache definition
cache = z.read('xl/pivotCache/pivotCacheDefinition1.xml').decode('utf-8')
print('=== FULL PIVOT CACHE DEFINITION ===')
print(cache[:1000])
print('...')

# Connections
try:
    conn = z.read('xl/connections.xml').decode('utf-8')
    print('\n=== CONNECTIONS ===')
    print(conn[:2000])
except:
    print('\nNo connections.xml')

# Query table
try:
    qt = z.read('xl/queryTables/queryTable1.xml').decode('utf-8')
    print('\n=== QUERY TABLE ===')
    print(qt[:1000])
except:
    print('\nNo queryTable1.xml')

# Pivot table locations
for i in range(1, 5):
    pt = z.read(f'xl/pivotTables/pivotTable{i}.xml').decode('utf-8')
    loc = re.search(r'<pivotTableDefinition[^>]*>', pt)
    location = re.search(r'location="([^"]+)"', pt)
    if location:
        print(f'\nPivot {i} location: {location.group(1)}')

z.close()
