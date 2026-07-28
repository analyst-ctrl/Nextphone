import zipfile

z = zipfile.ZipFile(r'excels\Iphone (1).xlsx')

for i in range(1, 5):
    pt = z.read(f'xl/pivotTables/pivotTable{i}.xml').decode('utf-8')
    fname = f'xl/pivotTables/pivotTable{i}.xml'
    with open(f'pruebas\\pivot{i}_iphone1.xml', 'w', encoding='utf-8') as f:
        f.write(pt)
    print(f'Saved {fname} -> pivot{i}_iphone1.xml ({len(pt)} bytes)')

# Save cache
cache = z.read('xl/pivotCache/pivotCacheDefinition1.xml').decode('utf-8')
with open('pruebas\\cache_iphone1.xml', 'w', encoding='utf-8') as f:
    f.write(cache)
print(f'Saved cache ({len(cache)} bytes)')

# Save cache records
try:
    records = z.read('xl/pivotCache/pivotCacheRecords1.xml').decode('utf-8')
    with open('pruebas\\cache_records_iphone1.xml', 'w', encoding='utf-8') as f:
        f.write(records)
    print(f'Saved cache records ({len(records)} bytes)')
except:
    print('No cache records found')

# Save table
table = z.read('xl/tables/table1.xml').decode('utf-8')
with open('pruebas\\table_iphone1.xml', 'w', encoding='utf-8') as f:
    f.write(table)
print(f'Saved table ({len(table)} bytes)')

# Save sheet2 (Tabla) XML
sheet2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
with open('pruebas\\sheet2_iphone1.xml', 'w', encoding='utf-8') as f:
    f.write(sheet2)
print(f'Saved sheet2 ({len(sheet2)} bytes)')

# Save sheet2 rels
sheet2_rels = z.read('xl/worksheets/_rels/sheet2.xml.rels').decode('utf-8')
with open('pruebas\\sheet2_rels_iphone1.xml', 'w', encoding='utf-8') as f:
    f.write(sheet2_rels)
print(f'Saved sheet2 rels ({len(sheet2_rels)} bytes)')

# Save pivot cache rels
cache_rels = z.read('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels').decode('utf-8')
with open('pruebas\\cache_rels_iphone1.xml', 'w', encoding='utf-8') as f:
    f.write(cache_rels)
print(f'Saved cache rels ({len(cache_rels)} bytes)')

z.close()
