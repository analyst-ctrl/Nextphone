import zipfile, re
z = zipfile.ZipFile(r'excels\hoy.xlsx')
cache = z.read('xl/pivotCache/pivotCacheDefinition1.xml').decode('utf-8')
wss = re.search(r'<worksheetSource[^>]*/>', cache)
rc = re.search(r'recordCount="\d+"', cache)
print('WorksheetSource:', wss.group(0))
print('RecordCount:', rc.group(0))
z.close()
