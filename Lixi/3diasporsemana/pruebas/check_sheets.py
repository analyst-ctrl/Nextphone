import zipfile
z = zipfile.ZipFile(r'excels\hoy.xlsx')
for f in z.namelist():
    if 'sheet' in f.lower():
        print(f)
z.close()
