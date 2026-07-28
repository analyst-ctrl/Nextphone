import zipfile, shutil, os, re

TGT = r'excels\hoy.xlsx'
TMP = r'excels\hoy_tmp2.xlsx'

shutil.copy2(TGT, TMP)

with zipfile.ZipFile(TMP, 'r') as zin:
    with zipfile.ZipFile(TGT, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            
            # Remove owssvr__28 sheet (sheet3.xml)
            if item.filename == 'xl/worksheets/sheet3.xml':
                continue
            if item.filename == 'xl/worksheets/_rels/sheet3.xml.rels':
                continue
            
            # Fix workbook.xml to remove owssvr__28 sheet
            if item.filename == 'xl/workbook.xml':
                content = data.decode('utf-8')
                # Remove the third sheet reference
                content = re.sub(r'<sheet[^>]*name="owssvr__28"[^/]*/>\s*', '', content)
                data = content.encode('utf-8')
            
            # Fix workbook.xml.rels to remove rId5 (which pointed to sheet3)
            if item.filename == 'xl/_rels/workbook.xml.rels':
                content = data.decode('utf-8')
                content = re.sub(r'<Relationship[^>]*Target="worksheets/sheet3.xml"[^/]*/>\s*', '', content)
                data = content.encode('utf-8')
            
            # Fix Content_Types to remove sheet3
            if item.filename == '[Content_Types].xml':
                content = data.decode('utf-8')
                content = re.sub(r'<Override[^>]*PartName="/xl/worksheets/sheet3.xml"[^/]*/>\s*', '', content)
                data = content.encode('utf-8')
            
            zout.writestr(item, data)

os.remove(TMP)

# Verify
with zipfile.ZipFile(TGT, 'r') as z:
    wb = z.read('xl/workbook.xml').decode('utf-8')
    sheets = re.findall(r'name="([^"]+)"', wb)
    print(f'Hojas: {sheets}')
print('Limpieza completada!')
