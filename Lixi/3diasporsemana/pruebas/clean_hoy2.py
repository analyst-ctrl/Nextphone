import zipfile, shutil, os, re

TGT = r'excels\hoy.xlsx'
TMP = r'excels\hoy_tmp3.xlsx'

# First, check what's in workbook.xml
with zipfile.ZipFile(TGT, 'r') as z:
    wb = z.read('xl/workbook.xml').decode('utf-8')
    print('BEFORE:')
    print(wb[:2000])

shutil.copy2(TGT, TMP)

with zipfile.ZipFile(TMP, 'r') as zin:
    with zipfile.ZipFile(TGT, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            
            # Skip owssvr__28 sheet files
            if item.filename in ('xl/worksheets/sheet3.xml', 'xl/worksheets/_rels/sheet3.xml.rels'):
                continue
            
            if item.filename == 'xl/workbook.xml':
                content = data.decode('utf-8')
                # Print the sheets section
                sheets_match = re.search(r'<sheets>.*?</sheets>', content, re.DOTALL)
                if sheets_match:
                    print(f'\nSheets section: {sheets_match.group(0)[:500]}')
                # Remove owssvr__28 sheet reference
                content = re.sub(r'\s*<sheet\s+name="owssvr__28"[^/]*/>', '', content)
                print(f'\nAFTER removal:')
                sheets_match2 = re.search(r'<sheets>.*?</sheets>', content, re.DOTALL)
                if sheets_match2:
                    print(f'Sheets section: {sheets_match2.group(0)[:500]}')
                data = content.encode('utf-8')
            
            if item.filename == 'xl/_rels/workbook.xml.rels':
                content = data.decode('utf-8')
                content = re.sub(r'\s*<Relationship[^>]*owssvr[^/]*/>', '', content)
                content = re.sub(r'\s*<Relationship[^>]*sheet3\.xml[^/]*/>', '', content)
                data = content.encode('utf-8')
            
            if item.filename == '[Content_Types].xml':
                content = data.decode('utf-8')
                content = re.sub(r'\s*<Override[^>]*sheet3[^/]*/>', '', content)
                data = content.encode('utf-8')
            
            zout.writestr(item, data)

os.remove(TMP)

with zipfile.ZipFile(TGT, 'r') as z:
    wb = z.read('xl/workbook.xml').decode('utf-8')
    sheets = re.findall(r'name="([^"]+)"', wb)
    print(f'\nHojas finales: {sheets}')
