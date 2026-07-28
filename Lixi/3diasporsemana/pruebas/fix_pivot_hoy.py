import zipfile, shutil, os, re

src = r'excels\hoy.xlsx'
tmp = r'excels\hoy_tmp.xlsx'

# Step 1: Copy xlsx to temp
shutil.copy2(src, tmp)

# Step 2: Read, modify pivot cache, write back
with zipfile.ZipFile(tmp, 'r') as zin:
    with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            
            # Fix pivot cache definition
            if item.filename == 'xl/pivotCache/pivotCacheDefinition1.xml':
                content = data.decode('utf-8')
                print('ANTES:', re.search(r'<worksheetSource[^>]*/>', content).group(0))
                # Change Tabla8 -> Tabla_owssvr__28
                content = content.replace('name="Tabla8"', 'name="Tabla_owssvr__28"')
                # Update record count to match actual data
                content = re.sub(r'recordCount="\d+"', 'recordCount="38816"', content)
                print('DESPUES:', re.search(r'<worksheetSource[^>]*/>', content).group(0))
                data = content.encode('utf-8')
            
            zout.writestr(item, data)

# Clean up temp
os.remove(tmp)
print('\nListo! Pivot cache actualizado en hoy.xlsx')
print('  - worksheetSource apunta a Tabla_owssvr__28')
print('  - recordCount = 38816')
print('  Abrir en Excel y refrescar tablas dinamicas (Ctrl+Alt+F5)')
