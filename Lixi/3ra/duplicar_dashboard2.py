# -*- coding: utf-8 -*-
"""
Duplica la hoja DASHBOARD2 como una nueva hoja 'DASHBOARD2 Copia'
dentro del mismo xlsx, manipulando el XML directamente para NO romper
pivots, graficos ni formulas.

- Hace backup del original antes de tocar nada.
- Copia xl/worksheets/sheet1.xml -> xl/worksheets/sheet16.xml
- Registra la nueva hoja en workbook.xml, workbook.xml.rels y [Content_Types].xml
"""
import os
import re
import shutil
import uuid
import zipfile

SRC = r"Lixi/3ra/reporte_cobro_dashboard (1).xlsx"
BAK = r"Lixi/3ra/reporte_cobro_dashboard (1) - BACKUP.xlsx"

NEW_SHEET_XML = "xl/worksheets/sheet16.xml"
NEW_SHEET_NAME = "DASHBOARD2 Copia"
NEW_SHEET_ID = "19"
NEW_RID = "rId26"
SRC_SHEET = "xl/worksheets/sheet1.xml"   # DASHBOARD2

# 1) Backup
if not os.path.exists(BAK):
    shutil.copy2(SRC, BAK)
    print(f"[OK] Backup creado: {BAK}")
else:
    print(f"[OK] Backup ya existia: {BAK}")

# 2) Leer todo el zip original
with zipfile.ZipFile(SRC, "r") as zin:
    items = zin.infolist()
    contents = {i.filename: zin.read(i.filename) for i in items}

sheet1 = contents[SRC_SHEET].decode("utf-8")
wbxml = contents["xl/workbook.xml"].decode("utf-8")
rels = contents["xl/_rels/workbook.xml.rels"].decode("utf-8")
ct = contents["[Content_Types].xml"].decode("utf-8")

# 3) Crear contenido de la hoja copia
newuid = "{%s}" % str(uuid.uuid4()).upper()
sheet16 = re.sub(r'xr:uid="\{[^}]*\}"', f'xr:uid="{newuid}"', sheet1, count=1)
# la copia no debe quedar como pestana seleccionada
sheet16 = sheet16.replace('tabSelected="1" ', '')

# 4) Registrar en [Content_Types].xml
if NEW_SHEET_XML not in ct:
    override = (f'<Override PartName="/{NEW_SHEET_XML}" '
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct = ct.replace("</Types>", override + "</Types>")

# 5) Registrar en workbook.xml.rels
if NEW_RID not in rels:
    rel = (f'<Relationship Id="{NEW_RID}" '
           'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
           f'Target="worksheets/sheet16.xml"/>')
    rels = rels.replace("</Relationships>", rel + "</Relationships>")

# 6) Registrar en workbook.xml (dentro de <sheets>)
if NEW_SHEET_NAME not in wbxml:
    sheet_entry = f'<sheet name="{NEW_SHEET_NAME}" sheetId="{NEW_SHEET_ID}" r:id="{NEW_RID}"/>'
    wbxml = wbxml.replace("</sheets>", sheet_entry + "</sheets>", 1)

# 7) Escribir el zip nuevo
contents[NEW_SHEET_XML] = sheet16.encode("utf-8")
contents["xl/workbook.xml"] = wbxml.encode("utf-8")
contents["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
contents["[Content_Types].xml"] = ct.encode("utf-8")

tmp = SRC + ".tmp"
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in items:
        zout.writestr(item, contents[item.filename])
    # la hoja nueva al final
    zi = zipfile.ZipInfo(NEW_SHEET_XML)
    zi.compress_type = zipfile.ZIP_DEFLATED
    zout.writestr(zi, sheet16.encode("utf-8"))

os.replace(tmp, SRC)
print("[OK] Hoja nueva añadida: DASHBOARD2 Copia")
print("[OK] Archivo reescrito:", SRC)
