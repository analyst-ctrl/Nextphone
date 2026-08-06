# -*- coding: utf-8 -*-
"""
===============================================================================
  PIPELINE RENOVACION SOHO -> dash_prod
  Gerente: Amir Josue Rodriguez Chavarria
  Campaña: RENOVACION (equipo Marinel Moreno)

  Lectura:
    - extrainfo/HC*.xlsx          -> plantillas de equipo (3 equipos)
    - Reporte_soho_junio_25 (1).xlsb -> Vici (marcaciones) + Ventas + Facturas
    - bases/Base para campaña cross-sell movil Nextphone.xlsx -> venta cruzada

  Genera:
    - data/renovacion_data.json  -> datos para el dashboard
    - index.html                 -> dashboard auto-contenido (HTML+JS+Chart.js)
      con navegación por meses (TODOS / ENERO..JUNIO)

  Uso:
    python dash_prod/extraer_datos.py
===============================================================================
"""
import glob, json, re, os, datetime
from collections import Counter, defaultdict
from datetime import date, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
PROD = os.path.dirname(BASE)

def norm(s):
    if s is None: return ''
    s = str(s).strip()
    # Arreglar mojibake: texto UTF-8 leído como cp1252 (ej: 'Ñ' -> 'Ã'+'\u2018')
    try:
        s = s.encode('cp1252').decode('utf-8')
    except Exception:
        pass
    s = s.upper()
    s = re.sub(r'\s+', ' ', s)
    for a, b in [('Á','A'),('É','E'),('Í','I'),('Ó','O'),('Ú','U'),('Ñ','N'),('Ü','U')]:
        s = s.replace(a, b)
    return s

def to_num(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '.').replace(' ', ''))
    except Exception:
        return 0.0

def cuenta12(v):
    """Normaliza una cuenta a 12 dígitos (o los últimos 8 como fallback)."""
    if v is None: return ''
    s = str(v).replace('.0', '').strip()
    s = re.sub(r'\D', '', s)
    if len(s) >= 12: return s[-12:]
    if len(s) >= 8: return s[-8:]
    return ''

def phone8(v):
    """Normaliza un teléfono a 8 dígitos (móvil panameño)."""
    if v is None: return ''
    s = str(v).replace('.0', '').strip()
    s = re.sub(r'\D', '', s)
    if len(s) >= 8: return s[-8:]
    if len(s) == 7: return '6' + s
    return ''

def tokens(name):
    return set(norm(name).split())

def match_member(name, roster):
    """Devuelve el miembro del roster que coincide con el nombre.

    Requisitos (para evitar falsos positivos por apellidos compartidos):
      - El PRIMER token del nombre (nombre de pila) debe estar en el miembro
      - Al menos 2 tokens en común
    """
    n = norm(name)
    if not n: return None
    toks = n.split()
    if not toks: return None
    first = toks[0]  # nombre de pila (orden real, no por longitud)
    t = set(toks)
    best, best_score = None, 0
    for member in roster:
        mt = tokens(member)
        if first not in mt:
            continue
        score = len(t & mt)
        if score >= 2 and score > best_score:
            best, best_score = member, score
    return best

def get_col(idx, *nombres):
    """Devuelve el índice de la primera columna cuyo nombre normalizado coincida."""
    for n in nombres:
        k = norm(n)
        if k in idx:
            return idx[k]
    return None

# =============================================================================
# 0. MESES
# =============================================================================
MESES_ORDEN = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO']
MES_ABREV = {'ENE':'ENERO','FEB':'FEBRERO','MAR':'MARZO','ABR':'ABRIL',
             'MAY':'MAYO','JUN':'JUNIO','JUL':'JULIO','AGO':'AGOSTO',
             'SEP':'SEPTIEMBRE','OCT':'OCTUBRE','NOV':'NOVIEMBRE','DIC':'DICIEMBRE'}
MES_NUM = {f'{i:02d}': m for i, m in enumerate(MESES_ORDEN, 1)}

def mes_de(v):
    """Convierte una fecha/mes (texto o serial Excel) a nombre de mes en MESES_ORDEN.
    Devuelve '' si está fuera de los meses del reporte (enero-junio)."""
    if v is None: return ''
    if isinstance(v, (int, float)):
        try:
            d = date(1899, 12, 30) + timedelta(days=int(v))
            if 1 <= d.month <= 6:
                return MESES_ORDEN[d.month - 1]
            return ''
        except Exception:
            return ''
    s = norm(v)
    # nombre completo del mes: 'junio', '12 DE ENERO'
    for m in MESES_ORDEN:
        if m in s:
            return m
    # abreviatura de 3 letras: 'may/26', '01/jun', '24/jun'
    ab = re.search(r'\b([A-Z]{3})\b', s)
    if ab and ab.group(1) in MES_ABREV:
        return MES_ABREV[ab.group(1)]
    # fecha ISO: '2026-01-02 00:00:00'
    num = re.search(r'(?:^|\D)(\d{1,2})(?=\D|$)', s)
    if num and num.group(1).zfill(2) in MES_NUM:
        return MES_NUM[num.group(1).zfill(2)]
    return ''

# =============================================================================
# 1. PLANTILLAS HC -> EQUIPOS
# =============================================================================
import openpyxl

EQUIPOS_RAW = [
    # (archivo, supervisor, nombre_campana)
    ('HC AGOSTO.xlsx',                'MARINEL MORENO',      'RENOVACION'),
]

equipos = {}
for archivo, sup, campana in EQUIPOS_RAW:
    f = glob.glob(os.path.join(PROD, 'extrainfo', archivo))
    if not f:
        print('HC no encontrado:', archivo); continue
    wb = openpyxl.load_workbook(f[0], read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    hdr = {norm(h): i for i, h in enumerate(header or [])}
    col_nombre = get_col(hdr, 'Nombre', 'Nombre del Agente', 'Agente')
    col_sup = get_col(hdr, 'Supervisor')
    col_canal = get_col(hdr, 'Canal')
    miembros = []
    for r in rows:
        if r is None or all(v is None for v in r): continue
        nombre = r[col_nombre] if col_nombre is not None and col_nombre < len(r) else None
        if not nombre:
            continue
        sup_col = r[col_sup] if col_sup is not None and col_sup < len(r) else None
        canal_col = r[col_canal] if col_canal is not None and col_canal < len(r) else None
        # Filtrar por canal: el equipo RENOVACION pertenece al canal SOHO
        if not (canal_col and norm(canal_col).startswith('SOHO')):
            continue
        # El supervisor como persona solo cuenta si además está listado como integrante;
        # en HC AGOSTO Marinel está listada como coordinadora (canal SOHO) -> se incluye
        miembros.append(str(nombre).strip())
    wb.close()
    # Eliminar duplicados conservando el orden
    unicos = sorted(set(miembros), key=lambda x: norm(x))
    equipos[campana] = {
        'supervisor': sup,
        'archivo': archivo,
        'miembros': unicos,
    }
    print(f'Equipo {campana} ({sup}): {len(unicos)} integrantes')

roster = {c: [norm(m) for m in e['miembros']] for c, e in equipos.items()}

# -----------------------------------------------------------------------------
# ASIGNACIONES MANUALES (agentes que no estan en el HC pero pertenecen a un
# equipo, confirmado por la columna Coordinador en 'lineas facturadas'):
#   - Coordinador AMIR JOSE RODRIGUEZ -> equipo RENOVACION (Marinel Moreno)
#   - Coordinador JULIO CESAR ARAUZ   -> equipo SOHO
# -----------------------------------------------------------------------------
OVERRIDES = [
    # (aliases_posibles, campana, nombre_canonico_unificado)
    (['MILENA MOJICA', 'LILIAN MILENA MOJICA DE LAS SALAS'],
     'RENOVACION', 'LILIAN MILENA MOJICA DE LAS SALAS'),
    (['LIZETH GONZALEZ', 'LIZETH MILENA GONZALEZ RODRIGUEZ'],
     'RENOVACION', 'LIZETH MILENA GONZALEZ RODRIGUEZ'),
]

def match_override(nombre, aliases):
    """Match de override igual de estricto que match_member:
    el PRIMER token (nombre de pila) debe estar en el alias y >=2 tokens en común.
    Evita falsos positivos tipo 'DANNA DE LAS SALAS' vs 'LILIAN MILENA MOJICA DE LAS SALAS'."""
    n = norm(nombre)
    if not n: return False
    toks = n.split()
    first = toks[0]
    t = set(toks)
    for al in aliases:
        at = tokens(al)
        if first in at and len(t & at) >= 2:
            return True
    return False

def equipo_de(nombre):
    """Devuelve (campana, miembro_normalizado) para un nombre de agente."""
    n = norm(nombre) if nombre else ''
    # 1) Overrides manuales (match estricto: primer token + >=2 tokens comunes)
    for aliases, camp, canonico in OVERRIDES:
        if match_override(nombre, aliases):
            return camp, norm(canonico)
    # 2) Match contra el roster de cada equipo
    for campana, r in roster.items():
        m = match_member(nombre, r)
        if m:
            return campana, m
    return None, n if n else 'SIN AGENTE'

# =============================================================================
# 2. VICI - MARCACIONES
# =============================================================================
import pyxlsb

xlsb = glob.glob(os.path.join(PROD, 'Reporte_soho_junio_25 (1).xlsb'))[0]
wb = pyxlsb.open_workbook(xlsb)

def read_sheet(sn):
    with wb.get_sheet(sn) as ws:
        header = None
        rows = []
        for r in ws.rows():
            vals = [c.v for c in r]
            if header is None:
                header = vals; continue
            rows.append(vals)
    return header, rows

vici_header, vici_rows = read_sheet('Vici')
vi = {norm(h): i for i, h in enumerate(vici_header)}
c_full = get_col(vi, 'full_name')
c_status = get_col(vi, 'status_name')
c_mes = get_col(vi, 'Mes')
c_phone = get_col(vi, 'phone_number_dialed')
c_seg = get_col(vi, 'length_in_sec')
c_src = get_col(vi, 'source_id', 'lead_id', 'account')
print('Vici:', len(vici_rows), 'llamadas')

# Definición alineada con el consolidado previo: no-contacto = No contesta + Buzón/Contestador
STATUS_NO_CONTACTO = {'NO CONTESTA', 'LLAMADA NO CONTESTA', 'CONTESTADOR AUTOMATICO',
                      'BUZON DE VOZ'}

llamadas = []
for v in vici_rows:
    nombre = v[c_full] if c_full is not None and c_full < len(v) else None
    status = norm(v[c_status]) if c_status is not None and c_status < len(v) else ''
    mes = norm(v[c_mes]) if c_mes is not None and c_mes < len(v) else ''
    phone = str(v[c_phone]).replace('.0','') if c_phone is not None and c_phone < len(v) and v[c_phone] is not None else ''
    seg = to_num(v[c_seg]) if c_seg is not None and c_seg < len(v) else 0
    src = cuenta12(v[c_src]) if c_src is not None and c_src < len(v) and v[c_src] is not None else ''
    campana, miembro = equipo_de(nombre)
    llamadas.append({
        'nombre': norm(nombre), 'campana': campana, 'miembro': miembro,
        'status': status, 'mes': mes, 'phone': phone, 'seg': seg, 'src': src,
    })

def stats_vici(rows):
    total = len(rows)
    no_contesta = sum(1 for r in rows if r['status'] in {'NO CONTESTA','LLAMADA NO CONTESTA'})
    contactados = sum(1 for r in rows if r['status'] and r['status'] not in STATUS_NO_CONTACTO)
    phones = set(r['phone'] for r in rows if r['phone'])
    seg_total = sum(r['seg'] for r in rows)
    meses = Counter(r['mes'] for r in rows)
    return {
        'llamadas': total,
        'no_contesta': no_contesta,
        'contactados': contactados,
        'tasa_contacto': round(contactados / total * 100, 1) if total else 0,
        'numeros_unicos': len(phones),
        'intentos_prom': round(total / len(phones), 2) if phones else 0,
        'duracion_prom_s': round(seg_total / total, 0) if total else 0,
        'por_mes': {m: meses.get(m, 0) for m in MESES_ORDEN},
    }

def vici_por_mes(rows):
    """Métricas de marcación desglosadas por mes (enero-junio)."""
    d = {m: {'llamadas': 0, 'contactados': 0, 'no_contesta': 0, 'numeros': 0,
             'seg': 0, 'duracion_prom_s': 0.0}
         for m in MESES_ORDEN}
    for r in rows:
        m = r['mes'] or ''
        if m not in d: continue
        d[m]['llamadas'] += 1
        d[m]['seg'] += r['seg']
        if r['status'] and r['status'] not in STATUS_NO_CONTACTO:
            d[m]['contactados'] += 1
        if r['status'] in {'NO CONTESTA','LLAMADA NO CONTESTA'}:
            d[m]['no_contesta'] += 1
    for m in MESES_ORDEN:
        d[m]['numeros'] = len(set(r['phone'] for r in rows
                                  if (r['mes'] or '') == m and r['phone']))
        d[m]['duracion_prom_s'] = round(d[m]['seg'] / d[m]['llamadas'], 0) if d[m]['llamadas'] else 0.0
    return d

# Global + por campaña
global_vici = stats_vici(llamadas)
global_por_mes = {m: stats_vici([r for r in llamadas if (r['mes'] or '') == m])
                  for m in MESES_ORDEN}
vici_por_campana = {}
for c in equipos:
    rows_c = [r for r in llamadas if r['campana'] == c]
    vici_por_campana[c] = stats_vici(rows_c)

# Por campaña y mes (para el selector de campaña del dashboard)
vici_campana_por_mes = {}
for c in equipos:
    vici_campana_por_mes[c] = {
        m: stats_vici([r for r in llamadas if r['campana'] == c and (r['mes'] or '') == m])
        for m in MESES_ORDEN
    }

# Por agente: clave (campana, miembro_normalizado) — unifica nombres cortos y completos
vici_por_agente = defaultdict(list)
for r in llamadas:
    vici_por_agente[(r['campana'] or 'SIN EQUIPO', r['miembro'] or 'SIN AGENTE')].append(r)

# Tipificaciones
tipif_global = Counter(r['status'] or '(SIN TIPIFICACION)' for r in llamadas)
tipif_renov = Counter(r['status'] or '(SIN TIPIFICACION)' for r in llamadas if r['campana'] == 'RENOVACION')
tipif_glob_por_mes = {}
tipif_renov_por_mes = {}
for m in MESES_ORDEN:
    tipif_glob_por_mes[m] = Counter(r['status'] or '(SIN TIPIFICACION)'
                                    for r in llamadas if (r['mes'] or '') == m).most_common()
    tipif_renov_por_mes[m] = Counter(r['status'] or '(SIN TIPIFICACION)'
                                     for r in llamadas if r['campana'] == 'RENOVACION' and (r['mes'] or '') == m).most_common()

# =============================================================================
# 3. VENTAS
# =============================================================================
def load_ventas_moviles():
    header, rows = read_sheet('Ventas Moviles')
    idx = {norm(h): i for i, h in enumerate(header)}
    c_ag = get_col(idx, 'Agentes', 'AGENTE', 'Vendedor')
    c_tram = get_col(idx, 'Tramite')
    c_mrc = get_col(idx, 'MRC Actual')
    c_mrc_ant = get_col(idx, 'MRC Anterior')
    c_mes = get_col(idx, 'Mes', 'MES')
    c_fecha = get_col(idx, 'Fecha')
    out = []
    for v in rows:
        ag = v[c_ag] if c_ag is not None and c_ag < len(v) else None
        campana, miembro = equipo_de(ag)
        mes = ''
        if c_mes is not None and c_mes < len(v):
            mes = mes_de(v[c_mes])
        if not mes and c_fecha is not None and c_fecha < len(v):
            mes = mes_de(v[c_fecha])
        out.append({'nombre': norm(ag), 'campana': campana, 'miembro': miembro, 'mes': mes,
                    'tramite': norm(v[c_tram]) if c_tram is not None and c_tram < len(v) else '',
                    'mrc': to_num(v[c_mrc]) if c_mrc is not None and c_mrc < len(v) else 0,
                    'mrc_ant': to_num(v[c_mrc_ant]) if c_mrc_ant is not None and c_mrc_ant < len(v) else 0})
    return out

def load_ventas_fijo():
    header, rows = read_sheet('Ventas Fijo')
    idx = {norm(h): i for i, h in enumerate(header)}
    c_ven = get_col(idx, 'Vendedor', 'VENDEDOR')
    c_rgu = get_col(idx, 'Rgu')
    c_mrc = get_col(idx, 'Mrc')
    c_st = get_col(idx, 'Status Final')
    c_mes = get_col(idx, 'Mes Reg', 'Mes', 'MES')
    out = []
    for v in rows:
        ven = v[c_ven] if c_ven is not None and c_ven < len(v) else None
        campana, miembro = equipo_de(ven)
        mes = mes_de(v[c_mes]) if c_mes is not None and c_mes < len(v) else ''
        out.append({'nombre': norm(ven), 'campana': campana, 'miembro': miembro, 'mes': mes,
                    'rgu': to_num(v[c_rgu]) if c_rgu is not None and c_rgu < len(v) else 0,
                    'mrc': to_num(v[c_mrc]) if c_mrc is not None and c_mrc < len(v) else 0,
                    'status': norm(v[c_st]) if c_st is not None and c_st < len(v) else ''})
    return out

def load_lineas_facturadas():
    header, rows = read_sheet('lineas facturadas')
    idx = {norm(h): i for i, h in enumerate(header)}
    c_ven = get_col(idx, 'Vendedor')
    c_tip = get_col(idx, 'Tipoventa', 'Tipo de Venta')
    c_val = get_col(idx, 'Valor plan')
    c_cant = get_col(idx, 'Cantidad lineas')
    c_conv = get_col(idx, 'Venta convergente')
    c_mes = get_col(idx, 'Fecha fact', 'Fecha Fact', 'Fecha')
    out = []
    for v in rows:
        ven = v[c_ven] if c_ven is not None and c_ven < len(v) else None
        campana, miembro = equipo_de(ven)
        mes = mes_de(v[c_mes]) if c_mes is not None and c_mes < len(v) else ''
        out.append({'nombre': norm(ven), 'campana': campana, 'miembro': miembro, 'mes': mes,
                    'tipoventa': norm(v[c_tip]) if c_tip is not None and c_tip < len(v) else '',
                    'valor': to_num(v[c_val]) if c_val is not None and c_val < len(v) else 0,
                    'cant': to_num(v[c_cant]) if c_cant is not None and c_cant < len(v) else 0,
                    'convergente': norm(v[c_conv]) if c_conv is not None and c_conv < len(v) else ''})
    return out

ventas_moviles = load_ventas_moviles()
ventas_fijo = load_ventas_fijo()
lineas_fact = load_lineas_facturadas()
print('Ventas moviles:', len(ventas_moviles), '| Ventas fijo:', len(ventas_fijo),
      '| Lineas facturadas:', len(lineas_fact))
print('  Ventas moviles por mes:', {m: sum(1 for s in ventas_moviles if s['mes'] == m) for m in MESES_ORDEN})
print('  Ventas fijo por mes:   ', {m: sum(1 for s in ventas_fijo if s['mes'] == m) for m in MESES_ORDEN})
print('  Lineas por mes:        ', {m: sum(1 for s in lineas_fact if s['mes'] == m) for m in MESES_ORDEN})

def ventas_por_agente(sales):
    d = defaultdict(lambda: {'n': 0, 'mrc': 0.0, 'rgu': 0.0, 'valor': 0.0,
                             'tramite': Counter(), 'tipoventa': Counter(), 'status': Counter()})
    for s in sales:
        k = (s['campana'] or 'SIN EQUIPO', s['miembro'] or 'SIN AGENTE')
        a = d[k]
        a['n'] += 1
        a['mrc'] += s.get('mrc', 0) or 0
        # RGU = columna Rgu (Ventas Fijo) o Cantidad lineas (lineas facturadas)
        a['rgu'] += (s.get('rgu', 0) or 0) + (s.get('cant', 0) or 0)
        a['valor'] += s.get('valor', 0) or 0
        if s.get('tramite'): a['tramite'][s['tramite']] += 1
        if s.get('tipoventa'): a['tipoventa'][s['tipoventa']] += 1
        if s.get('status'): a['status'][s['status']] += 1
    return d

def ventas_por_mes_agente():
    """ventas por agente y mes: {clave: {mes: {movil,fijo,lineas,mrc,rgu,tramite,tipoventa}}}"""
    d = defaultdict(lambda: {m: {'movil': 0, 'fijo': 0, 'lineas': 0, 'mrc': 0.0, 'rgu': 0.0,
                                 'tramite': {}, 'tipoventa': {}}
                             for m in MESES_ORDEN})
    for sales, tipo in [(ventas_moviles, 'movil'), (ventas_fijo, 'fijo'), (lineas_fact, 'lineas')]:
        for s in sales:
            k = (s['campana'] or 'SIN EQUIPO', s['miembro'] or 'SIN AGENTE')
            m = s.get('mes') or ''
            if m not in d[k]: continue
            v = d[k][m]
            v[tipo] += 1
            v['mrc'] += (s.get('mrc') or 0) + (s.get('valor') or 0)
            v['rgu'] += (s.get('rgu') or 0) + (s.get('cant') or 0)
            if tipo == 'movil' and s.get('tramite'):
                v['tramite'][s['tramite']] = v['tramite'].get(s['tramite'], 0) + 1
            if tipo == 'lineas' and s.get('tipoventa'):
                v['tipoventa'][s['tipoventa']] = v['tipoventa'].get(s['tipoventa'], 0) + 1
    return d

vm_por = ventas_por_agente(ventas_moviles)
vf_por = ventas_por_agente(ventas_fijo)
lf_por = ventas_por_agente(lineas_fact)
vpm = ventas_por_mes_agente()

# =============================================================================
# 4. VENTA CRUZADA (cross-sell base)
# =============================================================================
cs = glob.glob(os.path.join(PROD, 'bases', 'Base para campaña cross-sell movil Nextphone.xlsx'))
cross_por_agente = Counter()
cross_mes = defaultdict(Counter)   # miembro -> Counter(mes)
if cs:
    cwb = openpyxl.load_workbook(cs[0], read_only=True, data_only=True)
    for sn in ['Base para gestionar CC ', 'con contacto telefonos']:
        if sn not in cwb.sheetnames: continue
        ws = cwb[sn]
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        hdr = {norm(h): i for i, h in enumerate(header or [])}
        col = get_col(hdr, 'Ejecutivo', 'AGENTE', 'Agentes', 'Ejecutiva')
        col_fecha = get_col(hdr, 'FECHA', 'Fecha')
        if col is None: continue
        for r in rows:
            if r is None: continue
            ag = r[col] if col < len(r) else None
            if ag:
                _, miembro = equipo_de(ag)
                cross_por_agente[(norm(ag), miembro)] += 1
                mes = mes_de(r[col_fecha]) if col_fecha is not None and col_fecha < len(r) else ''
                if mes in MESES_ORDEN:
                    cross_mes[miembro][mes] += 1
    cwb.close()
print('Gestiones cross-sell:', sum(cross_por_agente.values()))
print('  Cross-sell por mes:', {m: sum(cm.get(m, 0) for cm in cross_mes.values()) for m in MESES_ORDEN})

# =============================================================================
# 5. ARMAR DATOS POR AGENTE (ranking top performance)
# =============================================================================
agentes = {}
all_keys = set(vici_por_agente.keys()) | set(vm_por.keys()) | set(vf_por.keys()) | set(lf_por.keys())

for k in all_keys:
    campana, miembro = k
    v = vici_por_agente.get(k, [])
    vs = stats_vici(v)
    vm_ag = vici_por_mes(v)
    vm_s = vm_por.get(k)
    vf_s = vf_por.get(k)
    lf_s = lf_por.get(k)
    ventas = (vm_s['n'] if vm_s else 0) + (vf_s['n'] if vf_s else 0) + (lf_s['n'] if lf_s else 0)
    rgu = (vf_s['rgu'] if vf_s else 0) + (lf_s['rgu'] if lf_s else 0)
    mrc_total = (vm_s['mrc'] if vm_s else 0) + (vf_s['mrc'] if vf_s else 0) + (lf_s['valor'] if lf_s else 0)
    cruzada = sum(n for (n_, m_), n in cross_por_agente.items() if m_ == miembro)
    agentes[k] = {
        'agente': miembro,
        'campana': campana,
        'llamadas': vs['llamadas'],
        'no_contesta': vs['no_contesta'],
        'contactados': vs['contactados'],
        'tasa_contacto': vs['tasa_contacto'],
        'numeros': vs['numeros_unicos'],
        'intentos_prom': vs['intentos_prom'],
        'duracion_prom_s': vs['duracion_prom_s'],
        'ventas_movil': vm_s['n'] if vm_s else 0,
        'ventas_fijo': vf_s['n'] if vf_s else 0,
        'lineas_fact': lf_s['n'] if lf_s else 0,
        'ventas_total': ventas,
        'rgu': round(rgu, 1),
        'mrc': round(mrc_total, 2),
        'venta_cruzada': cruzada,
        'tramites': dict(vm_s['tramite']) if vm_s else {},
        'tipoventa': dict(lf_s['tipoventa']) if lf_s else {},
        'status_fijo': dict(vf_s['status']) if vf_s else {},
        'por_mes': vs['por_mes'],
        'contactados_por_mes': {m: vm_ag[m]['contactados'] for m in MESES_ORDEN},
        'no_contesta_por_mes': {m: vm_ag[m]['no_contesta'] for m in MESES_ORDEN},
        'numeros_por_mes': {m: vm_ag[m]['numeros'] for m in MESES_ORDEN},
        'duracion_por_mes': {m: vm_ag[m]['duracion_prom_s'] for m in MESES_ORDEN},
        'ventas_por_mes': vpm[k],
        'cruzada_por_mes': {m: cross_mes[miembro].get(m, 0) for m in MESES_ORDEN},
        # Eficiencia (calculada al final para reflejar merges)
        'ef_ventas_x_llamada': 0.0,
        'ef_conv_pct': 0.0,
        'ef_mrc_x_venta': 0.0,
        'en_roster': False,
        'sin_actividad': False,
    }

# --- Merge de agentes SIN EQUIPO con formatos de nombre distintos ---
# Ej: 'MILENA MOJICA' y 'LILIAN MILENA MOJICA DE LAS SALAS' son la misma persona.
STOP = {'DE', 'DEL', 'LA', 'LAS', 'LOS', 'EL', 'Y', 'SAN', 'SANTA', 'MC', 'MA', 'VON'}

def sig_tokens(nombre):
    return {t for t in norm(nombre).split() if t not in STOP}

sin_keys = [k for k, a in agentes.items() if a['campana'] == 'SIN EQUIPO']
merged = set()
for i, k1 in enumerate(sin_keys):
    if k1 in merged: continue
    t1 = sig_tokens(k1[1])
    grupo = [k1]
    for k2 in sin_keys[i+1:]:
        if k2 in merged: continue
        t2 = sig_tokens(k2[1])
        if len(t1 & t2) >= 2:  # comparten al menos 2 apellidos significativos
            grupo.append(k2)
    if len(grupo) > 1:
        base = agentes[k1]
        for k2 in grupo[1:]:
            b = agentes[k2]
            for campo in ['llamadas','no_contesta','contactados','numeros','ventas_movil',
                          'ventas_fijo','lineas_fact','ventas_total','venta_cruzada']:
                base[campo] += b[campo]
            base['rgu'] += b['rgu']
            base['mrc'] += b['mrc']
            for m in MESES_ORDEN:
                base['por_mes'][m] = base['por_mes'].get(m, 0) + b['por_mes'].get(m, 0)
                base['contactados_por_mes'][m] += b['contactados_por_mes'].get(m, 0)
                base['no_contesta_por_mes'][m] += b['no_contesta_por_mes'].get(m, 0)
                base['numeros_por_mes'][m] += b['numeros_por_mes'].get(m, 0)
                base['cruzada_por_mes'][m] += b['cruzada_por_mes'].get(m, 0)
                # duración: promedio ponderado por llamadas del mes
                llam_bb = b['por_mes'].get(m, 0)
                llam_bb_antes = base['por_mes'].get(m, 0) - llam_bb
                if llam_bb and (llam_bb_antes + llam_bb):
                    base['duracion_por_mes'][m] = round(
                        (base['duracion_por_mes'].get(m, 0) * llam_bb_antes + b['duracion_por_mes'].get(m, 0) * llam_bb)
                        / (llam_bb_antes + llam_bb), 0) if llam_bb_antes else b['duracion_por_mes'].get(m, 0)
                bm = b['ventas_por_mes'].get(m, {})
                for kk in ('movil', 'fijo', 'lineas', 'mrc', 'rgu'):
                    base['ventas_por_mes'][m][kk] += bm.get(kk, 0)
                for kk in ('tramite', 'tipoventa'):
                    for k_, v_ in bm.get(kk, {}).items():
                        base['ventas_por_mes'][m][kk][k_] = base['ventas_por_mes'][m][kk].get(k_, 0) + v_
            for d, d2 in ((base['tramites'], b['tramites']), (base['tipoventa'], b['tipoventa']),
                          (base['status_fijo'], b['status_fijo'])):
                for k_, v_ in d2.items():
                    d[k_] = d.get(k_, 0) + v_
            merged.add(k2)
        # tasa de contacto recalculada
        base['tasa_contacto'] = round(base['contactados'] / base['llamadas'] * 100, 1) if base['llamadas'] else 0
        base['intentos_prom'] = round(base['llamadas'] / base['numeros'], 2) if base['numeros'] else 0
        # usar el nombre más completo del grupo
        base['agente'] = max([agentes[k]['agente'] for k in grupo], key=len)
        # BORRAR todas las claves del grupo y reinsertar con el nombre más completo
        for k in grupo:
            agentes.pop(k, None)
        agentes[(base['campana'], norm(base['agente']))] = base

# --- Metricas de eficiencia (despues de merges) ---
for a in agentes.values():
    a['ef_ventas_x_llamada'] = round(a['ventas_total'] / a['llamadas'], 3) if a['llamadas'] else 0.0
    a['ef_conv_pct'] = round(a['ventas_total'] / a['llamadas'] * 100, 2) if a['llamadas'] else 0.0
    a['ef_mrc_x_venta'] = round(a['mrc'] / a['ventas_total'], 2) if a['ventas_total'] else 0.0

# --- Roster completo POR CAMPAÑA (HC + overrides): los integrantes del roster
# sin actividad aparecen con ceros. Se arma para TODAS las campañas para que el
# selector de campaña del dashboard pueda mostrar cada equipo. ---
def armar_equipo(campana):
    """Devuelve (con_datos, sin_actividad) para una campaña, incluyendo los
    integrantes del roster sin actividad con ceros."""
    over = [o[2] for o in OVERRIDES if o[1] == campana]
    roster_camp = sorted(set(norm(m) for m in equipos[campana]['miembros']) | set(norm(o) for o in over))

    por_key = {k: a for k, a in agentes.items() if a['campana'] == campana}
    mapa = {}
    for k, a in por_key.items():
        mapa[norm(a['agente'])] = a
        mapa[norm(k[1])] = a

    for miembro_norm in roster_camp:
        if miembro_norm in mapa:
            mapa[miembro_norm]['en_roster'] = True
        else:
            k = (campana, miembro_norm)
            agentes[k] = {
                'agente': miembro_norm, 'campana': campana,
                'llamadas': 0, 'no_contesta': 0, 'contactados': 0, 'tasa_contacto': 0.0,
                'numeros': 0, 'intentos_prom': 0.0, 'duracion_prom_s': 0.0,
                'ventas_movil': 0, 'ventas_fijo': 0, 'lineas_fact': 0, 'ventas_total': 0,
                'rgu': 0.0, 'mrc': 0.0, 'venta_cruzada': 0,
                'tramites': {}, 'tipoventa': {}, 'status_fijo': {},
                'por_mes': {m: 0 for m in MESES_ORDEN},
                'contactados_por_mes': {m: 0 for m in MESES_ORDEN},
                'no_contesta_por_mes': {m: 0 for m in MESES_ORDEN},
                'numeros_por_mes': {m: 0 for m in MESES_ORDEN},
                'duracion_por_mes': {m: 0.0 for m in MESES_ORDEN},
                'ventas_por_mes': {m: {'movil':0,'fijo':0,'lineas':0,'mrc':0.0,'rgu':0.0,'tramite':{},'tipoventa':{}} for m in MESES_ORDEN},
                'cruzada_por_mes': {m: 0 for m in MESES_ORDEN},
                'ef_ventas_x_llamada': 0.0, 'ef_conv_pct': 0.0, 'ef_mrc_x_venta': 0.0,
                'en_roster': True,
                'sin_actividad': True,
            }
            mapa[miembro_norm] = agentes[k]

    lista = sorted(mapa.values(),
                   key=lambda x: (not x.get('en_roster', False),
                                  -(x['ventas_total'] + x['mrc']/1000),
                                  -x['llamadas']))
    lista = [a for a in lista if a['llamadas'] > 0 or a['ventas_total'] > 0 or a.get('en_roster', False)]
    con_datos = [a for a in lista if a['llamadas'] > 0 or a['ventas_total'] > 0 or a['venta_cruzada'] > 0]
    sin_actividad = [a for a in lista if a['llamadas'] == 0 and a['ventas_total'] == 0 and a['venta_cruzada'] == 0]
    return con_datos, sin_actividad

campanas_data = {}
for c in equipos:
    campanas_data[c] = dict(zip(['con_datos', 'sin_actividad'], armar_equipo(c)))

renov_con_datos = campanas_data['RENOVACION']['con_datos']
renov_sin_actividad = campanas_data['RENOVACION']['sin_actividad']

sin_eq = sorted([a for k, a in agentes.items() if a['campana'] == 'SIN EQUIPO'],
                key=lambda x: x['llamadas'], reverse=True)

# Totales del equipo RENOVACION por mes (para KPIs y gráfico de evolución)
# Usa renov_con_datos (solo agentes con actividad) para no depender de la
# variable 'renov' que ahora incluye entradas del roster sin actividad.
renov_por_mes = {}
for m in MESES_ORDEN:
    renov_por_mes[m] = {
        'llamadas': sum(a['por_mes'].get(m, 0) for a in renov_con_datos),
        'contactados': sum(a['contactados_por_mes'].get(m, 0) for a in renov_con_datos),
        'no_contesta': sum(a['no_contesta_por_mes'].get(m, 0) for a in renov_con_datos),
        'numeros': sum(a['numeros_por_mes'].get(m, 0) for a in renov_con_datos),
        'ventas_movil': sum(a['ventas_por_mes'][m]['movil'] for a in renov_con_datos),
        'ventas_fijo': sum(a['ventas_por_mes'][m]['fijo'] for a in renov_con_datos),
        'lineas_fact': sum(a['ventas_por_mes'][m]['lineas'] for a in renov_con_datos),
        'mrc': round(sum(a['ventas_por_mes'][m]['mrc'] for a in renov_con_datos), 2),
        'rgu': round(sum(a['ventas_por_mes'][m]['rgu'] for a in renov_con_datos), 1),
        'cruzada': sum(a['cruzada_por_mes'].get(m, 0) for a in renov_con_datos),
    }
    renov_por_mes[m]['tasa_contacto'] = round(
        renov_por_mes[m]['contactados'] / renov_por_mes[m]['llamadas'] * 100, 1) if renov_por_mes[m]['llamadas'] else 0

# =============================================================================
# 7. DATOS PARA EL DASHBOARD
# =============================================================================
data = {
    'generado': datetime.date.today().isoformat(),
    'fuente': 'Reporte_soho_junio_25 (1).xlsb + extrainfo/HC + bases',
    'global_vici': global_vici,
    'global_por_mes': global_por_mes,
    'tipif_global': tipif_global.most_common(),
    'tipif_renovacion': tipif_renov.most_common(),
    'tipif_glob_por_mes': tipif_glob_por_mes,
    'tipif_renov_por_mes': tipif_renov_por_mes,
    'meses': MESES_ORDEN,
    'renov_por_mes': renov_por_mes,
    'equipo_renovacion': renov_con_datos,
    'renov_sin_actividad': renov_sin_actividad,
    'totales_ventas': {
        'moviles': len(ventas_moviles),
        'fijo': len(ventas_fijo),
        'lineas': len(lineas_fact),
        'cross_sell': sum(cross_por_agente.values()),
    },
}

os.makedirs(os.path.join(BASE, 'data'), exist_ok=True)
json_path = os.path.join(BASE, 'data', 'renovacion_data.json')
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=1)
print('JSON generado:', json_path)

# =============================================================================
# 8. GENERAR HTML
# =============================================================================
embed = json.dumps(data, ensure_ascii=False)

html = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Renovación SOHO — Equipo Marinel Moreno</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Arial,sans-serif;background:#0f1220;color:#e8eaf6;padding:16px}
.header{background:linear-gradient(135deg,#7b1fa2,#1a1a2e 60%,#16213e);padding:22px 26px;border-radius:14px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;border:1px solid rgba(255,255,255,.08)}
.header h1{font-size:22px;letter-spacing:.5px}
.header .sub{font-size:12px;opacity:.75;margin-top:4px}
.badge-camp{background:#ffd54f;color:#1a1a2e;font-weight:700;font-size:11px;padding:3px 10px;border-radius:20px;letter-spacing:1px}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px}
.kpi{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:14px 16px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;left:0;top:0;bottom:0;width:4px}
.kpi .lbl{font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#9fa8da}
.kpi .val{font-size:24px;font-weight:800;margin-top:4px}
.kpi .sub{font-size:11px;color:#7986cb;margin-top:2px}
.kpi.c-purple::before{background:#ab47bc}.kpi.c-blue::before{background:#42a5f5}
.kpi.c-green::before{background:#66bb6a}.kpi.c-red::before{background:#ef5350}
.kpi.c-orange::before{background:#ffa726}.kpi.c-teal::before{background:#26a69a}
.tabs{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap}
.tab{padding:9px 22px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;background:rgba(255,255,255,.06);color:#b0bec5;border:1px solid transparent;transition:.2s}
.tab:hover{background:rgba(255,255,255,.12)}
.tab.active{background:#7b1fa2;color:#fff;border-color:#9c27b0}
.content{display:none}.content.active{display:block}
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px}
.card{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:16px}
.card h2{font-size:13px;margin-bottom:12px;color:#c5cae9;border-bottom:1px solid rgba(255,255,255,.08);padding-bottom:8px;letter-spacing:.5px}
.card.full{grid-column:1/-1}
.chart-box{position:relative;height:280px}
.chart-box-sm{position:relative;height:230px}
table{width:100%;border-collapse:collapse;font-size:12px}
th{text-align:left;padding:8px;background:rgba(255,255,255,.06);color:#9fa8da;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.5px;position:sticky;top:0}
td{padding:7px 8px;border-bottom:1px solid rgba(255,255,255,.06)}
tr:hover td{background:rgba(123,31,162,.12)}
.rt{text-align:right}.ct{text-align:center}
.pos{display:inline-flex;width:24px;height:24px;border-radius:50%;align-items:center;justify-content:center;font-weight:800;font-size:11px;background:rgba(255,255,255,.08)}
.pos-1{background:#ffd54f;color:#1a1a2e}.pos-2{background:#b0bec5;color:#1a1a2e}.pos-3{background:#ff8a65;color:#1a1a2e}
.bar{height:8px;border-radius:4px;background:linear-gradient(90deg,#ab47bc,#7b1fa2);min-width:2px}
@media(max-width:900px){.grid-2,.grid-3{grid-template-columns:1fr}}
.note{font-size:11px;color:#7986cb;margin-top:8px}
.filter-bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:14px;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);padding:10px 14px;border-radius:12px}
.mes-btn{padding:6px 14px;border-radius:7px;cursor:pointer;font-size:12px;font-weight:700;background:rgba(255,255,255,.06);color:#b0bec5;border:1px solid rgba(255,255,255,.12);transition:.18s}
.mes-btn:hover{background:rgba(255,255,255,.14);transform:translateY(-1px)}
.mes-btn.active{background:#ffd54f;color:#1a1a2e;border-color:#ffd54f}
.mes-lbl{font-size:11px;font-weight:800;letter-spacing:1px;color:#9fa8da;margin-right:4px}
.period-chip{font-size:12px;font-weight:800;color:#ffd54f;background:rgba(255,213,79,.12);border:1px solid rgba(255,213,79,.3);padding:4px 12px;border-radius:20px;letter-spacing:.5px}
.st-badge{display:inline-flex;align-items:center;gap:4px;font-size:10px;font-weight:800;padding:2px 9px;border-radius:20px;letter-spacing:.4px}
.st-top{background:rgba(255,213,79,.15);color:#ffd54f;border:1px solid rgba(255,213,79,.4)}
.st-act{background:rgba(102,187,106,.12);color:#81c784;border:1px solid rgba(102,187,106,.35)}
.st-low{background:rgba(239,83,80,.12);color:#ef9a9a;border:1px solid rgba(239,83,80,.35)}
.st-warn{background:rgba(255,167,38,.12);color:#ffb74d;border:1px solid rgba(255,167,38,.35)}
.tr-low td{background:rgba(239,83,80,.05)!important;opacity:.55}
.tr-top td{background:rgba(255,213,79,.04)!important}
.ef-bar{display:inline-block;height:8px;border-radius:4px;background:linear-gradient(90deg,#66bb6a,#26a69a);vertical-align:middle}
</style>
</head>
<body>

<div class="header">
  <div>
    <span class="badge-camp">CAMPAÑA RENOVACIÓN</span>
    <h1>Desempeño Equipo Marinel Moreno</h1>
    <div class="sub">SOHO · Iniciativa Nextphone · Marcaciones ViciDial + Ventas · Gerente: Amir Josue Rodriguez Chavarria · Datos: <span id="fechaGen"></span></div>
  </div>
  <div style="text-align:right;font-size:12px;opacity:.8;">
    <div id="totCall"></div>
  </div>
</div>

<div class="filter-bar">
  <span class="mes-lbl">PERIODO:</span>
  <button class="mes-btn active" data-mes="TODOS" onclick="setMes('TODOS')">TODOS</button>
  <button class="mes-btn" data-mes="ENERO" onclick="setMes('ENERO')">ENE</button>
  <button class="mes-btn" data-mes="FEBRERO" onclick="setMes('FEBRERO')">FEB</button>
  <button class="mes-btn" data-mes="MARZO" onclick="setMes('MARZO')">MAR</button>
  <button class="mes-btn" data-mes="ABRIL" onclick="setMes('ABRIL')">ABR</button>
  <button class="mes-btn" data-mes="MAYO" onclick="setMes('MAYO')">MAY</button>
  <button class="mes-btn" data-mes="JUNIO" onclick="setMes('JUNIO')">JUN</button>
  <span style="margin-left:auto;display:flex;gap:6px;align-items:center;">
    <button class="mes-btn" onclick="navMes(-1)" title="Mes anterior">◀</button>
    <span class="period-chip" id="mesActLbl">ENE — JUN 2026</span>
    <button class="mes-btn" onclick="navMes(1)" title="Mes siguiente">▶</button>
  </span>
</div>

<div class="kpi-row" id="kpiRow"></div>

<div class="tabs">
  <button class="tab active" onclick="sw('renov',this)">🏆 Equipo Renovación</button>
  <button class="tab" onclick="sw('rend',this)">📊 Rendimiento del Equipo</button>
  <button class="tab" onclick="sw('marc',this)">📞 Marcaciones (quién llama más/menos)</button>
  <button class="tab" onclick="sw('tipif',this)">🏷️ Tipificación</button>
  <button class="tab" onclick="sw('ventas',this)">💰 Ventas & RGU</button>
</div>


  <div class="card full">
    <h2>🏆 Top Performance — Equipo Renovación <span class="period-chip" id="topPeriod" style="font-size:10px">TODOS LOS MESES</span> <span style="font-weight:400;color:#7986cb;font-size:11px">(clic en columna para ordenar)</span></h2>
    <div style="overflow-x:auto;max-height:520px;overflow-y:auto;">
    <table id="tabRenov"><thead><tr>
      <th>#</th><th>Agente</th><th class="rt">Llamadas</th><th class="rt">Contactados</th><th class="rt">% Contacto</th>
      <th class="rt">Ventas</th><th class="rt">RGU</th><th class="rt">MRC $</th><th class="rt">Venta Cruzada</th><th>Progreso</th>
    </tr></thead><tbody></tbody></table>
    </div>
  </div>
  <div class="grid-2">
    <div class="card"><h2>📈 Marcaciones por Mes — EQUIPO RENOVACIÓN</h2><div class="chart-box-sm"><canvas id="chMesRenov"></canvas></div></div>
    <div class="card"><h2>🔢 Llamadas por Agente</h2><div class="chart-box-sm"><canvas id="chAgLlam"></canvas></div></div>
  </div>
</div>

<div id="v-marc" class="content">
  <div class="grid-3" style="margin-bottom:14px">
    <div class="card"><h2>🔥 Quién llama MÁS (marcaciones)</h2><div class="chart-box-sm"><canvas id="chLlamMas"></canvas></div></div>
    <div class="card"><h2>🌙 Quién llama MENOS</h2><div class="chart-box-sm"><canvas id="chLlamMenos"></canvas></div></div>
    <div class="card"><h2>💬 Contactados vs No Contesta</h2><div class="chart-box-sm"><canvas id="chContVsNc"></canvas></div></div>
  </div>
  <div class="card full">
    <h2>📞 Ranking de Marcaciones — de quien llama más a quien llama menos <span class="period-chip" id="marcPeriod" style="font-size:10px">TODOS LOS MESES</span></h2>
    <div style="overflow-x:auto;max-height:560px;overflow-y:auto;">
    <table id="tabMarc"><thead><tr>
      <th>#</th><th>Agente</th><th class="rt">Llamadas</th><th class="rt">Contactados</th>
      <th class="rt">% Contacto</th><th class="rt">No Contesta</th><th class="rt">Números únicos</th>
      <th class="rt">Intentos/núm</th><th class="rt">Duración prom (s)</th><th>% del total</th>
    </tr></thead><tbody></tbody></table>
    </div>
    <div class="note" id="marcNota"></div>
  </div>
</div>

<div id="v-rend" class="content">
  <div class="grid-3" style="margin-bottom:14px">
    <div class="card"><h2>🥇 Top Vendedores (por ventas + MRC)</h2><div class="chart-box-sm"><canvas id="chTopVen"></canvas></div></div>
    <div class="card"><h2>📉 Eficiencia — Ventas por Llamada</h2><div class="chart-box-sm"><canvas id="chEfVxL"></canvas></div></div>
    <div class="card"><h2>🎯 Conversión — Ventas / Llamadas %</h2><div class="chart-box-sm"><canvas id="chEfConv"></canvas></div></div>
  </div>
  <div class="card full">
    <h2>📊 Ranking Completo del Equipo (incluye integrantes sin actividad) <span class="period-chip" id="rendPeriod" style="font-size:10px">TODOS LOS MESES</span></h2>
    <div style="overflow-x:auto;max-height:560px;overflow-y:auto;">
    <table id="tabRend"><thead><tr>
      <th>#</th><th>Agente</th><th>Estado</th><th class="rt">Llamadas</th><th class="rt">Ventas</th>
      <th class="rt">MRC $</th><th class="rt">RGU</th><th class="rt">Ventas/Llamada</th>
      <th class="rt">% Conversión</th><th class="rt">MRC/Venta $</th>
    </tr></thead><tbody></tbody></table>
    </div>
    <div class="note" id="rendNota"></div>
  </div>
</div>

<div id="v-tipif" class="content">
  <div class="grid-3">
    <div class="card"><h2>Tipificación — Equipo EQUIPO RENOVACIÓN</h2><div class="chart-box-sm"><canvas id="chTipRenov"></canvas></div></div>
    <div class="card"><h2>Tipificación — Global</h2><div class="chart-box-sm"><canvas id="chTipGlob"></canvas></div></div>
    <div class="card"><h2>Métricas de Marcación</h2><div class="chart-box-sm"><canvas id="chMarc"></canvas></div></div>
  </div>
  <div class="card full">
    <h2>Detalle de Tipificaciones (Renovación)</h2>
    <table><thead><tr><th>#</th><th>Tipificación</th><th class="rt">Llamadas</th><th class="rt">% del Total</th><th>Interpretación</th></tr></thead>
    <tbody id="tabTip"></tbody></table>
  </div>
</div>

<div id="v-ventas" class="content">
  <div class="grid-3">
    <div class="card"><h2>Ventas por Agente (Móvil + Fijo + Líneas)</h2><div class="chart-box-sm"><canvas id="chVentas"></canvas></div></div>
    <div class="card"><h2>MRC Generado por Agente</h2><div class="chart-box-sm"><canvas id="chMrc"></canvas></div></div>
    <div class="card"><h2>RGU por Agente</h2><div class="chart-box-sm"><canvas id="chRgu"></canvas></div></div>
  </div>
  <div class="grid-2">
    <div class="card"><h2>Ventas Móviles por Trámite</h2><div class="chart-box-sm"><canvas id="chTramite"></canvas></div></div>
    <div class="card"><h2>Líneas Facturadas por Tipo de Venta</h2><div class="chart-box-sm"><canvas id="chTipoventa"></canvas></div></div>
  </div>
</div>


const D = __DATA__;
const fmt$ = v => '$' + Number(v).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});
const COLORS = ['#ab47bc','#42a5f5','#66bb6a','#ffa726','#ef5350','#26a69a','#ec407a','#5c6bc0','#ffca28','#8d6e63'];
const MESES = D.meses;
const charts = {};

document.getElementById('fechaGen').textContent = D.generado;

function fmtCtx(c){
  const el = document.getElementById(c);
  return el ? el.getContext('2d') : null;
}

function nuevoChart(id, cfg){
  const c = fmtCtx(id); if(!c) return;
  if(charts[id]) charts[id].destroy();
  charts[id] = new Chart(c, cfg);
}

let mesAct = 'TODOS';
function campData(){
  return {con: D.equipo_renovacion||[], sin: D.renov_sin_actividad||[]};
}

// --- Filtrado por mes ---
function agFiltrado(a){
  if(mesAct === 'TODOS'){
    return Object.assign({}, a);
  }
  const m = mesAct;
  const pm = a.por_mes[m]||0, cm = a.contactados_por_mes[m]||0;
  const nm = a.no_contesta_por_mes[m]||0, num = a.numeros_por_mes[m]||0;
  const vp = a.ventas_por_mes[m]||{};
  const mov = vp.movil||0, fij = vp.fijo||0, lin = vp.lineas||0;
  const vt = mov+fij+lin, mrc = vp.mrc||0, rgu = vp.rgu||0;
  return {
    agente: a.agente, campana: a.campana,
    llamadas: pm, contactados: cm, no_contesta: nm, numeros: num,
    tasa_contacto: pm ? +(cm/pm*100).toFixed(1) : 0,
    ventas_movil: mov, ventas_fijo: fij, lineas_fact: lin,
    ventas_total: vt, mrc: mrc, rgu: rgu,
    venta_cruzada: a.cruzada_por_mes[m]||0,
    tramites: vp.tramite||{}, tipoventa: vp.tipoventa||{},
    ef_ventas_x_llamada: pm ? +(vt/pm).toFixed(3) : 0,
    ef_conv_pct: pm ? +(vt/pm*100).toFixed(2) : 0,
    ef_mrc_x_venta: vt ? +(mrc/vt).toFixed(2) : 0,
    intentos_prom: num ? +(pm/num).toFixed(2) : 0,
    duracion_prom_s: (a.duracion_por_mes&&a.duracion_por_mes[m]) ? a.duracion_por_mes[m] : 0,
    en_roster: !!a.en_roster, sin_actividad: !!a.sin_actividad
  };
}
function agentesActuales(){
  return campData().con.map(agFiltrado);
}
// Roster completo: con datos + los del HC sin actividad (ceros)
function rosterActual(){
  const conDatos = agentesActuales();
  const sinAct = campData().sin.map(agFiltrado);
  return conDatos.concat(sinAct);
}
function tipifRenovActual(){
  return mesAct==='TODOS' ? (D.tipif_renovacion||[]) : ((D.tipif_renov_por_mes||{})[mesAct]||[]);
}
function tipifGlobActual(){
  return mesAct==='TODOS' ? D.tipif_global : (D.tipif_glob_por_mes[mesAct]||[]);
}
function mesLabel(){
  return mesAct==='TODOS' ? 'ENE — JUN 2026' : mesAct;
}

function setMes(m){
  mesAct = m;
  document.querySelectorAll('.mes-btn').forEach(b=>b.classList.toggle('active', b.getAttribute('data-mes')===m));
  renderAll();
}
function navMes(d){
  if(mesAct === 'TODOS'){
    setMes(d>0 ? MESES[0] : MESES[MESES.length-1]);
    return;
  }
  const i = MESES.indexOf(mesAct);
  const n = (i + d + MESES.length) % MESES.length;
  setMes(MESES[n]);
}

function kpis(){
  const g = mesAct==='TODOS' ? D.global_vici : D.global_por_mes[mesAct];
  const r = agentesActuales();
  const sum = f => r.reduce((a,x)=>a+f(x),0);
  const llamR = sum(x=>x.llamadas), contR = sum(x=>x.contactados);
  const ventasR = sum(x=>x.ventas_total), mrcR = sum(x=>x.mrc);
  const rguR = sum(x=>x.rgu), xcR = sum(x=>x.venta_cruzada);
  const lbl = 'RENOVACIÓN';
  const vsub = mesAct==='TODOS'
    ? D.totales_ventas.moviles+' móv. · '+D.totales_ventas.fijo+' fijo · '+D.totales_ventas.lineas+' líneas'
    : sum(x=>x.ventas_movil)+' móv. · '+sum(x=>x.ventas_fijo)+' fijo · '+sum(x=>x.lineas_fact)+' líneas';
  const arr = [
    ['Llamadas '+lbl, llamR.toLocaleString(), g.llamadas? (llamR/g.llamadas*100).toFixed(1)+'% del global ('+g.llamadas.toLocaleString()+')':'', 'purple'],
    ['Contactados', contR.toLocaleString(), llamR? (contR/llamR*100).toFixed(1)+'% tasa contacto':'-', 'green'],
    ['Ventas Subidas', ventasR, vsub, 'blue'],
    ['MRC Generado', fmt$(mrcR), 'móvil + fijo + líneas', 'orange'],
    ['RGU Total', rguR.toLocaleString(), 'servicios vendidos', 'teal'],
    ['Venta Cruzada', xcR, 'gestiones cross-sell', 'red'],
  ];
  document.getElementById('kpiRow').innerHTML = arr.map(a=>
    `<div class="kpi c-${a[3]}"><div class="lbl">${a[0]}</div><div class="val">${a[1]}</div><div class="sub">${a[2]}</div></div>`
  ).join('');
  document.getElementById('totCall').innerHTML = `<strong>${g.llamadas.toLocaleString()}</strong> llamadas ${lbl.toLowerCase()} · ${g.numeros_unicos.toLocaleString()} números únicos`;
  document.getElementById('mesActLbl').textContent = mesLabel();
  document.getElementById('topPeriod').textContent = mesAct==='TODOS' ? 'TODOS LOS MESES' : mesAct;
}

function sw(id, btn){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.content').forEach(c=>c.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('v-'+id).classList.add('active');
}

// ---------- RENOVACION ----------
let sortKey = 'llamadas', sortDir = -1;
function renderRenov(){
  const rows = agentesActuales();
  rows.sort((a,b)=> (sortKey==='agente' ? String(a[sortKey]).localeCompare(String(b[sortKey])) : a[sortKey]-b[sortKey])*sortDir);
  const maxL = Math.max(...rows.map(r=>r.llamadas),1);
  document.getElementById('tabRenov').querySelector('tbody').innerHTML = rows.map((r,i)=>{
    const pct = (r.llamadas/maxL*100).toFixed(0);
    return `<tr>
      <td><span class="pos pos-${i+1}">${i+1}</span></td>
      <td><strong>${r.agente}</strong></td>
      <td class="rt">${r.llamadas.toLocaleString()}</td>
      <td class="rt">${r.contactados.toLocaleString()}</td>
      <td class="rt">${r.tasa_contacto}%</td>
      <td class="rt"><strong>${r.ventas_total}</strong></td>
      <td class="rt">${r.rgu.toLocaleString()}</td>
      <td class="rt">${fmt$(r.mrc)}</td>
      <td class="rt">${r.venta_cruzada}</td>
      <td style="min-width:140px"><div class="bar" style="width:${pct}%"></div><span style="font-size:10px;color:#7986cb">${pct}% de máx</span></td>
    </tr>`;
  }).join('');
}
document.querySelector('#tabRenov thead').addEventListener('click', e=>{
  const th = e.target.closest('th'); if(!th) return;
  const map = {0:'agente',1:'agente',2:'llamadas',3:'contactados',4:'tasa_contacto',5:'ventas_total',6:'rgu',7:'mrc',8:'venta_cruzada'};
  const idx = Array.from(th.parentNode.children).indexOf(th);
  const k = map[idx]; if(!k) return;
  if(sortKey===k) sortDir*=-1; else {sortKey=k; sortDir = (k==='agente')?1:-1;}
  renderRenov();
});

function renderMesRenov(){
  const data = MESES.map(m=> (D.renov_por_mes[m]||{}).llamadas||0);
  const bg = MESES.map(m=> (mesAct==='TODOS'||mesAct===m) ? '#ab47bc' : 'rgba(171,71,188,.25)');
  nuevoChart('chMesRenov',{type:'bar',data:{labels:MESES,datasets:[{label:'Llamadas',data,backgroundColor:bg,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{y:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},x:{ticks:{color:'#9fa8da'}}}}});
}
function renderAgLlam(){
  const top = agentesActuales().sort((a,b)=>b.llamadas-a.llamadas).slice(0,10);
  nuevoChart('chAgLlam',{type:'bar',data:{labels:top.map(r=>r.agente),datasets:[{label:'Llamadas',data:top.map(r=>r.llamadas),backgroundColor:COLORS,borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{x:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},y:{ticks:{color:'#9fa8da',font:{size:10}}}}}});
}

// ---------- TIPIFICACION ----------
function renderTipif(){
  const tRen = tipifRenovActual();
  const c1 = fmtCtx('chTipRenov'); if(c1){
    const t = tRen.slice(0,10);
    const otros = tRen.slice(10).reduce((s,x)=>s+x[1],0);
    const labs = t.map(x=>x[0]), vals = t.map(x=>x[1]);
    if(otros>0){labs.push('Otros');vals.push(otros);}
    nuevoChart('chTipRenov',{type:'doughnut',data:{labels:labs,datasets:[{data:vals,backgroundColor:COLORS,borderWidth:0}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{color:'#9fa8da',boxWidth:10,font:{size:10}}}}}});
  }
  const c2 = fmtCtx('chTipGlob'); if(c2){
    const tGlob = tipifGlobActual();
    const t = tGlob.slice(0,10);
    const otros = tGlob.slice(10).reduce((s,x)=>s+x[1],0);
    const labs = t.map(x=>x[0]), vals = t.map(x=>x[1]);
    if(otros>0){labs.push('Otros');vals.push(otros);}
    nuevoChart('chTipGlob',{type:'bar',data:{labels:labs.map(l=>l.length>18?l.slice(0,18)+'…':l),datasets:[{label:'Llamadas',data:vals,backgroundColor:'#42a5f5',borderRadius:4}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{x:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},y:{ticks:{color:'#9fa8da',font:{size:9}}}}}});
  }
  const c3 = fmtCtx('chMarc'); if(c3){
    const r = agentesActuales();
    nuevoChart('chMarc',{type:'bar',data:{labels:['Llamadas','Contactados','No Contesta','Números Únicos'],
      datasets:[{label:'',data:[r.reduce((a,x)=>a+x.llamadas,0),r.reduce((a,x)=>a+x.contactados,0),r.reduce((a,x)=>a+x.no_contesta,0),r.reduce((a,x)=>a+x.numeros,0)],backgroundColor:['#ab47bc','#66bb6a','#ef5350','#42a5f5'],borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{y:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},x:{ticks:{color:'#9fa8da'}}}}});
  }
  const total = tRen.reduce((s,x)=>s+x[1],0) || 1;
  const interp = s=>{
    if(/NO CONTESTA|LLAMADA NO CONTESTA/.test(s)) return 'Cliente no respondió la llamada';
    if(/BUZON|CONTESTADOR/.test(s)) return 'Contestador / buzón de voz';
    if(/SEGUIMIENTO|VOLVER A LLAMAR|CALL BACK/.test(s)) return 'Requiere seguimiento / rellamada';
    if(/PROPUESTA/.test(s)) return 'Se envió propuesta al cliente';
    if(/NO ACEPTA|RECHAZA|NO INTERESADO/.test(s)) return 'Cliente rechazó la oferta';
    if(/VENTA|RENOVO|APROBACION/.test(s)) return 'Resultado positivo / venta';
    if(/TITULAR|TOMADOR/.test(s)) return 'No es la persona responsable';
    if(/CUELGA/.test(s)) return 'Cliente colgó';
    if(/DEUDA/.test(s)) return 'Cliente con deuda pendiente';
    if(/OCUPADO|BUSY|MUDA|LEAD|OUTBOUND/.test(s)) return 'Sin contacto efectivo';
    return 'Gestión completada';
  };
  document.getElementById('tabTip').innerHTML = tRen.map((x,i)=>{
    const p = (x[1]/total*100).toFixed(1);
    return `<tr><td>${i+1}</td><td><strong>${x[0]}</strong></td><td class="rt">${x[1].toLocaleString()}</td><td class="rt">${p}%</td><td style="font-size:11px;color:#9fa8da">${interp(x[0])}</td></tr>`;
  }).join('');
}

// ---------- VENTAS ----------
function renderVentas(){
  const r = agentesActuales().sort((a,b)=>b.ventas_total-a.ventas_total).slice(0,10);
  const c1 = fmtCtx('chVentas'); if(c1){
    nuevoChart('chVentas',{
      type:'bar',
      data:{
        labels: r.map(x=>x.agente),
        datasets:[
          {label:'Móvil', data:r.map(x=>x.ventas_movil), backgroundColor:'#42a5f5', borderRadius:3},
          {label:'Fijo', data:r.map(x=>x.ventas_fijo), backgroundColor:'#66bb6a', borderRadius:3},
          {label:'Líneas', data:r.map(x=>x.lineas_fact), backgroundColor:'#ffa726', borderRadius:3}
        ]
      },
      options:{
        responsive:true,
        maintainAspectRatio:false,
        plugins:{legend:{position:'top', labels:{color:'#9fa8da', boxWidth:10}}},
        scales:{
          y:{beginAtZero:true, stacked:true, grid:{color:'rgba(255,255,255,.06)'}, ticks:{color:'#9fa8da'}},
          x:{stacked:true, ticks:{color:'#9fa8da', font:{size:9}}}
        }
      }
    });
  }
  const c2 = fmtCtx('chMrc'); if(c2){
    nuevoChart('chMrc',{type:'bar',data:{labels:r.map(x=>x.agente),datasets:[{label:'MRC $',data:r.map(x=>x.mrc),backgroundColor:'#ab47bc',borderRadius:4}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{x:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},y:{ticks:{color:'#9fa8da',font:{size:9}}}}}});
  }
  const c3 = fmtCtx('chRgu'); if(c3){
    nuevoChart('chRgu',{type:'bar',data:{labels:r.map(x=>x.agente),datasets:[{label:'RGU',data:r.map(x=>x.rgu),backgroundColor:'#26a69a',borderRadius:4}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{x:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},y:{ticks:{color:'#9fa8da',font:{size:9}}}}}});
  }
  const c4 = fmtCtx('chTramite'); if(c4){
    const agg = {};
    agentesActuales().forEach(x=>{Object.entries(x.tramites||{}).forEach(([k,v])=>agg[k]=(agg[k]||0)+v);});
    const e = Object.entries(agg).sort((a,b)=>b[1]-a[1]);
    nuevoChart('chTramite',{type:'doughnut',data:{labels:e.map(x=>x[0]),datasets:[{data:e.map(x=>x[1]),backgroundColor:COLORS,borderWidth:0}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{color:'#9fa8da',boxWidth:10,font:{size:10}}}}}});
  }
  const c5 = fmtCtx('chTipoventa'); if(c5){
    const agg = {};
    agentesActuales().forEach(x=>{Object.entries(x.tipoventa||{}).forEach(([k,v])=>agg[k]=(agg[k]||0)+v);});
    const e = Object.entries(agg).sort((a,b)=>b[1]-a[1]);
    nuevoChart('chTipoventa',{type:'doughnut',data:{labels:e.map(x=>x[0]),datasets:[{data:e.map(x=>x[1]),backgroundColor:COLORS,borderWidth:0}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{color:'#9fa8da',boxWidth:10,font:{size:10}}}}}});
  }
}

// ---------- RENDIMIENTO ----------
function renderRend(){
  const roster = rosterActual();
  // Ranking: por ventas+MRC, luego llamadas
  const ranked = roster.slice().sort((a,b)=> (b.ventas_total + b.mrc/1000) - (a.ventas_total + a.mrc/1000) || b.llamadas - a.llamadas);
  // Top 3 vendedores
  const top3 = ranked.filter(x=>x.ventas_total>0).slice(0,3).map(x=>x.agente);
  // Grafico 1: Top vendedores
  const c1 = fmtCtx('chTopVen'); if(c1){
    const t = ranked.filter(x=>x.ventas_total>0).slice(0,8);
    if(!t.length){ nuevoChart('chTopVen',{type:'bar',data:{labels:['Sin ventas en este período'],datasets:[{data:[0],backgroundColor:'#ef5350'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},x:{ticks:{color:'#9fa8da'}}}}}); }
    else
    nuevoChart('chTopVen',{type:'bar',data:{labels:t.map(x=>x.agente.length>16?x.agente.slice(0,16)+'…':x.agente),
      datasets:[{label:'Ventas',data:t.map(x=>x.ventas_total),backgroundColor:COLORS,borderRadius:4},
                {label:'MRC $',data:t.map(x=>x.mrc),backgroundColor:'#42a5f5',borderRadius:4}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{color:'#9fa8da',boxWidth:10,font:{size:9}}}},
        scales:{x:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},y:{ticks:{color:'#9fa8da',font:{size:9}}}}}});
  }
  // Grafico 2: eficiencia ventas/llamada
  const c2 = fmtCtx('chEfVxL'); if(c2){
    const t = ranked.filter(x=>x.llamadas>0).slice(0,8);
    if(!t.length){ nuevoChart('chEfVxL',{type:'bar',data:{labels:['Sin llamadas en este período'],datasets:[{data:[0],backgroundColor:'#ef5350'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},x:{ticks:{color:'#9fa8da'}}}}}); }
    else
    nuevoChart('chEfVxL',{type:'bar',data:{labels:t.map(x=>x.agente.length>16?x.agente.slice(0,16)+'…':x.agente),
      datasets:[{label:'Ventas por llamada',data:t.map(x=>x.ef_ventas_x_llamada),backgroundColor:'#26a69a',borderRadius:4}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{x:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},y:{ticks:{color:'#9fa8da',font:{size:9}}}}}});
  }
  // Grafico 3: conversion %
  const c3 = fmtCtx('chEfConv'); if(c3){
    const t = ranked.filter(x=>x.llamadas>0).slice(0,8);
    if(!t.length){ nuevoChart('chEfConv',{type:'bar',data:{labels:['Sin llamadas en este período'],datasets:[{data:[0],backgroundColor:'#ef5350'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},x:{ticks:{color:'#9fa8da'}}}}}); }
    else
    nuevoChart('chEfConv',{type:'bar',data:{labels:t.map(x=>x.agente.length>16?x.agente.slice(0,16)+'…':x.agente),
      datasets:[{label:'% Conversión',data:t.map(x=>x.ef_conv_pct),backgroundColor:'#66bb6a',borderRadius:4}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{x:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},y:{ticks:{color:'#9fa8da',font:{size:9}}}}}});
  }
  // Tabla ranking completo
  const maxV = Math.max(...ranked.map(x=>x.ventas_total),1);
  document.getElementById('tabRend').querySelector('tbody').innerHTML = ranked.map((r,i)=>{
    let est, cls = '', badge;
    if(r.ventas_total === 0 && r.llamadas === 0){
      est = 'Sin actividad'; cls = 'tr-low'; badge = '<span class="st-badge st-low">⚠ SIN ACTIVIDAD</span>';
    } else if(top3.includes(r.agente)){
      est = 'Top vendedor'; cls = 'tr-top'; badge = '<span class="st-badge st-top">🥇 TOP '+(top3.indexOf(r.agente)+1)+'</span>';
    } else if(r.ef_conv_pct >= 2){
      est = 'Activo'; badge = '<span class="st-badge st-act">✓ ACTIVO</span>';
    } else {
      est = 'Bajo rendimiento'; badge = '<span class="st-badge st-warn">⚠ BAJO</span>';
    }
    const pct = (r.ventas_total/maxV*100).toFixed(0);
    return `<tr class="${cls}">
      <td><span class="pos">${i+1}</span></td>
      <td><strong>${r.agente}</strong></td>
      <td>${badge}</td>
      <td class="rt">${r.llamadas.toLocaleString()}</td>
      <td class="rt"><strong>${r.ventas_total}</strong></td>
      <td class="rt">${fmt$(r.mrc)}</td>
      <td class="rt">${r.rgu.toLocaleString()}</td>
      <td class="rt">${r.ef_ventas_x_llamada? r.ef_ventas_x_llamada.toFixed(3) : '—'}</td>
      <td class="rt">${r.ef_conv_pct? r.ef_conv_pct.toFixed(2)+'%' : '—'}</td>
      <td class="rt">${r.ef_mrc_x_venta? fmt$(r.ef_mrc_x_venta) : '—'}</td>
    </tr>`;
  }).join('');
  const activos = ranked.filter(x=>x.ventas_total>0||x.llamadas>0).length;
  const sinAct = ranked.length - activos;
  document.getElementById('rendNota').textContent =
    `${ranked.length} integrantes en total · ${activos} con actividad · ${sinAct} sin actividad · Los TOP son los 3 con más ventas+MRC · 'Sin actividad' = integrantes del HC sin llamadas ni ventas en el período.`;
  document.getElementById('rendPeriod').textContent = mesAct==='TODOS' ? 'TODOS LOS MESES' : mesAct;
}

// ---------- MARCACIONES (quien llama mas / menos) ----------
function renderMarc(){
  const roster = rosterActual();
  const activos = roster.filter(x=>x.llamadas>0).sort((a,b)=>b.llamadas-a.llamadas);
  const totalLlam = activos.reduce((s,x)=>s+x.llamadas,0) || 1;
  // Grafico 1: quien llama mas (top 8)
  const c1 = fmtCtx('chLlamMas'); if(c1){
    const t = activos.slice(0,8);
    if(!t.length){ nuevoChart('chLlamMas',{type:'bar',data:{labels:['Sin llamadas'],datasets:[{data:[0],backgroundColor:'#ef5350'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},x:{ticks:{color:'#9fa8da'}}}}}); }
    else
    nuevoChart('chLlamMas',{type:'bar',data:{labels:t.map(x=>x.agente.length>15?x.agente.slice(0,15)+'…':x.agente),
      datasets:[{label:'Llamadas',data:t.map(x=>x.llamadas),backgroundColor:'#42a5f5',borderRadius:4}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{x:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},y:{ticks:{color:'#9fa8da',font:{size:9}}}}}});
  }
  // Grafico 2: quien llama menos (bottom 8 con actividad)
  const c2 = fmtCtx('chLlamMenos'); if(c2){
    const t = activos.slice(-8).reverse();
    if(!t.length){ nuevoChart('chLlamMenos',{type:'bar',data:{labels:['Sin llamadas'],datasets:[{data:[0],backgroundColor:'#ef5350'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},x:{ticks:{color:'#9fa8da'}}}}}); }
    else
    nuevoChart('chLlamMenos',{type:'bar',data:{labels:t.map(x=>x.agente.length>15?x.agente.slice(0,15)+'…':x.agente),
      datasets:[{label:'Llamadas',data:t.map(x=>x.llamadas),backgroundColor:'#ffa726',borderRadius:4}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{x:{beginAtZero:true,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},y:{ticks:{color:'#9fa8da',font:{size:9}}}}}});
  }
  // Grafico 3: contactados vs no contesta
  const c3 = fmtCtx('chContVsNc'); if(c3){
    const t = activos.slice(0,8);
    if(!t.length){ nuevoChart('chContVsNc',{type:'bar',data:{labels:['Sin datos'],datasets:[{data:[0]},{data:[0]}]},options:{responsive:true,maintainAspectRatio:false}}); }
    else
    nuevoChart('chContVsNc',{type:'bar',data:{labels:t.map(x=>x.agente.length>15?x.agente.slice(0,15)+'…':x.agente),
      datasets:[
        {label:'Contactados',data:t.map(x=>x.contactados),backgroundColor:'#66bb6a',borderRadius:3},
        {label:'No contesta',data:t.map(x=>x.no_contesta),backgroundColor:'#ef5350',borderRadius:3}
      ]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{color:'#9fa8da',boxWidth:10,font:{size:9}}}},
        scales:{y:{beginAtZero:true,stacked:false,grid:{color:'rgba(255,255,255,.06)'},ticks:{color:'#9fa8da'}},x:{ticks:{color:'#9fa8da',font:{size:9}}}}}});
  }
  // Tabla ranking completo (incluye sin actividad al final)
  const todos = roster.slice().sort((a,b)=>b.llamadas-a.llamadas);
  document.getElementById('tabMarc').querySelector('tbody').innerHTML = todos.map((r,i)=>{
    const pctT = (r.llamadas/totalLlam*100).toFixed(1);
    const sinAct = r.llamadas===0 && r.ventas_total===0;
    const cls = sinAct ? 'tr-low' : '';
    const badge = sinAct ? '<span class="st-badge st-low">⚠ SIN ACTIVIDAD</span>' : '';
    return `<tr class="${cls}">
      <td><span class="pos">${i+1}</span></td>
      <td><strong>${r.agente}</strong> ${badge}</td>
      <td class="rt"><strong>${r.llamadas.toLocaleString()}</strong></td>
      <td class="rt">${r.contactados.toLocaleString()}</td>
      <td class="rt">${r.tasa_contacto}%</td>
      <td class="rt">${r.no_contesta.toLocaleString()}</td>
      <td class="rt">${r.numeros.toLocaleString()}</td>
      <td class="rt">${r.intentos_prom? r.intentos_prom.toFixed(2) : '—'}</td>
      <td class="rt">${r.duracion_prom_s? Math.round(r.duracion_prom_s)+' s' : '—'}</td>
      <td style="min-width:120px"><div class="bar" style="width:${Math.min(pctT*3,100)}%"></div><span style="font-size:10px;color:#7986cb">${pctT}%</span></td>
    </tr>`;
  }).join('');
  const sinActivos = todos.filter(x=>x.llamadas===0).length;
  document.getElementById('marcNota').textContent =
    `${activos.length} agentes con marcaciones · ${totalLlam.toLocaleString()} llamadas en total · ${sinActivos} integrantes sin ninguna llamada · Ordenado de quien más llama a quien menos.`;
  document.getElementById('marcPeriod').textContent = mesAct==='TODOS' ? 'TODOS LOS MESES' : mesAct;
}

function renderAll(){
  kpis();
  renderRenov(); renderMesRenov(); renderAgLlam();
  renderRend();
  renderMarc();
  renderTipif();
  renderVentas();
}

renderAll();
</script>
</body>
</html>
"""

html = html.replace('__DATA__', embed)
with open(os.path.join(BASE, 'index.html'), 'w', encoding='utf-8') as f:
    f.write(html)
print('Dashboard generado:', os.path.join(BASE, 'index.html'))
