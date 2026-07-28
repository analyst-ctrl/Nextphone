#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EXPORT_JSON.PY - Exporta SQLite a JSON para dashboard standalone (GitHub Pages)
================================================================================
Lee la base de datos calidad.db y genera un archivo JSON con todos los datos
necesarios para el dashboard HTML/JS/CSS (sin PHP).

Uso:
    python export_json.py

Genera:
    ../data/calidad_data.json   (datos para el dashboard)
"""

import sqlite3
import json
import os
import sys

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'calidad.db')
JSON_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'calidad_data.json')

def query_dict(db, sql, params=None):
    cur = db.execute(sql, params or [])
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]

def main():
    print(f"📂 Leyendo DB: {DB_PATH}")
    
    if not os.path.exists(DB_PATH):
        print(f"❌ No se encuentra {DB_PATH}")
        print("   Ejecuta primero: python scripts/extract_data.py")
        sys.exit(1)
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    
    # 1. Resumen general
    resumen = query_dict(conn, """
        SELECT 
            COUNT(*) as total_evaluaciones,
            COUNT(DISTINCT agente) as total_agentes,
            COUNT(DISTINCT supervisor) as total_supervisores,
            ROUND(AVG(score), 1) as score_promedio,
            ROUND(MIN(score), 1) as score_minimo,
            ROUND(MAX(score), 1) as score_maximo,
            SUM(es_critica) as total_criticas,
            ROUND(SUM(es_critica) * 100.0 / COUNT(*), 1) as pct_criticas
        FROM evaluaciones
    """)[0]
    
    # 2. Componentes promedio
    componentes = query_dict(conn, """
        SELECT 
            ROUND(AVG(penc), 1) as penc,
            ROUND(AVG(pecuf), 1) as pecuf,
            ROUND(AVG(pecne), 1) as pecne,
            ROUND(AVG(peccum), 1) as peccum
        FROM evaluaciones
    """)[0]
    
    # 3. Supervisores
    supervisores = query_dict(conn, """
        SELECT 
            supervisor,
            COUNT(*) as evaluaciones,
            COUNT(DISTINCT agente) as agentes,
            ROUND(AVG(score), 1) as score_promedio,
            ROUND(AVG(penc), 1) as penc,
            ROUND(AVG(pecuf), 1) as pecuf,
            ROUND(AVG(pecne), 1) as pecne,
            ROUND(AVG(peccum), 1) as peccum,
            SUM(es_critica) as criticas,
            ROUND(SUM(es_critica) * 100.0 / COUNT(*), 1) as pct_criticas
        FROM evaluaciones
        WHERE supervisor != ''
        GROUP BY supervisor
        ORDER BY score_promedio DESC
    """)
    
    # 4. Agentes
    agentes = query_dict(conn, """
        SELECT 
            agente,
            supervisor,
            COUNT(*) as evaluaciones,
            ROUND(AVG(score), 1) as score_promedio,
            ROUND(AVG(penc), 1) as penc,
            ROUND(AVG(pecuf), 1) as pecuf,
            ROUND(AVG(pecne), 1) as pecne,
            ROUND(AVG(peccum), 1) as peccum,
            SUM(es_critica) as criticas
        FROM evaluaciones
        WHERE agente != ''
        GROUP BY agente
        ORDER BY score_promedio DESC
    """)
    
    # Cuartiles para agentes
    scores = sorted([a['score_promedio'] for a in agentes])
    n = len(scores)
    q1 = scores[int(n * 0.25)] if n > 0 else 0
    q2 = scores[int(n * 0.50)] if n > 0 else 0
    q3 = scores[int(n * 0.75)] if n > 0 else 0
    
    for a in agentes:
        s = a['score_promedio']
        if s >= q3: a['cuartil'] = 'Q1'
        elif s >= q2: a['cuartil'] = 'Q2'
        elif s >= q1: a['cuartil'] = 'Q3'
        else: a['cuartil'] = 'Q4'
        a['cumple'] = 'Cumple' if s >= 85 else 'No Cumple'
    
    # 5. Campañas (Nextphone vs OJT)
    campanas = query_dict(conn, """
        SELECT 
            proyecto as campana,
            COUNT(*) as evaluaciones,
            COUNT(DISTINCT agente) as agentes,
            ROUND(AVG(score), 1) as score_promedio,
            ROUND(AVG(penc), 1) as penc,
            ROUND(AVG(pecuf), 1) as pecuf,
            ROUND(AVG(pecne), 1) as pecne,
            ROUND(AVG(peccum), 1) as peccum,
            SUM(es_critica) as criticas
        FROM evaluaciones
        GROUP BY proyecto
        ORDER BY score_promedio DESC
    """)
    
    # 6. Histórico mensual
    historico = query_dict(conn, """
        SELECT 
            mes,
            mes_texto,
            COUNT(*) as evaluaciones,
            COUNT(DISTINCT agente) as agentes,
            ROUND(AVG(score), 1) as score_promedio,
            ROUND(AVG(penc), 1) as penc,
            ROUND(AVG(pecuf), 1) as pecuf,
            ROUND(AVG(pecne), 1) as pecne,
            ROUND(AVG(peccum), 1) as peccum,
            SUM(es_critica) as criticas
        FROM evaluaciones
        WHERE mes != ''
        GROUP BY mes
        ORDER BY mes ASC
    """)
    
    # 7. Evaluaciones críticas
    criticas = query_dict(conn, """
        SELECT 
            agente, supervisor, fecha, proyecto, score,
            ROUND(penc, 1) as penc,
            ROUND(pecuf, 1) as pecuf,
            ROUND(pecne, 1) as pecne,
            ROUND(peccum, 1) as peccum
        FROM evaluaciones
        WHERE es_critica = 1
        ORDER BY score ASC
        LIMIT 50
    """)
    
    # 8. Filtros
    supervisores_list = sorted(set(s['supervisor'] for s in supervisores))
    proyectos_list = sorted(set(c['campana'] for c in campanas))
    meses_list = query_dict(conn, "SELECT DISTINCT mes, mes_texto FROM evaluaciones WHERE mes != '' ORDER BY mes DESC")
    
    # 9. Atributos
    atributos = [
        {'nombre': 'Precisión Error No Crítico', 'sigla': 'PENC', 'valor': componentes['penc']},
        {'nombre': 'Precisión Error Crítico Usuario Final', 'sigla': 'PECUF', 'valor': componentes['pecuf']},
        {'nombre': 'Precisión Error Crítico Negocio', 'sigla': 'PECNE', 'valor': componentes['pecne']},
        {'nombre': 'Precisión Error Crítico Cumplimiento', 'sigla': 'PECCUM', 'valor': componentes['peccum']},
    ]
    
    # 10. Pesos
    pesos = query_dict(conn, "SELECT factor, peso FROM pesos")
    
    # Armar JSON completo
    data = {
        'resumen': resumen,
        'componentes': componentes,
        'supervisores': supervisores,
        'agentes': agentes,
        'cuartiles': {'Q1': q3, 'Q2': q2, 'Q3': q1, 'Q4': 0},
        'campanas': campanas,
        'historico': historico,
        'criticas': criticas,
        'atributos': atributos,
        'pesos': pesos,
        'filtros': {
            'supervisores': supervisores_list,
            'proyectos': proyectos_list,
            'meses': meses_list
        },
        'metadata': {
            'generado': 'export_json.py',
            'total_evaluaciones': resumen['total_evaluaciones'],
            'total_agentes': resumen['total_agentes'],
            'db_path': DB_PATH
        }
    }
    
    # Guardar JSON
    os.makedirs(os.path.dirname(JSON_PATH), exist_ok=True)
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    print(f"✅ JSON generado: {JSON_PATH}")
    print(f"   {os.path.getsize(JSON_PATH) / 1024:.1f} KB")
    print(f"   {resumen['total_evaluaciones']} evaluaciones, {resumen['total_agentes']} agentes")
    
    conn.close()

if __name__ == '__main__':
    main()
