#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera graficos finales: Resultados por grupos de nuevo ingreso + Historico
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
os.makedirs(OUT_DIR, exist_ok=True)

plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.sans-serif': ['Segoe UI', 'DejaVu Sans'],
    'axes.edgecolor': '#2a2d3e',
    'axes.labelcolor': '#8b8fa3',
    'axes.titlecolor': '#e8eaf0',
    'xtick.color': '#8b8fa3',
    'ytick.color': '#8b8fa3',
    'text.color': '#e8eaf0',
    'figure.facecolor': '#0f1117',
    'axes.facecolor': '#1a1d27',
    'legend.facecolor': '#1e2130',
    'legend.edgecolor': '#2a2d3e',
    'legend.labelcolor': '#e8eaf0',
})

def save(fig, name):
    path = os.path.join(OUT_DIR, name)
    fig.savefig(path, dpi=200, bbox_inches='tight', facecolor='#0f1117')
    print(f"Saved: {path}")
    plt.close(fig)

# ═══════════════════════════════════════════════════════════
# 1. RESULTADOS POR GRUPOS DE NUEVO INGRESO
# ═══════════════════════════════════════════════════════════
data = [
    ('Wave 7',   '12-may', 70.0),
    ('Wave 6',   '22-may', 55.0),
    ('Wave 5',   '16-jun', 66.4),
    ('Wave 4',   '26-jun', 60.8),
    ('Jun-26',   'Grupo inicial', 75.7),
    ('Jul-26',   'Grupo inicial', 64.5),
]

labels = [d[0] for d in data]
sublabels = [d[1] for d in data]
values = [d[2] for d in data]
colors = ['#fbbf24', '#f87171', '#a78bfa', '#4f8cff', '#34d399', '#4f8cff']

fig, ax = plt.subplots(figsize=(12, 7))
bars = ax.bar(range(len(data)), values, color=colors, width=0.6, edgecolor='none', zorder=3)

for i, (bar, val) in enumerate(zip(bars, values)):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.2,
            f'{val:.1f}%', ha='center', va='bottom', fontweight='bold',
            fontsize=14, color='#e8eaf0')

ax.set_xticks(range(len(data)))
ax.set_xticklabels([f'{l}\n{s}' for l, s in zip(labels, sublabels)], fontsize=11, ha='center')
ax.axhline(y=90, color='#f87171', linestyle='--', linewidth=1.5, alpha=0.7, label='Umbral crítico (90%)')
avg = np.mean(values)
ax.axhline(y=avg, color='#fbbf24', linestyle=':', linewidth=1.2, alpha=0.6, label=f'Promedio ({avg:.1f}%)')
ax.set_ylim(0, 100)
ax.set_title('Resultados por grupos de nuevo ingreso', fontsize=18, fontweight='bold', pad=20)
ax.set_ylabel('Score (%)', fontsize=13)
ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
ax.legend(framealpha=0.3, fontsize=10, loc='upper right')
ax.grid(axis='y', alpha=0.15, zorder=0)
save(fig, 'resultados_nuevo_ingreso.png')

# ═══════════════════════════════════════════════════════════
# 2. HISTORICO - SCORE GENERAL
# ═══════════════════════════════════════════════════════════
# Pesos: PENC=10%, PECUF=30%, PECNE=30%, PECCUM=30%
def calc_score(penc, pecuf, pecne, peccum):
    return 0.1*penc + 0.3*pecuf + 0.3*pecne + 0.3*peccum

componentes = {
    'Mar-26': {'PENC': 78.0, 'PECUF': 64.29, 'PECNE': 38.0, 'PECCUM': 73.81},
    'Abr-26': {'PENC': 83.0, 'PECUF': 39.39, 'PECNE': 30.0, 'PECCUM': 66.67},
    'May-26': {'PENC': 86.0, 'PECUF': 40.0,  'PECNE': 48.0, 'PECCUM': 100.0},
    'Jun-26': {'PENC': 100.0,'PECUF': 85.7,  'PECNE': 33.3, 'PECCUM': 100.0},
    'Jul-26': {'PENC': 95.0, 'PECUF': 58.3,  'PECNE': 25.0, 'PECCUM': 100.0},
}

for m, d in componentes.items():
    d['Score'] = round(calc_score(d['PENC'], d['PECUF'], d['PECNE'], d['PECCUM']), 1)

fig, ax = plt.subplots(figsize=(12, 7))
meses = list(componentes.keys())
x = np.arange(len(meses))
width = 0.18

colores_comp = {'PENC': '#4f8cff', 'PECUF': '#a78bfa', 'PECNE': '#fbbf24', 'PECCUM': '#f87171'}

for i, comp in enumerate(['PENC', 'PECUF', 'PECNE', 'PECCUM']):
    vals = [componentes[m][comp] for m in meses]
    ax.bar(x + (i-1.5)*width, vals, width, label=comp, color=colores_comp[comp], zorder=3)

# Score line
score_vals = [componentes[m]['Score'] for m in meses]
ax.plot(x, score_vals, 'o-', color='#34d399', linewidth=3, markersize=9, zorder=5, label='Score General')

for i, val in enumerate(score_vals):
    ax.text(i, val + 3, f'{val:.1f}%', ha='center', fontweight='bold', fontsize=11, color='#34d399')

ax.set_xticks(x)
ax.set_xticklabels(meses, fontsize=12, fontweight='bold')
ax.set_ylim(0, 115)
ax.set_title('Histórico - Score General', fontsize=18, fontweight='bold', pad=20)
ax.set_ylabel('%', fontsize=13)
ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
ax.legend(framealpha=0.3, fontsize=10, ncol=5, loc='upper center', bbox_to_anchor=(0.5, -0.08))
ax.grid(axis='y', alpha=0.15, zorder=0)
ax.axhline(y=90, color='#f87171', linestyle='--', linewidth=1, alpha=0.5)
save(fig, 'historico_score_general.png')

# ═══════════════════════════════════════════════════════════
# 3. RESULTADOS POR CAMPAÑA
# ═══════════════════════════════════════════════════════════
campanas = [
    ('Nextphone\nPre2Pos', 68.8),
    ('Nextphone\nRenovacion', 63.6),
    ('Nextphone\nTotal', 65.9),
    ('OJT', 71.6),
    ('OJT\nPre2Pos', 75.7),
    ('OJT\nRenovacion', 64.5),
]

fig, ax = plt.subplots(figsize=(12, 7))
naranja = '#D04423'
bars = ax.bar(range(len(campanas)), [c[1] for c in campanas], color=naranja, width=0.6, edgecolor='none', zorder=3)

for bar, val in zip(bars, [c[1] for c in campanas]):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.2,
            f'{val:.1f}%', ha='center', va='bottom', fontweight='bold',
            fontsize=13, color='#e8eaf0')

ax.set_xticks(range(len(campanas)))
ax.set_xticklabels([c[0] for c in campanas], fontsize=10, ha='center')
ax.axhline(y=90, color='#f87171', linestyle='--', linewidth=1.5, alpha=0.7, label='Umbral crítico (90%)')
ax.set_ylim(0, 100)
ax.set_title('Resultados por Campaña', fontsize=18, fontweight='bold', pad=20)
ax.set_ylabel('Score (%)', fontsize=13)
ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
ax.legend(framealpha=0.3, fontsize=10, loc='upper right')
ax.grid(axis='y', alpha=0.15, zorder=0)
save(fig, 'resultados_campana.png')

print("Done! Todos los graficos generados.")
