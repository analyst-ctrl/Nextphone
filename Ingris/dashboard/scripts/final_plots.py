#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera grafico final: Resultados por grupos de nuevo ingreso
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os

OUT_DIR = r"C:\xampp\htdocs\Chamba Panama\Ingris\final_report"
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

# Data: (label, sublabel, value, date_sort)
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

bars = ax.bar(range(len(data)), values, color=colors, width=0.6,
              edgecolor='none', zorder=3)

for i, (bar, val) in enumerate(zip(bars, values)):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.2,
            f'{val:.1f}%', ha='center', va='bottom', fontweight='bold',
            fontsize=14, color='#e8eaf0')

# X-axis labels
ax.set_xticks(range(len(data)))
ax.set_xticklabels([f'{l}\n{s}' for l, s in zip(labels, sublabels)],
                   fontsize=11, ha='center')

# Umbral critical line
ax.axhline(y=90, color='#f87171', linestyle='--', linewidth=1.5, alpha=0.7,
           label='Umbral crítico (90%)')

ax.set_ylim(0, 100)
ax.set_title('Resultados por grupos de nuevo ingreso', fontsize=18,
             fontweight='bold', pad=20)
ax.set_ylabel('Score (%)', fontsize=13)
ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
ax.legend(framealpha=0.3, fontsize=10, loc='upper right')
ax.grid(axis='y', alpha=0.15, zorder=0)

# Add average line
avg = np.mean(values)
ax.axhline(y=avg, color='#fbbf24', linestyle=':', linewidth=1.2, alpha=0.6,
           label=f'Promedio ({avg:.1f}%)')

ax.legend(framealpha=0.3, fontsize=10, loc='upper right')

save(fig, 'resultados_nuevo_ingreso.png')

print("Done!")
