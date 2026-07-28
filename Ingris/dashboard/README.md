# Dashboard de Calidad — Ingris 📊

Dashboard interactivo para visualizar las evaluaciones de calidad de **Nextphone** y **OJT**.

## 🚀 Cómo usar en GitHub Pages

### 1. Generar los datos (solo 1 vez o cuando actualices el Excel)

```bash
cd Ingris/dashboard
python scripts/export_json.py
```

Esto lee `data/calidad.db` (generado desde los `.xlsm`) y crea `data/calidad_data.json`.

### 2. Subir a GitHub

```bash
git add Ingris/dashboard/data/calidad_data.json
git commit -m "data: actualizar datos de calidad"
git push
```

### 3. Abrir en GitHub Pages

Ve a `https://analyst-ctrl.github.io/Nextphone/Ingris/dashboard/index.html`

O si estás en local, abre directamente:

```
Ingris/dashboard/index.html
```

## 📁 Estructura

```
Ingris/dashboard/
├── index.html              ← Dashboard standalone (HTML+CSS+JS)
├── README.md               ← Este archivo
├── scripts/
│   ├── extract_data.py     ← Extrae datos del .xlsm → SQLite
│   └── export_json.py      ← Convierte SQLite → JSON
└── data/
    ├── calidad.db           ← Base de datos SQLite (no subir a git)
    └── calidad_data.json   ← Datos para el dashboard (¡subir a git!)
```

## ✅ Errores corregidos

- **Fórmula `+F17*1/G17/100`** → El `/100` extra dividía el score por 100. Ahora se multiplica por 100 al extraer.
- **Score leía columna equivocada** → Ahora lee 'SCORE CALIDAD' (Col 187) en vez de columnas de factores individuales.
- **Duplicación Nextphone+OJT** → Unificados en una sola DB con campo `proyecto`.
- **Referencias `#REF!`** → Las columnas borradas del Excel se omiten automáticamente.

## 🔧 Stack

- **Frontend:** HTML + CSS + JavaScript + [Chart.js](https://www.chartjs.org/)
- **Backend:** Python 3 (solo para generación de datos, no necesita servidor)
- **Base de datos:** SQLite

---

Hecho con ❤️ para Ingris
