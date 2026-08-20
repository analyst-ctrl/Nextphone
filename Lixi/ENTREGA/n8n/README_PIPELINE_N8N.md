# Pipeline de Cobros con n8n

Propuesta para automatizar el reporte mensual de cobros (cartera + gestión) usando n8n.
Reemplaza el trabajo manual mensual: buscar fuentes, importar TXT/Excel, regenerar el reporte y avisar.

## Flujo propuesto (resumen)

1. **Disparador programado** — se ejecuta cada mes (ej. día 1 a las 8:00).
2. **Leer fuentes** — toma los archivos del mes desde `Lixi/Cobros/` (los 3 TXT: `LIST_51124`, `EXPORT_CALL_REPORT`, `LIST_16021`) y los Excel de gestión (`Materials/`).
3. **Limpiar / normalizar** (Code node) — quita espacios, unifica nombres de columnas, tipos numéricos.
4. **Guardar en SQLite** — hace UPSERT en `cobros.db` (tablas `list_51124`, `call_report`, `list_16021`).
5. **Generar reportes** — ejecuta `importar.py` + `excel_report.py` (hojas con fórmulas) y el dashboard HTML.
6. **Enviar aviso** — manda correo con el resumen (y adjunta el Excel/dashboard si se quiere).

## Por qué n8n y no solo Python

- **Visual y auditable**: el flujo queda documentado en un JSON versionable (este archivo).
- **Reintentos y errores**: si una fuente no está, n8n puede fallar o esperar; se ve dónde se rompe.
- **Extensiones fáciles**: sumar envío por WhatsApp/Drive/Sheets después sin tocar código.

## Requisitos

- n8n corriendo en la misma máquina (o con acceso a las carpetas). Ej. con Docker:
  ```
  docker run -d --name n8n -p 5678:5678 -v n8n_data:/home/node/.n8n -v "C:\xampp\htdocs\Chamba Panama":/data n8nio/n8n
  ```
  En Windows sin Docker: `npx n8n` (Node 18+).
- Python 3 con `openpyxl` instalado (para `excel_report.py`).
- SQLite con `cobros.db` accesible.

## Importar el flujo

1. Abrir n8n (http://localhost:5678).
2. "Workflows" → "Import from File" → elegir `pipeline_cobros_mensual.json`.
3. Configurar credenciales:
   - **SQLite**: ruta a `cobros.db`.
   - **Execute Command**: verificar que la ruta del script y el `cwd` apunten a `Lixi/Cobros`.
   - **Email (SMTP)**: solo si se usa el nodo de correo; si no, desconectar el nodo final.
4. Activar el workflow.

## Archivos

- `pipeline_cobros_mensual.json` — workflow importable a n8n.
- `README_PIPELINE_N8N.md` — este documento.

## Nota

El diseño usa los datos COMPLETOS del pipeline de `Lixi/Cobros` (TXT), no los XLSX recortados del reporte manual, para no perder llamadas ni cartera (ver hallazgos de julio: ~519 llamadas y ~211 contratos fuera del XLSX).
