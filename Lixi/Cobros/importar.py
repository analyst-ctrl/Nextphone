import sqlite3, csv, os, re

DB = os.path.join(os.path.dirname(__file__), "cobros.db")
FILES = {
    "call_report": "EXPORT_CALL_REPORT_20260730-115503.txt",
    "list_16021": "LIST_16021_20260730-115912.txt",
    "list_51124": "LIST_51124_20260730-115626.txt",
}

os.chdir(os.path.dirname(__file__))

if os.path.exists(DB):
    os.remove(DB)

conn = sqlite3.connect(DB)
conn.execute("PRAGMA journal_mode=WAL")

def clean_name(n):
    n = n.strip().lower().replace(" ", "_")
    n = re.sub(r'[^a-z0-9_]', '', n)
    if n and n[0].isdigit():
        n = 'c' + n
    return n or 'col'

for table, fname in FILES.items():
    print(f"Processing {fname} -> {table} ...")
    with open(fname, "r", encoding="utf-8") as f:
        reader = csv.reader(f, delimiter="\t")
        headers = [clean_name(h) for h in next(reader)]
        cols = ", ".join(f'"{h}"' for h in headers)
        conn.execute(f'CREATE TABLE IF NOT EXISTS "{table}" ({cols})')
        placeholders = ", ".join("?" for _ in headers)
        sql = f'INSERT INTO "{table}" ({cols}) VALUES ({placeholders})'
        ncols = len(headers)
        rows = []
        for r in reader:
            if len(r) < ncols:
                r = r + [None] * (ncols - len(r))
            rows.append(r[:ncols])
        conn.executemany(sql, rows)
    print(f"  -> {len(rows)} rows imported")

conn.commit()
conn.close()
print(f"\nDatabase created: {DB}")
