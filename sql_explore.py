# -*- coding: utf-8 -*-
"""Шаг 4b: диагностика Document155 и поиск справочника механиков."""
"""Шаг 5: выгрузка механиков с исправленным фильтром дат (год 4026 = 2026 в 1С)."""
import pyodbc, pandas as pd

conn_str = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SRV1C;DATABASE=VOSTOK_UPR;"
    "UID=Diorbek;PWD=Vhgf5y$%^$56hg;"
    "TrustServerCertificate=yes;Encrypt=yes;"
)
conn = pyodbc.connect(conn_str, timeout=30)
cursor = conn.cursor()
OUT = r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\sql_diag2.txt"
log = []

# 1. Найти все RRef колонки в VT663 и определить что они указывают
cursor.execute("""
    SELECT TOP 5
        vt._Fld665RRef, vt._Fld671RRef, vt._Fld676RRef,
        vt._Fld666, vt._Fld667, vt._Fld670
    FROM _Document155 d
    JOIN _Document155_VT663 vt ON vt._Document155_IDRRef = d._IDRRef
    WHERE vt._Fld666 > 0 AND d._Posted = 0x01
""")
rows = cursor.fetchall()
log.append("VT663 RRef sample (5 strok):\n")
for r in rows:
    log.append(f"  665={r[0].hex()[:12] if r[0] else None}  671={r[1].hex()[:12] if r[1] else None}  676={r[2].hex()[:12] if r[2] else None}  h={r[3]}  rate={r[4]}  sum={r[5]}\n")

# 2. Матчим все три RRef против всех Reference (берём первый непустой)
all_refs_q = "SELECT name FROM sys.tables WHERE name LIKE '_Reference%' AND name NOT LIKE '%_VT%'"
cursor.execute(all_refs_q)
all_refs = [r[0] for r in cursor.fetchall()]

if rows:
    for fld_idx, fld_name in [(0, "Fld665RRef"), (1, "Fld671RRef"), (2, "Fld676RRef")]:
        sample = rows[0][fld_idx]
        if sample:
            log.append(f"\n--- {fld_name} ({sample.hex()[:16]}) -> References:\n")
            for tbl in all_refs:
                try:
                    cursor.execute(f"SELECT TOP 1 _Description FROM {tbl} WHERE _IDRRef = ?", sample)
                    found = cursor.fetchone()
                    if found and found[0]:
                        log.append(f"  MATCH {tbl}: {found[0]}\n")
                except Exception:
                    pass

# 3. Проверяем заголовок Document155 — ищем Reference117 в его полях
log.append("\n--- Mekhanik v zagolovke Document155? ---\n")
cursor.execute("""
    SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = '_Document155' AND DATA_TYPE = 'binary'
    AND COLUMN_NAME LIKE '%RRef'
""")
rref_cols = [r[0] for r in cursor.fetchall()]
cursor.execute("SELECT TOP 1 * FROM _Document155 WHERE _Posted=0x01")
doc_cols = [c[0] for c in cursor.description]
doc_row = cursor.fetchone()

if doc_row:
    doc_dict = dict(zip(doc_cols, doc_row))
    for col in rref_cols:
        val = doc_dict.get(col)
        if val and val != bytes(len(val)):  # не нулевой
            for tbl in ['_Reference117']:
                try:
                    cursor.execute(f"SELECT TOP 1 _Description FROM {tbl} WHERE _IDRRef = ?", val)
                    found = cursor.fetchone()
                    if found and found[0]:
                        log.append(f"  {col} -> {tbl}: {found[0]}\n")
                except Exception:
                    pass

# 4. Пробуем выборку за 4026 (= 2026 год в 1С)
cursor.execute("""
    SELECT TOP 5
        YEAR(d._Date_Time) AS stored_year,
        CONVERT(date, d._Date_Time) AS stored_date,
        d._Number,
        d._Posted,
        vt._Fld666, vt._Fld667, vt._Fld670
    FROM _Document155 d
    JOIN _Document155_VT663 vt ON vt._Document155_IDRRef = d._IDRRef
    WHERE YEAR(d._Date_Time) = 4026 AND vt._Fld666 > 0
""")
cols = [c[0] for c in cursor.description]
rows4026 = cursor.fetchall()
log.append(f"\nYear=4026 rows: {len(rows4026)}\n")
for r in rows4026:
    log.append("  " + "  ".join(str(v)[:20] if v is not None else "NULL" for v in r) + "\n")

with open(OUT, "w", encoding="utf-8") as f:
    f.writelines(log)
conn.close()
print("Done. See sql_diag2.txt")




