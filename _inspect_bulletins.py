"""Временный скрипт для чтения PDF бюллетеней и Excel Битрикса."""
import sys, json
from pathlib import Path
import pdfplumber
import pandas as pd

BASE = Path(r"c:\Users\D.Muldabaev\Desktop\Приложения для сервиса")
BULL = BASE / "бюллетени"
BX   = BASE / "bitrix_exports"

print("=" * 70)
print("PDF БЮЛЛЕТЕНИ")
print("=" * 70)

for pdf_path in sorted(BULL.glob("*.pdf")):
    print(f"\n{'─'*60}")
    print(f"FILE: {pdf_path.name}")
    print(f"{'─'*60}")
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                print(f"[Страница {i+1}]")
                # Попытка извлечь таблицы
                tables = page.extract_tables()
                if tables:
                    for t_i, tbl in enumerate(tables):
                        print(f"  Таблица {t_i+1}:")
                        for row in tbl:
                            if row and any(c for c in row if c):
                                print("   |", " | ".join(str(c or "").strip() for c in row), "|")
                else:
                    # Нет таблицы — вывести текст
                    text = page.extract_text() or ""
                    for line in text.split("\n")[:60]:
                        print("  ", line)
    except Exception as e:
        print(f"  ОШИБКА: {e}")

print("\n\n" + "=" * 70)
print("EXCEL БИТРИКС — структура и примеры")
print("=" * 70)

for xl_path in sorted(BX.glob("*.xlsx")):
    print(f"\n{'─'*60}")
    print(f"FILE: {xl_path.name}")
    print(f"{'─'*60}")
    try:
        xl = pd.ExcelFile(xl_path)
        print(f"Листы: {xl.sheet_names}")
        for sheet in xl.sheet_names[:3]:
            df = pd.read_excel(xl_path, sheet_name=sheet, nrows=5)
            print(f"\n  Лист '{sheet}' — колонки ({len(df.columns)}):")
            for c in df.columns:
                print(f"    • {c}")
            print(f"\n  Первые строки:")
            print(df.to_string(max_cols=10, max_colwidth=40))
    except Exception as e:
        print(f"  ОШИБКА: {e}")
