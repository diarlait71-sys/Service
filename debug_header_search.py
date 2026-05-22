import pandas as pd
from pathlib import Path

file_path = Path("Факт Услуги.xlsx")
df = pd.read_excel(file_path, sheet_name="Sheet1", header=None)

print("Первый столбец (для поиска заголовка):")
for i in range(5):
    print(f"  Строка {i}: {repr(df[0].iloc[i])}")

print("\nПоиск 'Название ДЦ' в столбце 0:")
for i, val in enumerate(df[0]):
    if isinstance(val, str) and "Название ДЦ" in val:
        print(f"  НАЙДЕНО на строке {i}")
        break
else:
    print("  НЕ НАЙДЕНО")

# Проверим столбец 0 целиком
print("\nВсе значения в столбце 0:")
for i, val in enumerate(df[0]):
    if pd.notna(val):
        print(f"  {i}: {repr(val)[:80]}")
