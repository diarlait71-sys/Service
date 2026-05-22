import pandas as pd
from pathlib import Path

file_path = Path("Факт Услуги.xlsx")
df = pd.read_excel(file_path, sheet_name="Sheet1", header=None)

print("=" * 120)
print("Первые 20 строк файла 'Факт Услуги.xlsx':")
print("=" * 120)
print(df.head(20).to_string())
print("\n")
print("Структура: столбцы =", df.shape[1], ", строки =", df.shape[0])
print("\nПервые 3 строки для анализа:")
for i in range(min(3, len(df))):
    print(f"Строка {i}: {df.iloc[i].tolist()}")
