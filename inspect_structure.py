"""
Детальная инспекция структуры колонок файла для plan/fact анализа.
Записывает результат в UTF-8 файл.
"""
import pandas as pd
import json

path = r'C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\Монит 28.04.2026) 2 — копия.xlsx'

# Читаем первые 5 строк без заголовка
df = pd.read_excel(path, sheet_name='ЛС', header=None, nrows=5, engine='openpyxl')

lines = []
lines.append(f"Форма: {df.shape}")
lines.append("")
lines.append("=== СТРОКИ 0-4 (заголовки) ===")

for row_idx in range(5):
    lines.append(f"\n--- Строка {row_idx} ---")
    for col_idx in range(df.shape[1]):
        val = df.iloc[row_idx, col_idx]
        if pd.notna(val) and str(val).strip() not in ('', 'nan'):
            lines.append(f"  col[{col_idx:3d}]: {val}")

# Теперь читаем данные начиная со строки 2 (3я строка) с заголовками из строк 0-1
lines.append("\n\n=== ПЕРВЫЕ 10 СТРОК ДАННЫХ (начиная с 3-й строки файла) ===")
df_data = pd.read_excel(path, sheet_name='ЛС', header=None, skiprows=2, nrows=10, engine='openpyxl')
for row_idx in range(min(10, len(df_data))):
    lines.append(f"\n--- Строка данных {row_idx} ---")
    for col_idx in range(min(20, df_data.shape[1])):
        val = df_data.iloc[row_idx, col_idx]
        if pd.notna(val) and str(val).strip() not in ('', 'nan'):
            lines.append(f"  col[{col_idx:3d}]: {val}")

output = "\n".join(lines)
with open(r'C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\inspect_structure.txt', 'w', encoding='utf-8') as f:
    f.write(output)

print("Done!")
