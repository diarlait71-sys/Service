"""Детальный анализ структуры Excel файла лекарственного мониторинга."""
import openpyxl
import pandas as pd

path = r'C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\Монит 28.04.2026) 2 — копия.xlsx'

# Сначала посмотрим на листы
wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
print('=== ЛИСТЫ ===')
for name in wb.sheetnames:
    print(f'  - "{name}"')
wb.close()

# Читаем лист ЛС через pandas
print('\n=== ЧТЕНИЕ ЛИСТА ЛС ===')
# Читаем сырые данные без заголовков, первые 10 строк
df_raw = pd.read_excel(path, sheet_name='ЛС', header=None, nrows=10)
print(f'Форма: {df_raw.shape}')
print('\n--- Строка 0 (возможно заголовки уровня 1) ---')
for i, val in enumerate(df_raw.iloc[0]):
    if val is not None and str(val).strip() not in ('', 'nan'):
        print(f'  col[{i}]: {val}')

print('\n--- Строка 1 (возможно заголовки уровня 2) ---')
for i, val in enumerate(df_raw.iloc[1]):
    if val is not None and str(val).strip() not in ('', 'nan'):
        print(f'  col[{i}]: {val}')

print('\n--- Строка 2 ---')
for i, val in enumerate(df_raw.iloc[2]):
    if val is not None and str(val).strip() not in ('', 'nan'):
        print(f'  col[{i}]: {val}')

print('\n--- Строки 3-9 (примеры данных) ---')
for r in range(3, 10):
    row = df_raw.iloc[r]
    non_null = [(i, v) for i, v in enumerate(row) if v is not None and str(v).strip() not in ('', 'nan')]
    print(f'  Строка {r}: {non_null[:15]}')
