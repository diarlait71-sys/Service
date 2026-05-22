import pandas as pd
import os

path = r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\финанализ\Актау\ОСВ 01-04 2026.xlsx"
df = pd.read_excel(path, header=None)

print("=" * 80)
print("СТРУКТУРА ОСВ ФАЙЛА")
print("=" * 80)
print(f"\nВсего строк: {len(df)}, столбцов: {len(df.columns)}")
print("\nПервые 10 строк:")
print(df.head(10).to_string())

print("\n" + "=" * 80)
print("СЧЕТА СЕРИИ 6 (Расходы) И 7 (Доходы)")
print("=" * 80)

accounts = {}
for i in range(6, len(df)):
    r = df.iloc[i]
    acc_str = str(r.iloc[0]).strip()
    if not acc_str or 'итого' in acc_str.lower():
        continue
    parts = acc_str.split(',')
    acc = parts[0].strip()
    if not acc or len(acc) != 4 or not acc.isdigit():
        continue
    
    desc = parts[1].strip() if len(parts) > 1 else ""
    
    # Значения
    ndt = pd.to_numeric(r.iloc[2], errors='coerce') or 0
    nkt = pd.to_numeric(r.iloc[3], errors='coerce') or 0 if len(r) > 3 else 0
    kdt = pd.to_numeric(r.iloc[4], errors='coerce') or 0
    kkt = pd.to_numeric(r.iloc[5], errors='coerce') or 0 if len(r) > 5 else 0
    
    accounts[acc] = {
        'desc': desc,
        'ndt': ndt,
        'nkt': nkt,
        'kdt': kdt,
        'kkt': kkt
    }

# Счета 6 (расходы)
print("\n📊 РАСХОДЫ (Счета 6xxx):")
expenses_6 = {k: v for k, v in accounts.items() if k.startswith('6')}
for acc in sorted(expenses_6.keys()):
    v = accounts[acc]
    print(f"  {acc}: {v['desc']:40} | НДт={v['ndt']:12.0f} НКт={v['nkt']:12.0f} | КДт={v['kdt']:12.0f} ККт={v['kkt']:12.0f}")

print(f"\n  ИТОГО по 6xxx (Кредит/Расходы): {sum(v['kdt'] for v in expenses_6.values()):12.0f}")

# Счета 7 (доходы)
print("\n📈 ДОХОДЫ (Счета 7xxx):")
income_7 = {k: v for k, v in accounts.items() if k.startswith('7')}
for acc in sorted(income_7.keys()):
    v = accounts[acc]
    print(f"  {acc}: {v['desc']:40} | НДт={v['ndt']:12.0f} НКт={v['nkt']:12.0f} | КДт={v['kdt']:12.0f} ККт={v['kkt']:12.0f}")

print(f"\n  ИТОГО по 7xxx (Кредит/Доходы): {sum(v['kkt'] for v in income_7.values()):12.0f}")

# Расчет EBITDA
expenses = sum(v['kdt'] for k, v in accounts.items() if k.startswith('6'))
income = sum(v['kkt'] for k, v in accounts.items() if k.startswith('7'))
ebitda = income - expenses

print(f"\n\n💰 EBITDA РАСЧЕТ:")
print(f"  Доходы (7xxx по кредиту):     {income:15.0f}")
print(f"  Расходы (6xxx по дебету):     {expenses:15.0f}")
print(f"  EBITDA = Доходы - Расходы:   {ebitda:15.0f}")

# Касса
print("\n\n💳 КАССА (1010, 1030, 1050):")
cash_accounts = ['1010', '1030', '1050']
for acc in cash_accounts:
    if acc in accounts:
        v = accounts[acc]
        print(f"  {acc}: {v['desc']:40}")
        print(f"       Начало (Дт-Кт): {v['ndt']-v['nkt']:12.0f}")
        print(f"       Конец  (Дт-Кт): {v['kdt']-v['kkt']:12.0f}")
        print(f"       Изменение:       {(v['kdt']-v['kkt']) - (v['ndt']-v['nkt']):12.0f}")

cash_start = sum(accounts.get(a, {}).get('ndt', 0) - accounts.get(a, {}).get('nkt', 0) for a in cash_accounts if a in accounts)
cash_end = sum(accounts.get(a, {}).get('kdt', 0) - accounts.get(a, {}).get('kkt', 0) for a in cash_accounts if a in accounts)
cash_delta = cash_end - cash_start

print(f"\n  Кэш начало (Дт-Кт): {cash_start:12.0f}")
print(f"  Кэш конец  (Дт-Кт): {cash_end:12.0f}")
print(f"  Δ Кэш:               {cash_delta:12.0f}")

print(f"\n\n🔍 РАЗНИЦА (EBITDA - Δ Кэш):")
gap = ebitda - cash_delta
print(f"  {gap:12.0f}")
print(f"  Это затраты на развитие (капиталовложения, запасы, дебиторка и т.д.)")
