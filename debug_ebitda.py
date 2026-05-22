import pandas as pd
import os

CASH_FOLDER = r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\финанализ\Актау"

def to_num(v):
    try:
        s = str(v).strip()
        return 0.0 if s.lower() in ("nan", "", "none", "null") else float(s.replace(" ", "").replace(",", ".").replace("\xa0", ""))
    except:
        return 0.0

def parse_osv(fn):
    p = os.path.join(CASH_FOLDER, fn)
    if not os.path.exists(p):
        return {}
    df = pd.read_excel(p, header=None)
    od = {}
    for i in range(6, len(df)):
        r = df.iloc[i]
        acc_str = str(r.iloc[0]).strip()
        if not acc_str or 'итого' in acc_str.lower():
            continue
        acc = acc_str.split(',')[0].strip()
        if not acc or len(acc) != 4 or not acc.isdigit():
            continue
        od[acc] = {'ndt': to_num(r.iloc[2]), 'nkt': to_num(r.iloc[3]) if len(r) > 3 else 0, 'kdt': to_num(r.iloc[4]), 'kkt': to_num(r.iloc[5]) if len(r) > 5 else 0}

    # Добавляем суммы по 6xxx и 7xxx
    expenses_deb = sum(v['kdt'] for k, v in od.items() if k.startswith('6'))
    income_crd = sum(v['kkt'] for k, v in od.items() if k.startswith('7'))
    od['_expenses_deb'] = expenses_deb
    od['_income_crd'] = income_crd
    
    print(f"Expenses (6xxx по дебету): {expenses_deb:15.2f}")
    print(f"Income (7xxx по кредиту): {income_crd:15.2f}")
    print(f"EBITDA: {(income_crd - expenses_deb) / 1e6:15.2f} млн")
    
    return od

fn = "ОСВ 01-04 2026.xlsx"
od = parse_osv(fn)
print(f"\nПроверка значений в od:")
print(f"'_expenses_deb' in od: {'_expenses_deb' in od}")
print(f"'_income_crd' in od: {'_income_crd' in od}")
print(f"od.get('_expenses_deb', 0): {od.get('_expenses_deb', 0)}")
print(f"od.get('_income_crd', 0): {od.get('_income_crd', 0)}")
