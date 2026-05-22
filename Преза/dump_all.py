# -*- coding: utf-8 -*-
import sys, openpyxl, datetime
sys.stdout.reconfigure(encoding='utf-8')

path = r"c:\Users\D.Muldabaev\Desktop\Приложения для сервиса\Преза\Маржа апрель — копия.xlsx"
wb = openpyxl.load_workbook(path, data_only=True)

# Все листы полностью
for sh in wb.sheetnames:
    ws = wb[sh]
    print(f"\n{'='*60}")
    print(f"ЛИСТ: {sh}  ({ws.max_row} строк x {ws.max_column} стол)")
    print(f"{'='*60}")
    for ri, row in enumerate(ws.iter_rows(values_only=True), 1):
        non_none = [v for v in row if v is not None]
        if non_none:
            # Форматируем datetime
            cleaned = []
            for v in non_none:
                if isinstance(v, datetime.datetime):
                    cleaned.append(v.strftime('%Y-%m-%d'))
                else:
                    cleaned.append(str(v))
            print(f"  r{ri:03d}: {cleaned}")
