import pandas as pd
import os

files = [r'c:\Users\D.Muldabaev\Desktop\Приложения для сервиса\bitrix_exports\Джак.xlsx',
         r'c:\Users\D.Muldabaev\Desktop\Приложения для сервиса\bitrix_exports\Джетур.xlsx',
         r'c:\Users\D.Muldabaev\Desktop\Приложения для сервиса\bitrix_exports\Киа.xlsx',
         r'c:\Users\D.Muldabaev\Desktop\Приложения для сервиса\bitrix_exports\Шеви.xlsx']

max_dates = []

for f in files:
    if not os.path.exists(f):
        print(f'File not found: {f}')
        continue
    try:
        xls = pd.ExcelFile(f)
        report_sheets = [s for s in xls.sheet_names if s.lower().startswith('report')]
        for sheet in report_sheets:
            df = pd.read_excel(f, sheet_name=sheet)
            date_cols = [c for c in df.columns if 'дата' in str(c).lower()]
            for col in date_cols:
                dates = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dropna()
                if not dates.empty:
                    max_dates.append(dates.max())
    except Exception as e:
        print(f'Error processing {f}: {e}')

if max_dates:
    overall_max = max(max_dates)
    print(f'RESULT_MAX_DATE:{overall_max.strftime("%Y-%m-%d")}')
else:
    print('RESULT_MAX_DATE:NONE')
