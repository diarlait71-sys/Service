import pandas as pd
import os
import sys

# Папка с файлами
folder = r"c:\Users\D.Muldabaev\Desktop\Приложения для сервиса"

# Список файлов для анализа
files_to_analyze = [
    "Дц,должность,репер.xlsx",
    "Формула расчета.xlsx",
    "Доля показателей для Руководителя отдела сервиса.xlsx",
    "План продаж услуг и запчастей.xlsx",
    "план по остальным показателям.xlsx",
    "Факт Услуги.xlsx",
    "Факт Запасные части.xlsx",
    "Факт маржанальность.xlsx",
    "Факт по остальным показателям.xlsx"
]

print("=" * 120)
print("АНАЛИЗ СТРУКТУРЫ EXCEL ФАЙЛОВ")
print("=" * 120)

for file_name in files_to_analyze:
    file_path = os.path.join(folder, file_name)
    
    if not os.path.exists(file_path):
        print(f"\n❌ Файл не найден: {file_name}")
        continue
    
    print(f"\n{'=' * 120}")
    print(f"📄 ФАЙЛ: {file_name}")
    print(f"{'=' * 120}")
    
    try:
        # Чтение файла
        excel_file = pd.ExcelFile(file_path)
        
        # Если есть несколько листов
        sheet_names = excel_file.sheet_names
        print(f"🗂️  Листы в файле: {sheet_names}")
        
        # Анализ первого листа (или всех листов)
        for sheet_name in sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            if len(sheet_names) > 1:
                print(f"\n📋 ЛИСТ: '{sheet_name}'")
            
            print(f"📊 Размер: {df.shape[0]} строк × {df.shape[1]} столбцов")
            
            # Столбцы
            print(f"\n📌 Столбцы ({len(df.columns)}):")
            for i, col in enumerate(df.columns, 1):
                dtype = df[col].dtype
                non_null = df[col].notna().sum()
                print(f"   {i}. '{col}' ({dtype}) - заполнено: {non_null}/{df.shape[0]}")
            
            # Примеры данных (первые 5 строк)
            print(f"\n📋 Примеры данных (первые {min(5, len(df))} строк):")
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', 120)
            pd.set_option('display.max_colwidth', 40)
            print(df.head(5).to_string())
            
            print("\n" + "-" * 120)
    
    except Exception as e:
        print(f"❌ Ошибка при чтении файла: {e}")
        import traceback
        traceback.print_exc()

print("\n" + "=" * 120)
print("✅ Анализ завершён!")
