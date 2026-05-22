"""
Модуль для расчета бонусов отдела сервиса
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional
import sys


class BonusCalculator:
    """Калькулятор бонусов на основе показателей"""
    
    def __init__(self, input_file: str, output_file: str = None):
        """
        Args:
            input_file: Путь к Excel файлу с исходными данными
            output_file: Путь для сохранения результатов (опционально)
        """
        self.input_file = input_file
        self.output_file = output_file or "bonus_results.xlsx"
        self.df = None
        self.results = None
        
    def load_data(self, sheet_name: int = 0) -> pd.DataFrame:
        """Загрузить данные из Excel файла"""
        try:
            self.df = pd.read_excel(self.input_file, sheet_name=sheet_name)
            print(f"✓ Данные загружены успешно: {self.df.shape[0]} строк, {self.df.shape[1]} столбцов")
            return self.df
        except FileNotFoundError:
            print(f"✗ Ошибка: Файл '{self.input_file}' не найден")
            raise
        except Exception as e:
            print(f"✗ Ошибка при загрузке файла: {e}")
            raise
    
    def display_data(self):
        """Показать загруженные данные"""
        if self.df is None:
            print("Данные не загружены. Вызовите load_data() сначала.")
            return
        
        print("\n" + "="*80)
        print("СТРУКТУРА ДАННЫХ:")
        print("="*80)
        print(f"\nКолонки: {list(self.df.columns)}\n")
        print(self.df.head(10).to_string())
        print(f"\n... всего {len(self.df)} строк\n")
    
    def calculate_bonus(self, 
                       plan_col: str, 
                       fact_col: str, 
                       rate_col: str,
                       hours_col: str,
                       threshold: float = 1.0,
                       coeff_col: Optional[str] = None) -> pd.DataFrame:
        """
        Расчет бонусов
        
        Args:
            plan_col: Название колонки с плановыми показателями
            fact_col: Название колонки с фактическими показателями
            rate_col: Название колонки со ставкой за час
            hours_col: Название колонки с нормированными часами
            threshold: Коэффициент выполнения плана (1.0 = 100%)
            coeff_col: Название колонки с коэффициентом (опционально)
        
        Returns:
            DataFrame с расчетными данными
        """
        if self.df is None:
            raise ValueError("Данные не загружены")
        
        result = self.df.copy()
        
        # Расчет процента выполнения плана
        result['Выполнение_%'] = (result[fact_col] / result[plan_col] * 100).round(1)
        
        # Определение коэффициента бонуса
        result['Коэфф_бонуса'] = result['Выполнение_%'].apply(
            lambda x: 1.0 if x >= threshold * 100 else 0.0
        )
        
        # Применить дополнительный коэффициент если указан
        if coeff_col and coeff_col in result.columns:
            result['Коэфф_бонуса'] = result['Коэфф_бонуса'] * result[coeff_col]
        
        # Расчет бонуса
        result['Бонус'] = (
            result[rate_col] * result[hours_col] * result['Коэфф_бонуса']
        ).round(2)
        
        self.results = result
        return result
    
    def save_results(self, template_file: Optional[str] = None):
        """
        Сохранить результаты
        
        Args:
            template_file: Путь к шаблону для форматирования (опционально)
        """
        if self.results is None:
            raise ValueError("Результаты не рассчитаны. Вызовите calculate_bonus() сначала")
        
        try:
            self.results.to_excel(self.output_file, sheet_name='Results', index=False)
            print(f"✓ Результаты сохранены в: {self.output_file}")
        except Exception as e:
            print(f"✗ Ошибка при сохранении: {e}")
            raise
    
    def get_summary(self) -> Dict:
        """Получить сводку по результатам"""
        if self.results is None:
            raise ValueError("Результаты не рассчитаны")
        
        return {
            'Всего_строк': len(self.results),
            'Выполнили_план': len(self.results[self.results['Коэфф_бонуса'] > 0]),
            'Средний_бонус': self.results['Бонус'].mean(),
            'Сумма_бонусов': self.results['Бонус'].sum(),
            'Max_бонус': self.results['Бонус'].max(),
            'Min_бонус': self.results['Бонус'].min(),
        }


def main():
    """Пример использования"""
    
    # НАСТРОЙКИ - УДАЛОСЬ МОДИФИЦИРОВАТЬ ДЛЯ ТВОИХ ДАННЫХ
    INPUT_FILE = "Ертыс.xls"  # Твой файл с данными
    OUTPUT_FILE = "bonus_results.xlsx"  # Результаты
    
    print("="*80)
    print("КАЛЬКУЛЯТОР БОНУСОВ - СЕРВИС")
    print("="*80 + "\n")
    
    try:
        # Создаем калькулятор
        calc = BonusCalculator(INPUT_FILE, OUTPUT_FILE)
        
        # Загружаем данные
        calc.load_data()
        
        # Показываем структуру
        calc.display_data()
        
        # ЗДЕСЬ НУЖНО ЗАДАТЬ ИМЕНА КОЛОНОК ИЗ ТВОЕГО ФАЙЛА
        # Замени на правильные названия после того как увидишь структуру:
        
        # calc.calculate_bonus(
        #     plan_col='План',
        #     fact_col='Факт', 
        #     rate_col='Ставка',
        #     hours_col='Нормо_часы',
        #     threshold=1.0  # 100% выполнения плана
        # )
        
        # calc.save_results()
        
        # summary = calc.get_summary()
        # print("\n" + "="*80)
        # print("СВОДКА:")
        # print("="*80)
        # for key, value in summary.items():
        #     print(f"{key}: {value}")
        
    except Exception as e:
        print(f"Ошибка: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
