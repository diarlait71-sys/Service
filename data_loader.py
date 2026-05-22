"""
Загрузчик данных планов и фактов для расчёта бонусов
"""

import pandas as pd
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from datetime import datetime
import json


@dataclass
class KPIData:
    """Данные KPI для сотрудника"""
    employee_fio: str
    position: str
    dealer_center: str
    period_start: str
    period_end: str
    metric_code: str
    metric_name: str
    plan_value: float
    actual_value: float
    source: str  # '1C', 'PowerBI', 'manual'


@dataclass
class EmployeeKPIData:
    """Все KPI данные сотрудника за период"""
    employee_fio: str
    position: str
    dealer_center: str
    period_start: str
    period_end: str
    kpi_data: Dict[str, KPIData]  # metric_code -> data


class DataLoader:
    """Загрузчик данных планов и фактов"""

    def __init__(self):
        self.kpi_data: Dict[str, EmployeeKPIData] = {}  # employee_fio -> data

    def load_from_excel(self, file_path: str, sheet_name: str = 'KPI_Data') -> Dict[str, EmployeeKPIData]:
        """
        Загружает данные планов/фактов из Excel

        Ожидаемая структура Excel:
        - Employee_FIO: ФИО сотрудника
        - Position: Должность
        - Dealer_Center: Дилерский центр
        - Period_Start: Начало периода (YYYY-MM-DD)
        - Period_End: Конец периода (YYYY-MM-DD)
        - Metric_Code: Код показателя (normo_hours, parts_sales, etc.)
        - Metric_Name: Название показателя
        - Plan_Value: Плановое значение
        - Actual_Value: Фактическое значение
        - Source: Источник данных (1C, PowerBI, manual)
        """
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Проверяем обязательные колонки
            required_columns = [
                'Employee_FIO', 'Position', 'Dealer_Center',
                'Period_Start', 'Period_End', 'Metric_Code',
                'Metric_Name', 'Plan_Value', 'Actual_Value', 'Source'
            ]

            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Отсутствуют обязательные колонки: {missing_columns}")

            # Группируем по сотрудникам
            for _, row in df.iterrows():
                try:
                    # Создаём ключ сотрудника
                    employee_key = str(row['Employee_FIO']).strip()

                    # Создаём данные KPI
                    kpi_data = KPIData(
                        employee_fio=employee_key,
                        position=str(row['Position']),
                        dealer_center=str(row['Dealer_Center']),
                        period_start=str(row['Period_Start']),
                        period_end=str(row['Period_End']),
                        metric_code=str(row['Metric_Code']),
                        metric_name=str(row['Metric_Name']),
                        plan_value=float(row['Plan_Value']) if pd.notna(row.get('Plan_Value')) else 0,
                        actual_value=float(row['Actual_Value']) if pd.notna(row.get('Actual_Value')) else 0,
                        source=str(row['Source']) if pd.notna(row.get('Source')) else 'manual'
                    )

                    # Добавляем к данным сотрудника
                    if employee_key not in self.kpi_data:
                        self.kpi_data[employee_key] = EmployeeKPIData(
                            employee_fio=employee_key,
                            position=kpi_data.position,
                            dealer_center=kpi_data.dealer_center,
                            period_start=kpi_data.period_start,
                            period_end=kpi_data.period_end,
                            kpi_data={}
                        )

                    self.kpi_data[employee_key].kpi_data[kpi_data.metric_code] = kpi_data

                except Exception as e:
                    print(f"Ошибка при обработке строки {row.name}: {e}")
                    continue

            print(f"✓ Загружено данных для {len(self.kpi_data)} сотрудников")
            return self.kpi_data

        except Exception as e:
            raise ValueError(f"Ошибка при загрузке данных: {e}")

    def create_sample_data_excel(self, output_file: str = 'sample_kpi_data.xlsx'):
        """
        Создаёт пример Excel файла с данными планов/фактов
        """
        sample_data = [
            {
                'Employee_FIO': 'Иванов Иван Иванович',
                'Position': 'Менеджер сервиса',
                'Dealer_Center': 'ДЦ Алматы',
                'Period_Start': '2026-04-01',
                'Period_End': '2026-04-30',
                'Metric_Code': 'normo_hours',
                'Metric_Name': 'Нормо-часы',
                'Plan_Value': 160,
                'Actual_Value': 158,
                'Source': '1C'
            },
            {
                'Employee_FIO': 'Иванов Иван Иванович',
                'Position': 'Менеджер сервиса',
                'Dealer_Center': 'ДЦ Алматы',
                'Period_Start': '2026-04-01',
                'Period_End': '2026-04-30',
                'Metric_Code': 'parts_sales',
                'Metric_Name': 'Запчасти',
                'Plan_Value': 200000,
                'Actual_Value': 210000,
                'Source': '1C'
            },
            {
                'Employee_FIO': 'Иванов Иван Иванович',
                'Position': 'Менеджер сервиса',
                'Dealer_Center': 'ДЦ Алматы',
                'Period_Start': '2026-04-01',
                'Period_End': '2026-04-30',
                'Metric_Code': 'upsell',
                'Metric_Name': 'Доп. оборудование',
                'Plan_Value': 150000,
                'Actual_Value': 160000,
                'Source': 'PowerBI'
            },
            {
                'Employee_FIO': 'Иванов Иван Иванович',
                'Position': 'Менеджер сервиса',
                'Dealer_Center': 'ДЦ Алматы',
                'Period_Start': '2026-04-01',
                'Period_End': '2026-04-30',
                'Metric_Code': 'customer_loyalty',
                'Metric_Name': 'Лояльность клиентов',
                'Plan_Value': 90,
                'Actual_Value': 92,
                'Source': 'manual'
            },
            {
                'Employee_FIO': 'Иванов Иван Иванович',
                'Position': 'Менеджер сервиса',
                'Dealer_Center': 'ДЦ Алматы',
                'Period_Start': '2026-04-01',
                'Period_End': '2026-04-30',
                'Metric_Code': 'service_quality',
                'Metric_Name': 'Качество сервиса',
                'Plan_Value': 95,
                'Actual_Value': 96,
                'Source': '1C'
            },
            # Второй сотрудник
            {
                'Employee_FIO': 'Петров Петр Петрович',
                'Position': 'Менеджер сервиса',
                'Dealer_Center': 'ДЦ Астана',
                'Period_Start': '2026-04-01',
                'Period_End': '2026-04-30',
                'Metric_Code': 'normo_hours',
                'Metric_Name': 'Нормо-часы',
                'Plan_Value': 160,
                'Actual_Value': 165,
                'Source': '1C'
            },
            {
                'Employee_FIO': 'Петров Петр Петрович',
                'Position': 'Менеджер сервиса',
                'Dealer_Center': 'ДЦ Астана',
                'Period_Start': '2026-04-01',
                'Period_End': '2026-04-30',
                'Metric_Code': 'parts_sales',
                'Metric_Name': 'Запчасти',
                'Plan_Value': 180000,
                'Actual_Value': 175000,
                'Source': '1C'
            }
        ]

        df = pd.DataFrame(sample_data)
        df.to_excel(output_file, sheet_name='KPI_Data', index=False)
        print(f"✓ Пример файла создан: {output_file}")

        return df

    def get_employee_data(self, employee_fio: str) -> Optional[EmployeeKPIData]:
        """Получить данные сотрудника"""
        return self.kpi_data.get(employee_fio)

    def list_employees(self) -> List[str]:
        """Получить список всех сотрудников"""
        return list(self.kpi_data.keys())

    def get_periods(self) -> List[Tuple[str, str]]:
        """Получить список уникальных периодов"""
        periods = set()
        for emp_data in self.kpi_data.values():
            periods.add((emp_data.period_start, emp_data.period_end))
        return sorted(list(periods))

    def to_json(self) -> str:
        """Экспорт данных в JSON"""
        data_dict = {}
        for employee, emp_data in self.kpi_data.items():
            data_dict[employee] = {
                'employee_fio': emp_data.employee_fio,
                'position': emp_data.position,
                'dealer_center': emp_data.dealer_center,
                'period_start': emp_data.period_start,
                'period_end': emp_data.period_end,
                'kpi_data': {
                    code: {
                        'metric_code': data.metric_code,
                        'metric_name': data.metric_name,
                        'plan_value': data.plan_value,
                        'actual_value': data.actual_value,
                        'source': data.source
                    } for code, data in emp_data.kpi_data.items()
                }
            }

        return json.dumps(data_dict, indent=2, ensure_ascii=False)


# ============================================================================
# ТЕСТОВЫЕ ФУНКЦИИ
# ============================================================================

def test_data_loader():
    """Тест загрузчика данных"""
    loader = DataLoader()

    # Создаём пример файла
    loader.create_sample_data_excel()

    # Загружаем данные
    try:
        data = loader.load_from_excel('sample_kpi_data.xlsx')
        print(f"✓ Данные загружены для {len(data)} сотрудников")

        for employee, emp_data in data.items():
            print(f"\n👤 {employee}")
            print(f"   Должность: {emp_data.position}")
            print(f"   ДЦ: {emp_data.dealer_center}")
            print(f"   Период: {emp_data.period_start} - {emp_data.period_end}")
            print(f"   KPI показателей: {len(emp_data.kpi_data)}")

            for code, kpi in emp_data.kpi_data.items():
                print(f"     - {code}: план={kpi.plan_value}, факт={kpi.actual_value} ({kpi.source})")

        # Экспорт в JSON
        json_output = loader.to_json()
        with open('kpi_data_export.json', 'w', encoding='utf-8') as f:
            f.write(json_output)
        print(f"\n✓ Данные экспортированы в kpi_data_export.json")

    except Exception as e:
        print(f"✗ Ошибка: {e}")


if __name__ == "__main__":
    test_data_loader()