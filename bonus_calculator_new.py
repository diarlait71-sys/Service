"""
Расчётчик бонусов на основе шаблонов и фактических данных
"""

import json
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field
from template_parser import TemplateParser, BonusTemplate, KPITemplate, DemotivatorTemplate
from data_loader import DataLoader, EmployeeKPIData, KPIData


@dataclass
class KPICalculationResult:
    """Результат расчёта KPI"""
    metric_code: str
    metric_name: str
    planned_value: float
    actual_value: float
    weight: float
    formula_type: str
    result_value: float  # рассчитанное значение (0-1 или абсолютное)
    weighted_value: float  # взвешенное значение


@dataclass
class DemotivatorResult:
    """Результат применения демотиватора"""
    code: str
    name: str
    demotivator_type: str
    rule: str
    value: float
    applied_value: float  # применённое значение


@dataclass
class BonusCalculation:
    """Полный расчёт бонуса сотрудника"""
    employee_fio: str
    position: str
    dealer_center: str
    period_start: str
    period_end: str

    # Базовые компоненты
    base_salary: float
    budget_bonus: float
    base_total: float = field(init=False)

    # Результаты KPI
    kpi_results: List[KPICalculationResult] = field(default_factory=list)
    kpi_factor: float = field(init=False)  # общий фактор KPI

    # Результаты демотиваторов
    demotivator_results: List[DemotivatorResult] = field(default_factory=list)
    demotivator_factor: float = field(init=False)  # общий фактор демотиваторов

    # Итоговый бонус
    total_bonus: float = field(init=False)

    def __post_init__(self):
        self.base_total = self.base_salary + self.budget_bonus
        self.kpi_factor = 1.0
        self.demotivator_factor = 1.0
        self.total_bonus = 0.0


class BonusCalculator:
    """Расчётчик бонусов"""

    def __init__(self):
        self.template_parser = TemplateParser()
        self.data_loader = DataLoader()
        self.templates: Dict[str, BonusTemplate] = {}
        self.kpi_data: Dict[str, EmployeeKPIData] = {}

    def load_templates(self, template_file: str):
        """Загрузить шаблоны расчёта"""
        self.templates = self.template_parser.parse_excel_template(template_file)
        print(f"✓ Загружено {len(self.templates)} шаблонов")

    def load_kpi_data(self, data_file: str):
        """Загрузить данные планов/фактов"""
        self.kpi_data = self.data_loader.load_from_excel(data_file)

    def calculate_employee_bonus(self, employee_fio: str) -> Optional[BonusCalculation]:
        """
        Рассчитать бонус для сотрудника

        Args:
            employee_fio: ФИО сотрудника

        Returns:
            Результат расчёта или None если данные не найдены
        """
        # Получаем данные сотрудника
        emp_data = self.kpi_data.get(employee_fio)
        if not emp_data:
            print(f"✗ Данные для сотрудника {employee_fio} не найдены")
            return None

        # Получаем шаблон для должности
        template = self.templates.get(emp_data.position)
        if not template:
            print(f"✗ Шаблон для должности {emp_data.position} не найден")
            return None

        # Создаём объект расчёта
        calculation = BonusCalculation(
            employee_fio=employee_fio,
            position=emp_data.position,
            dealer_center=emp_data.dealer_center,
            period_start=emp_data.period_start,
            period_end=emp_data.period_end,
            base_salary=template.base_salary,
            budget_bonus=template.budget_bonus
        )

        # Рассчитываем KPI
        self._calculate_kpis(calculation, template, emp_data)

        # Применяем демотиваторы
        self._apply_demotivators(calculation, template, emp_data)

        # Рассчитываем итоговый бонус
        self._calculate_total_bonus(calculation)

        return calculation

    def _calculate_kpis(self, calculation: BonusCalculation,
                        template: BonusTemplate,
                        emp_data: EmployeeKPIData):
        """Рассчитать KPI показатели"""
        total_weighted_result = 0.0
        total_weight = 0.0

        for kpi_template in template.kpis:
            # Получаем фактические данные
            kpi_data = emp_data.kpi_data.get(kpi_template.code)

            if kpi_data:
                actual_value = kpi_data.actual_value
                planned_value = kpi_data.plan_value
            else:
                # Если данных нет, используем плановое значение из шаблона
                actual_value = kpi_template.plan
                planned_value = kpi_template.plan
                print(f"⚠ Для {kpi_template.code} используются плановые значения")

            # Рассчитываем результат по формуле
            result_value = self._calculate_kpi_formula(
                kpi_template.formula_type,
                kpi_template.formula_params,
                planned_value,
                actual_value
            )

            # Взвешенное значение
            weighted_value = result_value * kpi_template.weight

            # Добавляем к общему результату
            total_weighted_result += weighted_value
            total_weight += kpi_template.weight

            # Создаём результат KPI
            kpi_result = KPICalculationResult(
                metric_code=kpi_template.code,
                metric_name=kpi_template.name,
                planned_value=planned_value,
                actual_value=actual_value,
                weight=kpi_template.weight,
                formula_type=kpi_template.formula_type,
                result_value=result_value,
                weighted_value=weighted_value
            )

            calculation.kpi_results.append(kpi_result)

        # Общий фактор KPI
        if total_weight > 0:
            calculation.kpi_factor = total_weighted_result / total_weight
        else:
            calculation.kpi_factor = 1.0

    def _calculate_kpi_formula(self, formula_type: str,
                              formula_params: Dict[str, Any],
                              planned: float, actual: float) -> float:
        """Рассчитать значение по формуле KPI"""

        if formula_type == 'ratio':
            # Простое отношение факт/план с ограничениями
            min_ratio = formula_params.get('min_ratio', 0.8)
            max_ratio = formula_params.get('max_ratio', 1.2)

            ratio = actual / planned if planned > 0 else 0
            return max(min_ratio, min(max_ratio, ratio))

        elif formula_type == 'threshold':
            # Пороговые значения
            thresholds = formula_params.get('thresholds', [])
            ratio = actual / planned if planned > 0 else 0

            for threshold in thresholds:
                if ratio >= threshold['value']:
                    return threshold['multiplier']

            return 0.0  # Если не достигнут минимальный порог

        elif formula_type == 'linear':
            # Линейная функция: a * (actual - planned) + b
            a = formula_params.get('a', 0.05)
            b = formula_params.get('b', 0)
            return a * (actual - planned) + b

        elif formula_type == 'score':
            # Оценка: (actual / scale) * bonus_per_point
            scale = formula_params.get('scale', 100)
            bonus_per_point = formula_params.get('bonus_per_point', 50)
            return (actual / scale) * bonus_per_point

        else:
            # По умолчанию - простое отношение
            return actual / planned if planned > 0 else 0

    def _apply_demotivators(self, calculation: BonusCalculation,
                           template: BonusTemplate,
                           emp_data: EmployeeKPIData):
        """Применить демотиваторы"""

        for demo_template in template.demotivators:
            # Получаем фактические данные для демотиватора
            demo_data = emp_data.kpi_data.get(demo_template.code)

            actual_value = demo_data.actual_value if demo_data else 0
            planned_value = demo_data.plan_value if demo_data else 0

            # Оцениваем правило
            rule_result = self._evaluate_rule(
                demo_template.rule,
                planned_value,
                actual_value
            )

            # Применяем значение если правило истинно
            applied_value = 1.0  # по умолчанию
            if rule_result:
                if demo_template.value_type == 'coefficient':
                    applied_value = demo_template.value
                else:  # fixed
                    applied_value = demo_template.value

            # Создаём результат демотиватора
            demo_result = DemotivatorResult(
                code=demo_template.code,
                name=demo_template.name,
                demotivator_type=demo_template.demotivator_type,
                rule=demo_template.rule,
                value=demo_template.value,
                applied_value=applied_value
            )

            calculation.demotivator_results.append(demo_result)

            # Применяем к общему фактору
            if demo_template.demotivator_type == 'penalty':
                calculation.demotivator_factor *= applied_value
            else:  # reward
                calculation.demotivator_factor *= applied_value

    def _evaluate_rule(self, rule: str, planned: float, actual: float) -> bool:
        """Оценить правило демотиватора"""
        try:
            # Простая оценка условий
            if 'actual < plan' in rule.lower():
                return actual < planned
            elif 'actual >= plan' in rule.lower():
                return actual >= planned
            elif 'actual >' in rule.lower():
                # Извлекаем значение
                parts = rule.lower().split('actual >')
                if len(parts) > 1:
                    threshold = float(parts[1].split()[0])
                    return actual > threshold
            elif 'actual <' in rule.lower():
                # Извлекаем значение
                parts = rule.lower().split('actual <')
                if len(parts) > 1:
                    threshold = float(parts[1].split()[0])
                    return actual < threshold

            return False  # по умолчанию

        except Exception as e:
            print(f"Ошибка при оценке правила '{rule}': {e}")
            return False

    def _calculate_total_bonus(self, calculation: BonusCalculation):
        """Рассчитать итоговый бонус"""
        calculation.total_bonus = (
            calculation.base_total *
            calculation.kpi_factor *
            calculation.demotivator_factor
        )

    def calculate_all_bonuses(self) -> Dict[str, BonusCalculation]:
        """Рассчитать бонусы для всех сотрудников"""
        results = {}

        for employee_fio in self.kpi_data.keys():
            calculation = self.calculate_employee_bonus(employee_fio)
            if calculation:
                results[employee_fio] = calculation

        return results

    def export_results_to_excel(self, results: Dict[str, BonusCalculation],
                               output_file: str = 'bonus_calculation_results.xlsx'):
        """Экспорт результатов в Excel"""
        data = []

        for employee, calc in results.items():
            for kpi in calc.kpi_results:
                data.append({
                    'Employee_FIO': employee,
                    'Position': calc.position,
                    'Dealer_Center': calc.dealer_center,
                    'Period_Start': calc.period_start,
                    'Period_End': calc.period_end,
                    'Base_Salary': calc.base_salary,
                    'Budget_Bonus': calc.budget_bonus,
                    'Base_Total': calc.base_total,
                    'KPI_Code': kpi.metric_code,
                    'KPI_Name': kpi.metric_name,
                    'KPI_Plan': kpi.planned_value,
                    'KPI_Actual': kpi.actual_value,
                    'KPI_Weight': kpi.weight,
                    'KPI_Result': kpi.result_value,
                    'KPI_Weighted': kpi.weighted_value,
                    'KPI_Factor': calc.kpi_factor,
                    'Demotivator_Factor': calc.demotivator_factor,
                    'Total_Bonus': calc.total_bonus
                })

        df = pd.DataFrame(data)
        df.to_excel(output_file, index=False)
        print(f"✓ Результаты экспортированы в {output_file}")


# ============================================================================
# ТЕСТОВЫЕ ФУНКЦИИ
# ============================================================================

def test_calculator():
    """Тест расчётчика бонусов"""
    calculator = BonusCalculator()

    try:
        # Загружаем шаблоны
        calculator.load_templates('normalized_bonus_template.xlsx')

        # Создаём и загружаем тестовые данные
        calculator.data_loader.create_sample_data_excel()
        calculator.load_kpi_data('sample_kpi_data.xlsx')

        # Рассчитываем бонусы
        results = calculator.calculate_all_bonuses()
        print(f"✓ Рассчитано бонусов для {len(results)} сотрудников")

        for employee, calc in results.items():
            print(f"\n💰 {employee}")
            print(f"   Должность: {calc.position}")
            print(f"   Базовый: {calc.base_total}")
            print(f"   Фактор KPI: {calc.kpi_factor:.3f}")
            print(f"   Фактор демотиваторов: {calc.demotivator_factor:.3f}")
            print(f"   Итоговый бонус: {calc.total_bonus:.2f}")

            print("   KPI результаты:")
            for kpi in calc.kpi_results:
                print(f"     - {kpi.metric_code}: {kpi.result_value:.3f} (взвешен: {kpi.weighted_value:.3f})")

            print("   Демотиваторы:")
            for demo in calc.demotivator_results:
                print(f"     - {demo.code}: {demo.applied_value:.3f}")

        # Экспорт результатов
        calculator.export_results_to_excel(results)

    except Exception as e:
        print(f"✗ Ошибка: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_calculator()