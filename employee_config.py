"""
Конфигурация сотрудников и их показателей для расчета бонусов
"""

from dataclasses import dataclass
from typing import Dict, List
from enum import Enum


class EmployeeType(Enum):
    """Тип сотрудника"""
    MECHANIC = "Механик"  # Тариф × Выработка
    SALES = "Продажи"     # % от репера
    SERVICE = "Сервис"    # % от репера
    WARRANTY = "Гарантия" # % от репера


@dataclass
class KPIIndicator:
    """КПИ показатель"""
    name: str                    # Название показателя (например "Отдел Сервиса")
    plan: float = 0              # Плановое значение
    fact: float = 0              # Фактическое значение
    rate_per_hour: float = 0     # Тариф за норм-час (для механиков)
    weight: float = 1.0          # Вес в расчете бонуса (0-1)
    
    def execution_percent(self) -> float:
        """Процент выполнения плана"""
        if self.plan == 0:
            return 0
        return (self.fact / self.plan) * 100
    
    def calculate_amount(self) -> float:
        """Расчет суммы для этого показателя"""
        if self.rate_per_hour > 0:
            # Для механиков: тариф * выработка
            return self.rate_per_hour * self.fact
        else:
            # Для других: процент от плана
            return self.fact


@dataclass
class Employee:
    """Конфигурация сотрудника"""
    fio: str                              # ФИО
    employee_type: EmployeeType           # Тип сотрудника
    department: str                        # Отдел
    base_salary: float = 0                # Базовая зарплата/репер
    kpi_indicators: Dict[str, KPIIndicator] = None  # КПИ показатели
    bonus_coefficient: float = 1.0        # Коэффициент бонуса (1.0 = 100%)
    
    def __post_init__(self):
        if self.kpi_indicators is None:
            self.kpi_indicators = {}
    
    def add_kpi(self, kpi: KPIIndicator):
        """Добавить КПИ показатель"""
        self.kpi_indicators[kpi.name] = kpi
    
    def calculate_bonus(self, min_execution_threshold: float = 1.0) -> float:
        """
        Расчет бонуса
        
        Args:
            min_execution_threshold: Минимальное выполнение плана для бонуса (1.0 = 100%)
        
        Returns:
            Сумма бонуса
        """
        if not self.kpi_indicators:
            return 0
        
        if self.employee_type == EmployeeType.MECHANIC:
            # Для механиков: сумма тариф*выработка
            return sum(
                kpi.rate_per_hour * kpi.fact 
                for kpi in self.kpi_indicators.values()
            ) * self.bonus_coefficient
        else:
            # Для остальных: % от репера в зависимости от выполнения
            total_execution = 0
            total_weight = 0
            
            for kpi in self.kpi_indicators.values():
                execution = kpi.execution_percent() / 100  # Переводим в доли
                total_execution += execution * kpi.weight
                total_weight += kpi.weight
            
            avg_execution = total_execution / total_weight if total_weight > 0 else 0
            
            if avg_execution >= min_execution_threshold:
                # Бонус = репер * (выполнение - пороговое значение) * коэффициент
                bonus_percent = min(avg_execution - (min_execution_threshold - 1), 1.0)
                return self.base_salary * bonus_percent * self.bonus_coefficient
            else:
                return 0
    
    def get_summary(self) -> Dict:
        """Получить сводку по сотруднику"""
        return {
            'ФИО': self.fio,
            'Отдел': self.department,
            'Тип': self.employee_type.value,
            'Репер': self.base_salary,
            'КПИ_всего': len(self.kpi_indicators),
            'Бонус': self.calculate_bonus(),
        }


# ============================================================================
# КОНФИГУРАЦИЯ СОТРУДНИКОВ (ЗАПОЛНЯЕТСЯ ВРУЧНУЮ)
# ============================================================================

EMPLOYEES_CONFIG: List[Employee] = [
    # ПРИМЕР 1: Механик
    Employee(
        fio="Баткошев Адильжан",
        employee_type=EmployeeType.MECHANIC,
        department="Сервис",
        base_salary=119737,
    ),
    
    # ПРИМЕР 2: Сотрудник сервиса (% от репера)
    Employee(
        fio="Егинбаев Аслан Асемович",
        employee_type=EmployeeType.SERVICE,
        department="Сервис",
        base_salary=119737,
    ),
]


def get_employee_config(fio: str) -> Employee:
    """Получить конфигурацию сотрудника по ФИО или создать новую"""
    for emp in EMPLOYEES_CONFIG:
        if emp.fio.lower() == fio.lower():
            return emp
    
    # Если не найден, создаем новую запись
    new_emp = Employee(
        fio=fio,
        employee_type=EmployeeType.SERVICE,  # По умолчанию
        department="Unknown",
        base_salary=0,
    )
    return new_emp


def update_employee_config(employee: Employee):
    """Обновить конфигурацию сотрудника"""
    for i, emp in enumerate(EMPLOYEES_CONFIG):
        if emp.fio.lower() == employee.fio.lower():
            EMPLOYEES_CONFIG[i] = employee
            return True
    
    EMPLOYEES_CONFIG.append(employee)
    return True
