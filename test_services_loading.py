#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Тест загрузки и расчёта бонуса с новой структурой услуг/запчастей"""

from pathlib import Path
from real_bonus_calculator import BonusDataLoader, BonusCalculator

# Инициализация
data_folder = Path(".")
loader = BonusDataLoader(str(data_folder))

print("=" * 100)
print("ТЕСТ ЗАГРУЗКИ ДАННЫХ")
print("=" * 100)

# Загружаем настройки первыми чтобы узнать ДЦ
settings = loader.load_bonus_settings()
print(f"\n⚙️  Настройки:")
print(f"   ДЦ: {settings.dealer_center}")
print(f"   Должность: {settings.position}")
print(f"   Репер: {settings.reper:,.0f} тг")

# Загружаем услуги
services_plan, services_fact = loader.load_services_facts(settings.dealer_center)
print(f"\n📦 Услуги для {settings.dealer_center}:")
print(f"   План:  {services_plan:,.2f}")
print(f"   Факт:  {services_fact:,.2f}")
print(f"   Выполнение: {services_fact/services_plan*100:.1f}%" if services_plan > 0 else "   Выполнение: -")

# Загружаем запчасти
spare_plan, spare_fact = loader.load_spare_parts_facts(settings.dealer_center)
print(f"\n🔧 Запчасти для {settings.dealer_center}:")
print(f"   План:  {spare_plan:,.2f}")
print(f"   Факт:  {spare_fact:,.2f}")
print(f"   Выполнение: {spare_fact/spare_plan*100:.1f}%" if spare_plan > 0 else "   Выполнение: -")

# Загружаем веса
weights = loader.load_metric_weights()
print(f"\n⚖️  Веса показателей (первые 5):")
for i, (name, weight) in enumerate(list(weights.items())[:5]):
    print(f"   {name}: {weight.weight}")

print("\n" + "=" * 100)
print("РАСЧЁТ БОНУСА")
print("=" * 100)

calculator = BonusCalculator(loader)
result = calculator.calculate_for_month("Март")

print(f"\n✅ Результат расчёта за Март:")
print(f"   Оклад (база): {result.base_salary_bonus:,.2f} тг")
print(f"   Услуги ({result.services_execution*100:.1f}%): {result.services_plan_bonus:,.2f} тг")
print(f"   Запчасти ({result.spare_parts_execution*100:.1f}%): {result.spare_parts_bonus:,.2f} тг")
print(f"   Маржинальность: {result.marginality_bonus:,.2f} тг")
print(f"   Неликвид: {result.negliquidity_deduction:,.2f} тг")
print(f"   Другие KPI: {sum(result.other_metric_bonuses.values()):,.2f} тг")
print(f"\n💰 ИТОГО БОНУС: {result.total_bonus:,.2f} тг")

# Таблица показателей
print("\n" + "=" * 100)
print("ТАБЛИЦА ПОКАЗАТЕЛЕЙ")
print("=" * 100)

indicator_df = calculator.get_indicator_df("Март")
print(indicator_df.to_string(index=False))

print("\n✅ Тест завершён успешно!")
