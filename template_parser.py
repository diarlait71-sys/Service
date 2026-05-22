"""
Парсер шаблонов расчёта бонусов из Excel
"""

import pandas as pd
import json
from typing import Dict, List, Optional, Any
from dataclasses import dataclass
from pathlib import Path


@dataclass
class KPITemplate:
    """Шаблон KPI показателя"""
    code: str
    name: str
    plan: float
    weight: float
    formula_type: str
    formula_params: Dict[str, Any]


@dataclass
class DemotivatorTemplate:
    """Шаблон демотиватора"""
    code: str
    name: str
    demotivator_type: str  # 'penalty' or 'reward'
    value_type: str  # 'coefficient' or 'fixed'
    rule: str
    value: float


@dataclass
class BonusTemplate:
    """Полный шаблон расчёта бонусов для должности"""
    position: str
    dealer_center_level: str
    base_salary: float
    budget_bonus: float
    kpis: List[KPITemplate]
    demotivators: List[DemotivatorTemplate]


class TemplateParser:
    """Парсер шаблонов из Excel"""

    def __init__(self):
        self.templates: Dict[str, BonusTemplate] = {}

    def parse_excel_template(self, file_path: str) -> Dict[str, BonusTemplate]:
        """
        Парсит Excel файл с шаблоном расчёта бонусов

        Args:
            file_path: Путь к Excel файлу

        Returns:
            Словарь шаблонов по должностям
        """
        try:
            df = pd.read_excel(file_path, sheet_name=0)

            # Группируем по должностям
            positions = df['Position'].dropna().unique()

            for position in positions:
                position_data = df[df['Position'] == position]

                # Основная информация (первая строка)
                main_row = position_data.iloc[0]
                dealer_level = main_row.get('Dealer Center Level', 'A')
                base_salary = main_row.get('Base Salary', 0)
                budget_bonus = main_row.get('Budget Bonus', 0)

                # KPI показатели
                kpis = self._parse_kpis(position_data)

                # Демотиваторы
                demotivators = self._parse_demotivators(position_data)

                # Создаём шаблон
                template = BonusTemplate(
                    position=str(position),
                    dealer_center_level=str(dealer_level),
                    base_salary=float(base_salary),
                    budget_bonus=float(budget_bonus),
                    kpis=kpis,
                    demotivators=demotivators
                )

                self.templates[position] = template

            return self.templates

        except Exception as e:
            raise ValueError(f"Ошибка при парсинге шаблона: {e}")

    def _parse_kpis(self, df: pd.DataFrame) -> List[KPITemplate]:
        """Парсит KPI показатели из DataFrame"""
        kpis = []

        # Фильтруем строки с KPI
        kpi_rows = df[df['KPI Code'].notna()]

        for _, row in kpi_rows.iterrows():
            try:
                # Парсим параметры формулы из JSON
                formula_params = {}
                if pd.notna(row.get('KPI Formula Params')):
                    try:
                        formula_params = json.loads(str(row['KPI Formula Params']))
                    except json.JSONDecodeError:
                        formula_params = {}

                kpi = KPITemplate(
                    code=str(row['KPI Code']),
                    name=str(row['KPI Name']),
                    plan=float(row['KPI Plan']) if pd.notna(row.get('KPI Plan')) else 0,
                    weight=float(row['KPI Weight']) if pd.notna(row.get('KPI Weight')) else 0,
                    formula_type=str(row['KPI Formula Type']) if pd.notna(row.get('KPI Formula Type')) else 'ratio',
                    formula_params=formula_params
                )
                kpis.append(kpi)

            except Exception as e:
                print(f"Ошибка при парсинге KPI {row.get('KPI Code', 'unknown')}: {e}")
                continue

        return kpis

    def _parse_demotivators(self, df: pd.DataFrame) -> List[DemotivatorTemplate]:
        """Парсит демотиваторы из DataFrame"""
        demotivators = []

        # Фильтруем строки с демотиваторами
        demo_rows = df[df['Demotivator Code'].notna()]

        for _, row in demo_rows.iterrows():
            try:
                demotivator = DemotivatorTemplate(
                    code=str(row['Demotivator Code']),
                    name=str(row['Demotivator Name']),
                    demotivator_type=str(row['Demotivator Type']) if pd.notna(row.get('Demotivator Type')) else 'penalty',
                    value_type=str(row['Demotivator Value Type']) if pd.notna(row.get('Demotivator Value Type')) else 'coefficient',
                    rule=str(row['Demotivator Rule']) if pd.notna(row.get('Demotivator Rule')) else '',
                    value=float(row['Demotivator Value']) if pd.notna(row.get('Demotivator Value')) else 0
                )
                demotivators.append(demotivator)

            except Exception as e:
                print(f"Ошибка при парсинге демотиватора {row.get('Demotivator Code', 'unknown')}: {e}")
                continue

        return demotivators

    def get_template(self, position: str) -> Optional[BonusTemplate]:
        """Получить шаблон по должности"""
        return self.templates.get(position)

    def list_positions(self) -> List[str]:
        """Получить список всех должностей в шаблонах"""
        return list(self.templates.keys())

    def to_json(self) -> str:
        """Экспорт всех шаблонов в JSON"""
        templates_dict = {}
        for position, template in self.templates.items():
            templates_dict[position] = {
                'position': template.position,
                'dealer_center_level': template.dealer_center_level,
                'base_components': {
                    'base_salary': template.base_salary,
                    'budget_bonus': template.budget_bonus
                },
                'kpis': [
                    {
                        'code': kpi.code,
                        'name': kpi.name,
                        'weight': kpi.weight,
                        'plan': kpi.plan,
                        'formula': {
                            'type': kpi.formula_type,
                            **kpi.formula_params
                        }
                    } for kpi in template.kpis
                ],
                'demotivators': [
                    {
                        'code': demo.code,
                        'name': demo.name,
                        'type': demo.demotivator_type,
                        'value_type': demo.value_type,
                        'rule': demo.rule,
                        'value': demo.value
                    } for demo in template.demotivators
                ]
            }

        return json.dumps(templates_dict, indent=2, ensure_ascii=False)


# ============================================================================
# ТЕСТОВЫЕ ФУНКЦИИ
# ============================================================================

def test_parser():
    """Тест парсера"""
    parser = TemplateParser()

    try:
        templates = parser.parse_excel_template('normalized_bonus_template.xlsx')
        print(f"✓ Загружено шаблонов: {len(templates)}")

        for position, template in templates.items():
            print(f"\n📋 Шаблон: {position}")
            print(f"   Уровень ДЦ: {template.dealer_center_level}")
            print(f"   Базовая зарплата: {template.base_salary}")
            print(f"   Бюджетный бонус: {template.budget_bonus}")
            print(f"   KPI показателей: {len(template.kpis)}")
            print(f"   Демотиваторов: {len(template.demotivators)}")

            print("   KPI:")
            for kpi in template.kpis:
                print(f"     - {kpi.code}: {kpi.name} (вес: {kpi.weight})")

            print("   Демотиваторы:")
            for demo in template.demotivators:
                print(f"     - {demo.code}: {demo.name} ({demo.demotivator_type})")

        # Экспорт в JSON
        json_output = parser.to_json()
        with open('templates_export.json', 'w', encoding='utf-8') as f:
            f.write(json_output)
        print(f"\n✓ Шаблоны экспортированы в templates_export.json")

    except Exception as e:
        print(f"✗ Ошибка: {e}")


if __name__ == "__main__":
    test_parser()