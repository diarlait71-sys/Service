from pathlib import Path

from real_bonus_calculator import BonusCalculator, BonusSettings, MetricWeight


class MockBonusDataLoader:
    def load_bonus_settings(self, file_name="Дц,должность,репер.xlsx", dealer_center=None, position=None):
        return BonusSettings(
            dealer_center=dealer_center or "01-CJ-3S",
            position=position or "РОС",
            reper=100000,
            manager_name="Тест",
            weights_profile="",
            logic_profile="",
        )

    def resolve_weights_file(self, position: str, weights_profile: str = ""):
        return Path("weights.xlsx")

    def resolve_weights_sheet(self, file_path, position: str = "", weights_profile: str = "") -> str:
        return "Лист1"

    def resolve_logic_file(self, logic_profile: str = ""):
        return Path("logic.xlsx")

    def load_metric_weights(self, position=None, weights_profile=""):
        return {
            "Оклад (База)": MetricWeight("Оклад (База)", 0.50, ""),
            "Выполнение плана продаж по продажи услуг": MetricWeight("Выполнение плана продаж по продажи услуг", 0.23, ""),
            "Выполнение плана продаж по Запасным Частям": MetricWeight("Выполнение плана продаж по Запасным Частям", 0.23, ""),
            "Выполнение  маржанальности": MetricWeight("Выполнение  маржанальности", 0.04, ""),
            "Соблюдение уровня неликвида на складе запасных частей": MetricWeight("Соблюдение уровня неликвида на складе запасных частей", -0.10, ""),
        }

    def load_calculation_logic(self, file_name="Формула расчета.xlsx", logic_profile=""):
        return {}

    def load_other_plans(self, file_name="план по остальным показателям.xlsx"):
        return {
            "Выполнение  маржанальности": "0.3",
            "Соблюдение уровня неликвида на складе запасных частей": "0.07",
        }

    def load_services_facts(self, dealer_center, file_name="Факт Услуги.xlsx"):
        return 100000.0, 100000.0

    def load_spare_parts_facts(self, dealer_center, file_name="Факт Запасные части.xlsx"):
        return 100000.0, 100000.0

    def load_marginality_facts(self, file_name="Факт маржанальность.xlsx"):
        return {"01-CJ-3S": 0.30}

    def load_other_facts(self, dealer_center, file_name="Факт по остальным показателям.xlsx"):
        return {"Соблюдение уровня неликвида на складе запасных частей": "0.05"}


def test_ros_full_plan():
    loader = MockBonusDataLoader()
    calc = BonusCalculator(loader, dealer_center="01-CJ-3S", position="РОС")
    result = calc.calculate_for_month("Январь")

    expected_bonus = loader.load_bonus_settings().reper * (0.23 + 0.23 + 0.04)
    assert abs(result.total_bonus - expected_bonus) < 0.01


def test_negliquidity_penalty():
    class PenalizedLoader(MockBonusDataLoader):
        def load_other_facts(self, dealer_center, file_name="Факт по остальным показателям.xlsx"):
            return {"Соблюдение уровня неликвида на складе запасных частей": "0.10"}

    loader = PenalizedLoader()
    calc = BonusCalculator(loader, dealer_center="01-CJ-3S", position="РОС")
    result = calc.calculate_for_month("Январь")

    assert result.negliquidity_deduction == -loader.load_bonus_settings().reper * 0.10