"""
Модуль для работы с реальной структурой данных из PowerBI/1C
Парсит данные о бонусах руководителей отдела сервиса
"""

from dataclasses import dataclass, field
import re
import time
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from pathlib import Path
from pydantic import ValidationError

from config_loader import load_app_config
from models import BonusCalcRequest, KPIMetricInput
from utils.logger import audit_event, setup_bonus_logger


CONFIG = load_app_config()
LOGGER = setup_bonus_logger(CONFIG["logging"]["file"])


@dataclass
class BonusSettings:
    """Настройки расчёта бонуса для должности"""
    dealer_center: str
    position: str
    reper: float  # Реперный показатель - базовая сумма
    manager_name: str = ""
    weights_profile: str = ""
    logic_profile: str = ""
    fixed_salary: float = 0.0
    bonus_pool: float = 0.0


@dataclass
class MetricWeight:
    """Вес показателя в общей сумме"""
    metric_name: str
    weight: float
    calculation_logic: str
    raw_weight: str = ""
    description: str = ""
    penalty_flag: str = ""


@dataclass
class MonthlyPlan:
    """Месячный план"""
    metric_name: str
    month: str
    plan_value: float


@dataclass
class MonthlyFact:
    """Месячный факт"""
    dealer_center: str
    brand: str
    metric_name: str
    plan_value: float
    fact_value: float
    execution_percent: float


@dataclass
class BonusCalculationResult:
    """Результат расчёта бонуса"""
    dealer_center: str
    position: str
    reper: float
    month: str
    
    # Компоненты бонуса
    base_salary_bonus: float  # Оклад (50%)
    services_plan_bonus: float  # План услуг (23%)
    spare_parts_bonus: float  # План запчастей (23%)
    marginality_bonus: float  # Маржинальность (4%)
    negliquidity_deduction: float  # Неликвид (-10% демотиватор)
    
    # Итого
    total_bonus: float
    
    # Справочная информация
    services_plan_value: float  # План услуг (сумма)
    services_fact_value: float  # Факт услуг (сумма)
    services_execution: float  # % выполнения плана услуг
    spare_parts_plan_value: float  # План запчастей (сумма)
    spare_parts_fact_value: float  # Факт запчастей (сумма)
    spare_parts_execution: float  # % выполнения плана запчастей
    marginality_fact: float  # Фактическая маржинальность
    negliquidity_fact: float  # Фактический уровень неликвида
    other_metric_bonuses: Dict[str, float] = field(default_factory=dict)


class DataLoadingError(Exception):
    """Базовое исключение для ошибок загрузки/валидации данных."""


class MissingColumnError(DataLoadingError):
    def __init__(self, file: str, expected_columns: List[str], found_columns: List[str]):
        self.file_name = file
        self.expected_columns = expected_columns
        self.found_columns = found_columns
        super().__init__(
            f"Файл '{file}': не найдены колонки {expected_columns}. Найдено: {found_columns}"
        )


@st.cache_data(ttl=CONFIG["cache"]["static_ttl_seconds"])
def _read_excel_cached(file_path: str, sheet_name: str, header: Any = 0) -> pd.DataFrame:
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header)


@st.cache_data(ttl=CONFIG["cache"]["static_ttl_seconds"])
def load_cached_settings(data_folder: str) -> List[BonusSettings]:
    loader = BonusDataLoader(data_folder)
    return loader.load_all_settings()


@st.cache_data(ttl=CONFIG["cache"]["facts_ttl_seconds"])
def load_cached_facts(data_folder: str, dealer_center: str, file_type: str) -> dict:
    loader = BonusDataLoader(data_folder)
    if file_type == "services":
        plan, fact = loader.load_services_facts(dealer_center)
        return {"plan": plan, "fact": fact}
    if file_type == "spare_parts":
        plan, fact = loader.load_spare_parts_facts(dealer_center)
        return {"plan": plan, "fact": fact}
    if file_type == "marginality":
        return loader.load_marginality_facts()
    if file_type == "other":
        return loader.load_other_facts(dealer_center)
    if file_type == "service_consultant":
        fio_facts = loader.load_service_consultant_fio_facts(dealer_center)
        fio_count = len(fio_facts)
        services_total = sum(v.get("services_fact", 0.0) for v in fio_facts.values())
        spare_total = sum(v.get("spare_parts_fact", 0.0) for v in fio_facts.values())
        services_avg = services_total / fio_count if fio_count > 0 else 0.0
        spare_avg = spare_total / fio_count if fio_count > 0 else 0.0
        return {
            "services_fact": services_avg,
            "spare_parts_fact": spare_avg,
            "fio_count": fio_count,
            "fio_facts": fio_facts,
        }
    return {}


@st.cache_data(ttl=CONFIG["cache"]["facts_ttl_seconds"])
def load_cached_marginality(data_folder: str) -> Dict[str, float]:
    """Кэшированная маржинальность для всех ДЦ — один вызов на папку, без привязки к конкретному ДЦ."""
    loader = BonusDataLoader(data_folder)
    return loader.load_marginality_facts()


class BonusDataLoader:
    """Загружает данные о бонусах из структуры файлов"""
    
    def __init__(self, data_folder: str):
        self.data_folder = Path(data_folder)

    def _read_excel(self, file_name: str, sheet_name: str, header: Any = 0) -> pd.DataFrame:
        file_path = self.data_folder / file_name
        try:
            return _read_excel_cached(str(file_path), sheet_name, header).copy()
        except FileNotFoundError as exc:
            raise DataLoadingError(f"Файл не найден: '{file_name}' в папке '{self.data_folder}'") from exc

    @staticmethod
    def _pick_column(df: pd.DataFrame, candidates: List[str], fallback_index: int = 0) -> str:
        """Возвращает первое найденное имя колонки из списка кандидатов."""
        normalized = {str(c).strip().lower(): c for c in df.columns}
        for candidate in candidates:
            key = candidate.strip().lower()
            if key in normalized:
                return normalized[key]
        return df.columns[fallback_index]

    @staticmethod
    def _norm_text(value) -> str:
        """Нормализует текстовое значение для сравнений и валидации."""
        if pd.isna(value):
            return ""
        return str(value).strip()

    @staticmethod
    def _parse_float(value) -> Optional[float]:
        """Парсит число из Excel-ячейки с поддержкой запятых и пробелов."""
        if pd.isna(value):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        cleaned = (
            str(value)
            .strip()
            .replace("\u00a0", "")
            .replace(" ", "")
            .replace("%", "")
            .replace(",", ".")
        )
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None

    @staticmethod
    def _extract_profile_name(value: str) -> str:
        """Преобразует значение профиля в чистое имя без расширения .xlsx."""
        text = str(value).strip()
        if not text:
            return ""
        if text.lower().endswith(".xlsx"):
            text = text[:-5]
        return text.strip()

    @staticmethod
    def _normalize_lookup(text: str) -> str:
        return " ".join(str(text).strip().lower().split())

    def _match_named_sheet(self, sheet_names: List[str], position: str = "", profile_name: str = "") -> Optional[str]:
        normalized_map = {self._normalize_lookup(s): s for s in sheet_names}
        normalized_profile = self._normalize_lookup(self._extract_profile_name(profile_name))
        normalized_position = self._normalize_lookup(position)

        if normalized_profile and normalized_profile in normalized_map:
            return normalized_map[normalized_profile]
        if normalized_position and normalized_position in normalized_map:
            return normalized_map[normalized_position]

        if normalized_position:
            for norm_name, original in normalized_map.items():
                if normalized_position in norm_name or norm_name in normalized_position:
                    return original
            if "сервис" in normalized_position and "консульт" in normalized_position:
                for norm_name, original in normalized_map.items():
                    if "сервис" in norm_name and "консульт" in norm_name:
                        return original

        return None

    def _find_profile_column(self, df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
        """Находит колонку профиля, если она есть в таблице настроек."""
        normalized = {str(c).strip().lower(): c for c in df.columns}
        for candidate in candidates:
            key = candidate.strip().lower()
            if key in normalized:
                return normalized[key]
        return None

    def _pick_column_or_raise(self, df: pd.DataFrame, file_name: str, candidates: List[str]) -> str:
        column = self._find_profile_column(df, candidates)
        if column is None:
            raise MissingColumnError(file_name, candidates, [str(c) for c in df.columns])
        return column

    def count_position_staff(self, dealer_center: str, position: str,
                             file_name: str = "Дц,должность,репер.xlsx") -> int:
        """Считает количество сотрудников по ДЦ+должности из исходного справочника (без дедупликации)."""
        file_path = self.data_folder / file_name
        if not file_path.exists():
            return 1

        try:
            df = _read_excel_cached(str(file_path), "Лист1", 0).copy()
        except Exception:
            # В некоторых файлах лист может называться иначе.
            xl = pd.ExcelFile(file_path)
            if not xl.sheet_names:
                return 1
            df = _read_excel_cached(str(file_path), xl.sheet_names[0], 0).copy()

        dc_col = self._find_profile_column(df, ["Диллерский центр", "Дилерский центр", "Название ДЦ", "ДЦ"])
        pos_col = self._find_profile_column(df, ["Должность", "Должность общий для свода", "Должность для свода"])
        if dc_col is None or pos_col is None:
            return 1

        dc_norm = self._norm_text(dealer_center).lower()
        pos_norm = self._norm_text(position).lower()
        if not dc_norm or not pos_norm:
            return 1

        is_service_consultant = ("сервис" in pos_norm and "консульт" in pos_norm)
        count = 0
        for _, row in df.iterrows():
            row_dc = self._norm_text(row.get(dc_col, "")).lower()
            row_pos = self._norm_text(row.get(pos_col, "")).lower()
            if row_dc != dc_norm:
                continue

            if is_service_consultant:
                # Для сервис-консультантов считаем все варианты написания в рамках ДЦ.
                if "сервис" in row_pos and "консульт" in row_pos:
                    count += 1
            elif row_pos == pos_norm:
                count += 1

        return max(count, 1)

    def load_all_settings(self, file_name: str = "Дц,должность,репер.xlsx") -> List[BonusSettings]:
        """Загружает настройки для всех должностей/ДЦ из файла (все строки)"""
        df = self._read_excel(file_name, sheet_name="Лист1")
        dc_col = self._pick_column_or_raise(df, file_name, ["Диллерский центр", "Дилерский центр", "Название ДЦ", "ДЦ"])
        pos_col = self._pick_column_or_raise(df, file_name, ["Должность", "Должность общий для свода", "Должность для свода"])
        reper_col = self._find_profile_column(df, ["Репер", "Оклад", "База", "Фикс", "Бонус"])
        if reper_col is None:
            raise MissingColumnError(file_name, ["Репер", "Фикс", "Бонус"], [str(c) for c in df.columns])
        fix_col = self._find_profile_column(df, ["Фикс"])
        bonus_col = self._find_profile_column(df, ["Бонус"])
        fio_col = self._find_profile_column(df, ["ФИО руководителя отдела сервиса", "ФИО", "Руководитель", "ФИО руководителя"])
        weights_profile_col = self._find_profile_column(df, ["Профиль весов", "Ключ весов", "Шаблон весов"])
        logic_profile_col = self._find_profile_column(df, ["Профиль логики", "Ключ логики", "Шаблон логики"])
        result = []
        seen_pairs = set()
        for _, row in df.iterrows():
            try:
                dealer_center = self._norm_text(row[dc_col])
                position = self._norm_text(row[pos_col])
                manager_name = self._norm_text(row[fio_col]) if fio_col else ""
                reper_value = self._parse_float(row[reper_col])
                fix_val = self._parse_float(row[fix_col]) if fix_col else None
                bonus_val = self._parse_float(row[bonus_col]) if bonus_col else None

                # Если используется колонка Фикс/Бонус вместо Репера.
                if reper_value is None and str(reper_col).strip().lower() in {"фикс", "бонус"}:
                    if fix_val is not None or bonus_val is not None:
                        reper_value = (fix_val or 0.0) + (bonus_val or 0.0)

                if reper_value is None and fix_val is not None and bonus_val is not None:
                    reper_value = fix_val + bonus_val
                weights_profile = self._extract_profile_name(self._norm_text(row[weights_profile_col])) if weights_profile_col else ""
                logic_profile = self._extract_profile_name(self._norm_text(row[logic_profile_col])) if logic_profile_col else ""

                # Пропускаем пустые/служебные строки и строки без репера.
                if dealer_center.lower() in ("", "nan", "none"):
                    continue
                if position.lower() in ("", "nan", "none"):
                    continue
                if reper_value is None:
                    continue

                # Защита от дублей комбинации ДЦ+должность в расширенном справочнике.
                pair_key = (dealer_center.lower(), position.lower())
                if pair_key in seen_pairs:
                    continue
                seen_pairs.add(pair_key)

                result.append(BonusSettings(
                    dealer_center=dealer_center,
                    position=position,
                    reper=reper_value,
                    manager_name=manager_name,
                    weights_profile=weights_profile,
                    logic_profile=logic_profile,
                    fixed_salary=fix_val or 0.0,
                    bonus_pool=bonus_val or 0.0,
                ))
            except KeyError:
                continue
        return result

    def load_bonus_settings(self, file_name: str = "Дц,должность,репер.xlsx",
                            dealer_center: str = None, position: str = None) -> BonusSettings:
        """Загружает настройки (ДЦ, должность, репер). 
        Если dealer_center и position указаны — ищет конкретную строку."""
        all_settings = self.load_all_settings(file_name=file_name)
        if not all_settings:
            raise ValueError("В файле настроек нет валидных строк для расчёта")

        if dealer_center and position:
            for s in all_settings:
                if s.dealer_center == dealer_center and s.position == position:
                    return s

        return all_settings[0]

    def resolve_weights_file(self, position: str, weights_profile: str = "") -> Path:
        """Определяет файл весов по профилю или должности."""
        matches = sorted(self.data_folder.glob("Доля показателей для*.xlsx"))
        if not matches:
            raise FileNotFoundError("Не найден файл(ы) весов: 'Доля показателей для*.xlsx'")

        default_path = self.data_folder / "Доля показателей для Руководителя отдела сервиса.xlsx"
        fallback = default_path if default_path.exists() else matches[0]

        profile = self._extract_profile_name(weights_profile)
        if profile:
            direct_name = self.data_folder / f"{profile}.xlsx"
            if direct_name.exists():
                return direct_name
            prefixed_name = self.data_folder / f"Доля показателей для {profile}.xlsx"
            if prefixed_name.exists():
                return prefixed_name

            normalized_profile = " ".join(profile.lower().split())
            for candidate in matches:
                candidate_name = candidate.stem.lower()
                if normalized_profile and normalized_profile in candidate_name:
                    return candidate

        normalized_position = " ".join(str(position).strip().lower().split())
        for candidate in matches:
            suffix = candidate.stem.replace("Доля показателей для", "", 1).strip().lower()
            suffix = " ".join(suffix.split())
            if suffix == normalized_position:
                return candidate

        for candidate in matches:
            if normalized_position and normalized_position in candidate.stem.lower():
                return candidate

        return fallback

    def resolve_weights_sheet(self, file_path: Path, position: str, weights_profile: str = "") -> str:
        """Определяет лист весов внутри файла: приоритет профиль -> должность -> Лист1 -> первый лист."""
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        if not sheet_names:
            raise DataLoadingError(f"Файл весов '{file_path.name}' не содержит листов")

        matched_sheet = self._match_named_sheet(sheet_names, position=position, profile_name=weights_profile)
        normalized_map = {self._normalize_lookup(s): s for s in sheet_names}

        if matched_sheet:
            return matched_sheet
        if "лист1" in normalized_map:
            return normalized_map["лист1"]
        return sheet_names[0]

    def has_explicit_weights_for_position(self, position: str = "", weights_profile: str = "") -> bool:
        file_path = self.resolve_weights_file(position=position, weights_profile=weights_profile)
        if not file_path.exists():
            return False

        if self._extract_profile_name(weights_profile):
            return True

        normalized_position = self._normalize_lookup(position)
        file_suffix = self._normalize_lookup(file_path.stem.replace("Доля показателей для", "", 1))
        if normalized_position and (file_suffix == normalized_position or normalized_position in file_suffix):
            return True

        xl = pd.ExcelFile(file_path)
        return self._match_named_sheet(xl.sheet_names, position=position) is not None
    
    def load_metric_weights(self, position: str = None, weights_profile: str = "") -> Dict[str, MetricWeight]:
        """Загружает веса показателей.
        Ищет файл по имени должности: 'Доля показателей для {position}.xlsx'.
        Если не найден — использует файл руководителя отдела сервиса."""
        file_path = self.resolve_weights_file(position=position or "", weights_profile=weights_profile)
        sheet_name = self.resolve_weights_sheet(file_path, position=position or "", weights_profile=weights_profile)
        file_name = file_path.name
        df = _read_excel_cached(str(file_path), sheet_name, 0).copy()
        metric_col = self._pick_column_or_raise(df, file_name, ["Показатели", "KPI", "Показатель", "Описание"])
        weight_col = self._pick_column_or_raise(df, file_name, ["Доля от репера", "Вес", "Доля", "Ставка (к начислению)"])
        weight_col_name = str(weight_col).strip().lower()
        
        weights = {}
        metrics_seen = 0
        for _, row in df.iterrows():
            metric_name = self._norm_text(row.get(metric_col, ""))
            raw_weight = self._norm_text(row.get(weight_col, ""))
            weight = self._parse_float(raw_weight)
            if not metric_name:
                continue
            metrics_seen += 1

            if weight is None:
                weight = 0.0

            # Если вес задан в процентах (например 50%), приводим к доле (0.5).
            # Для колонок ставок (фикс в тенге) деление НЕ применяем.
            looks_like_percent = ("%" in raw_weight) or ("доля" in weight_col_name) or ("вес" in weight_col_name)
            if weight > 1 and looks_like_percent:
                weight = weight / 100.0

            weights[metric_name] = MetricWeight(
                metric_name=metric_name,
                weight=weight,
                calculation_logic="",
                raw_weight=raw_weight,
                description=self._norm_text(row.get("Описание", "")),
                penalty_flag=self._norm_text(row.get("Штрафы", "")),
            )

        if metrics_seen == 0:
            raise DataLoadingError(
                f"Лист '{sheet_name}' в файле '{file_name}' не содержит метрик KPI (пустой лист)."
            )
        
        return weights

    def resolve_logic_file(self, logic_profile: str = "") -> Path:
        """Определяет файл логики по профилю или использует базовый."""
        profile = self._extract_profile_name(logic_profile)
        default_logic = self.data_folder / "Формула расчета.xlsx"

        if profile:
            candidates = [
                self.data_folder / f"{profile}.xlsx",
                self.data_folder / f"Формула расчета - {profile}.xlsx",
                self.data_folder / f"Формула расчета {profile}.xlsx",
            ]
            for candidate in candidates:
                if candidate.exists():
                    return candidate

        return default_logic

    def resolve_logic_sheet(self, file_path: Path, position: str = "", logic_profile: str = "") -> str:
        """Определяет лист логики: профиль -> должность -> Лист1 -> Рос -> первый лист."""
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        if not sheet_names:
            raise DataLoadingError(f"Файл логики '{file_path.name}' не содержит листов")

        matched_sheet = self._match_named_sheet(sheet_names, position=position, profile_name=logic_profile)
        normalized_map = {self._normalize_lookup(s): s for s in sheet_names}

        if matched_sheet:
            return matched_sheet

        if "лист1" in normalized_map:
            return normalized_map["лист1"]
        if "рос" in normalized_map:
            return normalized_map["рос"]
        return sheet_names[0]

    def has_explicit_logic_for_position(self, position: str = "", logic_profile: str = "") -> bool:
        file_path = self.resolve_logic_file(logic_profile=logic_profile)
        if not file_path.exists():
            return False

        if self._extract_profile_name(logic_profile):
            return True

        xl = pd.ExcelFile(file_path)
        return self._match_named_sheet(xl.sheet_names, position=position) is not None
    
    def load_calculation_logic(self, file_name: str = "Формула расчета.xlsx",
                               logic_profile: str = "", position: str = "") -> Dict[str, str]:
        """Загружает логику расчёта для каждого показателя"""
        file_path = self.resolve_logic_file(logic_profile=logic_profile)
        if not file_path.exists():
            file_path = self.data_folder / file_name
            if not file_path.exists():
                return {}
        sheet_name = self.resolve_logic_sheet(file_path, position=position, logic_profile=logic_profile)
        df = _read_excel_cached(str(file_path), sheet_name, 0).copy()
        metric_col = self._pick_column_or_raise(df, file_path.name, ["Показатели"])
        logic_col = self._pick_column_or_raise(df, file_path.name, ["Логика расчета"])
        
        logic = {}
        for _, row in df.iterrows():
            metric_name = self._norm_text(row.get(metric_col, ""))
            calc_logic = self._norm_text(row.get(logic_col, ""))
            if not metric_name:
                continue
            logic[metric_name] = calc_logic
        
        return logic
    
    def load_monthly_plans(self, file_name: str = "План продаж услуг и запчастей.xlsx") -> Dict[str, Dict[str, float]]:
        """Загружает месячные планы"""
        df = self._read_excel(file_name, sheet_name="План")
        metric_col = self._pick_column_or_raise(df, file_name, ["Показатель"])
        
        plans = {}
        
        for _, row in df.iterrows():
            metric_name = self._norm_text(row.get(metric_col, ""))
            if not metric_name:
                continue
            plans[metric_name] = {}
            
            # Считываем по месяцам
            months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                     "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
            
            for month in months:
                if month in df.columns:
                    value = self._parse_float(row[month])
                    plans[metric_name][month] = value if value is not None else 0.0
        
        return plans
    
    def _find_header_row(self, df: pd.DataFrame, header_name: str) -> Optional[int]:
        """Находит строку с заголовками, сравнивая точные значения ячеек."""
        for i, row in df.iterrows():
            for cell in row.values:
                if isinstance(cell, str) and cell.strip() == header_name:
                    return i
        return None
    
    def load_other_plans(self, file_name: str = "план по остальным показателям.xlsx") -> Dict[str, str]:
        """Загружает планы для остальных показателей (маржинальность, неликвид и т.д.)"""
        df = self._read_excel(file_name, sheet_name="Лист1")
        name_col = self._pick_column_or_raise(df, file_name, ["Наименование"])
        value_col = self._pick_column_or_raise(df, file_name, ["Плановый значение", "Плановое значение"])
        
        plans = {}
        for _, row in df.iterrows():
            name = self._norm_text(row.get(name_col, ""))
            value = self._norm_text(row.get(value_col, ""))
            if not name:
                continue
            plans[name] = value
        
        return plans
    
    def load_services_facts(self, dealer_center: str, file_name: str = "Факт Услуги.xlsx") -> Tuple[float, float]:
        """Загружает фактические данные по услугам для конкретного ДЦ, суммирует план и факт.
        Возвращает кортеж (plan_total, fact_total)"""
        df = self._read_excel(file_name, sheet_name="Sheet1", header=None)
        
        # Находим строку с заголовками (ищем точное совпадение "Название ДЦ")
        header_row = None
        for i, val in enumerate(df[0]):
            if isinstance(val, str) and val.strip() == "Название ДЦ":
                header_row = i
                break
        
        if header_row is None:
            raise MissingColumnError(file_name, ["Название ДЦ"], [str(v) for v in df.iloc[0].tolist()])
        
        # Устанавливаем заголовки и берём данные после header_row
        df.columns = df.iloc[header_row].values
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        dc_col = self._pick_column_or_raise(df, file_name, ["Название ДЦ"])
        plan_col = self._pick_column_or_raise(df, file_name, ["План текущий"])
        fact_col = self._pick_column_or_raise(df, file_name, ["Факт"])
        
        # Суммируем план и факт только для нужного ДЦ
        plan_total = 0.0
        fact_total = 0.0
        
        for _, row in df.iterrows():
            dc = str(row.get(dc_col, "")).strip()
            if dc != dealer_center:
                continue
            
            try:
                plan_val = row.get(plan_col, 0)
                fact_val = row.get(fact_col, 0)
                
                if pd.notna(plan_val):
                    plan_total += float(plan_val)
                if pd.notna(fact_val):
                    fact_total += float(fact_val)
            except (ValueError, TypeError):
                continue
        
        return plan_total, fact_total
    
    def load_spare_parts_facts(self, dealer_center: str, file_name: str = "Факт Запасные части.xlsx") -> Tuple[float, float]:
        """Загружает фактические данные по запчастям для конкретного ДЦ, суммирует план и факт.
        Возвращает кортеж (plan_total, fact_total)"""
        df = self._read_excel(file_name, sheet_name="Sheet1", header=None)
        
        # Находим строку с заголовками (ищем точное совпадение "Название ДЦ")
        header_row = None
        for i, val in enumerate(df[0]):
            if isinstance(val, str) and val.strip() == "Название ДЦ":
                header_row = i
                break
        
        if header_row is None:
            raise MissingColumnError(file_name, ["Название ДЦ"], [str(v) for v in df.iloc[0].tolist()])
        
        # Устанавливаем заголовки и берём данные после header_row
        df.columns = df.iloc[header_row].values
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        dc_col = self._pick_column_or_raise(df, file_name, ["Название ДЦ"])
        plan_col = self._pick_column_or_raise(df, file_name, ["План текущий"])
        fact_col = self._pick_column_or_raise(df, file_name, ["Факт"])
        
        # Суммируем план и факт только для нужного ДЦ
        plan_total = 0.0
        fact_total = 0.0
        
        for _, row in df.iterrows():
            dc = str(row.get(dc_col, "")).strip()
            if dc != dealer_center:
                continue
            
            try:
                plan_val = row.get(plan_col, 0)
                fact_val = row.get(fact_col, 0)
                
                if pd.notna(plan_val):
                    plan_total += float(plan_val)
                if pd.notna(fact_val):
                    fact_total += float(fact_val)
            except (ValueError, TypeError):
                continue
        
        return plan_total, fact_total
    
    def load_marginality_facts(self, file_name: str = "Факт маржанальность.xlsx") -> Dict[str, float]:
        """Загружает фактические данные по маржинальности.
        Маржинальность = sum(столбец I, маржа) / sum(столбец G, сумма реализации) по каждому ДЦ.
        """
        df = self._read_excel(file_name, sheet_name="Sheet1", header=None)

        # Находим строку с заголовками (где в столбце A есть "Название ДЦ")
        header_row = self._find_header_row(df, "Название ДЦ")
        if header_row is None:
            raise MissingColumnError(file_name, ["Название ДЦ"], [str(v) for v in df.iloc[0].tolist()])

        # Переделываем датафрейм
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)

        # Используем позиционные индексы: G = 6, I = 8 (0-based)
        # Получаем имена столбцов по позиции
        col_names = list(df.columns)
        if len(col_names) < 9:
            return {}

        dc_col = col_names[0]    # столбец A — Название ДЦ
        revenue_col = col_names[6]  # столбец G — Сумма реализации
        margin_col = col_names[8]   # столбец I — Маржа

        # Оставляем только нужные столбцы, приводим к числам
        df[revenue_col] = pd.to_numeric(df[revenue_col], errors='coerce').fillna(0)
        df[margin_col] = pd.to_numeric(df[margin_col], errors='coerce').fillna(0)
        df[dc_col] = df[dc_col].astype(str).str.strip()

        # Убираем пустые строки
        df = df[df[dc_col].notna() & (df[dc_col] != '') & (df[dc_col] != 'nan')]

        # Группируем по ДЦ, суммируем G и I
        grouped = df.groupby(dc_col).agg(
            total_revenue=(revenue_col, 'sum'),
            total_margin=(margin_col, 'sum')
        )

        # Маржинальность = sum(I) / sum(G)
        marginality = {}
        for dc, row in grouped.iterrows():
            if row['total_revenue'] != 0:
                marginality[dc] = row['total_margin'] / row['total_revenue']
            else:
                marginality[dc] = 0.0

        return marginality
    
    def load_other_facts(self, dealer_center: str, file_name: str = "Факт по остальным показателям.xlsx") -> Dict[str, str]:
        """Загружает фактические данные для остальных показателей.
        Если в файле есть колонки по ДЦ, выбирает колонку конкретного ДЦ.
        """
        df = self._read_excel(file_name, sheet_name="Лист1")

        # Нормализуем названия колонок: убираем пробелы, сохраняем соответствие.
        normalized_columns = {str(col).strip(): col for col in df.columns}

        # Первый столбец всегда содержит название показателя (даже если заголовок с пробелами/кодировкой).
        metric_col = df.columns[0]

        # Приоритет: колонка конкретного ДЦ -> колонка "Факт" -> вторая колонка в файле.
        dealer_center_key = str(dealer_center).strip()
        if dealer_center_key in normalized_columns:
            fact_col = normalized_columns[dealer_center_key]
        elif "Факт" in normalized_columns:
            fact_col = normalized_columns["Факт"]
        elif len(df.columns) > 1:
            fact_col = df.columns[1]
        else:
            return {}
        
        facts = {}
        for _, row in df.iterrows():
            name = str(row.get(metric_col, "")).strip()
            if not name:
                continue
            value = str(row.get(fact_col, "")).strip()
            facts[name] = value
        
        return facts

    def load_service_consultant_fio_facts(self, dealer_center: str,
                                          file_name: str = "Сервис конультант факт.xlsx") -> Dict[str, Dict[str, float]]:
        """Загружает факт сервис-консультантов по ДЦ и возвращает разрез по ФИО.
        Результат: {ФИО: {"services_fact": float, "spare_parts_fact": float}}.
        """
        file_path = self.data_folder / file_name
        if not file_path.exists():
            return {}

        raw = _read_excel_cached(str(file_path), "Sheet1", None).copy()

        header_row = None
        required_headers = {"Название ДЦ", "Ответственный", "Сумма", "Тип"}
        for i, row in raw.iterrows():
            values = {str(v).strip() for v in row.values}
            if required_headers.issubset(values):
                header_row = i
                break

        if header_row is None:
            return {}

        raw.columns = raw.iloc[header_row].values
        df = raw.iloc[header_row + 1:].reset_index(drop=True)

        dc_col = self._pick_column_or_raise(df, file_name, ["Название ДЦ", "ДЦ"])
        fio_col = self._pick_column_or_raise(df, file_name, ["Ответственный", "ФИО", "Консультант"])
        sum_col = self._pick_column_or_raise(df, file_name, ["Сумма"])
        type_col = self._pick_column_or_raise(df, file_name, ["Тип"])

        dc_norm = self._norm_text(dealer_center).lower()
        if not dc_norm:
            return {}

        filtered = df[df[dc_col].astype(str).str.strip().str.lower() == dc_norm].copy()
        if filtered.empty:
            return {}

        filtered[fio_col] = filtered[fio_col].astype(str).str.strip()
        filtered = filtered[(filtered[fio_col] != "") & (filtered[fio_col].str.lower() != "nan")]
        if filtered.empty:
            return {}

        filtered[sum_col] = pd.to_numeric(filtered[sum_col], errors="coerce").fillna(0.0)
        filtered[type_col] = filtered[type_col].astype(str).str.strip().str.lower()

        def classify(metric_type: str) -> str:
            text = " ".join(str(metric_type).split())
            if "товар" in text or "запчаст" in text:
                return "spare_parts"
            if "услуг" in text or "услуга" in text or "работ" in text:
                return "services"
            return "other"

        filtered["_kind"] = filtered[type_col].apply(classify)
        grouped = (
            filtered[filtered["_kind"].isin(["services", "spare_parts"])]
            .groupby([fio_col, "_kind"], as_index=False)[sum_col]
            .sum()
        )
        if grouped.empty:
            return {}

        pivot = grouped.pivot(index=fio_col, columns="_kind", values=sum_col).fillna(0.0)

        result = {}
        for fio, row in pivot.iterrows():
            result[str(fio)] = {
                "services_fact": float(row.get("services", 0.0)),
                "spare_parts_fact": float(row.get("spare_parts", 0.0)),
            }
        return result

    def load_service_consultant_facts(self, dealer_center: str,
                                      file_name: str = "Сервис конультант факт.xlsx") -> Tuple[float, float, int]:
        """Агрегированный факт по сервис-консультантам для ДЦ (среднее на одного консультанта)."""
        fio_facts = self.load_service_consultant_fio_facts(dealer_center=dealer_center, file_name=file_name)
        fio_count = len(fio_facts)
        if fio_count == 0:
            return 0.0, 0.0, 0

        services_total = sum(v.get("services_fact", 0.0) for v in fio_facts.values())
        spare_parts_total = sum(v.get("spare_parts_fact", 0.0) for v in fio_facts.values())
        return services_total / fio_count, spare_parts_total / fio_count, fio_count


class BonusCalculator:
    """Рассчитывает бонус на основе загруженных данных"""
    
    def __init__(self, loader: BonusDataLoader, dealer_center: str = None, position: str = None,
                 consultant_fio: str = "", staff_count_override: int = 0):
        self.loader = loader
        self.consultant_fio = str(consultant_fio or "").strip()
        self.staff_count_override = int(staff_count_override) if staff_count_override and int(staff_count_override) > 0 else 0
        self.settings = loader.load_bonus_settings(dealer_center=dealer_center, position=position)
        self.has_explicit_weights = loader.has_explicit_weights_for_position(
            position=self.settings.position,
            weights_profile=self.settings.weights_profile,
        ) if hasattr(loader, "has_explicit_weights_for_position") else True
        self.has_explicit_logic = loader.has_explicit_logic_for_position(
            position=self.settings.position,
            logic_profile=self.settings.logic_profile,
        ) if hasattr(loader, "has_explicit_logic_for_position") else True
        self.weights_file_path = loader.resolve_weights_file(
            position=self.settings.position,
            weights_profile=self.settings.weights_profile
        )
        self.weights_sheet_name = loader.resolve_weights_sheet(
            self.weights_file_path,
            position=self.settings.position,
            weights_profile=self.settings.weights_profile,
        )
        self.logic_file_path = loader.resolve_logic_file(logic_profile=self.settings.logic_profile)
        try:
            self.weights = loader.load_metric_weights(
                position=self.settings.position,
                weights_profile=self.settings.weights_profile
            )
        except Exception:
            self.weights = {}

        # Если веса удалось загрузить (включая fallback-файл), считаем, что KPI-правила доступны.
        self.has_position_rules = bool(self.weights)

        try:
            self.logic = loader.load_calculation_logic(
                logic_profile=self.settings.logic_profile,
                position=self.settings.position,
            ) if self.has_position_rules else {}
        except TypeError:
            self.logic = loader.load_calculation_logic(logic_profile=self.settings.logic_profile) if self.has_position_rules else {}
        except Exception:
            self.logic = {}
        self.other_plans = loader.load_other_plans()
        
        # Загружаем план/факт с кэшем: это ускоряет повторные расчёты в рамках сессии.
        dealer_center = self.settings.dealer_center
        if hasattr(loader, "data_folder"):
            services = load_cached_facts(str(loader.data_folder), dealer_center, "services")
            spare_parts = load_cached_facts(str(loader.data_folder), dealer_center, "spare_parts")
            self.services_plan = services.get("plan", 0.0)
            self.services_fact = services.get("fact", 0.0)
            self.spare_parts_plan = spare_parts.get("plan", 0.0)
            self.spare_parts_fact = spare_parts.get("fact", 0.0)
            self.marginality_facts = load_cached_marginality(str(loader.data_folder))
            self.other_facts = load_cached_facts(str(loader.data_folder), dealer_center, "other")
        else:
            # Fallback для моков в unit-тестах без data_folder.
            self.services_plan, self.services_fact = loader.load_services_facts(dealer_center)
            self.spare_parts_plan, self.spare_parts_fact = loader.load_spare_parts_facts(dealer_center)
            self.marginality_facts = loader.load_marginality_facts()
            self.other_facts = loader.load_other_facts(dealer_center)

        self._apply_position_plan_rules()
        self._apply_service_consultant_facts()

    def _apply_position_plan_rules(self) -> None:
        """Применяет позиционные правила к планам. Для сервис-консультанта план персонализируется."""
        pos_norm = self._normalize_text(self.settings.position)
        if not ("сервис" in pos_norm and "консульт" in pos_norm):
            return

        if not hasattr(self.loader, "count_position_staff"):
            return

        staff_count = self.loader.count_position_staff(
            dealer_center=self.settings.dealer_center,
            position=self.settings.position,
        )
        # Переопределение из UI (ручная правка пользователем).
        if self.staff_count_override > 0:
            staff_count = self.staff_count_override
        if staff_count <= 1:
            return

        # План по сервис-консультанту: общий план ДЦ / количество консультантов.
        self.services_plan = self.services_plan / staff_count
        self.spare_parts_plan = self.spare_parts_plan / staff_count

    def _apply_service_consultant_facts(self) -> None:
        """Подменяет факт услуг/запчастей для сервис-консультанта данными из отдельного файла факта."""
        pos_norm = self._normalize_text(self.settings.position)
        if not ("сервис" in pos_norm and "консульт" in pos_norm):
            return

        facts_payload = {}
        if hasattr(self.loader, "data_folder"):
            facts_payload = load_cached_facts(
                str(self.loader.data_folder),
                self.settings.dealer_center,
                "service_consultant",
            )
        elif hasattr(self.loader, "load_service_consultant_fio_facts"):
            fio_facts = self.loader.load_service_consultant_fio_facts(self.settings.dealer_center)
            fio_count = len(fio_facts)
            services_total = sum(v.get("services_fact", 0.0) for v in fio_facts.values())
            spare_total = sum(v.get("spare_parts_fact", 0.0) for v in fio_facts.values())
            facts_payload = {
                "services_fact": (services_total / fio_count if fio_count > 0 else 0.0),
                "spare_parts_fact": (spare_total / fio_count if fio_count > 0 else 0.0),
                "fio_count": fio_count,
                "fio_facts": fio_facts,
            }
        elif hasattr(self.loader, "load_service_consultant_facts"):
            services_fact, spare_parts_fact, fio_count = self.loader.load_service_consultant_facts(
                self.settings.dealer_center
            )
            facts_payload = {
                "services_fact": services_fact,
                "spare_parts_fact": spare_parts_fact,
                "fio_count": fio_count,
            }

        if not facts_payload:
            return

        fio_facts = facts_payload.get("fio_facts", {}) or {}
        if self.consultant_fio and self.consultant_fio in fio_facts:
            services_fact = float(fio_facts[self.consultant_fio].get("services_fact", 0.0) or 0.0)
            spare_parts_fact = float(fio_facts[self.consultant_fio].get("spare_parts_fact", 0.0) or 0.0)
        else:
            services_fact = float(facts_payload.get("services_fact", 0.0) or 0.0)
            spare_parts_fact = float(facts_payload.get("spare_parts_fact", 0.0) or 0.0)

        # Подменяем факт только если нашли хотя бы одно ненулевое значение.
        if services_fact > 0 or spare_parts_fact > 0:
            self.services_fact = services_fact
            self.spare_parts_fact = spare_parts_fact

    @staticmethod
    def _parse_number(value) -> Optional[float]:
        if value is None:
            return None
        text = str(value).strip().replace("\u00a0", "").replace(" ", "").replace("%", "").replace(",", ".")
        if text in ("", "-", "—"):
            return None
        try:
            return float(text)
        except ValueError:
            return None

    @staticmethod
    def _extract_percent(text: str) -> Optional[float]:
        m = re.search(r"(-?\d+(?:[\.,]\d+)?)\s*%", text or "")
        if not m:
            return None
        return float(m.group(1).replace(",", ".")) / 100.0

    @staticmethod
    def _extract_money(text: str) -> Optional[float]:
        m = re.search(r"(-?\d[\d\s]*)\s*(?:₸|тг)", text or "", re.IGNORECASE)
        if not m:
            return None
        return float(m.group(1).replace(" ", ""))

    @staticmethod
    def _normalize_text(value: str) -> str:
        return " ".join(str(value).strip().lower().split())

    def _find_fact_value_by_metric(self, metric_name: str) -> str:
        metric_key = self._normalize_text(metric_name)
        # 1) Точное совпадение
        for k, v in self.other_facts.items():
            if self._normalize_text(k) == metric_key:
                return str(v)
        # 2) Частичное совпадение
        for k, v in self.other_facts.items():
            k_norm = self._normalize_text(k)
            if metric_key in k_norm or k_norm in metric_key:
                return str(v)
        return ""

    def _detect_formula_mode(self) -> bool:
        standard_keys = {
            "Оклад (База)",
            "Выполнение плана продаж по продажи услуг",
            "Выполнение плана продаж по Запасным Частям",
            "Выполнение  маржанальности",
            "Соблюдение уровня неликвида на складе запасных частей",
        }
        if all(k in self.weights for k in standard_keys):
            return False
        # Если хотя бы в одном показателе есть формула в сыром виде — это формульный лист.
        return any((w.raw_weight or w.description or w.penalty_flag) for w in self.weights.values())

    def _calculate_for_formula_sheet(self, month: str) -> BonusCalculationResult:
        reper = self.settings.reper

        base_salary_bonus = 0.0
        services_plan_bonus = 0.0
        spare_parts_bonus = 0.0
        marginality_bonus = 0.0
        negliquidity_deduction = 0.0
        other_metric_bonuses: Dict[str, float] = {}

        services_execution = (self.services_fact / self.services_plan) if self.services_plan > 0 else 0.0
        spare_parts_execution = (self.spare_parts_fact / self.spare_parts_plan) if self.spare_parts_plan > 0 else 0.0
        marginality_fact = self.marginality_facts.get(self.settings.dealer_center, 0.0)

        negliquidity_metric = "Соблюдение уровня неликвида на складе запасных частей"
        negliquidity_plan_str = self.other_plans.get(
            negliquidity_metric,
            str(CONFIG["thresholds"]["negliquidity_max"])
        )
        try:
            negliquidity_plan = float(negliquidity_plan_str) if negliquidity_plan_str != "нет" else 0
        except ValueError:
            negliquidity_plan = CONFIG["thresholds"]["negliquidity_max"]
        try:
            negliquidity_fact = float(self.other_facts.get(negliquidity_metric, CONFIG["thresholds"]["negliquidity_max"]))
        except (TypeError, ValueError):
            negliquidity_fact = 0.0

        for metric_name, metric in self.weights.items():
            metric_low = self._normalize_text(metric_name)
            raw_weight = metric.raw_weight or ""
            description = metric.description or ""
            penalty_flag = metric.penalty_flag or ""

            # До утверждения отдельной логики/источника факта этот KPI не начисляем.
            if "доп" in metric_low and ("оборудован" in metric_low or "допоборуд" in metric_low):
                other_metric_bonuses[metric_name] = 0.0
                continue

            value = 0.0
            # 1) Числовой вес как в старой модели (доля от репера)
            if metric.weight != 0:
                if "оклад" in metric_low and "база" in metric_low:
                    value = reper * metric.weight
                    base_salary_bonus += value
                    continue
                if "услуг" in metric_low:
                    value = reper * metric.weight * services_execution
                    services_plan_bonus += value
                    continue
                if "запчаст" in metric_low:
                    value = reper * metric.weight * spare_parts_execution
                    spare_parts_bonus += value
                    continue
                if "маржаналь" in metric_low or "маржиналь" in metric_low:
                    marginality_plan_str = self.other_plans.get("Выполнение  маржанальности", str(CONFIG["thresholds"]["marginality_min"]))
                    try:
                        marginality_plan = float(marginality_plan_str) if marginality_plan_str != "нет" else 0
                    except ValueError:
                        marginality_plan = CONFIG["thresholds"]["marginality_min"]
                    value = reper * metric.weight if marginality_fact >= marginality_plan else 0.0
                    marginality_bonus += value
                    continue
                if "неликвид" in metric_low:
                    value = reper * metric.weight if negliquidity_fact > negliquidity_plan else 0.0
                    negliquidity_deduction += value
                    continue

                other_metric_bonuses[metric_name] = reper * metric.weight
                continue

            # 2) Формулы вида "(Репер-оклад) *50%" или "...Бонус*50%" или просто "50%"
            percent = self._extract_percent(raw_weight)
            if percent is not None:
                raw_norm = self._normalize_text(raw_weight)
                if "репер" in raw_norm and "оклад" in raw_norm:
                    # Явная формула (Репер - Оклад) * X%
                    base_amount = max(reper - self.settings.fixed_salary, 0.0)
                elif self.settings.bonus_pool > 0:
                    # Формула только с процентом, база = переменная часть (Бонус)
                    base_amount = self.settings.bonus_pool
                else:
                    base_amount = reper
                execution = 1.0
                if "услуг" in metric_low:
                    execution = services_execution
                elif "запчаст" in metric_low:
                    execution = spare_parts_execution
                else:
                    plan_raw = self.other_plans.get(metric_name, "")
                    fact_raw = self._find_fact_value_by_metric(metric_name)
                    plan_num = self._parse_number(plan_raw)
                    fact_num = self._parse_number(fact_raw)
                    if plan_num and plan_num > 0 and fact_num is not None:
                        execution = fact_num / plan_num

                value = max(base_amount, 0) * percent * execution
                if "услуг" in metric_low:
                    services_plan_bonus += value
                elif "запчаст" in metric_low:
                    spare_parts_bonus += value
                elif "оклад" in metric_low:
                    base_salary_bonus += value
                else:
                    other_metric_bonuses[metric_name] = value
                continue

            # 3) Фикс за штуку: "5000 ₸/шт"
            fixed_amount = self._extract_money(raw_weight)
            if fixed_amount is not None and ("/шт" in self._normalize_text(raw_weight) or "за шт" in self._normalize_text(raw_weight)):
                fact_raw = self._find_fact_value_by_metric(metric_name)
                count = self._parse_number(fact_raw) or 0.0
                value = fixed_amount * count
                other_metric_bonuses[metric_name] = value
                continue

            # 4) Штрафы Да/Нет с суммой в описании
            penalty_amount = self._extract_money(description) or self._extract_money(raw_weight)
            if penalty_amount is not None and (
                "да/нет" in self._normalize_text(penalty_flag)
                or "при невыполнении" in self._normalize_text(description)
                or "штраф" in self._normalize_text(description)
            ):
                fact_raw = self._find_fact_value_by_metric(metric_name)
                fact_norm = self._normalize_text(fact_raw)
                failed = fact_norm in {"нет", "no", "false", "0", "не выполнено", "не соблюдено", "не соблюден"}
                if failed:
                    value = -abs(penalty_amount)
                else:
                    value = 0.0
                other_metric_bonuses[metric_name] = value
                continue

            # 5) Неподдержанная формула пока даёт 0, но остаётся в детализации.
            other_metric_bonuses[metric_name] = 0.0

        # Если в формульном листе нет явной строки оклада, берём фикс из настроек ДЦ/должности.
        if base_salary_bonus <= 0 and self.settings.fixed_salary > 0:
            base_salary_bonus = float(self.settings.fixed_salary)

        other_bonus_total = sum(other_metric_bonuses.values())
        total_bonus = services_plan_bonus + spare_parts_bonus + marginality_bonus + negliquidity_deduction + other_bonus_total

        return BonusCalculationResult(
            dealer_center=self.settings.dealer_center,
            position=self.settings.position,
            reper=reper,
            month=month,
            base_salary_bonus=base_salary_bonus,
            services_plan_bonus=services_plan_bonus,
            spare_parts_bonus=spare_parts_bonus,
            marginality_bonus=marginality_bonus,
            negliquidity_deduction=negliquidity_deduction,
            total_bonus=total_bonus,
            services_plan_value=self.services_plan,
            services_fact_value=self.services_fact,
            services_execution=services_execution,
            spare_parts_plan_value=self.spare_parts_plan,
            spare_parts_fact_value=self.spare_parts_fact,
            spare_parts_execution=spare_parts_execution,
            marginality_fact=marginality_fact,
            negliquidity_fact=negliquidity_fact,
            other_metric_bonuses=other_metric_bonuses,
        )

    def _calculate_settings_only(self, month: str) -> BonusCalculationResult:
        return BonusCalculationResult(
            dealer_center=self.settings.dealer_center,
            position=self.settings.position,
            reper=self.settings.reper,
            month=month,
            base_salary_bonus=float(self.settings.fixed_salary or 0.0),
            services_plan_bonus=0.0,
            spare_parts_bonus=0.0,
            marginality_bonus=0.0,
            negliquidity_deduction=0.0,
            total_bonus=float(self.settings.bonus_pool or 0.0),
            services_plan_value=self.services_plan,
            services_fact_value=self.services_fact,
            services_execution=0.0,
            spare_parts_plan_value=self.spare_parts_plan,
            spare_parts_fact_value=self.spare_parts_fact,
            spare_parts_execution=0.0,
            marginality_fact=self.marginality_facts.get(self.settings.dealer_center, 0.0),
            negliquidity_fact=0.0,
            other_metric_bonuses={},
        )
    
    def calculate_for_month(self, month: str) -> BonusCalculationResult:
        """Рассчитывает бонус для конкретного месяца"""
        start_time = time.time()
        audit_event(
            LOGGER,
            "calc_start",
            dealer_center=self.settings.dealer_center,
            position=self.settings.position,
            month=month,
            weights_file=self.weights_file_path.name,
            logic_file=self.logic_file_path.name,
        )

        try:
            if not self.has_position_rules:
                result = self._calculate_settings_only(month)
                audit_event(
                    LOGGER,
                    "calc_end",
                    dealer_center=self.settings.dealer_center,
                    position=self.settings.position,
                    month=month,
                    total_bonus=round(result.total_bonus, 2),
                    duration_seconds=round(time.time() - start_time, 3),
                    mode="settings_only",
                )
                return result

            if self._detect_formula_mode():
                result = self._calculate_for_formula_sheet(month)
                audit_event(
                    LOGGER,
                    "calc_end",
                    dealer_center=self.settings.dealer_center,
                    position=self.settings.position,
                    month=month,
                    total_bonus=round(result.total_bonus, 2),
                    duration_seconds=round(time.time() - start_time, 3),
                    mode="formula_sheet",
                )
                return result

        # Получаем репер
            reper = self.settings.reper
        
        # 1. Оклад (база) = Репер × доля
            base_weight = self.weights.get("Оклад (База)", MetricWeight("Оклад (База)", 0.0, "")).weight
            base_salary_bonus = reper * base_weight
        
        # 2. План услуг = Репер × доля × (факт / план)
            services_metric = "Выполнение плана продаж по продажи услуг"
            services_weight = self.weights.get(services_metric, MetricWeight(services_metric, 0, "")).weight
        
            services_plan_total = self.services_plan
            services_fact_total = self.services_fact
        
            services_execution = (services_fact_total / services_plan_total) if services_plan_total > 0 else 0
            services_plan_bonus = reper * services_weight * services_execution
        
        # 3. План запчастей = Репер × доля × (факт / план)
            spare_parts_metric = "Выполнение плана продаж по Запасным Частям"
            spare_parts_weight = self.weights.get(spare_parts_metric, MetricWeight(spare_parts_metric, 0, "")).weight
        
            spare_parts_plan_total = self.spare_parts_plan
            spare_parts_fact_total = self.spare_parts_fact
        
            spare_parts_execution = (spare_parts_fact_total / spare_parts_plan_total) if spare_parts_plan_total > 0 else 0
            spare_parts_bonus = reper * spare_parts_weight * spare_parts_execution
        
        # 4. Маржинальность = Репер × доля, но если < плана то 0
            marginality_metric = "Выполнение  маржанальности"
            marginality_weight = self.weights.get(marginality_metric, MetricWeight(marginality_metric, 0, "")).weight
        
            marginality_plan_str = self.other_plans.get(
                "Выполнение  маржанальности",
                str(CONFIG["thresholds"]["marginality_min"])
            )
            marginality_plan = float(marginality_plan_str) if marginality_plan_str != "нет" else 0
        
        # Берём маржинальность только выбранного ДЦ:
        # sum(I маржа) / sum(G сумма реализации) по конкретному дилерскому центру.
            marginality_fact = self.marginality_facts.get(self.settings.dealer_center, 0.0)

            BonusCalcRequest(
                dealer_center=self.settings.dealer_center,
                position=self.settings.position,
                month=month,
                metrics={
                    services_metric: KPIMetricInput(metric_name=services_metric, plan=services_plan_total, fact=services_fact_total),
                    spare_parts_metric: KPIMetricInput(metric_name=spare_parts_metric, plan=spare_parts_plan_total, fact=spare_parts_fact_total),
                    marginality_metric: KPIMetricInput(metric_name=marginality_metric, plan=marginality_plan, fact=marginality_fact),
                }
            )
        
            if marginality_fact >= marginality_plan:
                marginality_bonus = reper * marginality_weight
            else:
                marginality_bonus = 0
        
        # 5. Неликвид (демотиватор) = -Репер × доля если факт > плана
            negliquidity_metric = "Соблюдение уровня неликвида на складе запасных частей"
            negliquidity_weight = self.weights.get(negliquidity_metric, MetricWeight(negliquidity_metric, 0, "")).weight
        
            negliquidity_plan_str = self.other_plans.get(
                negliquidity_metric,
                str(CONFIG["thresholds"]["negliquidity_max"])
            )
            negliquidity_plan = float(negliquidity_plan_str) if negliquidity_plan_str != "нет" else 0
        
            negliquidity_fact_str = self.other_facts.get(negliquidity_metric, str(CONFIG["thresholds"]["negliquidity_max"]))
            try:
                negliquidity_fact = float(negliquidity_fact_str)
            except (TypeError, ValueError):
                negliquidity_fact = 0

            KPIMetricInput(metric_name=negliquidity_metric, plan=negliquidity_plan, fact=negliquidity_fact)
        
            if negliquidity_fact > negliquidity_plan:
                negliquidity_deduction = reper * negliquidity_weight  # Это отрицательное значение
            else:
                negliquidity_deduction = 0
        
        # 6. Остальные KPI по плану/факту
            other_metric_bonuses = {}
            for metric_name, weight_obj in self.weights.items():
                if metric_name in {
                    "Оклад (База)",
                    services_metric,
                    spare_parts_metric,
                    marginality_metric,
                    negliquidity_metric
                }:
                    continue

                plan_value = self.other_plans.get(metric_name, "-")
                fact_value = self.other_facts.get(metric_name, "-")
                other_metric_bonuses[metric_name] = self._calculate_other_metric_bonus(
                    metric_name,
                    reper,
                    weight_obj.weight,
                    plan_value,
                    fact_value
                )

            other_bonus_total = sum(other_metric_bonuses.values())

        # Итого бонус (без оклада — оклад фиксированный и не входит в бонус)
            total_bonus = (services_plan_bonus + spare_parts_bonus +
                          marginality_bonus + negliquidity_deduction + other_bonus_total)

            result = BonusCalculationResult(
                dealer_center=self.settings.dealer_center,
                position=self.settings.position,
                reper=reper,
                month=month,
                base_salary_bonus=base_salary_bonus,
                services_plan_bonus=services_plan_bonus,
                spare_parts_bonus=spare_parts_bonus,
                marginality_bonus=marginality_bonus,
                negliquidity_deduction=negliquidity_deduction,
                total_bonus=total_bonus,
                services_plan_value=services_plan_total,
                services_fact_value=services_fact_total,
                services_execution=services_execution,
                spare_parts_plan_value=spare_parts_plan_total,
                spare_parts_fact_value=spare_parts_fact_total,
                spare_parts_execution=spare_parts_execution,
                marginality_fact=marginality_fact,
                negliquidity_fact=negliquidity_fact,
                other_metric_bonuses=other_metric_bonuses
            )

            audit_event(
                LOGGER,
                "calc_end",
                dealer_center=self.settings.dealer_center,
                position=self.settings.position,
                month=month,
                total_bonus=round(result.total_bonus, 2),
                duration_seconds=round(time.time() - start_time, 3),
            )
            return result
        except ValidationError as exc:
            audit_event(
                LOGGER,
                "calc_validation_error",
                dealer_center=self.settings.dealer_center,
                position=self.settings.position,
                month=month,
                errors=exc.errors(),
            )
            raise DataLoadingError(f"Ошибка валидации входных данных: {exc.errors()}") from exc
        except Exception as exc:
            audit_event(
                LOGGER,
                "calc_error",
                dealer_center=self.settings.dealer_center,
                position=self.settings.position,
                month=month,
                error=str(exc),
            )
            raise

    def _calculate_other_metric_bonus(
        self,
        metric_name: str,
        reper: float,
        weight: float,
        plan_value: str,
        fact_value: str
    ) -> float:
        """Расчитывает бонус для дополнительных KPI по правилу показателя."""
        def normalize_text(value: str) -> str:
            return str(value).strip().lower()

        def parse_number(value: str) -> Optional[float]:
            try:
                normalized = str(value).strip().replace(',', '.')
                return float(normalized)
            except (ValueError, TypeError):
                return None

        plan_str = normalize_text(plan_value)
        fact_str = normalize_text(fact_value)

        # Неликвид: штраф только при превышении плана
        if "неликвид" in metric_name.lower():
            plan_num = parse_number(plan_value)
            fact_num = parse_number(fact_value)
            if plan_num is not None and fact_num is not None and fact_num > plan_num:
                return reper * weight
            return 0.0

        # Для булевых показателей да/нет: если факт отличается от плана, применяем вес
        if plan_str in {"да", "нет"} or fact_str in {"да", "нет"}:
            if plan_str and fact_str and plan_str != fact_str:
                return reper * weight
            return 0.0

        # Для остальных показателей: если факт не равен плану, применяем вес
        plan_num = parse_number(plan_value)
        fact_num = parse_number(fact_value)
        if plan_num is not None and fact_num is not None:
            if plan_num != fact_num:
                return reper * weight
        elif plan_str and fact_str and plan_str != fact_str:
            return reper * weight

        return 0.0

    def get_indicator_df(self, month: str) -> pd.DataFrame:
        """Собирает таблицу всех показателей по шаблону"""
        result = self.calculate_for_month(month)

        services_metric = "Выполнение плана продаж по продажи услуг"
        spare_metric = "Выполнение плана продаж по Запасным Частям"
        marginality_metric = "Выполнение  маржанальности"
        negliquidity_metric = "Соблюдение уровня неликвида на складе запасных частей"

        marginality_plan = self.other_plans.get(marginality_metric, "0.3")
        try:
            marginality_plan_val = float(marginality_plan) if marginality_plan != "нет" else 0.0
        except ValueError:
            marginality_plan_val = 0.0
        negliquidity_plan = self.other_plans.get(negliquidity_metric, "0.07")
        try:
            negliquidity_plan_val = float(negliquidity_plan) if negliquidity_plan != "нет" else 0.0
        except ValueError:
            negliquidity_plan_val = 0.0

        rows = []

        rows.append({
            "Показатель": "Оклад (База)",
            "Вес": self.weights.get("Оклад (База)", MetricWeight("Оклад (База)", 0.0, "")).weight,
            "План": "-",
            "Факт": "-",
            "Выполнение": "-",
            "Бонус": result.base_salary_bonus,
            "Логика": self.logic.get("Оклад (База)", "Репер × доля")
        })

        rows.append({
            "Показатель": services_metric,
            "Вес": self.weights.get(services_metric, MetricWeight(services_metric, 0.0, "")).weight,
            "План": self.services_plan,
            "Факт": self.services_fact,
            "Выполнение": f"{result.services_execution * 100:.1f}%" if self.services_plan > 0 else "-",
            "Бонус": result.services_plan_bonus,
            "Логика": self.logic.get(services_metric, "Репер × доля × факт/план")
        })

        rows.append({
            "Показатель": spare_metric,
            "Вес": self.weights.get(spare_metric, MetricWeight(spare_metric, 0.0, "")).weight,
            "План": self.spare_parts_plan,
            "Факт": self.spare_parts_fact,
            "Выполнение": f"{result.spare_parts_execution * 100:.1f}%" if self.spare_parts_plan > 0 else "-",
            "Бонус": result.spare_parts_bonus,
            "Логика": self.logic.get(spare_metric, "Репер × доля × факт/план")
        })

        rows.append({
            "Показатель": marginality_metric,
            "Вес": self.weights.get(marginality_metric, MetricWeight(marginality_metric, 0.0, "")).weight,
            "План": marginality_plan,
            "Факт": result.marginality_fact,
            "Выполнение": f"{result.marginality_fact / marginality_plan_val * 100:.1f}%" if marginality_plan_val > 0 else "-",
            "Бонус": result.marginality_bonus,
            "Логика": self.logic.get(marginality_metric, "Репер × доля при достижении плана")
        })

        rows.append({
            "Показатель": negliquidity_metric,
            "Вес": self.weights.get(negliquidity_metric, MetricWeight(negliquidity_metric, 0.0, "")).weight,
            "План": negliquidity_plan,
            "Факт": result.negliquidity_fact,
            "Выполнение": f"{result.negliquidity_fact / negliquidity_plan_val * 100:.1f}%" if negliquidity_plan_val > 0 else "-",
            "Бонус": result.negliquidity_deduction,
            "Логика": self.logic.get(negliquidity_metric, "-Репер × доля при превышении плана")
        })

        for metric_name, weight in self.weights.items():
            if metric_name in {"Оклад (База)", services_metric, spare_metric, marginality_metric, negliquidity_metric}:
                continue
            plan_value = self.other_plans.get(metric_name, "-")
            fact_value = self.other_facts.get(metric_name, "-")
            bonus_value = result.other_metric_bonuses.get(metric_name, 0.0)
            rows.append({
                "Показатель": metric_name,
                "Вес": weight.weight,
                "План": plan_value,
                "Факт": fact_value,
                "Выполнение": "-",
                "Бонус": bonus_value,
                "Логика": self.logic.get(metric_name, "-")
            })

        return pd.DataFrame(rows)
