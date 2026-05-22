from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


APRIL_START = pd.Timestamp("2026-04-01")
APRIL_END = pd.Timestamp("2026-05-01")

WORKBOOK_PATH = Path(r"c:\Users\D.Muldabaev\Desktop\Приложения для сервиса\Бюллитень апрель\Маржа апрель.xlsx")
OUTPUT_PATH = Path(r"c:\Users\D.Muldabaev\Desktop\Приложения для сервиса\Бюллитень апрель\bulletin_margin_check_april_2026.xlsx")


@dataclass(frozen=True)
class BulletinRule:
    brand: str
    model: str
    trim: str
    reward: int
    source: str
    year: int | None = None
    drive: str | None = None
    engine: int | None = None


def rule(
    brand: str,
    model: str,
    trim: str,
    reward: int,
    source: str,
    *,
    year: int | None = None,
    drive: str | None = None,
    engine: int | None = None,
) -> BulletinRule:
    return BulletinRule(
        brand=brand,
        model=model,
        trim=trim,
        reward=int(reward),
        source=source,
        year=year,
        drive=drive,
        engine=engine,
    )


RULES: list[BulletinRule] = [
    # KIA
    rule("Киа", "Carnival", "Luxe+", 1429350, "Kia SD DC 17-03-2026; Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2025),
    rule("Киа", "Carnival", "Limousine+", 1884350, "Kia SD DC 17-03-2026; Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2025),
    rule("Киа", "Ceed", "Comfort", 436050, "Kia bulletins", year=2025),
    rule("Киа", "Ceed", "Luxe", 481050, "Kia bulletins", year=2025),
    rule("Киа", "Ceed SW", "Comfort", 494550, "Kia bulletins", year=2025),
    rule("Киа", "Ceed SW", "Luxe", 517050, "Kia bulletins", year=2025),
    rule("Киа", "Ceed SW", "Prestige", 562050, "Kia bulletins", year=2025),
    rule("Киа", "Cerato", "Comfort", 399600, "Kia bulletins", year=2025, engine=1591),
    rule("Киа", "Cerato", "Comfort", 399600, "Kia bulletins", year=2026, engine=1591),
    rule("Киа", "Cerato", "Luxe", 439600, "Kia bulletins", year=2025, engine=1999),
    rule("Киа", "Cerato", "Luxe+", 419600, "Kia bulletins", year=2025, engine=1591),
    rule("Киа", "Cerato", "Luxe+", 419600, "Kia bulletins", year=2026, engine=1591),
    rule("Киа", "Cerato", "Luxe+", 439600, "Kia bulletins", year=2025, engine=2),
    rule("Киа", "Cerato", "Luxe+", 439600, "Kia bulletins", year=2026, engine=2),
    rule("Киа", "K5", "Comfort", 604450, "Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2025),
    rule("Киа", "K5", "Comfort", 637450, "Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2026),
    rule("Киа", "K5", "Comfort", 774950, "Kia SD DC 17-03-2026", year=2026),
    rule("Киа", "K5", "Luxe", 697950, "Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2025),
    rule("Киа", "K5", "Luxe", 752950, "Kia SD DC 17-03-2026", year=2025),
    rule("Киа", "K5", "Luxe", 741950, "Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2026),
    rule("Киа", "K5", "Luxe", 879450, "Kia SD DC 17-03-2026", year=2026),
    rule("Киа", "K5", "Prestige", 769450, "Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2025),
    rule("Киа", "K5", "Prestige", 813450, "Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2026),
    rule("Киа", "K5", "Premium", 868450, "Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2026),
    rule("Киа", "K5", "Premium", 1005950, "Kia SD DC 17-03-2026", year=2026),
    rule("Киа", "K5", "Tulpar", 736450, "Kia SD DC 03-04-2026; Kia SD DC 20-04-2026", year=2025),
    rule("Киа", "K5", "Tulpar", 791450, "Kia SD DC 17-03-2026", year=2025),
    rule("Киа", "K8", "Prestige", 1721850, "Kia bulletins", year=2025),
    rule("Киа", "Seltos", "Comfort", 494550, "Kia bulletins", year=2025, drive="2WD", engine=1999),
    rule("Киа", "Seltos", "Comfort", 521550, "Kia bulletins", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Seltos", "Comfort", 544050, "Kia bulletins", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Seltos", "Luxe", 539550, "Kia bulletins", year=2025, drive="2WD", engine=1999),
    rule("Киа", "Seltos", "Luxe", 634050, "Kia SD DC 17-03-2026", year=2025, drive="2WD", engine=1999),
    rule("Киа", "Seltos", "Luxe", 566550, "Kia bulletins", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Seltos", "Style", 607050, "Kia bulletins", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Seltos", "Tulpar", 584550, "Kia bulletins", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Seltos", "Tulpar", 607050, "Kia bulletins", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Soluto", "Comfort", 300000, "Kia bulletins", year=2026),
    rule("Киа", "Soluto", "Prestige", 300000, "Kia bulletins", year=2025),
    rule("Киа", "Soluto", "Prestige", 300000, "Kia bulletins", year=2026),
    rule("Киа", "Sorento", "Comfort", 809550, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Comfort", 934450, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Comfort", 854550, "Kia bulletins", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Luxe", 872550, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Luxe", 989450, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Luxe", 1011450, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Luxe", 899550, "Kia bulletins", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Luxe", 917550, "Kia bulletins", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Premium", 1034550, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Premium", 1209450, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Premium", 1079550, "Kia bulletins", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Premium", 1147050, "Kia bulletins", year=2026, drive="4WD", engine=3470),
    rule("Киа", "Sorento", "Style", 1082950, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Style", 1099450, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento", "Tulpar", 1044450, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Soul", "Luxe", 526050, "Kia bulletins", year=2025),
    rule("Киа", "Soul", "Style", 584550, "Kia bulletins", year=2025),
    rule("Киа", "Sportage", "Comfort", 383700, "Kia bulletins", year=2026, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Luxe", 796950, "Kia bulletins", year=2025, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Luxe", 818950, "Kia bulletins", year=2025, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Luxe", 824450, "Kia bulletins", year=2025, drive="4WD", engine=1975),
    rule("Киа", "Sportage", "Luxe", 857450, "Kia bulletins", year=2025, drive="4WD", engine=1975),
    rule("Киа", "Sportage", "Luxe", 741950, "Kia bulletins", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Sportage", "Luxe", 769450, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "Luxe", 824450, "Kia bulletins", year=2026, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Luxe", 873950, "Kia SD DC 17-03-2026", year=2026, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Luxe", 857450, "Kia bulletins", year=2026, drive="4WD", engine=1975),
    rule("Киа", "Sportage", "Luxe", 741950, "Kia bulletins", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Sportage", "Style", 912450, "Kia bulletins", year=2025, drive="4WD", engine=1975),
    rule("Киа", "Sportage", "Style", 829950, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "Style", 950950, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "Style", 1005950, "Kia SD DC 17-03-2026", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "X-Line+", 1022450, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "X-Line+", 1077450, "Kia bulletins", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "X-Line+", 1077450, "Kia bulletins", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "X-Line+", 906950, "Kia bulletins", year=2026, drive="4WD", engine=2497),

    # Chevrolet
    rule("Шеви", "Cobalt", "Optimum AT", 200000, "Шеви.jpeg", year=2025),
    rule("Шеви", "Cobalt", "Elegant AT", 200000, "Шеви.jpeg", year=2025),
    rule("Шеви", "Cobalt", "Optimum AT", 200000, "Шеви.jpeg", year=2026),
    rule("Шеви", "Cobalt", "Elegant AT", 200000, "Шеви.jpeg", year=2026),
    rule("Шеви", "Cobalt", "Optimum AT", 250000, "Шеви.jpeg", year=2026),
    rule("Шеви", "Cobalt", "Elegant AT", 250000, "Шеви.jpeg", year=2026),
    rule("Шеви", "Cobalt", "Optimum AT", 300000, "30.04 Шеви.jpeg", year=2025),
    rule("Шеви", "Cobalt", "Elegant AT", 300000, "30.04 Шеви.jpeg", year=2025),
    rule("Шеви", "Cobalt", "Optimum AT", 300000, "30.04 Шеви.jpeg", year=2026),
    rule("Шеви", "Cobalt", "Elegant AT", 300000, "30.04 Шеви.jpeg", year=2026),
    rule("Шеви", "Onix", "4LT AT (LTZ)", 200000, "Шеви.jpeg", year=2024),
    rule("Шеви", "Onix", "Premier 1", 200000, "Шеви.jpeg", year=2024),
    rule("Шеви", "Onix", "Premier 2", 200000, "Шеви.jpeg", year=2024),
    rule("Шеви", "Onix", "4LT AT (LTZ)", 300000, "Шеви.jpeg; 30.04 Шеви.jpeg", year=2025),
    rule("Шеви", "Onix", "Premier 1", 300000, "Шеви.jpeg; 30.04 Шеви.jpeg", year=2025),
    rule("Шеви", "Onix", "Premier 2", 300000, "Шеви.jpeg; 30.04 Шеви.jpeg", year=2025),
    rule("Шеви", "Onix", "4LT AT (LTZ)", 350000, "Шеви.jpeg", year=2026),
    rule("Шеви", "Onix", "Premier 1", 350000, "Шеви.jpeg", year=2026),
    rule("Шеви", "Onix", "Premier 2", 350000, "Шеви.jpeg", year=2026),
    rule("Шеви", "Tracker MY22", "1LT AT", 400000, "Шеви.jpeg", year=2025),
    rule("Шеви", "Tracker MY22", "Premier AT", 400000, "Шеви.jpeg", year=2025),
    rule("Шеви", "Tahoe MY25", "High Country", 1200000, "Шеви.jpeg", year=2026),
    rule("Шеви", "Labo", "LABO", 200000, "Шеви.jpeg", year=2025),

    # JAC
    rule("Джак", "S3 Pro", "Intelligent 1.6VVT CVT", 267600, "Джак 8-ДП 23.04.2026", year=2025),
    rule("Джак", "S3 Pro", "Intelligent 1.6VVT CVT", 279600, "Джак 8-ДП 23.04.2026", year=2026),
    rule("Джак", "J7", "Luxury 1.5T CVT NEW", 279600, "Джак 5-ДП 20.03.2026", year=2024),
    rule("Джак", "J7", "Comfort Plus 1.5T", 291600, "Джак 5-ДП 20.03.2026", year=2025),
    rule("Джак", "J7", "Luxury 1.5T CVT NEW JL", 315600, "Джак 5-ДП 20.03.2026", year=2025),
    rule("Джак", "J7 PLUS", "Flagship 1.5TGDI 6DCT", 407600, "Джак 5-ДП 20.03.2026", year=2025),
    rule("Джак", "J7 PLUS", "Flagship 1.5TGDI 6DCT BR", 415600, "Джак 5-ДП 20.03.2026", year=2025),
    rule("Джак", "JS4", "Luxury 1.5T CVT", 279600, "Джак 5-ДП 20.03.2026", year=2024),
    rule("Джак", "JS4", "Intelligent 1.5T CVT BR", 367600, "Джак 5-ДП 20.03.2026", year=2025),
    rule("Джак", "JS8", "Luxury 1.5TGDI 7DCT 7-seats JL", 475600, "Джак 5-ДП 20.03.2026", year=2025),
    rule("Джак", "T6", "Luxury 2.0 Ti 4x4", 887400, "Джак 8-ДП 23.04.2026", year=2025),
    rule("Джак", "T8 Pro", "Luxury 2.4� MT 4x4", 947400, "Джак 8-ДП 23.04.2026", year=2025),
    rule("Джак", "T8 Pro", "Luxury 2.4 MT 4x4", 947400, "Джак 8-ДП 23.04.2026", year=2025),
    rule("Джак", "T8 Pro", "Luxury 2.4Т MT 4x4", 947400, "Джак 8-ДП 23.04.2026", year=2025),
    rule("Джак", "T9", "Luxury 2.0T AT 4x4", 1049400, "Джак 8-ДП 23.04.2026", year=2025),

    # Jetour
    rule("Джетур", "X50", "PRESTIGE", 259600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "X50", "PREMIUM", 299600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "X50", "PRESTIGE", 279600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2026),
    rule("Джетур", "X50", "PRESTIGE", 319600, "Приказ РПЦ 3-ОД Джетур", year=2026),
    rule("Джетур", "X50", "PREMIUM", 319600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2026),
    rule("Джетур", "X50", "PREMIUM", 359600, "Приказ РПЦ 3-ОД Джетур", year=2026),
    rule("Джетур", "X70", "COMFORT", 351600, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "X70", "PREMIUM", 319600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "X70", "PREMIUM", 339600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2026),
    rule("Джетур", "X70FL", "PRESTIGE", 299600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "X70FL", "PRESTIGE", 359600, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "X70FL", "PRESTIGE", 399500, "Приказ РРЦ 32-ОД 17.04 джетур", year=2026),
    rule("Джетур", "X70FL", "PRESTIGE", 474500, "Приказ РПЦ 3-ОД Джетур", year=2026),
    rule("Джетур", "X70FL", "PREMIUM", 339600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "X70FL", "PREMIUM", 398000, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "X70FL", "PREMIUM", 399600, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "X70FL", "PREMIUM", 449500, "Приказ РРЦ 32-ОД 17.04 джетур", year=2026),
    rule("Джетур", "X70FL", "PREMIUM", 524500, "Приказ РПЦ 3-ОД Джетур", year=2026),
    rule("Джетур", "Dashing", "PRESTIGE", 359600, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "Dashing", "PRESTIGE", 379600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "Dashing", "PREMIUM", 399600, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "Dashing", "PREMIUM", 419600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "X70Plus", "PREMIUM PRO", 479600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "X70Plus", "PREMIUM PRO", 649500, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "X70Plus", "PREMIUM PRO", 624500, "Приказ РРЦ 32-ОД 17.04 джетур", year=2026),
    rule("Джетур", "X90Plus", "1.6T PREMIUM", 459600, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "X90Plus", "1.6T PREMIUM", 479600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "X90Plus", "2.0T PREMIUM", 519600, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "X90Plus", "2.0T PREMIUM", 539600, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "T1", "ADVENTURE", 869400, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "T1", "ADVENTURE", 899400, "Приказ РРЦ 32-ОД 17.04 джетур; Приказ РПЦ 3-ОД Джетур", year=2026),
    rule("Джетур", "T1", "EXPEDITION", 774500, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "T1", "EXPEDITION", 959400, "Приказ РРЦ 32-ОД 17.04 джетур; Приказ РПЦ 3-ОД Джетур", year=2026),
    rule("Джетур", "T1", "EXPEDITION", 989400, "Приказ РПЦ 3-ОД Джетур", year=2026),
    rule("Джетур", "T2", "ADVENTURE", 724500, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "T2", "ADVENTURE", 959400, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "T2", "EXPEDITION", 774500, "Приказ РРЦ 32-ОД 17.04 джетур", year=2025),
    rule("Джетур", "T2", "EXPEDITION", 1049400, "Приказ РПЦ 3-ОД Джетур", year=2025),
    rule("Джетур", "T2", "EXPEDITION", 959400, "Приказ РРЦ 32-ОД 17.04 джетур", year=2026),
    rule("Джетур", "T2", "EXPEDITION", 1073400, "Приказ РПЦ 3-ОД Джетур", year=2026),
    rule("Джетур", "T2", "EXPEDITION", 1079400, "Приказ РПЦ 3-ОД Джетур", year=2026),
]


SHEET_TO_BRAND = {
    "Киа": "Киа",
    "Шеви": "Шеви",
    "Джак": "Джак",
    "Джетур": "Джетур",
}


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().upper()
    text = text.replace("Ё", "Е").replace("—", "-")
    text = text.replace("\xa0", " ").replace("�", "")
    text = " ".join(text.split())
    return text


def normalize_year(value: Any) -> int | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def normalize_engine(value: Any) -> int | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def normalize_reward(value: Any) -> int | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return None


def match_rule(
    current_brand: str,
    current_model: str,
    current_trim: str,
    current_year: int | None,
    current_drive: str,
    current_engine: int | None,
    candidate: BulletinRule,
) -> bool:
    if current_brand != candidate.brand:
        return False
    if normalize_text(current_model) != normalize_text(candidate.model):
        return False
    if normalize_text(current_trim) != normalize_text(candidate.trim):
        return False
    if candidate.year is not None and current_year != candidate.year:
        return False
    if candidate.drive is not None and normalize_text(current_drive) != normalize_text(candidate.drive):
        return False
    if candidate.engine is not None and current_engine != candidate.engine:
        return False
    return True


def collect_sheet_rows(sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(WORKBOOK_PATH, sheet_name=sheet_name)
    brand = SHEET_TO_BRAND[sheet_name]
    model_col = next(col for col in df.columns if "Модель" in str(col))
    trim_col = next(col for col in df.columns if "Комплектац" in str(col))
    date_col = next(col for col in df.columns if "Дата договора" in str(col))
    reward_col = next(col for col in df.columns if "Максимальное вознаграждение дилера" in str(col))
    vin_col = next(col for col in df.columns if "VIN" in str(col))
    year_cols = [col for col in df.columns if "Год выпуска" in str(col)]
    drive_cols = [col for col in df.columns if "Привод" in str(col)]
    engine_cols = [col for col in df.columns if "Объем двигателя" in str(col)]

    year_col = year_cols[0] if year_cols else None
    drive_col = drive_cols[0] if drive_cols else None
    engine_col = engine_cols[0] if engine_cols else None

    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df[(df[date_col] >= APRIL_START) & (df[date_col] < APRIL_END)].copy()

    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        current_model = row[model_col]
        current_trim = row[trim_col]
        current_year = normalize_year(row[year_col]) if year_col else None
        current_drive = row[drive_col] if drive_col else ""
        current_engine = normalize_engine(row[engine_col]) if engine_col else None
        current_reward = normalize_reward(row[reward_col])
        matching_rules = [
            item
            for item in RULES
            if match_rule(
                brand,
                str(current_model),
                str(current_trim),
                current_year,
                str(current_drive or ""),
                current_engine,
                item,
            )
        ]

        allowed_rewards = sorted({item.reward for item in matching_rules})
        reward_match = [item for item in matching_rules if item.reward == current_reward]

        if reward_match:
            status = "ok"
            source_text = "; ".join(sorted({item.source for item in reward_match}))
        elif matching_rules:
            status = "reward_not_in_bulletin"
            source_text = "; ".join(sorted({item.source for item in matching_rules}))
        else:
            status = "combo_not_found"
            source_text = ""

        rows.append(
            {
                "Бренд": brand,
                "Дата договора": row[date_col],
                "VIN": row[vin_col],
                "Модель": current_model,
                "Комплектация": current_trim,
                "Год выпуска": current_year,
                "Привод": current_drive,
                "Объем двигателя": current_engine,
                "Макс вознаграждение в файле": current_reward,
                "Найденные значения в бюллетенях": ", ".join(f"{value:,}".replace(",", " ") for value in allowed_rewards),
                "Источники": source_text,
                "Статус": status,
            }
        )

    return pd.DataFrame(rows)


def main() -> None:
    all_rows = pd.concat([collect_sheet_rows(sheet_name) for sheet_name in SHEET_TO_BRAND], ignore_index=True)
    summary = (
        all_rows.groupby(["Бренд", "Статус"], as_index=False)
        .size()
        .rename(columns={"size": "Количество"})
        .sort_values(["Бренд", "Статус"])
    )

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        all_rows.sort_values(["Бренд", "Статус", "Дата договора", "Модель", "Комплектация"]).to_excel(
            writer,
            sheet_name="Сверка",
            index=False,
        )
        summary.to_excel(writer, sheet_name="Сводка", index=False)
        all_rows[all_rows["Статус"] != "ok"].to_excel(writer, sheet_name="Проблемы", index=False)

    print(f"Saved report: {OUTPUT_PATH}")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()