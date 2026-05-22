"""
daily_margin_report.py
======================
Ежедневный отчёт по марже продаж авто.

Использование
-------------
1. Скопируйте выгрузки из Битрикса (xlsx) в папку  bitrix_exports/
   (по одному файлу на бренд, имя файла должно содержать название бренда:
    kia / киа / chevrolet / шеви / jac / джак / jetour / джетур)

2. Запустите скрипт:
       python daily_margin_report.py

   Или выберите диапазон дат:
       python daily_margin_report.py --from 2026-05-01 --to 2026-05-19

3. Готовый отчёт появится в папке  отчёты_маржа/

Структура отчёта
----------------
  • Все сделки   — каждая строка продажи + статус по бюллетеню
  • Сводка       — итоги по бренду
  • Проблемы     — только строки с расхождением или без бюллетеня
  • По дням      — динамика продаж / маржи по дням
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import date, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

# Импорт базы правил
try:
    from bulletin_rules_db import find_active_rewards, normalize_int, normalize_text
except ImportError:
    sys.exit(
        "Ошибка: не найден файл bulletin_rules_db.py рядом со скриптом."
    )

# ---------------------------------------------------------------------------
# Настройки путей
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent
EXPORTS_DIR = BASE_DIR / "bitrix_exports"
REPORTS_DIR = BASE_DIR / "отчёты_маржа"

# ---------------------------------------------------------------------------
# Определение бренда по имени файла
# ---------------------------------------------------------------------------

BRAND_KEYWORDS: dict[str, list[str]] = {
    "Киа":    ["kia", "киа"],
    "Шеви":   ["chevrolet", "шеви", "chevy", "chevi", "чеви"],
    "Джак":   ["jac", "джак"],
    "Джетур": ["jetour", "джетур"],
}

# ---------------------------------------------------------------------------
# Белый список наших ДЦ по бренду.
# Для Киа и Шеви учитываем и коды, и алиасы в названии (например, AMQ).
# ---------------------------------------------------------------------------

OWN_DC_CODES: dict[str, set[int]] = {
    "Киа": {
        1069,   # Актау TOO «DCG Aktau»
        1151,   # Алматы ТОО «DCG Qulja»
        331,    # Шымкент ТОО «DCG ONTUSTIK»
        363,    # Алматы ТОО «DCG ALMATY»
        365,    # Тараз ТОО «DCGTARAZ»
        371,    # Астана ТОО «DCG ASTANA»
        735,    # Усть-Каменогорск ТОО «DCG VOSTOK»
        863,    # Павлодар ТОО «DCG ERTIS»
        1169, 1173, 1174, 1177, 1181, 1184, 1190, 1192,  # AMQ-алиасы тех же ДЦ
    },
    "Шеви": {
        1085,   # Актау ТОО «AUTOCENTER ONTUSTIK»
        1120,   # Атырау ТОО «AUTOCENTER ONTUSTIK»
        1159,   # Алматы ТОО «AUTOCENTER ONTUSTIK»
        119,    # Шымкент ТОО «AUTOCENTER ONTUSTIK»
        1267,   # Алматы TOO «DCG Zhetysu»
        397,    # Тараз ТОО «AUTOCENTER ONTUSTIK»
        858,    # Астана ТОО «DCG Astana»
        866,    # Усть-Каменогорск ТОО «DCG VOSTOK»
        872,    # Павлодар ТОО «DCG ERTIS»
        920,    # Уральск ТОО «AUTOCENTER ONTUSTIK»
        922,    # Кызылорда ТОО «AUTOCENTER ONTUSTIK»
        935,    # Шымкент TOO «AUTOCENTER ONTUSTIK» Байдибек би
    },
    "Джак": {
        1057,   # Шымкент ТОО «DCG Baidybek»
        1070,   # Тараз ТОО «DCG TARAZ»
        1083,   # Актау TOO «DCG Aktau»
        1123,   # Атырау ТОО «DCG Atyrau»
        327,    # Шымкент ТОО «AUTOCENTER ONTUSTIK»
        747,    # Усть-Каменогорск ТОО «DCG VOSTOK»
        762,    # Шымкент ТОО «AUTOCENTER ONTUSTIK» Темирлановское шоссе
        824,    # Астана ТОО «DCG Astana»
        860,    # Павлодар ТОО «DCG ERTIS»
        951,    # Кызылорда ТОО «DCG Orda»
        952,    # Уральск ТОО «DCG Oral»
    },
    "Джетур": {
        1005,   # Алматы ТОО «DCG Alatau»
        1082,   # Актау TOO «DCG Aktau»
        1125,   # Атырау ТОО «DCG Atyrau»
        1167,   # Алматы ТОО «DCG Zhetysu»
        908,    # Усть-Каменогорск ТОО «DCG VOSTOK»
        911,    # Шымкент ТОО «Ontustik Auto»
        926,    # Павлодар ТОО «DCG ERTIS»
        938,    # Кызылорда ТОО «DCG Orda»
        948,    # Тараз ТОО «DCG TARAZ»
        955,    # Шымкент ТОО «DCG Baidybek»
        998,    # Уральск ТОО «DCG Oral»
    },
}

OWN_DC_PATTERNS: dict[str, list[str]] = {
    "Киа": [
        "dcg aktau",
        "dcg qulja",
        "dcg ontustik",
        "dcg almaty",
        "dcgtaraz",
        "dcg astana",
        "dcg vostok",
        "dcg ertis",
    ],
    "Шеви": [
        "autocenter ontustik",
        "dcg zhetysu",
        "dcg astana",
        "dcg vostok",
        "dcg ertis",
        "autocenter ontustik байдибек би",
    ],
    "Джак": [
        "dcg baidybek",
        "dcg taraz",
        "dcg aktau",
        "dcg atryau",
        "autocenter ontustik",
        "dcg vostok",
        "dcg astana",
        "dcg ertis",
        "dcg orda",
        "dcg oral",
    ],
    "Джетур": [
        "dcg alatau",
        "dcg aktau",
        "dcg atryau",
        "dcg zhetysu",
        "dcg vostok",
        "ontustik auto",
        "dcg ertis",
        "dcg orda",
        "dcg taraz",
        "dcg baidybek",
        "dcg oral",
    ],
}


def is_own_dc(brand: str, dc_name: str) -> bool:
    """True, если ДЦ входит в белый список для данного бренда."""
    if not dc_name:
        return False
    text = normalize_text(str(dc_name)).lower()

    patterns = OWN_DC_PATTERNS.get(brand, [])
    if any(pattern in text for pattern in patterns):
        return True

    m = re.match(r"^(\d+)", str(dc_name).strip())
    if not m:
        return False
    code = int(m.group(1))
    allowed = OWN_DC_CODES.get(brand)
    if allowed is None:
        return True   # бренд без фильтра — пропускаем всё
    return code in allowed


def detect_brand(filename: str) -> str | None:
    lower = filename.lower()
    for brand, keywords in BRAND_KEYWORDS.items():
        if any(kw in lower for kw in keywords):
            return brand
    return None


# ---------------------------------------------------------------------------
# Поиск нужных столбцов
# ---------------------------------------------------------------------------

# Ключевые слова для поиска нужных колонок (ищем в названии без учёта регистра)
COLUMN_PATTERNS: dict[str, list[str]] = {
    "model":   ["модел"],
    "trim":    ["комплектац"],
    "date":    ["дата договора", "дата продажи", "дата закрыти", "дата сделки", "дата создан"],
    "vin":     ["vin"],
    "reward":  ["максимал", "вознаграждени", "маржа", "margin"],
    "reward_no_discount": ["вознаграждение без скид", "без скид"],
    "year":    ["год выпуска", "год авто", "год изготовл"],
    "drive":   ["привод"],
    "engine":  ["объем двигателя", "объём двигателя", "объем дв", "объём дв"],
    "amount":  ["цена реализац", "реализац", "сумма", "стоимость", "цена продаж", "итого"],
    "rrc":     ["рекомендуемая розничная цена", "ррц", "rrc"],
    "discount": ["сумма скид", "скидка"],
    "manager": ["фио", "кто продал", "менеджер", "ответствен", "сотрудник"],
    "status":  ["стату", "состояние сделки", "этап"],
    "period_date": ["дата подписания акта", "дата акта", "дата продажи", "дата договора"],
    # Дополнительные поля для аналитики по ДЦ
    "dc":      ["название дц"],
    "extra_comp": ["дополнительная компенсация дилера"],
    "comp2":   ["компенсация дилера"],
    "transfer": ["вознаграждение дилера (за перемещ"],
}


def find_column(df: pd.DataFrame, field: str) -> str | None:
    patterns = COLUMN_PATTERNS.get(field, [])
    cols_lower = {c.lower(): c for c in df.columns}
    for pat in patterns:
        for low, orig in cols_lower.items():
            if pat in low:
                return orig
    return None


def choose_date_column(df: pd.DataFrame, brand: str, fallback: str | None) -> str | None:
    if brand == "Джак":
        sale_col = next(
            (column for column in df.columns if "дата продажи" in column.lower()),
            None,
        )
        if sale_col:
            return sale_col
    return fallback


def parse_date_series(series: pd.Series) -> pd.Series:
    sample = next(
        (str(value).strip() for value in series if pd.notna(value) and str(value).strip()),
        "",
    )
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}(?: \d{2}:\d{2}:\d{2})?", sample):
        return pd.to_datetime(series, errors="coerce", dayfirst=False)
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def parse_date_value(value: Any) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    if not text:
        return None
    dayfirst = not bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}(?: \d{2}:\d{2}:\d{2})?", text))
    parsed = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
    if pd.isna(parsed):
        return None
    return parsed.date()


def detect_sale_contract_columns(df: pd.DataFrame) -> tuple[str | None, str | None]:
    cols = list(df.columns)
    cols_lower = [c.lower() for c in cols]

    sale_col = None
    contract_col = None

    for i, low in enumerate(cols_lower):
        if sale_col is None and ("дата продажи" in low or "дата закрыти" in low):
            sale_col = cols[i]
        if contract_col is None and ("дата договора" in low or "дата создан" in low):
            contract_col = cols[i]

    return sale_col, contract_col


def detect_period_column(df: pd.DataFrame, fallback: str | None) -> str | None:
    """Колонка для отбора периода отчёта: приоритет у даты подписания акта."""
    for column in df.columns:
        low = column.lower()
        if "дата подписания акта" in low or "дата акта" in low:
            return column
    return fallback


def pick_bulletin_by_dates(
    brand: str,
    model: str,
    trim: str,
    year: int | None,
    drive: str,
    engine: int | None,
    reward_val: int | None,
    sale_dt: date | None,
    contract_dt: date | None,
) -> tuple[list[int], str, date | None, str]:
    candidates: list[tuple[str, date]] = []
    if sale_dt:
        candidates.append(("дата продажи", sale_dt))
    if contract_dt and contract_dt != sale_dt:
        candidates.append(("дата договора", contract_dt))

    if not candidates:
        return [], "", None, "дата не определена"

    checks: list[tuple[str, date, list[int], str, date | None, bool]] = []
    for label, ref_date in candidates:
        allowed, source, bulletin_date = find_active_rewards(
            brand=brand,
            model=model,
            trim=trim,
            year=year,
            drive=drive,
            engine=engine,
            sale_date=ref_date,
        )
        matched = reward_val is not None and reward_val in allowed
        checks.append((label, ref_date, allowed, source, bulletin_date, matched))

    matched_checks = [c for c in checks if c[5]]
    if matched_checks:
        # Если обе даты дают совпадение, приоритет у даты продажи.
        matched_checks.sort(key=lambda c: (c[0] != "дата продажи", c[1]))
        label, _, allowed, source, bulletin_date, _ = matched_checks[0]
        return allowed, source, bulletin_date, label

    # Если совпадения нет, используем приоритет даты продажи, затем договора.
    checks.sort(key=lambda c: (c[0] != "дата продажи", c[1]))
    label, _, allowed, source, bulletin_date, _ = checks[0]
    return allowed, source, bulletin_date, label


def classify_mismatch_reason(
    allowed: list[int],
    reward_val: int | None,
    reward_no_discount_val: int | None,
    discount_val: int | None,
    amount_val: int | None,
    rrc_val: int | None,
) -> str:
    if not allowed:
        return "Нет правила в бюллетене"
    if reward_val in allowed:
        return "Совпадает"

    bulletin_reward = max(allowed)

    # Частый кейс: по дате уже новый бюллетень, но в сделке использована старая/сниженная база РРЦ,
    # из-за чего даже "без скидки" выплата ниже актуального бюллетеня.
    if reward_no_discount_val is not None and reward_no_discount_val < bulletin_reward:
        if amount_val is not None and rrc_val is not None and amount_val < rrc_val:
            return "Маржа ниже из-за цены ниже РРЦ/сниженной РРЦ"
        return "Маржа ниже актуального бюллетеня (возможна старая РРЦ/база)"

    if discount_val and discount_val > 0:
        if reward_no_discount_val in allowed:
            return "Расхождение из-за скидки"
        return "Есть скидка, но сумма всё равно не совпала"
    if reward_no_discount_val in allowed:
        return "Расхождение без скидки"
    return "Расхождение по вознаграждению"


# ---------------------------------------------------------------------------
# Чтение одного файла выгрузки
# ---------------------------------------------------------------------------


def read_brand_file(
    path: Path,
    brand: str,
    date_from: date,
    date_to: date,
) -> pd.DataFrame:
    """Читает Excel-файл бренда, возвращает DataFrame с унифицированными колонками."""
    try:
        xl = pd.ExcelFile(path)
        # Выбираем лист, чьё имя начинается с "report" (регистронезависимо)
        data_sheet = next(
            (s for s in xl.sheet_names if s.lower().startswith("report")),
            xl.sheet_names[-1],  # fallback: последний лист
        )
        print(f"         лист: {data_sheet!r} (из {xl.sheet_names})")
        raw = pd.read_excel(xl, sheet_name=data_sheet, dtype=str)
    except Exception as exc:
        print(f"  [!] Не удалось открыть {path.name}: {exc}")
        return pd.DataFrame()

    if raw.empty:
        return pd.DataFrame()

    # --- Найти нужные колонки ---
    col = {field: find_column(raw, field) for field in COLUMN_PATTERNS}
    sale_col, contract_col = detect_sale_contract_columns(raw)
    col["date"] = choose_date_column(raw, brand, col["date"])

    if col["date"] is None:
        col["date"] = sale_col or contract_col

    if col["model"] is None:
        print(f"  [!] {path.name}: не найдена колонка 'Модель' — пропускаем")
        return pd.DataFrame()
    if col["date"] is None:
        print(f"  [!] {path.name}: не найдена колонка с датой договора — пропускаем")
        return pd.DataFrame()

    # --- Преобразование дат ---
    raw["__date__"] = parse_date_series(raw[col["date"]])
    period_col = detect_period_column(raw, col["date"])
    if period_col is None:
        period_col = col["date"]
    raw["__period_date__"] = parse_date_series(raw[period_col])

    # Фильтр по диапазону: по дате периода (дата акта, если есть)
    mask = (
        (raw["__period_date__"].dt.date >= date_from)
        & (raw["__period_date__"].dt.date <= date_to)
    )
    raw = raw[mask].copy()
    if raw.empty:
        return pd.DataFrame()

    def get(row: pd.Series, field: str) -> Any:
        c = col[field]
        return row[c] if c else None

    rows = []
    for _, row in raw.iterrows():
        row_date_raw = row["__date__"]
        if pd.isna(row_date_raw):
            continue

        sale_dt = parse_date_value(row[sale_col]) if sale_col else None
        contract_dt = parse_date_value(row[contract_col]) if contract_col else None

        if sale_dt is None:
            sale_dt = row_date_raw.date()
        if contract_dt is None and contract_col is None:
            contract_dt = row_date_raw.date()

        model_val  = get(row, "model")
        trim_val   = get(row, "trim")
        year_val   = normalize_int(get(row, "year"))
        drive_val  = str(get(row, "drive") or "")
        engine_val = normalize_int(get(row, "engine"))
        reward_val = normalize_int(get(row, "reward"))
        reward_no_discount_val = normalize_int(get(row, "reward_no_discount"))
        amount_val = normalize_int(get(row, "amount"))
        rrc_val = normalize_int(get(row, "rrc"))
        discount_val = normalize_int(get(row, "discount"))
        vin_val    = str(get(row, "vin") or "")
        manager_val= str(get(row, "manager") or "")
        status_val = str(get(row, "status") or "")
        dc_val     = str(get(row, "dc") or "")
        extra_comp_val = normalize_int(get(row, "extra_comp"))
        comp2_val      = normalize_int(get(row, "comp2"))
        transfer_val   = normalize_int(get(row, "transfer"))
        # ДО = сумма всех дополнительных вознаграждений
        do_parts = [v for v in [extra_comp_val, comp2_val, transfer_val] if v is not None]
        do_val = sum(do_parts) if do_parts else None

        # --- Сверка с бюллетенями ---
        allowed, source, bulletin_date, date_used_for_bulletin = pick_bulletin_by_dates(
            brand=brand,
            model=str(model_val or ""),
            trim=str(trim_val or ""),
            year=year_val,
            drive=drive_val,
            engine=engine_val,
            reward_val=reward_val,
            sale_dt=sale_dt,
            contract_dt=contract_dt,
        )

        if not allowed:
            check_status = "❌ Не в бюллетене"
            bulletin_reward = None
            diff = None
        elif reward_val in allowed:
            check_status = "✅ Совпадает"
            bulletin_reward = reward_val
            diff = 0
        else:
            check_status = "⚠️ Расхождение"
            bulletin_reward = allowed[-1]  # берём максимально допустимое
            diff = (reward_val or 0) - bulletin_reward

        allowed_str = (
            ", ".join(f"{v:,.0f}".replace(",", " ") for v in allowed)
            if allowed else ""
        )
        mismatch_reason = classify_mismatch_reason(
            allowed=allowed,
            reward_val=reward_val,
            reward_no_discount_val=reward_no_discount_val,
            discount_val=discount_val,
            amount_val=amount_val,
            rrc_val=rrc_val,
        )

        rrc_delta = None
        if amount_val is not None and rrc_val is not None:
            rrc_delta = amount_val - rrc_val

        lost_margin = None
        if allowed and reward_val is not None:
            lost_margin = max(allowed) - reward_val

        rows.append({
            "Дата продажи":            sale_dt,
            "Дата договора":           contract_dt,
            "Бренд":                   brand,
            "Модель":                  model_val,
            "Комплектация":            trim_val,
            "Год выпуска":             year_val,
            "Привод":                  drive_val if drive_val else None,
            "Объём двигателя":         engine_val,
            "VIN":                     vin_val,
            "Менеджер":                manager_val,
            "Статус сделки":           status_val,
            "Сумма продажи":           amount_val,
            "РРЦ":                     rrc_val,
            "Отклонение цены от РРЦ":  rrc_delta,
            "Сумма скидки":            discount_val,
            "Вознаграждение (Битрикс)":reward_val,
            "Вознаграждение без скидки": reward_no_discount_val,
            "Вознаграждение (бюллет.)":bulletin_reward,
            "Потеря маржи к бюллетеню": lost_margin,
            "Разрешённые значения":    allowed_str,
            "Разница":                 diff,
            "Статус проверки":         check_status,
            "Причина расхождения":     mismatch_reason,
            "Дата для сверки бюллетеня": date_used_for_bulletin,
            "Источник бюллетеня":      source,
            "Дата бюллетеня":          bulletin_date,
            "Файл":                    path.name,
            "ДЦ":                      dc_val if dc_val else None,
            "ДО":                      do_val,
        })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Основная логика
# ---------------------------------------------------------------------------


def build_report(date_from: date, date_to: date) -> None:
    REPORTS_DIR.mkdir(exist_ok=True)
    EXPORTS_DIR.mkdir(exist_ok=True)

    # --- Собрать все xlsx из папки выгрузок ---
    xlsx_files = sorted(EXPORTS_DIR.glob("*.xlsx"))
    if not xlsx_files:
        print(f"\n[!] Папка '{EXPORTS_DIR}' пуста или не содержит .xlsx файлов.")
        print("    Скопируйте туда выгрузки из Битрикса и запустите снова.\n")
        return

    all_frames: list[pd.DataFrame] = []

    for path in xlsx_files:
        brand = detect_brand(path.stem)
        if brand is None:
            print(f"  [?] {path.name}: не удалось определить бренд — пропускаем")
            print(f"      Добавьте в имя файла: kia/киа/chevrolet/шеви/jac/джак/jetour/джетур")
            continue

        print(f"  Читаю: {path.name}  →  {brand}")
        df = read_brand_file(path, brand, date_from, date_to)
        if not df.empty:
            all_frames.append(df)
            print(f"         {len(df)} строк за {date_from} – {date_to}")
        else:
            print(f"         Нет данных за выбранный период")

    if not all_frames:
        print("\n[!] Нет данных для отчёта. Проверьте файлы и диапазон дат.\n")
        return

    all_df = pd.concat(all_frames, ignore_index=True)
    all_df.sort_values(["Дата продажи", "Бренд", "Модель"], inplace=True)

    # --- Листы отчёта ---

    # 1. Анализ по брендам (как «Анализ по мес» в апрельском файле)
    def _brand_analytics(df: pd.DataFrame) -> pd.DataFrame:
        grp = df.groupby("Бренд").agg(
            Количество=("VIN", "count"),
            Маржа_всего=("Вознаграждение (Битрикс)", "sum"),
            Сумма_РРЦ=("РРЦ", "sum"),
            ДО_всего=("ДО", "sum"),
            Сумма_скидок=("Сумма скидки", "sum"),
            Совпадает=("Статус проверки",
                       lambda x: (x == "✅ Совпадает").sum()),
            Расхождение=("Статус проверки",
                         lambda x: (x == "⚠️ Расхождение").sum()),
            Не_найдено=("Статус проверки",
                        lambda x: (x == "❌ Не в бюллетене").sum()),
        ).reset_index()
        grp["Ср маржа"]    = (grp["Маржа_всего"] / grp["Количество"]).round(0)
        grp["К получению"] = grp["Маржа_всего"] + grp["ДО_всего"]
        grp["% маржи"]     = (grp["Маржа_всего"] / grp["Сумма_РРЦ"]).round(6)
        total = pd.DataFrame([{
            "Бренд": "Итого",
            "Количество":  grp["Количество"].sum(),
            "Маржа_всего": grp["Маржа_всего"].sum(),
            "Сумма_РРЦ":   grp["Сумма_РРЦ"].sum(),
            "ДО_всего":    grp["ДО_всего"].sum(),
            "Сумма_скидок":grp["Сумма_скидок"].sum(),
            "Совпадает":   grp["Совпадает"].sum(),
            "Расхождение": grp["Расхождение"].sum(),
            "Не_найдено":  grp["Не_найдено"].sum(),
            "Ср маржа":    (grp["Маржа_всего"].sum() / grp["Количество"].sum()).round(0),
            "К получению": grp["К получению"].sum(),
            "% маржи":     (grp["Маржа_всего"].sum() / grp["Сумма_РРЦ"].sum()).round(6),
        }])
        result = pd.concat([grp, total], ignore_index=True)
        return result[["Бренд", "Количество", "Маржа_всего", "Ср маржа",
                        "ДО_всего", "К получению", "% маржи",
                        "Сумма_скидок", "Сумма_РРЦ",
                        "Совпадает", "Расхождение", "Не_найдено"]].rename(columns={
            "Маржа_всего": "Маржа всего",
            "ДО_всего":    "ДО всего",
            "Сумма_скидок": "Скидки всего",
            "Сумма_РРЦ":   "РРЦ всего",
        })

    # Данные только по нашим ДЦ — для аналитических листов
    own_df = all_df[
        all_df.apply(lambda r: is_own_dc(r["Бренд"], r.get("ДЦ", "")), axis=1)
    ].copy()

    brand_analytics = _brand_analytics(own_df)

    # 2. Анализ по ДЦ (разбивка внутри каждого бренда)
    def _dc_analytics(df: pd.DataFrame) -> pd.DataFrame:
        df2 = df.dropna(subset=["ДЦ"]).copy()
        df2 = df2[df2["ДЦ"].str.strip() != ""]
        if df2.empty:
            return pd.DataFrame()
        grp = df2.groupby(["Бренд", "ДЦ"]).agg(
            Количество=("VIN", "count"),
            Ср_РРЦ=("РРЦ", "mean"),
            Маржа_всего=("Вознаграждение (Битрикс)", "sum"),
            Сумма_скидок=("Сумма скидки", "sum"),
            ДО_всего=("ДО", "sum"),
        ).reset_index()
        rrc_sum = df2.groupby(["Бренд", "ДЦ"])["РРЦ"].sum().reset_index(name="Сумма_РРЦ")
        grp = grp.merge(rrc_sum, on=["Бренд", "ДЦ"])
        grp["Ср маржа"]        = (grp["Маржа_всего"] / grp["Количество"]).round(0)
        grp["Маржа - скидки"]  = grp["Маржа_всего"] - grp["Сумма_скидок"]
        grp["К получению"]     = grp["Маржа_всего"] + grp["ДО_всего"]
        grp["% маржи"]         = (grp["Маржа_всего"] / grp["Сумма_РРЦ"]).round(6)
        grp["Ср РРЦ"]          = grp["Ср_РРЦ"].round(0)
        совп = df2.groupby(["Бренд", "ДЦ"])["Статус проверки"].apply(
            lambda x: (x == "✅ Совпадает").sum()).reset_index(name="Совпадает")
        расх = df2.groupby(["Бренд", "ДЦ"])["Статус проверки"].apply(
            lambda x: (x == "⚠️ Расхождение").sum()).reset_index(name="Расхождение")
        grp = grp.merge(совп, on=["Бренд", "ДЦ"]).merge(расх, on=["Бренд", "ДЦ"])
        return grp[["Бренд", "ДЦ", "Количество", "Ср РРЦ",
                    "Маржа_всего", "Сумма_скидок", "Маржа - скидки",
                    "ДО_всего", "К получению", "Ср маржа", "% маржи",
                    "Совпадает", "Расхождение"]].rename(columns={
            "Маржа_всего":  "Маржа всего",
            "Сумма_скидок": "Скидки всего",
            "ДО_всего":     "ДО всего",
        }).sort_values(["Бренд", "Маржа всего"], ascending=[True, False])

    dc_analytics = _dc_analytics(own_df)

    # 3. Сводка по бренду + статус проверки (только свои ДЦ)
    summary = (
        own_df.groupby(["Бренд", "Статус проверки"])
        .agg(
            Количество=("VIN", "count"),
            Вознаграждение_Битрикс=("Вознаграждение (Битрикс)", "sum"),
            Вознаграждение_Бюллетень=("Вознаграждение (бюллет.)", "sum"),
            Разница_итого=("Разница", "sum"),
        )
        .reset_index()
        .sort_values(["Бренд", "Статус проверки"])
    )

    # 4. Динамика по дням
    by_day = (
        all_df.groupby(["Дата продажи", "Бренд"])
        .agg(
            Продаж=("VIN", "count"),
            Сумма_продаж=("Сумма продажи", "sum"),
            Вознаграждение=("Вознаграждение (Битрикс)", "sum"),
            Расхождений=("Разница", lambda x: (x != 0).sum()),
        )
        .reset_index()
        .sort_values(["Дата продажи", "Бренд"])
    )

    # 5. Только проблемы
    problems = all_df[all_df["Статус проверки"] != "✅ Совпадает"].copy()

    # --- Запись отчёта ---
    label = (
        date_from.strftime("%Y%m%d")
        if date_from == date_to
        else f"{date_from.strftime('%Y%m%d')}_{date_to.strftime('%Y%m%d')}"
    )
    out_path = REPORTS_DIR / f"маржа_{label}.xlsx"

    def _write_all(writer: pd.ExcelWriter) -> None:
        _write_sheet(writer, brand_analytics, "Анализ по брендам", freeze=False)
        if not dc_analytics.empty:
            _write_sheet(writer, dc_analytics, "По ДЦ",           freeze=True)
        _write_sheet(writer, all_df,          "Все сделки",        freeze=True)
        _write_sheet(writer, summary,         "Сводка",            freeze=False)
        _write_sheet(writer, problems,        "Проблемы",          freeze=True)
        _write_sheet(writer, by_day,          "По дням",           freeze=False)

    final_out_path = out_path
    try:
        with pd.ExcelWriter(final_out_path, engine="openpyxl") as writer:
            _write_all(writer)
    except PermissionError:
        final_out_path = REPORTS_DIR / f"маржа_{label}_new.xlsx"
        with pd.ExcelWriter(final_out_path, engine="openpyxl") as writer:
            _write_all(writer)

    print(f"\n✅ Отчёт сохранён: {final_out_path}")
    print(f"   Всего строк: {len(all_df)}")
    print(f"   Совпадает:   {(all_df['Статус проверки'] == '✅ Совпадает').sum()}")
    print(f"   Расхождение: {(all_df['Статус проверки'] == '⚠️ Расхождение').sum()}")
    print(f"   Не найдено:  {(all_df['Статус проверки'] == '❌ Не в бюллетене').sum()}")
    _print_summary(all_df)


def _write_sheet(
    writer: pd.ExcelWriter,
    df: pd.DataFrame,
    sheet_name: str,
    freeze: bool,
) -> None:
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]

    # Авто-ширина колонок
    for col_cells in ws.columns:
        max_len = max(
            (len(str(c.value or "")) for c in col_cells),
            default=8,
        )
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 50)

    # Закрепить первую строку
    if freeze:
        ws.freeze_panes = "A2"


def _print_summary(df: pd.DataFrame) -> None:
    print()
    grp = df.groupby("Бренд")["Статус проверки"].value_counts().unstack(fill_value=0)
    print(grp.to_string())
    print()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Ежедневный отчёт по марже продаж авто"
    )
    parser.add_argument(
        "--from", dest="date_from", metavar="ГГГГ-ММ-ДД",
        help="Начало периода (по умолчанию — сегодня)",
    )
    parser.add_argument(
        "--to", dest="date_to", metavar="ГГГГ-ММ-ДД",
        help="Конец периода (по умолчанию — сегодня)",
    )
    parser.add_argument(
        "--month", metavar="ГГГГ-ММ",
        help="Весь месяц, например: 2026-04",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    today = date.today()

    if args.month:
        y, m = map(int, args.month.split("-"))
        date_from = date(y, m, 1)
        next_month = date(y + (m // 12), (m % 12) + 1, 1)
        date_to = next_month - timedelta(days=1)
    else:
        date_from = date.fromisoformat(args.date_from) if args.date_from else today
        date_to   = date.fromisoformat(args.date_to)   if args.date_to   else today

    print(f"\n{'='*60}")
    print(f"  Отчёт по марже: {date_from} → {date_to}")
    print(f"{'='*60}\n")

    build_report(date_from, date_to)


if __name__ == "__main__":
    main()
