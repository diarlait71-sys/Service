import os
import re
import json
import requests
import pandas as pd
import pdfplumber
from pathlib import Path
from datetime import datetime

# ─── НАСТРОЙКИ ────────────────────────────────────────────────────────────────
SUPABASE_URL = "https://stefapqdbtjisvujxulr.supabase.co"
SUPABASE_KEY = "ВСТАВЬТЕ_ВАШ_ANON_KEY_ИЛИ_SERVICE_ROLE_KEY"  # Settings → API

# Папки с файлами (укажите ваши пути)
PDF_FOLDER   = r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\бюллетени"
EXCEL_FOLDER = r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\bitrix_exports"

BATCH_SIZE = 100  # записей за один запрос (Supabase лимит)

# ─── АЛИАСЫ ДЦ ───────────────────────────────────────────────────────────────
DC_ALIASES = {
    "735":  "Киа ДЦ 735/1181",
    "1181": "Киа ДЦ 735/1181",
    # добавьте свои
}

# ─── БРЕНДЫ по ключевым словам в имени файла ─────────────────────────────────
BRAND_KEYWORDS = {
    "kia": "kia", "киа": "kia",
    "chev": "chev", "шеви": "chev", "chevrolet": "chev",
    "jac": "jac", "джак": "jac",
    "jetr": "jetr", "jetour": "jetr", "джетур": "jetr",
}

# ─── HTTP ХЕЛПЕРЫ ─────────────────────────────────────────────────────────────
HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal",
}

def upsert(table: str, rows: list[dict]) -> None:
    """Batch upsert в Supabase."""
    if not rows:
        return
    for i in range(0, len(rows), BATCH_SIZE):
        batch = rows[i:i + BATCH_SIZE]
        r = requests.post(
            f"{SUPABASE_URL}/rest/v1/{table}",
            headers={**HEADERS, "Prefer": "resolution=merge-duplicates,return=minimal"},
            json=batch,
            timeout=30,
        )
        if r.status_code not in (200, 201):
            print(f"  ⚠️  Ошибка {r.status_code}: {r.text[:300]}")
        else:
            print(f"  ✓  {table}: загружено {len(batch)} строк")

def detect_brand(filename: str) -> str:
    name = filename.lower()
    for kw, brand in BRAND_KEYWORDS.items():
        if kw in name:
            return brand
    return "unknown"

# ─── ПАРСЕР PDF БЮЛЛЕТЕНЕЙ ────────────────────────────────────────────────────
def parse_date(text: str):
    """Пытается распознать дату в разных форматах."""
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text.strip(), fmt).date().isoformat()
        except ValueError:
            continue
    return None

def parse_bulletin_pdf(path: Path) -> list[dict]:
    """
    Читает PDF бюллетеня и извлекает строки правил.

    Ожидаемая структура таблицы в PDF:
    | Модель | Год | Комплектация | Двигатель | Привод | ДЦ | ДО | Действует с | По |

    Если структура другая — подстройте маппинг колонок ниже.
    """
    brand = detect_brand(path.name)
    rules = []

    # Пытаемся вытащить период действия из имени файла
    # Пример: "bulletin_kia_2025_04.pdf" → valid_from=2025-04-01
    date_match = re.search(r"(\d{4})[_\-](\d{2})", path.name)
    file_valid_from = None
    if date_match:
        y, m = date_match.group(1), date_match.group(2)
        file_valid_from = f"{y}-{m}-01"

    try:
        with pdfplumber.open(path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables()
                if not tables:
                    # Страница без таблиц — пробуем текст (fallback)
                    continue

                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    # Первая строка = заголовки
                    headers = [str(h).strip().lower() if h else "" for h in table[0]]

                    # ── Маппинг колонок (подстройте под свой PDF) ──────────
                    # Ключ = что ищем в заголовке, значение = наше поле
                    COL_MAP = {
                        "модел": "model",
                        "год":   "year",
                        "компл": "trim",
                        "двигат": "engine",
                        "привод": "drive",
                        "дц":    "dc_pay",
                        "до":    "do_pay",
                        "с":     "valid_from",
                        "по":    "valid_to",
                        "источн": "source",
                    }
                    col_idx = {}
                    for i, h in enumerate(headers):
                        for kw, field in COL_MAP.items():
                            if kw in h and field not in col_idx:
                                col_idx[field] = i

                    for row in table[1:]:
                        if not row or all(not c for c in row):
                            continue

                        def get(field):
                            idx = col_idx.get(field)
                            if idx is None or idx >= len(row):
                                return None
                            val = row[idx]
                            return str(val).strip() if val else None

                        # Пытаемся распарсить числовые поля
                        def to_num(s):
                            if not s:
                                return None
                            s = re.sub(r"[^\d.,]", "", s).replace(",", ".")
                            try:
                                return float(s)
                            except ValueError:
                                return None

                        valid_from = parse_date(get("valid_from") or "") or file_valid_from
                        valid_to   = parse_date(get("valid_to") or "")

                        rule = {
                            "brand":      brand,
                            "model":      get("model"),
                            "year":       int(get("year")) if get("year") and get("year").isdigit() else None,
                            "trim":       get("trim"),
                            "engine":     get("engine"),
                            "drive":      get("drive"),
                            "dc_pay":     to_num(get("dc_pay")),
                            "do_pay":     to_num(get("do_pay")),
                            "valid_from": valid_from,
                            "valid_to":   valid_to,
                            "source":     get("source") or path.name,
                            "is_disabled": False,
                        }

                        # Пропускаем пустые строки
                        if rule["model"] or rule["dc_pay"]:
                            rules.append(rule)

    except Exception as e:
        print(f"  ⚠️  Не удалось прочитать {path.name}: {e}")

    return rules

# ─── ПАРСЕР EXCEL BITRIX ──────────────────────────────────────────────────────

# Маппинг: как называются колонки в вашей выгрузке → наши поля
# Подстройте под реальные названия колонок в вашем Excel!
BITRIX_COL_MAP = {
    # Ваша колонка     : наше поле
    "VIN":              "vin",
    "Марка":            "brand_raw",
    "Модель":           "model",
    "Год":              "year",
    "Комплектация":     "trim",
    "Двигатель":        "engine",
    "Привод":           "drive",
    "Дилерский центр":  "dc_name",
    "Дата продажи":     "date_sale",
    "Дата договора":    "date_contract",
    "Дата акта":        "date_act",
    "Выплата ДЦ":       "dc_pay_fact",
    "Выплата ДО":       "do_pay_fact",
}

BRAND_NORMALIZE = {
    "kia": "kia", "киа": "kia",
    "chevrolet": "chev", "шевроле": "chev", "шеви": "chev",
    "jac": "jac", "джак": "jac",
    "jetour": "jetr", "джетур": "jetr",
}

def normalize_brand(raw: str) -> str:
    if not raw:
        return "unknown"
    for kw, brand in BRAND_NORMALIZE.items():
        if kw in str(raw).lower():
            return brand
    return str(raw).lower()

def normalize_dc(name: str) -> str:
    if not name:
        return name
    name = str(name).strip()
    return DC_ALIASES.get(name, name)

def parse_bitrix_excel(path: Path) -> list[dict]:
    """
    Читает Excel-выгрузку Bitrix.
    Пробует листы: 'report', 'report (2)', бренд-named листы.
    """
    deals = []
    brand_from_file = detect_brand(path.name)

    try:
        xl = pd.ExcelFile(path)
        sheet_names = xl.sheet_names

        # Ищем нужный лист: сначала 'report*', потом любой
        target_sheets = [s for s in sheet_names if "report" in s.lower()] or sheet_names

        for sheet in target_sheets:
            df = xl.parse(sheet, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]

            # Применяем маппинг колонок
            rename = {}
            for orig_col in df.columns:
                for pattern, field in BITRIX_COL_MAP.items():
                    if pattern.lower() in orig_col.lower():
                        rename[orig_col] = field
                        break
            df = df.rename(columns=rename)

            # Убираем строки без VIN
            if "vin" in df.columns:
                df = df.dropna(subset=["vin"])
            else:
                print(f"  ⚠️  Лист '{sheet}': нет колонки VIN, пропускаем")
                continue

            for _, row in df.iterrows():
                def g(field):
                    v = row.get(field)
                    if isinstance(v, float) and pd.isna(v):
                        return None
                    return str(v).strip() if v else None

                def to_num(field):
                    v = g(field)
                    if not v:
                        return None
                    v = re.sub(r"[^\d.,]", "", v).replace(",", ".")
                    try:
                        return float(v)
                    except ValueError:
                        return None

                def to_date(field):
                    v = g(field)
                    if not v:
                        return None
                    return parse_date(v)

                brand_raw = g("brand_raw") or brand_from_file
                dc_raw = g("dc_name") or ""

                deal = {
                    "vin":           g("vin"),
                    "brand":         normalize_brand(brand_raw),
                    "model":         g("model"),
                    "year":          int(g("year")) if g("year") and g("year").isdigit() else None,
                    "trim":          g("trim"),
                    "engine":        g("engine"),
                    "drive":         g("drive"),
                    "dc_name":       normalize_dc(dc_raw),
                    "date_sale":     to_date("date_sale"),
                    "date_contract": to_date("date_contract"),
                    "date_act":      to_date("date_act"),
                    "dc_pay_fact":   to_num("dc_pay_fact"),
                    "do_pay_fact":   to_num("do_pay_fact"),
                }
                deals.append(deal)

            print(f"  → Лист '{sheet}': {len(deals)} сделок прочитано")

    except Exception as e:
        print(f"  ⚠️  Ошибка при чтении {path.name}: {e}")

    return deals

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  Загрузка данных в Supabase")
    print("=" * 60)

    # 1. Бюллетени (PDF)
    pdf_path = Path(PDF_FOLDER)
    if pdf_path.exists():
        pdf_files = list(pdf_path.glob("*.pdf"))
        print(f"\n📄 Найдено PDF: {len(pdf_files)}")
        all_rules = []
        for f in pdf_files:
            print(f"  Читаю {f.name}...")
            rules = parse_bulletin_pdf(f)
            print(f"  → {len(rules)} правил")
            all_rules.extend(rules)
        print(f"\nИтого правил бюллетеней: {len(all_rules)}")
        upsert("bulletin_rules", all_rules)
    else:
        print(f"\n⚠️  Папка {PDF_FOLDER} не найдена, пропускаем PDF")

    # 2. Bitrix Excel
    xl_path = Path(EXCEL_FOLDER)
    if xl_path.exists():
        xl_files = list(xl_path.glob("*.xlsx")) + list(xl_path.glob("*.xls"))
        print(f"\n📊 Найдено Excel: {len(xl_files)}")
        all_deals = []
        for f in xl_files:
            print(f"  Читаю {f.name}...")
            deals = parse_bitrix_excel(f)
            all_deals.extend(deals)
        # Дедупликация по VIN
        seen = set()
        unique_deals = []
        for d in all_deals:
            if d["vin"] and d["vin"] not in seen:
                seen.add(d["vin"])
                unique_deals.append(d)
        print(f"\nИтого уникальных сделок: {len(unique_deals)}")
        upsert("deals", unique_deals)
    else:
        print(f"\n⚠️  Папка {EXCEL_FOLDER} не найдена, пропускаем Excel")

    print("\n✅ Готово!")

if __name__ == "__main__":
    main()
