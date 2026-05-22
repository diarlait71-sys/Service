"""
bulletin_rules_db.py
====================
База правил маржи по бюллетеням.

Как добавить новый бюллетень
----------------------------
1. Найдите дату бюллетеня (дата приказа/вступления в силу).
2. Добавьте строки rule(...) с этой датой в valid_from.
3. Если новый бюллетень ЗАМЕНЯЕТ предыдущий — просто добавьте новую строку;
   старая автоматически перестанет применяться для дат после valid_from нового.

Поля BulletinRule:
  brand      — бренд ("Киа", "Шеви", "Джак", "Джетур")
  model      — модель ("K5", "Sportage", ...)
  trim       — комплектация ("Comfort", "Luxe", ...)
  reward     — сумма дилерского вознаграждения в тенге
  source     — название документа (для справки)
  valid_from — дата вступления бюллетеня в силу ("ГГГГ-ММ-ДД")
    valid_to   — дата окончания действия бюллетеня (None = без окончания)
  year       — год выпуска авто (None = любой)
  drive      — привод ("2WD", "4WD", None = любой)
  engine     — объём двигателя в куб.см (None = любой)
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any

import pandas as pd


# ---------------------------------------------------------------------------
# Модель данных
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class BulletinRule:
    brand: str
    model: str
    trim: str
    reward: int
    source: str
    valid_from: date
    valid_to: date | None = None
    year: int | None = None
    drive: str | None = None
    engine: int | None = None


def rule(
    brand: str,
    model: str,
    trim: str,
    reward: int,
    source: str,
    valid_from: str,
    *,
    valid_to: str | None = None,
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
        valid_from=date.fromisoformat(valid_from),
        valid_to=date.fromisoformat(valid_to) if valid_to else None,
        year=year,
        drive=drive,
        engine=engine,
    )


# ---------------------------------------------------------------------------
# База правил — ДОБАВЛЯЙТЕ НОВЫЕ БЮЛЛЕТЕНИ СЮДА
# ---------------------------------------------------------------------------

RULES: list[BulletinRule] = [

    # ════════════════════════════════════════════════════════════════════════
    # КИА  (Kia SD DC)
    # ════════════════════════════════════════════════════════════════════════

    # --- Базовые правила (без конкретной даты бюллетеня) ---
    rule("Киа", "Carnival", "Luxe+",     1_429_350, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Carnival", "Limousine+",1_884_350, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Ceed",     "Comfort",     436_050, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Ceed",     "Luxe",        481_050, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Ceed SW",  "Comfort",     494_550, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Ceed SW",  "Luxe",        517_050, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Ceed SW",  "Prestige",    562_050, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Cerato",   "Comfort",     399_600, "Kia bulletins", "2025-01-01", year=2025, engine=1591),
    rule("Киа", "Cerato",   "Comfort",     399_600, "Kia bulletins", "2025-01-01", year=2026, engine=1591),
    rule("Киа", "Cerato",   "Luxe",        439_600, "Kia bulletins", "2025-01-01", year=2025, engine=1999),
    rule("Киа", "Cerato",   "Luxe+",       419_600, "Kia bulletins", "2025-01-01", year=2025, engine=1591),
    rule("Киа", "Cerato",   "Luxe+",       419_600, "Kia bulletins", "2025-01-01", year=2026, engine=1591),
    rule("Киа", "Cerato",   "Luxe+",       439_600, "Kia bulletins", "2025-01-01", year=2025, engine=2),
    rule("Киа", "Cerato",   "Luxe+",       439_600, "Kia bulletins", "2025-01-01", year=2026, engine=2),
    rule("Киа", "K8",       "Prestige",  1_721_850, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Seltos",   "Comfort",     494_550, "Kia bulletins", "2025-01-01", year=2025, drive="2WD", engine=1999),
    rule("Киа", "Seltos",   "Comfort",     521_550, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Seltos",   "Comfort",     544_050, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Seltos",   "Luxe",        539_550, "Kia bulletins", "2025-01-01", year=2025, drive="2WD", engine=1999),
    rule("Киа", "Seltos",   "Luxe",        566_550, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Seltos",   "Style",       607_050, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Seltos",   "Tulpar",      584_550, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Seltos",   "Tulpar",      607_050, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Seltos",   "Tulpar",      607_050, "Kia bulletins", "2026-04-20", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Soluto",   "Comfort",     300_000, "Kia bulletins", "2025-01-01", year=2026),
    rule("Киа", "Soluto",   "Prestige",    300_000, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Soluto",   "Prestige",    300_000, "Kia bulletins", "2025-01-01", year=2026),
    rule("Киа", "Sorento",  "Comfort",     809_550, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Comfort",     934_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Comfort",     854_550, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Luxe",        872_550, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Luxe",        989_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Luxe",      1_011_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Luxe",        899_550, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Luxe",        917_550, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Premium",   1_034_550, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Premium",   1_209_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Premium",   1_079_550, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Premium",   1_147_050, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=3470),
    rule("Киа", "Sorento",  "Style",     1_082_950, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Style",     1_099_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sorento",  "Tulpar",    1_044_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Soul",     "Luxe",        526_050, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Soul",     "Style",       584_550, "Kia bulletins", "2025-01-01", year=2025),
    rule("Киа", "Sportage", "Comfort",     383_700, "Kia bulletins", "2025-01-01", year=2026, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Luxe",        796_950, "Kia bulletins", "2025-01-01", year=2025, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Luxe",        818_950, "Kia bulletins", "2025-01-01", year=2025, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Luxe",        824_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=1975),
    rule("Киа", "Sportage", "Luxe",        857_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=1975),
    rule("Киа", "Sportage", "Luxe",        741_950, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=1999),
    rule("Киа", "Sportage", "Luxe",        769_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "Luxe",        824_450, "Kia bulletins", "2025-01-01", year=2026, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Luxe",        857_450, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=1975),
    rule("Киа", "Sportage", "Luxe",        741_950, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Sportage", "Style",       912_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=1975),
    rule("Киа", "Sportage", "Style",       829_950, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "Style",       950_950, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "X-Line+",  1_022_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "X-Line+",  1_077_450, "Kia bulletins", "2025-01-01", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "X-Line+",  1_077_450, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "X-Line+",    906_950, "Kia bulletins", "2025-01-01", year=2026, drive="4WD", engine=2497),

    # --- Бюллетень 17-03-2026 ---
    rule("Киа", "Seltos",   "Luxe",        634_050, "Kia SD DC 17-03-2026", "2026-03-17", year=2025, drive="2WD", engine=1999),
    rule("Киа", "Sportage", "Luxe",        873_950, "Kia SD DC 17-03-2026", "2026-03-17", year=2026, drive="2WD", engine=1975),
    rule("Киа", "Sportage", "Style",     1_005_950, "Kia SD DC 17-03-2026", "2026-03-17", year=2025, drive="4WD", engine=2497),
    rule("Киа", "Sportage", "X-Line+",  1_005_950, "Kia SD DC 17-03-2026", "2026-03-17", year=2025, drive="4WD", engine=2497),
    rule("Киа", "K5",       "Comfort",     774_950, "Kia SD DC 17-03-2026", "2026-03-17", year=2026),
    rule("Киа", "K5",       "Luxe",        752_950, "Kia SD DC 17-03-2026", "2026-03-17", year=2025),
    rule("Киа", "K5",       "Luxe",        879_450, "Kia SD DC 17-03-2026", "2026-03-17", year=2026),
    rule("Киа", "K5",       "Premium",   1_005_950, "Kia SD DC 17-03-2026", "2026-03-17", year=2026),
    rule("Киа", "K5",       "Tulpar",      791_450, "Kia SD DC 17-03-2026", "2026-03-17", year=2025),
    rule("Киа", "Carnival", "Luxe+",     1_429_350, "Kia SD DC 17-03-2026", "2026-03-17", year=2025),
    rule("Киа", "Carnival", "Limousine+",1_884_350, "Kia SD DC 17-03-2026", "2026-03-17", year=2025),

    # --- Бюллетень 03-04-2026 ---
    rule("Киа", "K5",       "Comfort",     604_450, "Kia SD DC 03-04-2026", "2026-04-03", year=2025),
    rule("Киа", "K5",       "Comfort",     637_450, "Kia SD DC 03-04-2026", "2026-04-03", year=2026),
    rule("Киа", "K5",       "Luxe",        697_950, "Kia SD DC 03-04-2026", "2026-04-03", year=2025),
    rule("Киа", "K5",       "Luxe",        741_950, "Kia SD DC 03-04-2026", "2026-04-03", year=2026),
    rule("Киа", "K5",       "Prestige",    769_450, "Kia SD DC 03-04-2026", "2026-04-03", year=2025),
    rule("Киа", "K5",       "Prestige",    813_450, "Kia SD DC 03-04-2026", "2026-04-03", year=2026),
    rule("Киа", "K5",       "Premium",     868_450, "Kia SD DC 03-04-2026", "2026-04-03", year=2026),
    rule("Киа", "K5",       "Tulpar",      736_450, "Kia SD DC 03-04-2026", "2026-04-03", year=2025),
    rule("Киа", "Carnival", "Luxe+",     1_429_350, "Kia SD DC 03-04-2026", "2026-04-03", year=2025),
    rule("Киа", "Carnival", "Limousine+",1_884_350, "Kia SD DC 03-04-2026", "2026-04-03", year=2025),

    # --- Бюллетень 20-04-2026 ---
    rule("Киа", "K5",       "Comfort",     604_450, "Kia SD DC 20-04-2026", "2026-04-20", year=2025),
    rule("Киа", "K5",       "Comfort",     637_450, "Kia SD DC 20-04-2026", "2026-04-20", year=2026),
    rule("Киа", "K5",       "Luxe",        697_950, "Kia SD DC 20-04-2026", "2026-04-20", year=2025),
    rule("Киа", "K5",       "Luxe",        741_950, "Kia SD DC 20-04-2026", "2026-04-20", year=2026),
    rule("Киа", "K5",       "Prestige",    769_450, "Kia SD DC 20-04-2026", "2026-04-20", year=2025),
    rule("Киа", "K5",       "Prestige",    813_450, "Kia SD DC 20-04-2026", "2026-04-20", year=2026),
    rule("Киа", "K5",       "Premium",     868_450, "Kia SD DC 20-04-2026", "2026-04-20", year=2026),
    rule("Киа", "K5",       "Tulpar",      736_450, "Kia SD DC 20-04-2026", "2026-04-20", year=2025),
    rule("Киа", "Carnival", "Luxe+",     1_429_350, "Kia SD DC 20-04-2026", "2026-04-20", year=2025),
    rule("Киа", "Carnival", "Limousine+",1_884_350, "Kia SD DC 20-04-2026", "2026-04-20", year=2025),

    # --- Бюллетень 15-05-2026 ---
    rule("Киа", "Seltos",   "Comfort",     499_600, "Kia SD DC 15-05-2026", "2026-05-15", year=2026, drive="2WD", engine=1999),
    rule("Киа", "Seltos",   "Luxe",        523_600, "Kia SD DC 15-05-2026", "2026-05-15", year=2026, drive="2WD", engine=1999),
    rule("Киа", "Seltos",   "Luxe",        547_600, "Kia SD DC 15-05-2026", "2026-05-15", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Seltos",   "Tulpar",      559_600, "Kia SD DC 15-05-2026", "2026-05-15", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Seltos",   "Style",       559_600, "Kia SD DC 15-05-2026", "2026-05-15", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Soluto",   "Comfort",     285_000, "Kia SD DC 15-05-2026", "2026-05-15", year=2025, drive="2WD"),
    rule("Киа", "Soul",     "Comfort",     439_600, "Kia SD DC 15-05-2026", "2026-05-15", year=2025, drive="2WD", engine=1591),
    rule("Киа", "Sportage", "FIFA",        834_500, "Kia SD DC 15-05-2026", "2026-05-15", year=2026, drive="4WD", engine=1999),
    rule("Киа", "Sportage", "FIFA",        869_500, "Kia SD DC 15-05-2026", "2026-05-15", year=2026, drive="4WD", engine=2497),

    # ════════════════════════════════════════════════════════════════════════
    # ШЕВИ  (Chevrolet)
    # ════════════════════════════════════════════════════════════════════════

    # --- Бюллетень до 30-04-2026 ---
    rule("Шеви", "Cobalt",      "Optimum AT",       200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),
    rule("Шеви", "Cobalt",      "Elegant AT",       200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),
    rule("Шеви", "Cobalt",      "Optimum AT",       250_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2026),
    rule("Шеви", "Cobalt",      "Elegant AT",       250_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2026),
    rule("Шеви", "Onix",        "1LT MT",           200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2024),
    rule("Шеви", "Onix",        "3LT MT",           200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2024),
    rule("Шеви", "Onix",        "4LT AT (LTZ)",     200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2024),
    rule("Шеви", "Onix",        "Premier 1",        200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2024),
    rule("Шеви", "Onix",        "Premier 2",        200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2024),
    rule("Шеви", "Onix",        "4LT AT (LTZ)",     300_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),
    rule("Шеви", "Onix",        "Premier 1",        300_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),
    rule("Шеви", "Onix",        "Premier 2",        300_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),
    rule("Шеви", "Onix",        "4LT AT (LTZ)",     350_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2026),
    rule("Шеви", "Onix",        "Premier 1",        350_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2026),
    rule("Шеви", "Onix",        "Premier 2",        350_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2026),
    rule("Шеви", "Tracker MY22","1LT AT",           400_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),
    rule("Шеви", "Tracker MY22","Premier AT",       400_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),
    rule("Шеви", "Captiva",     "2LT",              400_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),
    rule("Шеви", "Captiva",     "Premier AT",       400_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),
    rule("Шеви", "Traverse MY26","Z71",             800_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2026),
    rule("Шеви", "Tahoe MY25",  "High Country",   1_200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2026),
    rule("Шеви", "Damas",       "VAN",              200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29"),
    rule("Шеви", "Damas",       "DLX",              150_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29"),
    rule("Шеви", "Labo",        "LABO",             200_000, "до 30.04 Chevrolet.jpeg", "2026-01-01", valid_to="2026-04-29", year=2025),

    rule("Шеви", "Cobalt",      "Optimum AT",       200_000, "Шеви.jpeg",           "2026-01-01", year=2025),
    rule("Шеви", "Cobalt",      "Elegant AT",       200_000, "Шеви.jpeg",           "2026-01-01", year=2025),
    rule("Шеви", "Cobalt",      "Optimum AT",       200_000, "Шеви.jpeg",           "2026-01-01", year=2026),
    rule("Шеви", "Cobalt",      "Elegant AT",       200_000, "Шеви.jpeg",           "2026-01-01", year=2026),
    rule("Шеви", "Cobalt",      "Optimum AT",       250_000, "Шеви.jpeg",           "2026-01-01", year=2026),
    rule("Шеви", "Cobalt",      "Elegant AT",       250_000, "Шеви.jpeg",           "2026-01-01", year=2026),
    rule("Шеви", "Onix",        "4LT AT (LTZ)",     200_000, "Шеви.jpeg",           "2026-01-01", year=2024),
    rule("Шеви", "Onix",        "Premier 1",        200_000, "Шеви.jpeg",           "2026-01-01", year=2024),
    rule("Шеви", "Onix",        "Premier 2",        200_000, "Шеви.jpeg",           "2026-01-01", year=2024),
    rule("Шеви", "Onix",        "4LT AT (LTZ)",     300_000, "Шеви.jpeg",           "2026-01-01", year=2025),
    rule("Шеви", "Onix",        "Premier 1",        300_000, "Шеви.jpeg",           "2026-01-01", year=2025),
    rule("Шеви", "Onix",        "Premier 2",        300_000, "Шеви.jpeg",           "2026-01-01", year=2025),
    rule("Шеви", "Onix",        "4LT AT (LTZ)",     350_000, "Шеви.jpeg",           "2026-01-01", year=2026),
    rule("Шеви", "Onix",        "Premier 1",        350_000, "Шеви.jpeg",           "2026-01-01", year=2026),
    rule("Шеви", "Onix",        "Premier 2",        350_000, "Шеви.jpeg",           "2026-01-01", year=2026),
    rule("Шеви", "Tracker MY22","1LT AT",           400_000, "Шеви.jpeg",           "2026-01-01", year=2025),
    rule("Шеви", "Tracker MY22","Premier AT",       400_000, "Шеви.jpeg",           "2026-01-01", year=2025),
    rule("Шеви", "Tahoe MY25",  "High Country",   1_200_000, "Шеви.jpeg",           "2026-01-01", year=2026),
    rule("Шеви", "Labo",        "LABO",             200_000, "Шеви.jpeg",           "2026-01-01", year=2025),

    # --- Бюллетень 30-04-2026 ---
    rule("Шеви", "Cobalt",      "Optimum AT",       300_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2025),
    rule("Шеви", "Cobalt",      "Elegant AT",       300_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2025),
    rule("Шеви", "Cobalt",      "Optimum AT",       300_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2026),
    rule("Шеви", "Cobalt",      "Elegant AT",       300_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2026),
    rule("Шеви", "Onix",        "4LT AT (LTZ)",     300_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2025),
    rule("Шеви", "Onix",        "Premier 1",        300_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2025),
    rule("Шеви", "Onix",        "Premier 2",        300_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2025),
    rule("Шеви", "Captiva",     "2LT",              400_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2025),
    rule("Шеви", "Captiva",     "Premier AT",       400_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2025),
    rule("Шеви", "Traverse MY26","Z71",             800_000, "30.04 Шеви.jpeg",     "2026-04-30", year=2026),

    # ════════════════════════════════════════════════════════════════════════
    # ДЖАК  (JAC)
    # ════════════════════════════════════════════════════════════════════════

    # --- Бюллетень 20-03-2026 (Джак 5-ДП) ---
    rule("Джак", "S3 Pro",  "Intelligent 1.6VVT CVT",          267_600, "Джак 5-ДП 20.03.2026", "2026-03-20", year=2025),
    rule("Джак", "J7",      "Luxury 1.5T CVT NEW",              279_600, "Джак 5-ДП 20.03.2026", "2026-03-20", year=2024),
    rule("Джак", "J7",      "Comfort Plus 1.5T",                291_600, "Джак 5-ДП 20.03.2026", "2026-03-20", year=2025),
    rule("Джак", "J7",      "Luxury 1.5T CVT NEW JL",           315_600, "Джак 5-ДП 20.03.2026", "2026-03-20", year=2025),
    rule("Джак", "J7 PLUS", "Flagship 1.5TGDI 6DCT",           407_600, "Джак 5-ДП 20.03.2026", "2026-03-20", year=2025),
    rule("Джак", "J7 PLUS", "Flagship 1.5TGDI 6DCT BR",        415_600, "Джак 5-ДП 20.03.2026", "2026-03-20", year=2025),
    rule("Джак", "JS4",     "Luxury 1.5T CVT",                  279_600, "Джак 5-ДП 20.03.2026", "2026-03-20", year=2024),
    rule("Джак", "JS4",     "Intelligent 1.5T CVT BR",          367_600, "Джак 5-ДП 20.03.2026", "2026-03-20", year=2025),
    rule("Джак", "JS8",     "Luxury 1.5TGDI 7DCT 7-seats JL",  475_600, "Джак 5-ДП 20.03.2026", "2026-03-20", year=2025),

    # --- Бюллетень 23-04-2026 (Джак 8-ДП) ---
    rule("Джак", "S3 Pro",  "Intelligent 1.6VVT CVT",          279_600, "Джак 8-ДП 23.04.2026", "2026-04-23", year=2026),
    rule("Джак", "T6",      "Luxury 2.0 Ti 4x4",               887_400, "Джак 8-ДП 23.04.2026", "2026-04-23", year=2025),
    rule("Джак", "T8 Pro",  "Luxury 2.4 MT 4x4",               947_400, "Джак 8-ДП 23.04.2026", "2026-04-23", year=2025),
    rule("Джак", "T8 Pro",  "Luxury 2.4Т MT 4x4",              947_400, "Джак 8-ДП 23.04.2026", "2026-04-23", year=2025),
    rule("Джак", "T9",      "Luxury 2.0T AT 4x4",            1_049_400, "Джак 8-ДП 23.04.2026", "2026-04-23", year=2025),

    # --- Бюллетень 01-05-2026 (Джак 9-ДП) ---
    rule("Джак", "J7 PLUS", "Luxury 1.5TGDI 6DCT",            239_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "J7 PLUS", "Flagship 1.5TGDI 6DCT",          257_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "J7 PLUS", "Flagship 1.5TGDI 6DCT BR",       263_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "J7 PLUS", "Luxury 1.5TGDI 6DCT",            317_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2026),
    rule("Джак", "J7 PLUS", "Flagship 1.5TGDI 6DCT",          335_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2026),
    rule("Джак", "J7 PLUS", "Flagship 1.5TGDI 6DCT BR",       341_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2026),
    rule("Джак", "J7",      "Comfort Plus 1.5T",               194_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "J7",      "Luxury 1.5T CVT NEW",             203_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "J7",      "Luxury 1.5T CVT NEW JL",          209_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "J7",      "Comfort Plus 1.5T",               245_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2026),
    rule("Джак", "J7",      "Luxury 1.5T CVT NEW",             254_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2026),
    rule("Джак", "J7",      "Luxury 1.5T CVT NEW JL",          260_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2026),
    rule("Джак", "S3 Pro",  "Intelligent 1.6VVT CVT",          179_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "S3 Pro",  "Intelligent 1.6VVT CVT",          209_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2026),
    rule("Джак", "JS4",     "Luxury 1.5T MT/CVT",              200_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "JS4",     "Intelligent 1.5T CVT",            218_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "JS4",     "Intelligent 1.5T CVT BR",         224_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2025),
    rule("Джак", "JS4",     "Luxury 1.5T MT/CVT",              245_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2026),
    rule("Джак", "JS4",     "Luxury 1.5T CVT",                 260_700, "Джак 9-ДП 01.05.2026", "2026-05-01", year=2026),

    # --- Бюллетень 06-05-2026 (Джак 11-ДП) ---
    rule("Джак", "S3 Pro",  "Intelligent 1.6VVT CVT",          164_700, "Джак 11-ДП 06.05.2026", "2026-05-06", year=2025),
    rule("Джак", "S3 Pro",  "Intelligent 1.6VVT CVT",          179_700, "Джак 11-ДП 06.05.2026", "2026-05-06", year=2026),
    rule("Джак", "JS8",     "Intelligent 1.5TGDI 7DCT 6-seats",514_500, "Джак 11-ДП 06.05.2026", "2026-05-06", year=2025),
    rule("Джак", "JS8",     "Luxury 1.5TGDI 7DCT 7-seats JL",  509_500, "Джак 11-ДП 06.05.2026", "2026-05-06", year=2025),
    rule("Джак", "T6",      "Luxury 2.0 Ti 4x4",               827_400, "Джак 11-ДП 06.05.2026", "2026-05-06", year=2025),
    rule("Джак", "T9",      "Luxury 2.0T AT 4x4",            1_019_400, "Джак 11-ДП 06.05.2026", "2026-05-06", year=2025),

    # ════════════════════════════════════════════════════════════════════════
    # ДЖЕТУР  (Jetour)
    # ════════════════════════════════════════════════════════════════════════

    # --- Бюллетень ~03-2026 (Приказ РПЦ 3-ОД) ---
    rule("Джетур", "X70",    "COMFORT",       351_600, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "X70FL",  "PRESTIGE",      359_600, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "X70FL",  "PRESTIGE",      474_500, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2026),
    rule("Джетур", "X70FL",  "PREMIUM",       398_000, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "X70FL",  "PREMIUM",       399_600, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "X70FL",  "PREMIUM",       524_500, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2026),
    rule("Джетур", "X70Plus","PREMIUM PRO",   649_500, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "X90Plus","1.6T PREMIUM",  459_600, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "X90Plus","2.0T PREMIUM",  519_600, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "Dashing","PRESTIGE",      359_600, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "Dashing","PREMIUM",       399_600, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "T1",     "ADVENTURE",     869_400, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "T1",     "ADVENTURE",     899_400, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2026),
    rule("Джетур", "T1",     "EXPEDITION",    959_400, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2026),
    rule("Джетур", "T1",     "EXPEDITION",    989_400, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2026),
    rule("Джетур", "T2",     "ADVENTURE",     959_400, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "T2",     "EXPEDITION",  1_049_400, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2025),
    rule("Джетур", "T2",     "EXPEDITION",  1_073_400, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2026),
    rule("Джетур", "T2",     "EXPEDITION",  1_079_400, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2026),
    rule("Джетур", "X50",    "PRESTIGE",      319_600, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2026),
    rule("Джетур", "X50",    "PREMIUM",       359_600, "Приказ РПЦ 3-ОД Джетур",   "2026-03-01", year=2026),

    # --- Бюллетень 17-04-2026 (Приказ РРЦ 32-ОД) ---
    rule("Джетур", "X50",    "PRESTIGE",      259_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "X50",    "PRESTIGE",      279_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2026),
    rule("Джетур", "X50",    "PREMIUM",       299_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "X50",    "PREMIUM",       319_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2026),
    rule("Джетур", "X70",    "PREMIUM",       319_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "X70",    "PREMIUM",       339_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2026),
    rule("Джетур", "X70FL",  "PRESTIGE",      299_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "X70FL",  "PRESTIGE",      399_500, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2026),
    rule("Джетур", "X70FL",  "PREMIUM",       339_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "X70FL",  "PREMIUM",       449_500, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2026),
    rule("Джетур", "X70Plus","PREMIUM PRO",   479_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "X70Plus","PREMIUM PRO",   624_500, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2026),
    rule("Джетур", "X90Plus","1.6T PREMIUM",  479_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "X90Plus","2.0T PREMIUM",  539_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "Dashing","PRESTIGE",      379_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "Dashing","PREMIUM",       419_600, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "T1",     "EXPEDITION",    774_500, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "T1",     "ADVENTURE",     899_400, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2026),
    rule("Джетур", "T1",     "EXPEDITION",    959_400, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2026),
    rule("Джетур", "T2",     "ADVENTURE",     724_500, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "T2",     "EXPEDITION",    774_500, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2025),
    rule("Джетур", "T2",     "EXPEDITION",    959_400, "Приказ РРЦ 32-ОД 17.04",   "2026-04-17", year=2026),
]

# Источники, которых нет в локальной папке бюллетеней.
# Держим правила в базе для истории, но не используем в сверке,
# чтобы не получать ссылки на "невидимые" бюллетени.
DISABLED_SOURCES: set[str] = {
    "Kia SD DC 17-03-2026",
    "Kia SD DC 03-04-2026",
}


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().upper()
    text = text.replace("Ё", "Е").replace("—", "-")
    text = text.replace("\xa0", " ").replace("₽", "").replace("�", "")
    text = text.replace("FIFIA", "FIFA")
    return " ".join(text.split())


def normalize_int(value: Any) -> int | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return None


def find_active_rewards(
    brand: str,
    model: str,
    trim: str,
    year: int | None,
    drive: str,
    engine: int | None,
    sale_date: date,
) -> tuple[list[int], str, date | None]:
    """
    Возвращает (список разрешённых вознаграждений, источник, дата бюллетеня)
    для самого актуального бюллетеня на дату продажи.
    """
    nm = normalize_text

    def by_level(level: int) -> list[BulletinRule]:
        # 1: strict (model+trim+year+drive+engine)
        # 2: ignore drive/engine
        # 3: ignore trim
        # 4: model+year only (most tolerant fallback)
        return [
            r for r in RULES
            if r.brand == brand
            and r.source not in DISABLED_SOURCES
            and nm(r.model) == nm(model)
            and r.valid_from <= sale_date
                and (r.valid_to is None or sale_date <= r.valid_to)
            and (
                (level <= 2 and nm(r.trim) == nm(trim))
                or (level >= 3)
            )
            and (
                (r.year is None or r.year == year)
                if year is not None else True
            )
            and (
                level == 1 and (r.drive is None or nm(r.drive) == nm(drive)) and (r.engine is None or r.engine == engine)
                or level == 2
                or level == 3 and (r.drive is None or nm(r.drive) == nm(drive) or nm(drive) == "")
                or level == 4
            )
        ]

    fuzzy_labels = {
        1: "точное совпадение",
        2: "ослаблено: без учёта привода/двигателя",
        3: "ослаблено: без учёта комплектации",
        4: "ослаблено: только модель и год",
    }

    for level in (1, 2, 3, 4):
        matching = by_level(level)
        if not matching:
            continue

        max_date = max(r.valid_from for r in matching)
        active = [r for r in matching if r.valid_from == max_date]
        rewards = sorted({r.reward for r in active})
        sources = "; ".join(sorted({r.source for r in active}))
        if level == 1:
            return rewards, f"{sources} (точное совпадение)", max_date
        return rewards, f"{sources} ({fuzzy_labels[level]})", max_date

    return [], "", None
