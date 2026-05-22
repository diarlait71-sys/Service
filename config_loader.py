from functools import lru_cache
from pathlib import Path

import yaml


DEFAULT_CONFIG = {
    "files": {
        "static": [
            "Дц,должность,репер.xlsx",
            "Доля показателей для Руководителя отдела сервиса.xlsx",
            "план по остальным показателям.xlsx",
        ],
        "monthly": [
            "Факт Услуги.xlsx",
            "Факт Запасные части.xlsx",
            "Факт маржанальность.xlsx",
            "Факт по остальным показателям.xlsx",
        ],
        "logic_default": "Формула расчета.xlsx",
    },
    "thresholds": {
        "marginality_min": 0.30,
        "negliquidity_max": 0.07,
        "fact_cap_ratio": 1.5,
    },
    "cache": {
        "static_ttl_seconds": 3600,
        "facts_ttl_seconds": 1800,
    },
    "app": {
        "host_port": 8506,
        "default_year": "2026",
    },
    "logging": {
        "file": "logs/bonus_calc.log",
    },
}


def _deep_merge(base: dict, override: dict) -> dict:
    result = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


@lru_cache(maxsize=1)
def load_app_config() -> dict:
    config_path = Path(__file__).parent / "config" / "settings.yaml"
    if not config_path.exists():
        return DEFAULT_CONFIG

    loaded = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    return _deep_merge(DEFAULT_CONFIG, loaded)
