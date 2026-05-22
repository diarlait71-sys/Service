"""Локальное хранилище профилей SQL-подключений."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Any

PROFILES_FILE = Path("sql_profiles.json")


def load_sql_profiles() -> Dict[str, Dict[str, Any]]:
    """Загружает профили SQL из json-файла."""
    if not PROFILES_FILE.exists():
        return {}

    try:
        data = json.loads(PROFILES_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
        return {}
    except Exception:
        return {}


def save_sql_profiles(profiles: Dict[str, Dict[str, Any]]) -> None:
    """Сохраняет все профили SQL в json-файл."""
    PROFILES_FILE.write_text(
        json.dumps(profiles, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def upsert_sql_profile(profile_name: str, profile: Dict[str, Any]) -> None:
    """Добавляет новый профиль или обновляет существующий."""
    profiles = load_sql_profiles()
    profiles[profile_name] = profile
    save_sql_profiles(profiles)


def delete_sql_profile(profile_name: str) -> None:
    """Удаляет профиль по имени, если существует."""
    profiles = load_sql_profiles()
    if profile_name in profiles:
        del profiles[profile_name]
        save_sql_profiles(profiles)
