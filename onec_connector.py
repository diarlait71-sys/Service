"""
Коннектор к 1С через OData HTTP API.
Требует: объекты опубликованы в 1С через Конфигуратор → Публикация → OData.
"""

from __future__ import annotations

import base64
from typing import Optional
from pathlib import Path

import pandas as pd
import requests
import yaml


CONFIG_PATH = Path(__file__).parent / "config" / "onec_connection.yaml"


def _load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return yaml.safe_load(f)


def _make_headers(username: str, password: str) -> dict:
    """Формирует заголовок Basic Auth с поддержкой кириллических логинов."""
    creds = base64.b64encode(f"{username}:{password}".encode("utf-8")).decode()
    return {"Authorization": f"Basic {creds}", "Accept": "application/json"}


class OneCClient:
    """Клиент для работы с 1С OData API."""

    def __init__(self):
        cfg = _load_config()
        self._base_url = cfg["onec"]["url"].rstrip("/")
        self._odata_url = f"{self._base_url}/odata/standard.odata"
        self._headers = _make_headers(cfg["onec"]["username"], cfg["onec"]["password"])
        self._timeout = cfg["onec"].get("timeout", 30)
        self._entities = cfg.get("odata_entities", {})

    def check_connection(self) -> tuple[bool, str]:
        """Проверяет доступность OData сервиса."""
        try:
            r = requests.get(self._odata_url, headers=self._headers, timeout=self._timeout)
            if r.status_code == 200:
                return True, "Подключение успешно"
            return False, f"HTTP {r.status_code}: {r.text[:200]}"
        except Exception as e:
            return False, str(e)

    def get_available_entities(self) -> list[str]:
        """Возвращает список опубликованных объектов OData."""
        import xml.etree.ElementTree as ET
        r = requests.get(
            f"{self._odata_url}/$metadata",
            headers={**self._headers, "Accept": "application/xml"},
            timeout=self._timeout,
        )
        r.raise_for_status()
        ns = {"edm": "http://schemas.microsoft.com/ado/2009/11/edm"}
        root = ET.fromstring(r.text)
        entities = [el.get("Name") for el in root.iter("{http://schemas.microsoft.com/ado/2009/11/edm}EntityType")]
        return entities

    def _fetch_entity(
        self,
        entity_name: str,
        date_from: str,
        date_to: str,
        select_fields: Optional[list[str]] = None,
        top: int = 50000,
    ) -> pd.DataFrame:
        """Загружает данные из регистра/документа за период."""
        params = {
            "$format": "json",
            "$top": str(top),
            "$filter": f"Period ge datetime'{date_from}T00:00:00' and Period le datetime'{date_to}T23:59:59'",
        }
        if select_fields:
            params["$select"] = ",".join(select_fields)

        url = f"{self._odata_url}/{entity_name}"
        r = requests.get(url, headers=self._headers, params=params, timeout=self._timeout)
        r.raise_for_status()
        data = r.json().get("value", [])
        return pd.DataFrame(data)

    # ── Финансовые данные ────────────────────────────────────────────────────

    def get_sales(self, date_from: str, date_to: str) -> pd.DataFrame:
        """Продажи (выручка и себестоимость)."""
        entity = self._entities.get("sales", "AccumulationRegister_Продажи")
        return self._fetch_entity(entity, date_from, date_to)

    def get_receivables(self, date_from: str, date_to: str) -> pd.DataFrame:
        """Дебиторская задолженность (расчёты с покупателями)."""
        entity = self._entities.get("receivables", "AccumulationRegister_РасчетыСПокупателями")
        return self._fetch_entity(entity, date_from, date_to)

    def get_payables(self, date_from: str, date_to: str) -> pd.DataFrame:
        """Кредиторская задолженность (расчёты с поставщиками)."""
        entity = self._entities.get("payables", "AccumulationRegister_РасчетыСПоставщиками")
        return self._fetch_entity(entity, date_from, date_to)

    def get_inventory(self, date_from: str, date_to: str) -> pd.DataFrame:
        """Запасы (товары на складах)."""
        entity = self._entities.get("inventory", "AccumulationRegister_ТоварыНаСкладах")
        return self._fetch_entity(entity, date_from, date_to)

    def get_fixed_assets(self, date_from: str, date_to: str) -> pd.DataFrame:
        """Основные средства / CAPEX."""
        entity = self._entities.get("fixed_assets", "AccumulationRegister_СтоимостьОС")
        return self._fetch_entity(entity, date_from, date_to)

    def build_ebitda_cash_bridge(self, year: int) -> dict:
        """
        Строит мост EBITDA → Кэш за год.
        Возвращает словарь со всеми компонентами.
        """
        d_from = f"{year}-01-01"
        d_to = f"{year}-12-31"
        d_from_prev = f"{year - 1}-01-01"
        d_to_prev = f"{year - 1}-12-31"

        result = {}

        # Дебиторка: изменение = конец - начало
        try:
            rec_curr = self.get_receivables(d_from, d_to)
            rec_prev = self.get_receivables(d_from_prev, d_to_prev)
            result["delta_receivables"] = (
                rec_curr["СуммаОборот"].sum() - rec_prev["СуммаОборот"].sum()
                if "СуммаОборот" in rec_curr.columns else None
            )
        except Exception as e:
            result["delta_receivables_error"] = str(e)

        # Кредиторка
        try:
            pay_curr = self.get_payables(d_from, d_to)
            pay_prev = self.get_payables(d_from_prev, d_to_prev)
            result["delta_payables"] = (
                pay_curr["СуммаОборот"].sum() - pay_prev["СуммаОборот"].sum()
                if "СуммаОборот" in pay_curr.columns else None
            )
        except Exception as e:
            result["delta_payables_error"] = str(e)

        # Запасы
        try:
            inv_curr = self.get_inventory(d_from, d_to)
            inv_prev = self.get_inventory(d_from_prev, d_to_prev)
            result["delta_inventory"] = (
                inv_curr["СуммаОборот"].sum() - inv_prev["СуммаОборот"].sum()
                if "СуммаОборот" in inv_curr.columns else None
            )
        except Exception as e:
            result["delta_inventory_error"] = str(e)

        return result
