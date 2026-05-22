"""Утилиты для подключения к SQL Server и загрузки данных в pandas."""

from __future__ import annotations

from typing import Optional, Tuple
from urllib.parse import quote_plus

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine


def build_mssql_connection_url(
    server: str,
    database: str,
    driver: str = "ODBC Driver 18 for SQL Server",
    username: Optional[str] = None,
    password: Optional[str] = None,
    use_windows_auth: bool = True,
    encrypt: str = "yes",
    trust_server_certificate: bool = True,
    timeout: int = 30,
) -> str:
    """Собирает URL подключения SQLAlchemy для SQL Server через pyodbc."""
    if not server or not database:
        raise ValueError("Укажите server и database")

    odbc_parts = {
        "DRIVER": "{" + driver + "}",
        "SERVER": server,
        "DATABASE": database,
        "Encrypt": encrypt,
        "TrustServerCertificate": "yes" if trust_server_certificate else "no",
        "Connection Timeout": str(timeout),
    }

    if use_windows_auth:
        odbc_parts["Trusted_Connection"] = "yes"
    else:
        if not username or password is None:
            raise ValueError("Для SQL-аутентификации укажите username и password")
        odbc_parts["UID"] = username
        odbc_parts["PWD"] = password

    odbc_str = ";".join(f"{k}={v}" for k, v in odbc_parts.items())
    return f"mssql+pyodbc:///?odbc_connect={quote_plus(odbc_str)}"


def create_sql_engine(connection_url: str) -> Engine:
    """Создает SQLAlchemy engine с защитой от разрыва соединения."""
    return create_engine(connection_url, pool_pre_ping=True)


def test_sql_connection(engine: Engine) -> Tuple[bool, str]:
    """Проверяет подключение простым запросом SELECT 1."""
    try:
        with engine.connect() as connection:
            connection.execute(text("SELECT 1"))
        return True, "Подключение успешно"
    except Exception as exc:  # pragma: no cover
        return False, str(exc)


def query_to_dataframe(engine: Engine, query: str) -> pd.DataFrame:
    """Выполняет SQL-запрос и возвращает DataFrame."""
    if not query or not query.strip():
        raise ValueError("SQL-запрос пуст")

    with engine.connect() as connection:
        return pd.read_sql_query(text(query), connection)
