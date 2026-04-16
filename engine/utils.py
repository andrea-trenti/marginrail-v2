from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import pandas as pd


def slugify(value: str) -> str:
    value = re.sub(r"[^a-zA-Z0-9]+", "_", str(value).strip())
    return value.strip("_").lower()


def safe_read_excel(file_path: Path, sheet_name: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if df is None:
            return pd.DataFrame()
        return df
    except ValueError:
        return pd.DataFrame()


def ensure_columns(df: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    for col in columns:
        if col not in df.columns:
            df[col] = pd.NA
    return df


def to_numeric(df: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def to_datetime(df: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def nz(value: object, default: float = 0.0) -> float:
    if pd.isna(value):
        return default
    try:
        return float(value)
    except Exception:
        return default


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def build_case_id(idx: int) -> str:
    return f"CASE-{idx:06d}"


def build_natural_case_key(row: pd.Series, rule_code: str) -> str:
    parts = [
        clean_text(row.get("NumeroOrdine")) or "NAORD",
        str(int(nz(row.get("RigaOrdine"), 0))),
        clean_text(row.get("TipoDocumento")) or "NADOC",
        clean_text(row.get("ClienteID")) or "NACLI",
        clean_text(row.get("ProdottoID")) or "NAPRD",
        rule_code,
    ]
    return "|".join(parts)
