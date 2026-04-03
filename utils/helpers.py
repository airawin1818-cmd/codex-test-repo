from __future__ import annotations

import os
import re
import sys
from datetime import datetime, date
from pathlib import Path
from typing import Any, Optional

import pandas as pd


def resource_path(relative_path: str) -> str:
    """Return resource path for normal/dev and PyInstaller runtime."""
    if hasattr(sys, "_MEIPASS"):
        return str(Path(getattr(sys, "_MEIPASS")) / relative_path)
    return str(Path(relative_path).resolve())


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\u3000", " ").replace("\t", " ").replace("\n", " ")
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"none", "nan", "nat"}:
        return ""
    text = re.sub(r"\s+", " ", text)
    return text


def parse_amount(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value)
    if not text:
        return None

    text = text.replace(",", "").replace(" ", "")
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]

    try:
        num = float(text)
    except ValueError:
        return None
    return -num if negative else num


def parse_date(value: Any) -> Optional[datetime]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, float) and pd.isna(value):
        return None

    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return None
    if isinstance(ts, pd.Timestamp):
        return ts.to_pydatetime()
    return None


def safe_filename(base_name: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]", "_", base_name)
    return cleaned or "output"


def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path
