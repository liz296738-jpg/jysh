"""Conservative formatting normalization for offline comparisons."""

from __future__ import annotations

import re
import unicodedata
from typing import Any


def _text(value: Any) -> str:
    return unicodedata.normalize("NFKC", str(value or "")).strip()


def normalize_credit_code(value: Any) -> str:
    """Remove presentation separators and uppercase; never infer a missing code."""
    return re.sub(r"[\s\-_/\\.]+", "", _text(value)).upper()


def normalize_company_name(value: Any) -> str:
    """Normalize spacing and punctuation without deleting legal-name components."""
    text = _text(value)
    text = text.replace("（", "(").replace("）", ")")
    return re.sub(r"\s+", " ", text).strip().casefold()


def normalize_phone(value: Any) -> str:
    """Normalize common domestic phone formatting for equality checks only."""
    phone = re.sub(r"[\s\-()]+", "", _text(value))
    if phone.startswith("+86"):
        phone = phone[3:]
    elif phone.startswith("0086"):
        phone = phone[4:]
    return phone


def normalize_text(value: Any) -> str:
    return re.sub(r"\s+", " ", _text(value)).strip().casefold()
