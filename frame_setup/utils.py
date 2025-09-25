from __future__ import annotations

import math
import os
import re
from typing import Iterable, Tuple


MM_PER_INCH = 25.4
PT_PER_INCH = 72.0
PT_PER_MM = PT_PER_INCH / MM_PER_INCH


def mm_to_pt(value_mm: float) -> float:
    return value_mm * PT_PER_MM


def inch_to_pt(value_in: float) -> float:
    return value_in * PT_PER_INCH


def pt_to_mm(value_pt: float) -> float:
    return value_pt / PT_PER_MM


def hex_to_rgb_floats(hex_value: str) -> tuple[float, float, float]:
    hex_value = hex_value.lstrip("#")
    if len(hex_value) == 3:
        hex_value = "".join(ch * 2 for ch in hex_value)
    if len(hex_value) != 6:
        raise ValueError("Hex color must be in #RRGGBB or #RGB format")
    r = int(hex_value[0:2], 16) / 255.0
    g = int(hex_value[2:4], 16) / 255.0
    b = int(hex_value[4:6], 16) / 255.0
    return r, g, b


def slugify(value: str) -> str:
    value = value.strip().lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    value = re.sub(r"_+", "_", value)
    return value.strip("_") or "job"


def ensure_directory(path: str) -> None:
    os.makedirs(path, exist_ok=True)


class ValidationError(Exception):
    """Raised when a numeric conversion fails."""

    def __init__(self, field: str, message: str) -> None:
        super().__init__(f"{field}: {message}")
        self.field = field
        self.message = message


def parse_positive_float(value: str, field_name: str) -> float:
    try:
        result = float(value)
    except ValueError as exc:
        raise ValidationError(field_name, "must be a number") from exc
    if result <= 0:
        raise ValidationError(field_name, "must be greater than zero")
    return result


def parse_non_negative_float(value: str, field_name: str) -> float:
    try:
        result = float(value)
    except ValueError as exc:
        raise ValidationError(field_name, "must be a number") from exc
    if result < 0:
        raise ValidationError(field_name, "must be zero or greater")
    return result


