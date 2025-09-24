from __future__ import annotations

from dataclasses import dataclass
from typing import Dict


@dataclass(frozen=True)
class OutlineColor:
    name: str
    hex_value: str


PRESET_COLORS: Dict[str, OutlineColor] = {
    "SRG Scarlet": OutlineColor("SRG Scarlet", "#D22630"),
    "SRG Charcoal": OutlineColor("SRG Charcoal", "#2D2926"),
    "SRG Silver": OutlineColor("SRG Silver", "#A7A8AA"),
    "Pure Black": OutlineColor("Pure Black", "#000000"),
    "Bright White": OutlineColor("Bright White", "#FFFFFF"),
}

DEFAULT_COLOR_NAME = "SRG Scarlet"

