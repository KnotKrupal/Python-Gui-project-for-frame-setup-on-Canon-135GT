from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, List

import fitz

from .models import JobParameters
from .utils import inch_to_pt, mm_to_pt


@dataclass
class FramePlacement:
    cluster_index: int
    row: int  # 0 bottom, 1 top
    column: int
    x: float
    y: float
    width: float
    height: float

    @property
    def rect(self) -> fitz.Rect:
        return fitz.Rect(self.x, self.y, self.x + self.width, self.y + self.height)


@dataclass
class Layout:
    page_width: float
    page_height: float
    glass_width: float
    glass_height: float
    cluster_gap: float
    placements: List[FramePlacement]

    @property
    def cluster_width(self) -> float:
        return self.glass_width * 2

    @property
    def cluster_height(self) -> float:
        return self.glass_height * 2


def calculate_cluster_capacity(params: JobParameters) -> int:
    cluster_width_mm = params.glass_width_mm * 2
    bed_width_mm = params.bed_width_in * 25.4
    if cluster_width_mm <= 0:
        return 0
    if params.cluster_gap_mm < 0:
        return 0
    numerator = bed_width_mm + params.cluster_gap_mm
    denominator = cluster_width_mm + params.cluster_gap_mm
    capacity = int(numerator // denominator)
    return max(capacity, 0)


def build_layout(params: JobParameters) -> Layout:
    page_width = inch_to_pt(params.bed_width_in)
    glass_width = mm_to_pt(params.glass_width_mm)
    glass_height = mm_to_pt(params.glass_height_mm)
    cluster_gap = mm_to_pt(params.cluster_gap_mm)
    placements: List[FramePlacement] = []
    for cluster_index in range(params.cluster_count):
        base_x = cluster_index * (2 * glass_width + cluster_gap)
        for row in range(2):  # 0 bottom, 1 top
            y = row * glass_height
            for column in range(2):
                x = base_x + column * glass_width
                placements.append(
                    FramePlacement(
                        cluster_index=cluster_index,
                        row=row,
                        column=column,
                        x=x,
                        y=y,
                        width=glass_width,
                        height=glass_height,
                    )
                )
    page_height = 2 * glass_height
    return Layout(
        page_width=page_width,
        page_height=page_height,
        glass_width=glass_width,
        glass_height=glass_height,
        cluster_gap=cluster_gap,
        placements=placements,
    )


def calculate_capacity_from_values(glass_width_mm: float, bed_width_in: float, cluster_gap_mm: float) -> int:
    cluster_width_mm = glass_width_mm * 2
    bed_width_mm = bed_width_in * 25.4
    if cluster_width_mm <= 0:
        return 0
    if cluster_gap_mm < 0:
        return 0
    numerator = bed_width_mm + cluster_gap_mm
    denominator = cluster_width_mm + cluster_gap_mm
    capacity = int(numerator // denominator)
    return max(capacity, 0)

