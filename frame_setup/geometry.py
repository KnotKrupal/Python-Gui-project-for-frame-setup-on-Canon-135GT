from __future__ import annotations

from dataclasses import dataclass
from typing import List

import fitz

from .models import JobParameters
from .utils import inch_to_pt, mm_to_pt


@dataclass
class FramePlacement:
    cluster_index: int
    cluster_row: int
    row: int  # 0 bottom, 1 top
    column: int
    x: float
    y: float
    width: float
    height: float
    active: bool = True

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
    row_gap: float
    placements: List[FramePlacement]
    frames_requested: int

    @property
    def cluster_width(self) -> float:
        return self.glass_width * 2

    @property
    def cluster_height(self) -> float:
        return self.glass_height * 2


@dataclass(frozen=True)
class MatteGeometry:
    bottom_margin_mm: float
    top_margin_mm: float
    side_margin_mm: float
    opening_width_mm: float
    opening_height_mm: float
    visible_band_mm: float


def calculate_matte_geometry(
    glass_width_mm: float,
    glass_height_mm: float,
    indent_mm: float,
    matte_total_mm: float,
) -> MatteGeometry:
    glass_width_mm = max(glass_width_mm, 0.0)
    glass_height_mm = max(glass_height_mm, 0.0)
    indent_mm = max(indent_mm, 0.0)
    matte_total_mm = max(matte_total_mm, 0.0)

    bottom_margin_mm = min(indent_mm, glass_height_mm / 2 if glass_height_mm else 0.0)
    max_opening_height = max(glass_height_mm - bottom_margin_mm, 0.0)
    opening_height_mm = min(matte_total_mm, max_opening_height)
    top_margin_mm = max(glass_height_mm - bottom_margin_mm - opening_height_mm, 0.0)

    side_margin_mm = min(indent_mm, glass_width_mm / 2 if glass_width_mm else 0.0)
    opening_width_mm = max(glass_width_mm - 2 * side_margin_mm, 0.0)

    visible_band_mm = max(matte_total_mm - bottom_margin_mm, 0.0)

    return MatteGeometry(
        bottom_margin_mm=bottom_margin_mm,
        top_margin_mm=top_margin_mm,
        side_margin_mm=side_margin_mm,
        opening_width_mm=opening_width_mm,
        opening_height_mm=opening_height_mm,
        visible_band_mm=visible_band_mm,
    )


def calculate_capacity_from_values(
    glass_width_mm: float,
    bed_width_in: float,
    cluster_gap_mm: float,
) -> int:
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


def calculate_vertical_capacity(
    glass_height_mm: float,
    bed_height_in: float,
    row_gap_mm: float,
) -> int:
    cluster_height_mm = glass_height_mm * 2
    bed_height_mm = bed_height_in * 25.4
    if cluster_height_mm <= 0:
        return 0
    if row_gap_mm < 0:
        return 0
    numerator = bed_height_mm + row_gap_mm
    denominator = cluster_height_mm + row_gap_mm
    capacity = int(numerator // denominator)
    return max(capacity, 0)


def build_layout(params: JobParameters) -> Layout:
    page_width = inch_to_pt(params.bed_width_in)
    page_height = inch_to_pt(params.bed_height_in)
    glass_width = mm_to_pt(params.glass_width_mm)
    glass_height = mm_to_pt(params.glass_height_mm)
    cluster_gap = mm_to_pt(params.cluster_gap_mm)
    row_gap = mm_to_pt(params.row_gap_mm)

    placements: List[FramePlacement] = []
    frames_requested = max(params.frame_quantity, 0)
    frame_counter = 0

    for cluster_row in range(params.cluster_rows):
        base_y = cluster_row * (2 * glass_height + row_gap)
        for cluster_index in range(params.cluster_count):
            base_x = cluster_index * (2 * glass_width + cluster_gap)
            for row in range(2):
                y = base_y + row * glass_height
                for column in range(2):
                    x = base_x + column * glass_width
                    frame_counter += 1
                    active = frame_counter <= frames_requested or frames_requested == 0
                    placements.append(
                        FramePlacement(
                            cluster_index=cluster_index,
                            cluster_row=cluster_row,
                            row=row,
                            column=column,
                            x=x,
                            y=y,
                            width=glass_width,
                            height=glass_height,
                            active=active,
                        )
                    )

    return Layout(
        page_width=page_width,
        page_height=page_height,
        glass_width=glass_width,
        glass_height=glass_height,
        cluster_gap=cluster_gap,
        row_gap=row_gap,
        placements=placements,
        frames_requested=frames_requested,
    )
