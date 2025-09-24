from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Callable, List, Optional, Tuple

import fitz
from pypdf import PdfReader, PdfWriter

from .geometry import (
    Layout,
    MatteGeometry,
    build_layout,
    calculate_matte_geometry,
)
from .models import JobParameters, LogoAsset
from .utils import mm_to_pt

ProgressCallback = Optional[Callable[[int, str], None]]
Matrix = Tuple[float, float, float, float, float, float]


@dataclass(frozen=True)
class ExportResult:
    outline_path: str
    artwork_path: str


@dataclass(frozen=True)
class LogoPlacement:
    rect: fitz.Rect
    rotate_bottom: bool


def _multiply(m1: Matrix, m2: Matrix) -> Matrix:
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (
        a1 * a2 + b1 * c2,
        a1 * b2 + b1 * d2,
        c1 * a2 + d1 * c2,
        c1 * b2 + d1 * d2,
        e1 * a2 + f1 * c2 + e2,
        e1 * b2 + f1 * d2 + f2,
    )


def _translation(tx: float, ty: float) -> Matrix:
    return (1.0, 0.0, 0.0, 1.0, tx, ty)


def _scale(sx: float, sy: float) -> Matrix:
    return (sx, 0.0, 0.0, sy, 0.0, 0.0)


def _build_transformation(rect: fitz.Rect, rotate_bottom: bool, flip_horizontal: bool) -> Matrix:
    width = rect.width
    height = rect.height
    if width <= 0 or height <= 0:
        raise ValueError("Logo rectangle must have positive size")
    left = rect.x0
    bottom = rect.y0
    center_x = left + width / 2.0
    center_y = bottom + height / 2.0

    matrix: Matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    matrix = _multiply(matrix, _translation(-width / 2.0, -height / 2.0))
    if flip_horizontal:
        matrix = _multiply(matrix, _scale(-1.0, 1.0))
    if rotate_bottom:
        matrix = _multiply(matrix, _scale(-1.0, -1.0))
    matrix = _multiply(matrix, _translation(center_x, center_y))
    return matrix


class PDFBuilder:
    def __init__(self, params: JobParameters, logo: LogoAsset) -> None:
        self.params = params
        self.logo = logo
        self.layout: Layout = build_layout(params)
        self._matte: MatteGeometry = calculate_matte_geometry(
            params.glass_width_mm,
            params.glass_height_mm,
            params.indent_mm,
            params.matte_total_height_mm,
        )

    def _open_logo(self) -> fitz.Document:
        return fitz.open(stream=self.logo.pdf_bytes, filetype="pdf")

    def generate_pdfs(self, progress: ProgressCallback = None) -> ExportResult:
        if progress:
            progress(5, "Preparing outlines")
        self._generate_outlines()
        placements = self._compute_logo_placements()
        if progress:
            progress(50, "Preparing artwork")
        try:
            self._export_artwork_vector(placements)
        except Exception:
            if not self.params.allow_raster_fallback:
                raise
            if progress:
                progress(75, "Vector placement failed – rasterising logo")
            self._export_artwork_raster(placements)
        if progress:
            progress(100, "Finished")
        return ExportResult(self.params.outline_path, self.params.artwork_path)

    def _generate_outlines(self) -> None:
        doc = fitz.open()
        try:
            page = doc.new_page(width=self.layout.page_width, height=self.layout.page_height)
            for placement in self.layout.placements:
                page.draw_rect(
                    placement.rect,
                    color=self.params.outline_color,
                    width=mm_to_pt(self.params.outline_thickness_mm),
                    fill=None,
                )
            doc.save(self.params.outline_path, deflate=True, garbage=4)
        finally:
            doc.close()

    def _compute_logo_placements(self) -> List[LogoPlacement]:
        visible_band_pt = mm_to_pt(
            min(self._matte.visible_band_mm, self.params.glass_height_mm)
        )
        indent_pt = mm_to_pt(self.params.indent_mm)
        logo_width = self.logo.width_pt
        logo_height = self.logo.height_pt
        if visible_band_pt <= 0 or logo_width <= 0 or logo_height <= 0:
            return []

        placements: List[LogoPlacement] = []
        for placement in self.layout.placements:
            if placement.row == 1:  # top row
                center_y = placement.y + visible_band_pt / 2.0
            else:
                center_y = placement.y + placement.height - visible_band_pt / 2.0
            center_x = placement.x + placement.width / 2.0
            min_center = placement.y + indent_pt + logo_height / 2.0
            max_center = placement.y + placement.height - indent_pt - logo_height / 2.0
            if min_center > max_center:
                min_center = placement.y + logo_height / 2.0
                max_center = placement.y + placement.height - logo_height / 2.0
            center_y = max(min(center_y, max_center), min_center)
            rect = fitz.Rect(
                center_x - logo_width / 2.0,
                center_y - logo_height / 2.0,
                center_x + logo_width / 2.0,
                center_y + logo_height / 2.0,
            )
            rotate_bottom = self.params.rotate_bottom and placement.row == 0
            placements.append(LogoPlacement(rect=rect, rotate_bottom=rotate_bottom))
        return placements

    def _export_artwork_vector(self, placements: List[LogoPlacement]) -> None:
        writer = PdfWriter()
        page = writer.add_blank_page(
            width=self.layout.page_width,
            height=self.layout.page_height,
        )
        if placements:
            logo_reader = PdfReader(BytesIO(self.logo.pdf_bytes))
            base_logo_page = logo_reader.pages[0]
            for placement in placements:
                matrix = _build_transformation(
                    placement.rect,
                    rotate_bottom=placement.rotate_bottom,
                    flip_horizontal=self.params.flip_in_app,
                )
                page.merge_transformed_page(base_logo_page.copy(), matrix, expand=False)
        with open(self.params.artwork_path, "wb") as handle:
            writer.write(handle)

    def _export_artwork_raster(self, placements: List[LogoPlacement]) -> None:
        doc = fitz.open()
        try:
            page = doc.new_page(width=self.layout.page_width, height=self.layout.page_height)
            if placements:
                logo_doc = self._open_logo()
                try:
                    logo_page = logo_doc[0]
                    base_matrix = fitz.Matrix(600 / 72.0, 0.0, 0.0, 600 / 72.0)
                    for placement in placements:
                        transform = base_matrix
                        if self.params.flip_in_app:
                            transform = fitz.Matrix(
                                -1.0,
                                0.0,
                                0.0,
                                1.0,
                                logo_page.rect.width,
                                0.0,
                            ) * transform
                        if placement.rotate_bottom:
                            transform = fitz.Matrix(
                                -1.0,
                                0.0,
                                0.0,
                                -1.0,
                                logo_page.rect.width,
                                logo_page.rect.height,
                            ) * transform
                        pix = logo_page.get_pixmap(matrix=transform)
                        page.insert_image(placement.rect, pixmap=pix)
                finally:
                    logo_doc.close()
            doc.save(self.params.artwork_path, deflate=True, garbage=4)
        finally:
            doc.close()
