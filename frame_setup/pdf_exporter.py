from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

import fitz

from .geometry import Layout, build_layout
from .models import JobParameters, LogoAsset
from .utils import mm_to_pt

ProgressCallback = Optional[Callable[[int, str], None]]


@dataclass
class ExportResult:
    outline_path: str
    artwork_path: str


class PDFBuilder:
    def __init__(self, params: JobParameters, logo: LogoAsset) -> None:
        self.params = params
        self.logo = logo
        self.layout = build_layout(params)

    def _open_logo(self) -> fitz.Document:
        return fitz.open(stream=self.logo.pdf_bytes, filetype="pdf")

    def generate_pdfs(self, progress: ProgressCallback = None) -> ExportResult:
        if progress:
            progress(5, "Preparing outlines")
        self._generate_outlines(progress)
        if progress:
            progress(50, "Preparing artwork")
        self._generate_artwork(progress)
        if progress:
            progress(100, "Finished")
        return ExportResult(self.params.outline_path, self.params.artwork_path)

    def _generate_outlines(self, progress: ProgressCallback) -> None:
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

    def _generate_artwork(self, progress: ProgressCallback) -> None:
        doc = fitz.open()
        try:
            page = doc.new_page(width=self.layout.page_width, height=self.layout.page_height)
            visible_band_pt = mm_to_pt(
                max(self.params.matte_total_height_mm - self.params.indent_mm, 0.0)
            )
            visible_band_pt = min(visible_band_pt, self.layout.glass_height)
            indent_pt = mm_to_pt(self.params.indent_mm)
            logo_doc = self._open_logo()
            try:
                logo_page = logo_doc[0]
                logo_rect = logo_page.rect
                logo_width = logo_rect.width
                logo_height = logo_rect.height
                for placement in self.layout.placements:
                    if visible_band_pt == 0:
                        continue
                    if placement.row == 1:  # top row
                        center_y = placement.y + visible_band_pt / 2
                    else:  # bottom row
                        center_y = placement.y + placement.height - visible_band_pt / 2
                    center_x = placement.x + placement.width / 2
                    min_center = placement.y + indent_pt + logo_height / 2
                    max_center = placement.y + placement.height - indent_pt - logo_height / 2
                    if min_center > max_center:
                        min_center = placement.y + logo_height / 2
                        max_center = placement.y + placement.height - logo_height / 2
                    center_y = max(center_y, min_center)
                    center_y = min(center_y, max_center)
                    dest_rect = fitz.Rect(
                        center_x - logo_width / 2,
                        center_y - logo_height / 2,
                        center_x + logo_width / 2,
                        center_y + logo_height / 2,
                    )
                    rotate = self.params.rotate_bottom and placement.row == 0
                    self._place_logo(page, logo_page, dest_rect, rotate)
            finally:
                logo_doc.close()
            doc.save(self.params.artwork_path, deflate=True, garbage=4)
        finally:
            doc.close()

    def _place_logo(
        self,
        page: fitz.Page,
        logo_page: fitz.Page,
        dest_rect: fitz.Rect,
        rotate_bottom: bool,
    ) -> None:
        matrix = fitz.Matrix(1, 0, 0, 1)
        if self.params.flip_in_app:
            matrix = fitz.Matrix(-1, 0, 0, 1) * matrix
        if rotate_bottom:
            matrix = fitz.Matrix(-1, 0, 0, -1) * matrix
        try:
            page.show_pdf_page(dest_rect, logo_page, 0, matrix=matrix)
        except Exception:
            if not self.params.allow_raster_fallback:
                raise
            pix = logo_page.get_pixmap(matrix=fitz.Matrix(600 / 72, 0, 0, 600 / 72))
            page.insert_image(dest_rect, pixmap=pix)

