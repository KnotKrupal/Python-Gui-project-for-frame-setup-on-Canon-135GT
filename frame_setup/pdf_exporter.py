from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

import fitz

from .geometry import Layout, build_layout
from .models import JobParameters, LogoAsset
from .utils import mm_to_pt

ProgressCallback = Optional[Callable[[int, str], None]]


def _show_pdf_page_transformed(
    page: fitz.Page,
    rect: fitz.Rect,
    src: fitz.Document,
    pno: int = 0,
    *,
    keep_proportion: bool = True,
    overlay: bool = True,
    oc: int = 0,
    rotate: int = 0,
    clip: Optional[fitz.Rect] = None,
    flip_horizontal: bool = False,
) -> int:
    """Embed a PDF page with optional rotation and horizontal mirroring."""

    def calc_matrix(
        src_rect: fitz.Rect,
        target_rect: fitz.Rect,
        *,
        keep: bool = True,
        rotation: int = 0,
    ) -> fitz.Matrix:
        src_center = (src_rect.tl + src_rect.br) / 2.0
        target_center = (target_rect.tl + target_rect.br) / 2.0
        matrix = fitz.Matrix(1, 0, 0, 1, -src_center.x, -src_center.y)
        matrix *= fitz.Matrix(rotation)
        transformed_source = src_rect * matrix
        scale_x = target_rect.width / transformed_source.width
        scale_y = target_rect.height / transformed_source.height
        if keep:
            scale_x = scale_y = min(scale_x, scale_y)
        matrix *= fitz.Matrix(scale_x, scale_y)
        matrix *= fitz.Matrix(1, 0, 0, 1, target_center.x, target_center.y)
        return matrix

    fitz.CheckParent(page)
    target_doc = page.parent
    if not target_doc.is_pdf or not src.is_pdf:
        raise ValueError("is no PDF")
    if rect.is_empty or rect.is_infinite:
        raise ValueError("rect must be finite and not empty")

    while pno < 0:
        pno += src.page_count
    source_page = src[pno]
    if source_page.get_contents() == []:
        raise ValueError("nothing to show - source page empty")

    target_rect = rect * ~page.transformation_matrix
    source_rect = source_page.rect if not clip else source_page.rect & clip
    if source_rect.is_empty or source_rect.is_infinite:
        raise ValueError("clip must be finite and not empty")
    source_rect = source_rect * ~source_page.transformation_matrix

    matrix = calc_matrix(
        source_rect,
        target_rect,
        keep=keep_proportion,
        rotation=rotate,
    )
    if flip_horizontal:
        center_x = target_rect.x0 + target_rect.width / 2.0
        flip_matrix = fitz.Matrix(-1, 0, 0, 1, 2 * center_x, 0)
        matrix = flip_matrix * matrix

    existing_objects = [item[1] for item in target_doc.get_page_xobjects(page.number)]
    existing_objects += [item[7] for item in target_doc.get_page_images(page.number)]
    existing_objects += [item[4] for item in target_doc.get_page_fonts(page.number)]

    base_name = "fzFrm"
    suffix = 0
    object_name = f"{base_name}{suffix}"
    while object_name in existing_objects:
        suffix += 1
        object_name = f"{base_name}{suffix}"

    source_id = src._graft_id
    if target_doc._graft_id == source_id:
        raise ValueError("source document must not equal target")

    graft_map = target_doc.Graftmaps.get(source_id)
    if graft_map is None:
        graft_map = fitz.Graftmap(target_doc)
        target_doc.Graftmaps[source_id] = graft_map

    page_id = (source_id, pno)
    xref = target_doc.ShownPages.get(page_id, 0)

    xref = page._show_pdf_page(
        source_page,
        overlay=overlay,
        matrix=fitz.JM_TUPLE(matrix),
        xref=xref,
        oc=oc,
        clip=source_rect,
        graftmap=graft_map,
        _imgname=object_name,
    )
    target_doc.ShownPages[page_id] = xref
    return xref


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
                    self._place_logo(page, logo_doc, logo_page, dest_rect, rotate)
            finally:
                logo_doc.close()
            doc.save(self.params.artwork_path, deflate=True, garbage=4)
        finally:
            doc.close()

    def _place_logo(
        self,
        page: fitz.Page,
        logo_doc: fitz.Document,
        logo_page: fitz.Page,
        dest_rect: fitz.Rect,
        rotate_bottom: bool,
    ) -> None:
        rotate_degrees = 180 if rotate_bottom else 0
        try:
            _show_pdf_page_transformed(
                page,
                dest_rect,
                logo_doc,
                rotate=rotate_degrees,
                flip_horizontal=self.params.flip_in_app,
            )
        except Exception:
            if not self.params.allow_raster_fallback:
                raise
            scale = fitz.Matrix(600 / 72, 600 / 72)
            transform = scale
            if self.params.flip_in_app:
                transform = fitz.Matrix(
                    -1,
                    0,
                    0,
                    1,
                    logo_page.rect.width,
                    0,
                ) * transform
            if rotate_degrees:
                transform = fitz.Matrix(
                    -1,
                    0,
                    0,
                    -1,
                    logo_page.rect.width,
                    logo_page.rect.height,
                ) * transform
            pix = logo_page.get_pixmap(matrix=transform)
            page.insert_image(dest_rect, pixmap=pix)

