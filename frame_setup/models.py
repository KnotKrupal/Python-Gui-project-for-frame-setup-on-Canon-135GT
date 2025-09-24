from __future__ import annotations

from dataclasses import dataclass


@dataclass
class LogoAsset:
    """Represents an imported logo converted to PDF."""

    name: str
    pdf_bytes: bytes
    width_pt: float
    height_pt: float

    def aspect_ratio(self) -> float:
        return self.width_pt / self.height_pt if self.height_pt else 1.0


@dataclass
class JobParameters:
    """Container for all numeric and configuration inputs of a job."""

    job_name: str
    glass_width_mm: float
    glass_height_mm: float
    indent_mm: float
    matte_total_height_mm: float
    cluster_gap_mm: float
    bed_width_in: float
    outline_color: tuple[float, float, float]
    outline_thickness_mm: float
    cluster_count: int
    cluster_rows: int
    frame_quantity: int
    flip_in_app: bool
    allow_raster_fallback: bool
    rotate_bottom: bool
    row_gap_mm: float
    bed_height_in: float
    output_directory: str
    outline_filename: str
    artwork_filename: str

    @property
    def outline_path(self) -> str:
        return f"{self.output_directory}/{self.outline_filename}"

    @property
    def artwork_path(self) -> str:
        return f"{self.output_directory}/{self.artwork_filename}"

    @property
    def frames_per_cluster(self) -> int:
        return 4

    @property
    def max_frames(self) -> int:
        return self.cluster_count * self.cluster_rows * self.frames_per_cluster

