from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

import fitz

from .colors import DEFAULT_COLOR_NAME, PRESET_COLORS
from .ghostscript import GhostscriptError, convert_eps_to_pdf_bytes
from .geometry import (
    MatteGeometry,
    calculate_capacity_from_values,
    calculate_matte_geometry,
    calculate_vertical_capacity,
)
from .models import JobParameters, LogoAsset
from .pdf_exporter import PDFBuilder
from .utils import (
    ValidationError,
    ensure_directory,
    hex_to_rgb_floats,
    parse_non_negative_float,
    parse_positive_float,
    pt_to_mm,
    slugify,
)


SPLASH_DURATION_MS = 1800
APP_TITLE = "Arizona 135 GT Frame Setup"
AUTHOR_NAME = "Developed for SRG by ChatGPT"

class SplashScreen:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.window = tk.Toplevel(root)
        self.window.title("SRG Automation")
        self.window.overrideredirect(True)
        self.window.configure(bg="#1f1f1f")
        self._build_contents()
        self._center()

    def _build_contents(self) -> None:
        frame = ttk.Frame(self.window, padding=30)
        frame.pack(fill="both", expand=True)
        title = ttk.Label(
            frame,
            text="SmartImprint Arizona Suite",
            font=("Segoe UI", 18, "bold"),
        )
        title.pack(pady=(0, 8))
        subtitle = ttk.Label(frame, text="SRG Production Tools", font=("Segoe UI", 11))
        subtitle.pack()
        author = ttk.Label(frame, text=AUTHOR_NAME, font=("Segoe UI", 9, "italic"))
        author.pack(pady=(15, 0))

    def _center(self) -> None:
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def close(self) -> None:
        self.window.destroy()


class FrameSetupApp(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master)
        self.master.title(APP_TITLE)
        self.master.geometry("960x640")
        self.master.minsize(820, 580)
        self.pack(fill="both", expand=True)
        self.logo_asset: Optional[LogoAsset] = None
        self._create_style()
        self._create_variables()
        self._build_layout()
        self._bind_variable_updates()
        self.update_cluster_capacity()
        self.update_preview()
        self.set_status("Ready")

    def _create_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9))

    def _create_variables(self) -> None:
        self.job_name_var = tk.StringVar(value="SRG_Frame_Run")
        self.glass_width_var = tk.StringVar(value="180")
        self.glass_height_var = tk.StringVar(value="240")
        self.indent_var = tk.StringVar(value="10")
        self.matte_total_var = tk.StringVar(value="120")
        self.cluster_gap_var = tk.StringVar(value="15")
        self.bed_width_var = tk.StringVar(value="96")
        self.bed_height_var = tk.StringVar(value="48")
        self.outline_thickness_var = tk.StringVar(value="0.5")
        self.cluster_var = tk.IntVar(value=1)
        self.cluster_rows_var = tk.IntVar(value=1)
        self.frame_quantity_var = tk.IntVar(value=4)
        self.color_var = tk.StringVar(value=DEFAULT_COLOR_NAME)
        self.flip_var = tk.BooleanVar(value=False)
        self.raster_fallback_var = tk.BooleanVar(value=False)
        self.rotate_bottom_var = tk.BooleanVar(value=True)
        self.export_outlines_var = tk.BooleanVar(value=True)
        self.export_artwork_var = tk.BooleanVar(value=True)
        self.output_dir_var = tk.StringVar(value=os.getcwd())
        self.cluster_info_var = tk.StringVar(value="")
        self.frame_info_var = tk.StringVar(value="")
        self.logo_info_var = tk.StringVar(value="Logo size: --")
        self.selected_logo_var = tk.StringVar(value="No logo loaded")
        self.preview_message_var = tk.StringVar(
            value="Preview updates as you adjust frame dimensions."
        )
        self.status_var = tk.StringVar()
        self.progress_var = tk.DoubleVar(value=0.0)
        self._preview_after_id: Optional[str] = None
        self.preview_canvas: Optional[tk.Canvas] = None
        self.reference_canvas: Optional[tk.Canvas] = None
        self.row_gap_var = tk.StringVar(value="25")

    def _build_layout(self) -> None:
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=16, pady=16)

        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)

        self._build_inputs(container)
        self._build_options(container)
        self._build_footer()

    def _build_inputs(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Job Parameters")
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=(0, 12))
        for i in range(4):
            frame.columnconfigure(i, weight=1, uniform="params")

        row = 0
        ttk.Label(frame, text="Job name:").grid(row=row, column=0, sticky="w", pady=4)
        entry_job = ttk.Entry(frame, textvariable=self.job_name_var)
        entry_job.grid(row=row, column=1, columnspan=3, sticky="ew", pady=4)

        row += 1
        ttk.Label(frame, text="Glass width (mm):").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.glass_width_var).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Glass height (mm):").grid(row=row, column=2, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.glass_height_var).grid(row=row, column=3, sticky="ew", pady=4)

        row += 1
        ttk.Label(frame, text="Frame indent (mm):").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.indent_var).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Matte total height (mm):").grid(row=row, column=2, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.matte_total_var).grid(row=row, column=3, sticky="ew", pady=4)

        row += 1
        ttk.Label(frame, text="Cluster gap (mm):").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.cluster_gap_var).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Bed width (inches):").grid(row=row, column=2, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.bed_width_var).grid(row=row, column=3, sticky="ew", pady=4)

        row += 1
        ttk.Label(frame, text="Row gap (mm):").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.row_gap_var).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Bed height (inches):").grid(row=row, column=2, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.bed_height_var).grid(row=row, column=3, sticky="ew", pady=4)

        row += 1
        ttk.Label(frame, text="Outline color:").grid(row=row, column=0, sticky="w", pady=4)
        color_combo = ttk.Combobox(
            frame,
            textvariable=self.color_var,
            values=list(PRESET_COLORS.keys()),
            state="readonly",
        )
        color_combo.grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Outline thickness (mm):").grid(row=row, column=2, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.outline_thickness_var).grid(row=row, column=3, sticky="ew", pady=4)

        row += 1
        ttk.Label(frame, text="Cluster count:").grid(row=row, column=0, sticky="w", pady=4)
        spin = ttk.Spinbox(
            frame,
            from_=1,
            to=10,
            textvariable=self.cluster_var,
            width=6,
            command=self._on_cluster_spin,
        )
        spin.grid(row=row, column=1, sticky="w", pady=4)
        self.cluster_spin = spin
        cluster_info = ttk.Label(frame, textvariable=self.cluster_info_var, style="Header.TLabel")
        cluster_info.grid(row=row, column=2, columnspan=2, sticky="w", pady=4)

        row += 1
        ttk.Label(frame, text="Cluster rows:").grid(row=row, column=0, sticky="w", pady=4)
        row_spin = ttk.Spinbox(
            frame,
            from_=1,
            to=5,
            textvariable=self.cluster_rows_var,
            width=6,
            command=self._on_cluster_spin,
        )
        row_spin.grid(row=row, column=1, sticky="w", pady=4)
        self.cluster_rows_spin = row_spin
        ttk.Label(frame, text="Frame quantity:").grid(row=row, column=2, sticky="w", pady=4)
        frame_spin = ttk.Spinbox(
            frame,
            from_=1,
            to=4,
            textvariable=self.frame_quantity_var,
            width=8,
            command=self._on_frame_quantity_spin,
        )
        frame_spin.grid(row=row, column=3, sticky="w", pady=4)
        self.frame_quantity_spin = frame_spin

        row += 1
        ttk.Label(frame, textvariable=self.frame_info_var).grid(
            row=row, column=0, columnspan=4, sticky="w", pady=(0, 4)
        )

        row += 1
        ttk.Label(frame, text="Output directory:").grid(row=row, column=0, sticky="w", pady=4)
        output_entry = ttk.Entry(frame, textvariable=self.output_dir_var)
        output_entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=4)
        ttk.Button(frame, text="Browse", command=self._choose_output_directory).grid(
            row=row, column=3, sticky="ew", pady=4
        )

        row += 1
        checkbox_row = ttk.Frame(frame)
        checkbox_row.grid(row=row, column=0, columnspan=4, sticky="w", pady=(12, 0))
        ttk.Checkbutton(
            checkbox_row,
            text="Export outlines",
            variable=self.export_outlines_var,
        ).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(
            checkbox_row,
            text="Export artwork",
            variable=self.export_artwork_var,
        ).pack(side="left")

        row += 1
        ttk.Button(frame, text="Load EPS Logo", command=self._load_logo).grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(12, 4)
        )
        ttk.Button(frame, text="Generate PDFs", command=self._generate_pdfs).grid(
            row=row, column=2, columnspan=2, sticky="ew", pady=(12, 4)
        )

        row += 1
        ttk.Label(frame, textvariable=self.selected_logo_var).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
        ttk.Label(frame, textvariable=self.logo_info_var).grid(
            row=row, column=2, columnspan=2, sticky="w", pady=(4, 0)
        )

        row += 1
        ttk.Separator(frame, orient="horizontal").grid(
            row=row, column=0, columnspan=4, sticky="ew", pady=(12, 6)
        )

        row += 1
        ttk.Label(
            frame,
            text="Live layout preview",
            style="Header.TLabel",
        ).grid(row=row, column=0, columnspan=4, sticky="w")

        row += 1
        preview_container = ttk.Frame(frame)
        preview_container.grid(row=row, column=0, columnspan=4, sticky="nsew", pady=(4, 0))
        frame.rowconfigure(row, weight=1)
        preview_container.columnconfigure(0, weight=3)
        preview_container.columnconfigure(1, weight=2)

        layout_preview = tk.Canvas(
            preview_container,
            height=240,
            background="#f9f9fb",
            highlightthickness=1,
            highlightbackground="#c6c6c6",
        )
        layout_preview.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        reference_preview = tk.Canvas(
            preview_container,
            height=240,
            background="#ffffff",
            highlightthickness=1,
            highlightbackground="#d1d1d1",
        )
        reference_preview.grid(row=0, column=1, sticky="nsew")
        preview_container.rowconfigure(0, weight=1)

        layout_preview.bind("<Configure>", lambda event: self.schedule_preview_update())
        reference_preview.bind("<Configure>", lambda event: self.schedule_preview_update())
        self.preview_canvas = layout_preview
        self.reference_canvas = reference_preview

        row += 1
        ttk.Label(
            frame,
            textvariable=self.preview_message_var,
            font=("Segoe UI", 9, "italic"),
        ).grid(row=row, column=0, columnspan=4, sticky="w", pady=(4, 0))

    def _build_options(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Options & Live Metrics")
        frame.grid(row=0, column=1, sticky="nsew", pady=(0, 12))
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Print behaviour", style="Header.TLabel").grid(row=0, column=0, sticky="w", pady=(4, 6))
        ttk.Checkbutton(frame, text="Flip logo in app (reverse print)", variable=self.flip_var).grid(
            row=1, column=0, sticky="w", pady=2
        )
        ttk.Checkbutton(frame, text="Allow raster fallback", variable=self.raster_fallback_var).grid(
            row=2, column=0, sticky="w", pady=2
        )
        ttk.Checkbutton(frame, text="Rotate bottom row 180°", variable=self.rotate_bottom_var).grid(
            row=3, column=0, sticky="w", pady=2
        )

        ttk.Separator(frame, orient="horizontal").grid(row=4, column=0, sticky="ew", pady=(12, 8))
        ttk.Label(frame, text="Capacity", style="Header.TLabel").grid(row=5, column=0, sticky="w")
        capacity_label = ttk.Label(frame, textvariable=self.cluster_info_var)
        capacity_label.grid(row=6, column=0, sticky="w", pady=4)
        ttk.Label(frame, textvariable=self.frame_info_var).grid(row=7, column=0, sticky="w")

        ttk.Label(frame, text="Guidance", style="Header.TLabel").grid(row=8, column=0, sticky="w", pady=(12, 4))
        guidance_text = (
            "Clusters are placed from the origin (0,0) in the lower-left corner.\n"
            "Artwork and outlines share the same sheet size for guaranteed alignment."
        )
        ttk.Label(frame, text=guidance_text, wraplength=280, justify="left").grid(row=9, column=0, sticky="w")

    def _build_footer(self) -> None:
        footer = ttk.Frame(self)
        footer.pack(fill="x", padx=16, pady=(0, 12))
        progress = ttk.Progressbar(footer, variable=self.progress_var, maximum=100)
        progress.pack(fill="x", side="top")
        status = ttk.Label(footer, textvariable=self.status_var, style="Status.TLabel")
        status.pack(fill="x", side="top", pady=(4, 0))

    def _bind_variable_updates(self) -> None:
        for var in [
            self.glass_width_var,
            self.bed_width_var,
            self.cluster_gap_var,
            self.bed_height_var,
            self.row_gap_var,
        ]:
            var.trace_add("write", lambda *_: self.update_cluster_capacity())
        self.cluster_var.trace_add("write", lambda *_: self.update_cluster_capacity())
        self.cluster_rows_var.trace_add("write", lambda *_: self.update_cluster_capacity())
        self.frame_quantity_var.trace_add("write", lambda *_: self._sync_frame_quantity())
        preview_vars = [
            self.glass_width_var,
            self.glass_height_var,
            self.indent_var,
            self.matte_total_var,
            self.cluster_gap_var,
            self.bed_width_var,
            self.bed_height_var,
            self.cluster_var,
            self.cluster_rows_var,
            self.rotate_bottom_var,
            self.flip_var,
            self.row_gap_var,
            self.frame_quantity_var,
        ]
        for var in preview_vars:
            var.trace_add("write", lambda *_: self.schedule_preview_update())

    def _on_cluster_spin(self) -> None:
        self.update_cluster_capacity()
        self.schedule_preview_update()

    def _on_frame_quantity_spin(self) -> None:
        self._sync_frame_quantity()
        self.schedule_preview_update()

    def set_status(self, message: str) -> None:
        self.status_var.set(message)
        self.master.update_idletasks()

    def set_progress(self, value: float) -> None:
        self.progress_var.set(value)
        self.master.update_idletasks()

    def update_cluster_capacity(self) -> None:
        try:
            glass_width = float(self.glass_width_var.get())
            bed_width = float(self.bed_width_var.get())
            cluster_gap = float(self.cluster_gap_var.get())
            glass_height = float(self.glass_height_var.get())
            bed_height = float(self.bed_height_var.get())
            row_gap = float(self.row_gap_var.get())
        except ValueError:
            self.cluster_info_var.set("Enter numeric values to compute capacity")
            self.frame_info_var.set("")
            self.schedule_preview_update()
            return
        horizontal_capacity = calculate_capacity_from_values(glass_width, bed_width, cluster_gap)
        vertical_capacity = calculate_vertical_capacity(glass_height, bed_height, row_gap)

        max_horizontal = max(horizontal_capacity, 1)
        max_vertical = max(vertical_capacity, 1)

        self.cluster_spin.configure(to=max_horizontal)
        self.cluster_rows_spin.configure(to=max_vertical)

        clusters = self.cluster_var.get()
        if clusters < 1:
            clusters = 1
        if clusters > max_horizontal:
            clusters = max_horizontal
        if clusters != self.cluster_var.get():
            self.cluster_var.set(clusters)

        cluster_rows = self.cluster_rows_var.get()
        if cluster_rows < 1:
            cluster_rows = 1
        if cluster_rows > max_vertical:
            cluster_rows = max_vertical
        if cluster_rows != self.cluster_rows_var.get():
            self.cluster_rows_var.set(cluster_rows)

        max_frames = clusters * cluster_rows * 4
        self._sync_frame_quantity(max_frames)

        if horizontal_capacity < 1 or vertical_capacity < 1:
            self.cluster_info_var.set("No clusters fit current bed size")
        else:
            info = (
                f"Clusters: {clusters}/{horizontal_capacity} across × "
                f"{cluster_rows}/{vertical_capacity} high"
            )
            self.cluster_info_var.set(info)
        self.schedule_preview_update()

    def _sync_frame_quantity(self, max_frames: Optional[int] = None) -> None:
        if max_frames is None:
            max_frames = max(self.cluster_var.get() * self.cluster_rows_var.get() * 4, 1)
        else:
            max_frames = max(max_frames, 1)
        try:
            requested = int(self.frame_quantity_var.get())
        except (ValueError, tk.TclError):
            requested = 1
        if requested < 1:
            requested = 1
        if requested > max_frames:
            requested = max_frames
            self.frame_quantity_var.set(requested)
        self.frame_quantity_spin.configure(to=max_frames)
        self.frame_info_var.set(f"Frames scheduled: {requested} / {max_frames}")

    def schedule_preview_update(self) -> None:
        if not self.preview_canvas or not self.reference_canvas:
            return
        if self._preview_after_id is not None:
            self.after_cancel(self._preview_after_id)
        self._preview_after_id = self.after(75, self.update_preview)

    def update_preview(self) -> None:
        layout_canvas = self.preview_canvas
        reference_canvas = self.reference_canvas
        if layout_canvas is None or reference_canvas is None:
            return
        if self._preview_after_id is not None:
            self.after_cancel(self._preview_after_id)
            self._preview_after_id = None

        for target in (layout_canvas, reference_canvas):
            target.delete("all")

        layout_width = max(
            int(layout_canvas.winfo_width()),
            int(float(layout_canvas.cget("width"))),
        )
        layout_height = max(
            int(layout_canvas.winfo_height()),
            int(float(layout_canvas.cget("height"))),
        )
        reference_width = max(
            int(reference_canvas.winfo_width()),
            int(float(reference_canvas.cget("width"))),
        )
        reference_height = max(
            int(reference_canvas.winfo_height()),
            int(float(reference_canvas.cget("height"))),
        )

        if (
            layout_width <= 2
            or layout_height <= 2
            or reference_width <= 2
            or reference_height <= 2
        ):
            self._preview_after_id = self.after(120, self.update_preview)
            return

        try:
            glass_width = float(self.glass_width_var.get())
            glass_height = float(self.glass_height_var.get())
            indent = float(self.indent_var.get())
            matte_total = float(self.matte_total_var.get())
            cluster_gap = float(self.cluster_gap_var.get())
            bed_width_in = float(self.bed_width_var.get())
            bed_height_in = float(self.bed_height_var.get())
            row_gap = float(self.row_gap_var.get())
        except ValueError:
            message = "Preview unavailable: provide numeric values for dimensions."
            self._render_preview_message(
                layout_canvas,
                layout_width,
                layout_height,
                "Enter numeric values",
            )
            self._render_preview_message(
                reference_canvas,
                reference_width,
                reference_height,
                "Enter numeric values",
            )
            self.preview_message_var.set(message)
            return

        if (
            glass_width <= 0
            or glass_height <= 0
            or bed_width_in <= 0
            or bed_height_in <= 0
        ):
            message = "Preview unavailable: dimensions must be positive."
            info_text = "Enter positive dimensions to view layout"
            self._render_preview_message(layout_canvas, layout_width, layout_height, info_text)
            self._render_preview_message(
                reference_canvas,
                reference_width,
                reference_height,
                info_text,
            )
            self.preview_message_var.set(message)
            return

        cluster_count = max(self.cluster_var.get(), 1)
        cluster_rows = max(self.cluster_rows_var.get(), 1)
        frame_quantity = max(self.frame_quantity_var.get(), 1)

        matte = calculate_matte_geometry(glass_width, glass_height, indent, matte_total)
        bed_width_mm = bed_width_in * 25.4
        bed_height_mm = bed_height_in * 25.4
        cluster_width = glass_width * 2
        cluster_height = glass_height * 2
        gap = max(cluster_gap, 0.0)
        row_gap_mm = max(row_gap, 0.0)
        total_clusters_width = cluster_count * cluster_width + max(cluster_count - 1, 0) * gap
        total_clusters_height = (
            cluster_rows * cluster_height + max(cluster_rows - 1, 0) * row_gap_mm
        )

        logo_width_mm = logo_height_mm = 0.0
        has_logo = False
        if self.logo_asset is not None:
            logo_width_mm = pt_to_mm(self.logo_asset.width_pt)
            logo_height_mm = pt_to_mm(self.logo_asset.height_pt)
            has_logo = logo_width_mm > 0 and logo_height_mm > 0

        placements: list[dict[str, float | int | bool]] = []
        frame_counter = 0
        for cluster_row in range(cluster_rows):
            base_y = cluster_row * (cluster_height + row_gap_mm)
            for cluster_index in range(cluster_count):
                base_x = cluster_index * (cluster_width + gap)
                for row in range(2):
                    frame_y = base_y + row * glass_height
                    for column in range(2):
                        frame_x = base_x + column * glass_width
                        frame_counter += 1
                        active = frame_counter <= frame_quantity
                        placements.append(
                            {
                                "cluster_index": cluster_index,
                                "cluster_row": cluster_row,
                                "row": row,
                                "column": column,
                                "x_mm": frame_x,
                                "y_mm": frame_y,
                                "active": active,
                                "index": frame_counter,
                            }
                        )

        active_count = sum(1 for p in placements if p["active"])

        arrangement_data = {
            "glass_width_mm": glass_width,
            "glass_height_mm": glass_height,
            "indent_mm": max(indent, 0.0),
            "cluster_count": cluster_count,
            "cluster_rows": cluster_rows,
            "cluster_gap_mm": gap,
            "row_gap_mm": row_gap_mm,
            "bed_width_mm": bed_width_mm,
            "bed_height_mm": bed_height_mm,
            "total_clusters_width_mm": total_clusters_width,
            "total_clusters_height_mm": total_clusters_height,
            "logo_width_mm": logo_width_mm,
            "logo_height_mm": logo_height_mm,
            "has_logo": has_logo,
            "rotate_bottom": self.rotate_bottom_var.get(),
            "flip_in_app": self.flip_var.get(),
            "visible_band_mm": matte.visible_band_mm,
            "placements": placements,
            "frame_quantity": frame_quantity,
            "active_count": active_count,
            "max_frames": len(placements),
        }

        self._draw_arrangement_preview(
            layout_canvas,
            layout_width,
            layout_height,
            arrangement_data,
            matte,
        )
        self._draw_reference_preview(
            reference_canvas,
            reference_width,
            reference_height,
            arrangement_data,
            matte,
        )

        width_text: str
        if bed_width_mm > 0:
            width_ratio = total_clusters_width / bed_width_mm
            width_prefix = "⚠️ " if width_ratio > 1.001 else ""
            width_text = (
                f"{width_prefix}Width usage: {total_clusters_width:.1f} / {bed_width_mm:.1f} mm"
                f" ({width_ratio * 100:.0f}%)"
            )
        else:
            width_text = f"Total layout width: {total_clusters_width:.1f} mm"

        if bed_height_mm > 0:
            height_ratio = total_clusters_height / bed_height_mm if bed_height_mm else 0.0
            height_prefix = "⚠️ " if height_ratio > 1.001 else ""
            height_text = (
                f"{height_prefix}Height usage: {total_clusters_height:.1f} / {bed_height_mm:.1f} mm"
                f" ({height_ratio * 100:.0f}%)"
            )
        else:
            height_text = f"Total layout height: {total_clusters_height:.1f} mm"

        rotation_text = (
            "Bottom row logos rotated 180°"
            if arrangement_data["rotate_bottom"]
            else "Bottom row logos upright"
        )
        flip_text = "Mirrored in app" if arrangement_data["flip_in_app"] else "Not mirrored in app"
        matte_text = (
            f"Matte opening: {matte.opening_width_mm:.1f} × {matte.opening_height_mm:.1f} mm"
        )
        borders_text = (
            f"Borders top/bottom/side: {matte.top_margin_mm:.1f} / {matte.bottom_margin_mm:.1f} / {matte.side_margin_mm:.1f} mm"
        )
        if has_logo:
            logo_text = f"Logo: {logo_width_mm:.1f} × {logo_height_mm:.1f} mm"
        else:
            logo_text = "Logo: load EPS to preview placement"
        frames_text = (
            f"Frames scheduled: {active_count} / {len(placements)}"
        )
        self.preview_message_var.set(
            f"{width_text}\n{height_text}\n{matte_text}\n{borders_text}\n"
            f"{rotation_text} • {flip_text}\n{frames_text}\n{logo_text}"
        )

    def _draw_arrangement_preview(
        self,
        canvas: tk.Canvas,
        width: int,
        height: int,
        data: dict[str, float | int | bool | list],
        matte: MatteGeometry,
    ) -> None:
        margin = 16
        usable_width = max(width - margin * 2, 1)
        usable_height = max(height - margin * 2, 1)

        bed_width_mm = float(data["bed_width_mm"])
        bed_height_mm = float(data["bed_height_mm"])
        total_width_mm = float(data["total_clusters_width_mm"])
        total_height_mm = float(data["total_clusters_height_mm"])
        glass_width_mm = float(data["glass_width_mm"])
        glass_height_mm = float(data["glass_height_mm"])
        row_gap_mm = float(data["row_gap_mm"])
        cluster_rows = int(data["cluster_rows"])

        content_width = max(bed_width_mm, total_width_mm, 1.0)
        content_height = max(bed_height_mm, total_height_mm, 1.0)
        scale = min(usable_width / content_width, usable_height / content_height)

        bed_left = margin
        bed_bottom = height - margin
        bed_right = bed_left + bed_width_mm * scale
        bed_top = bed_bottom - bed_height_mm * scale

        canvas.create_rectangle(
            bed_left,
            bed_top,
            bed_right,
            bed_bottom,
            outline="#4c5c78",
            width=1.4,
        )

        arrangement_right = bed_left + total_width_mm * scale
        for cluster_row in range(cluster_rows):
            seam_mm = cluster_row * (2 * glass_height_mm + row_gap_mm) + glass_height_mm
            seam_y = bed_bottom - seam_mm * scale
            canvas.create_line(
                bed_left,
                seam_y,
                arrangement_right,
                seam_y,
                fill="#9fa8b8",
                dash=(4, 3),
            )

        frame_fills = ["#fbe3d8", "#d8e5fb"]
        matte_color = "#f2ddc6"
        opening_outline = "#b9986a"
        visible_band_mm = float(data["visible_band_mm"])
        has_logo = bool(data["has_logo"])
        indent_mm = float(data["indent_mm"])
        logo_width_mm = float(data["logo_width_mm"])
        logo_height_mm = float(data["logo_height_mm"])
        rotate_bottom = bool(data["rotate_bottom"])

        for placement in data["placements"]:
            frame_x = float(placement["x_mm"])
            frame_y = float(placement["y_mm"])
            row = int(placement["row"])
            active = bool(placement["active"])
            index = int(placement.get("index", 0))

            left = bed_left + frame_x * scale
            right = left + glass_width_mm * scale
            top = bed_bottom - (frame_y + glass_height_mm) * scale
            bottom = bed_bottom - frame_y * scale

            fill_color = frame_fills[row] if active else "#e6e8ef"
            outline_color = "#596982" if active else "#b3b8c6"
            canvas.create_rectangle(
                left,
                top,
                right,
                bottom,
                fill=fill_color,
                outline=outline_color,
                width=1.2 if active else 1.0,
            )

            if index:
                label_color = "#2f3a4f" if active else "#7f8797"
                canvas.create_text(
                    left + 12,
                    top + 14,
                    text=str(index),
                    fill=label_color,
                    font=("Segoe UI", 9, "bold"),
                    anchor="w",
                )

            if not active:
                canvas.create_line(left, top, right, bottom, fill="#b6bbc7", width=1)
                canvas.create_line(left, bottom, right, top, fill="#b6bbc7", width=1)
                continue

            opening_left_mm = frame_x + matte.side_margin_mm
            opening_right_mm = frame_x + glass_width_mm - matte.side_margin_mm
            opening_bottom_mm = frame_y + matte.bottom_margin_mm
            opening_top_mm = opening_bottom_mm + matte.opening_height_mm

            opening_left = bed_left + opening_left_mm * scale
            opening_right = bed_left + opening_right_mm * scale
            opening_bottom = bed_bottom - opening_bottom_mm * scale
            opening_top = bed_bottom - opening_top_mm * scale

            if opening_left - left > 1:
                canvas.create_rectangle(
                    left,
                    top,
                    opening_left,
                    bottom,
                    fill=matte_color,
                    outline="",
                )
            if right - opening_right > 1:
                canvas.create_rectangle(
                    opening_right,
                    top,
                    right,
                    bottom,
                    fill=matte_color,
                    outline="",
                )
            if opening_top - top > 1:
                canvas.create_rectangle(
                    opening_left,
                    top,
                    opening_right,
                    opening_top,
                    fill=matte_color,
                    outline="",
                )
            if bottom - opening_bottom > 1:
                canvas.create_rectangle(
                    opening_left,
                    opening_bottom,
                    opening_right,
                    bottom,
                    fill=matte_color,
                    outline="",
                )
            if opening_right - opening_left > 2 and opening_bottom - opening_top > 2:
                canvas.create_rectangle(
                    opening_left,
                    opening_top,
                    opening_right,
                    opening_bottom,
                    outline=opening_outline,
                    width=1.1,
                    fill="#ffffff",
                )

            if matte.bottom_margin_mm > 0:
                bottom_seam = bed_bottom - (frame_y + matte.bottom_margin_mm) * scale
                canvas.create_line(
                    opening_left,
                    bottom_seam,
                    opening_right,
                    bottom_seam,
                    fill="#c6812f",
                    dash=(4, 2),
                )
            if matte.top_margin_mm > 0:
                top_seam = bed_bottom - (frame_y + glass_height_mm - matte.top_margin_mm) * scale
                canvas.create_line(
                    opening_left,
                    top_seam,
                    opening_right,
                    top_seam,
                    fill="#c6812f",
                    dash=(4, 2),
                )

            if has_logo and visible_band_mm > 0 and logo_width_mm > 0 and logo_height_mm > 0:
                center_x_mm = frame_x + glass_width_mm / 2
                if row == 1:
                    center_y_mm = frame_y + visible_band_mm / 2
                else:
                    center_y_mm = frame_y + glass_height_mm - visible_band_mm / 2
                min_center = frame_y + indent_mm + logo_height_mm / 2
                max_center = frame_y + glass_height_mm - indent_mm - logo_height_mm / 2
                if min_center > max_center:
                    min_center = frame_y + logo_height_mm / 2
                    max_center = frame_y + glass_height_mm - logo_height_mm / 2
                center_y_mm = max(min(center_y_mm, max_center), min_center)

                logo_left = bed_left + (center_x_mm - logo_width_mm / 2) * scale
                logo_right = bed_left + (center_x_mm + logo_width_mm / 2) * scale
                logo_top = bed_bottom - (center_y_mm + logo_height_mm / 2) * scale
                logo_bottom = bed_bottom - (center_y_mm - logo_height_mm / 2) * scale
                if logo_right - logo_left > 2 and logo_bottom - logo_top > 2:
                    canvas.create_rectangle(
                        logo_left,
                        logo_top,
                        logo_right,
                        logo_bottom,
                        outline="#1f5137",
                        width=1.4,
                        fill="#ffffff",
                    )

            arrow_center_x = (left + right) / 2
            arrow_center_y = (top + bottom) / 2
            arrow_height = min((bottom - top) * 0.45, 26)
            arrow_half = arrow_height / 2
            arrow_width = arrow_half * 0.75
            pointing_up = row == 1 or (row == 0 and not rotate_bottom)
            arrow_color = "#2b455d"
            if pointing_up:
                points = [
                    arrow_center_x,
                    arrow_center_y - arrow_half,
                    arrow_center_x - arrow_width,
                    arrow_center_y + arrow_half,
                    arrow_center_x + arrow_width,
                    arrow_center_y + arrow_half,
                ]
            else:
                points = [
                    arrow_center_x,
                    arrow_center_y + arrow_half,
                    arrow_center_x - arrow_width,
                    arrow_center_y - arrow_half,
                    arrow_center_x + arrow_width,
                    arrow_center_y - arrow_half,
                ]
            canvas.create_polygon(points, fill=arrow_color, outline="")

    def _draw_reference_preview(
        self,
        canvas: tk.Canvas,
        width: int,
        height: int,
        data: dict[str, float | bool],
        matte: MatteGeometry,
    ) -> None:
        margin_x = 28
        margin_y = 34
        available_width = max(width - margin_x * 2, 1)
        available_height = max(height - margin_y * 2, 1)
        scale = min(
            available_width / float(data["glass_width_mm"]),
            available_height / float(data["glass_height_mm"]),
        )

        frame_width = float(data["glass_width_mm"]) * scale
        frame_height = float(data["glass_height_mm"]) * scale
        left = (width - frame_width) / 2
        top = (height - frame_height) / 2
        right = left + frame_width
        bottom = top + frame_height

        canvas.create_text(
            left,
            top - 12,
            text="1-Up Reference",
            anchor="sw",
            font=("Segoe UI", 11, "bold"),
            fill="#303b4f",
        )

        canvas.create_rectangle(
            left,
            top,
            right,
            bottom,
            outline="#41526d",
            width=2,
            fill="#f5f7fb",
        )

        def x_mm_to_px(value_mm: float) -> float:
            return left + value_mm * scale

        def y_mm_to_px(value_mm: float) -> float:
            return bottom - value_mm * scale

        opening_left = x_mm_to_px(matte.side_margin_mm)
        opening_right = x_mm_to_px(float(data["glass_width_mm"]) - matte.side_margin_mm)
        opening_bottom = y_mm_to_px(matte.bottom_margin_mm)
        opening_top = y_mm_to_px(matte.bottom_margin_mm + matte.opening_height_mm)

        matte_color = "#f2ddc6"
        if opening_left - left > 1:
            canvas.create_rectangle(left, top, opening_left, bottom, fill=matte_color, outline="")
        if right - opening_right > 1:
            canvas.create_rectangle(opening_right, top, right, bottom, fill=matte_color, outline="")
        if opening_top - top > 1:
            canvas.create_rectangle(opening_left, top, opening_right, opening_top, fill=matte_color, outline="")
        if bottom - opening_bottom > 1:
            canvas.create_rectangle(opening_left, opening_bottom, opening_right, bottom, fill=matte_color, outline="")
        if opening_right - opening_left > 2 and opening_bottom - opening_top > 2:
            canvas.create_rectangle(
                opening_left,
                opening_top,
                opening_right,
                opening_bottom,
                outline="#b9986a",
                width=1.5,
                fill="#ffffff",
            )

        bottom_seam = y_mm_to_px(matte.bottom_margin_mm)
        top_seam = y_mm_to_px(float(data["glass_height_mm"]) - matte.top_margin_mm)
        canvas.create_line(
            opening_left,
            bottom_seam,
            opening_right,
            bottom_seam,
            fill="#c6812f",
            dash=(4, 2),
        )
        canvas.create_text(
            opening_left,
            min(bottom_seam + 12, height - 10),
            text="Bottom seam",
            anchor="w",
            font=("Segoe UI", 8, "italic"),
            fill="#c6812f",
        )
        canvas.create_line(
            opening_left,
            top_seam,
            opening_right,
            top_seam,
            fill="#c6812f",
            dash=(4, 2),
        )
        canvas.create_text(
            opening_left,
            max(top_seam - 12, top + 10),
            text="Top seam",
            anchor="w",
            font=("Segoe UI", 8, "italic"),
            fill="#c6812f",
        )

        width_line_y = min(bottom + 18, height - 18)
        canvas.create_line(left, width_line_y, right, width_line_y, arrow=tk.BOTH, fill="#41526d")
        canvas.create_text(
            (left + right) / 2,
            min(bottom + 32, height - 6),
            text=f"Glass width {float(data['glass_width_mm']):.1f} mm",
            font=("Segoe UI", 9),
            fill="#41526d",
        )
        height_arrow_x = max(left - 18, 12)
        canvas.create_line(height_arrow_x, top, height_arrow_x, bottom, arrow=tk.BOTH, fill="#41526d")
        canvas.create_text(
            max(left - 34, 16),
            (top + bottom) / 2,
            text=f"{float(data['glass_height_mm']):.1f} mm",
            font=("Segoe UI", 9),
            angle=90,
            fill="#41526d",
        )

        right_arrow_x = min(right + 14, width - 18)
        canvas.create_line(
            right_arrow_x,
            opening_top,
            right_arrow_x,
            opening_bottom,
            arrow=tk.BOTH,
            fill="#b17a2e",
        )
        canvas.create_text(
            min(right + 32, width - 8),
            (opening_top + opening_bottom) / 2,
            text=f"Opening {matte.opening_height_mm:.1f} mm",
            font=("Segoe UI", 8),
            angle=90,
            fill="#b17a2e",
        )
        side_arrow_y = max(top - 16, 12)
        canvas.create_line(
            opening_left,
            side_arrow_y,
            left,
            side_arrow_y,
            arrow=tk.BOTH,
            fill="#b17a2e",
        )
        canvas.create_text(
            (opening_left + left) / 2,
            max(top - 28, 10),
            text=f"Side border {matte.side_margin_mm:.1f} mm",
            font=("Segoe UI", 8),
            fill="#b17a2e",
        )

        canvas.create_text(
            min(right, width - 8),
            max(top - 12, 10),
            text=("Flip in app" if data["flip_in_app"] else "No app flip"),
            anchor="se",
            font=("Segoe UI", 8),
            fill="#41526d",
        )

        if data["has_logo"] and float(data["visible_band_mm"]) > 0:
            logo_width_mm = float(data["logo_width_mm"])
            logo_height_mm = float(data["logo_height_mm"])
            indent_mm = float(data["indent_mm"])
            visible_band_mm = float(data["visible_band_mm"])
            glass_height_mm = float(data["glass_height_mm"])
            glass_width_mm = float(data["glass_width_mm"])

            def draw_logo(center_y_mm: float, outline: str) -> None:
                min_center = indent_mm + logo_height_mm / 2
                max_center = glass_height_mm - indent_mm - logo_height_mm / 2
                if min_center > max_center:
                    min_center = logo_height_mm / 2
                    max_center = glass_height_mm - logo_height_mm / 2
                center_y = max(min(center_y_mm, max_center), min_center)
                left_px = x_mm_to_px(glass_width_mm / 2 - logo_width_mm / 2)
                right_px = x_mm_to_px(glass_width_mm / 2 + logo_width_mm / 2)
                top_px = y_mm_to_px(center_y + logo_height_mm / 2)
                bottom_px = y_mm_to_px(center_y - logo_height_mm / 2)
                if right_px - left_px > 2 and bottom_px - top_px > 2:
                    canvas.create_rectangle(
                        left_px,
                        top_px,
                        right_px,
                        bottom_px,
                        outline=outline,
                        width=1.5,
                        fill="#ffffff",
                    )
                    arrow_height = min((bottom_px - top_px) * 0.6, 28)
                    arrow_half = arrow_height / 2
                    arrow_width = arrow_half * 0.7
                    center_x = (left_px + right_px) / 2
                    pointing_up = outline == "#1f5137"
                    if not pointing_up and data["rotate_bottom"]:
                        pointing_up = False
                    elif not pointing_up:
                        pointing_up = True
                    arrow_color = "#2b455d"
                    if pointing_up:
                        points = [
                            center_x,
                            (top_px + bottom_px) / 2 - arrow_half,
                            center_x - arrow_width,
                            (top_px + bottom_px) / 2 + arrow_half,
                            center_x + arrow_width,
                            (top_px + bottom_px) / 2 + arrow_half,
                        ]
                    else:
                        points = [
                            center_x,
                            (top_px + bottom_px) / 2 + arrow_half,
                            center_x - arrow_width,
                            (top_px + bottom_px) / 2 - arrow_half,
                            center_x + arrow_width,
                            (top_px + bottom_px) / 2 - arrow_half,
                        ]
                    canvas.create_polygon(points, fill=arrow_color, outline="")

            top_center = visible_band_mm / 2
            bottom_center = glass_height_mm - visible_band_mm / 2
            draw_logo(top_center, "#1f5137")
            draw_logo(bottom_center, "#3b5e85")

    def _render_preview_message(self, canvas: tk.Canvas, width: int, height: int, message: str) -> None:
        canvas.create_text(
            width / 2,
            height / 2,
            text=message,
            fill="#8a8a8a",
            font=("Segoe UI", 11),
        )

    def _choose_output_directory(self) -> None:
        directory = filedialog.askdirectory(initialdir=self.output_dir_var.get() or os.getcwd())
        if directory:
            self.output_dir_var.set(directory)

    def _load_logo(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select EPS Logo",
            filetypes=[("Encapsulated PostScript", "*.eps"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            pdf_bytes = convert_eps_to_pdf_bytes(file_path)
            with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
                page = doc[0]
                width_pt = page.rect.width
                height_pt = page.rect.height
            self.logo_asset = LogoAsset(
                name=os.path.basename(file_path),
                pdf_bytes=pdf_bytes,
                width_pt=width_pt,
                height_pt=height_pt,
            )
            width_mm = pt_to_mm(width_pt)
            height_mm = pt_to_mm(height_pt)
            self.logo_info_var.set(f"Logo size: {width_mm:.2f} mm × {height_mm:.2f} mm")
            self.selected_logo_var.set(f"Loaded: {os.path.basename(file_path)}")
            self.set_status("Logo loaded successfully")
            self.schedule_preview_update()
        except GhostscriptError as exc:
            messagebox.showerror("Ghostscript error", str(exc))
            self.set_status("Ghostscript conversion failed")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Logo error", str(exc))
            self.set_status("Failed to load logo")

    def _validate_inputs(self) -> JobParameters:
        job_name = self.job_name_var.get().strip()
        if not job_name:
            raise ValidationError("Job name", "cannot be empty")
        glass_width = parse_positive_float(self.glass_width_var.get(), "Glass width")
        glass_height = parse_positive_float(self.glass_height_var.get(), "Glass height")
        indent = parse_non_negative_float(self.indent_var.get(), "Frame indent")
        matte_total = parse_positive_float(self.matte_total_var.get(), "Matte total height")
        cluster_gap = parse_non_negative_float(self.cluster_gap_var.get(), "Cluster gap")
        bed_width = parse_positive_float(self.bed_width_var.get(), "Bed width")
        bed_height = parse_positive_float(self.bed_height_var.get(), "Bed height")
        row_gap = parse_non_negative_float(self.row_gap_var.get(), "Row gap")
        outline_thickness = parse_positive_float(
            self.outline_thickness_var.get(), "Outline thickness"
        )
        cluster_count = self.cluster_var.get()
        if cluster_count < 1:
            raise ValidationError("Cluster count", "must be at least 1")
        cluster_rows = self.cluster_rows_var.get()
        if cluster_rows < 1:
            raise ValidationError("Cluster rows", "must be at least 1")
        horizontal_capacity = calculate_capacity_from_values(glass_width, bed_width, cluster_gap)
        vertical_capacity = calculate_vertical_capacity(glass_height, bed_height, row_gap)
        if horizontal_capacity < 1 or vertical_capacity < 1:
            raise ValidationError("Cluster layout", "does not fit on the bed with current settings")
        if cluster_count > horizontal_capacity:
            raise ValidationError(
                "Cluster count",
                f"cannot exceed {horizontal_capacity} across the bed width",
            )
        if cluster_rows > vertical_capacity:
            raise ValidationError(
                "Cluster rows",
                f"cannot exceed {vertical_capacity} within the bed height",
            )
        max_frames = cluster_count * cluster_rows * 4
        frame_quantity = self.frame_quantity_var.get()
        if frame_quantity < 1:
            raise ValidationError("Frame quantity", "must be at least 1")
        if frame_quantity > max_frames:
            raise ValidationError(
                "Frame quantity",
                f"cannot exceed {max_frames} for the selected clusters",
            )
        color_name = self.color_var.get()
        color = PRESET_COLORS.get(color_name)
        if not color:
            raise ValidationError("Outline color", "invalid selection")
        outline_color = hex_to_rgb_floats(color.hex_value)
        output_directory = self.output_dir_var.get().strip() or os.getcwd()
        ensure_directory(output_directory)
        job_slug = slugify(job_name)
        outline_filename = f"{job_slug}_ArizonaOutlines.pdf"
        artwork_filename = f"{job_slug}_ArizonaArtwork.pdf"
        return JobParameters(
            job_name=job_name,
            glass_width_mm=glass_width,
            glass_height_mm=glass_height,
            indent_mm=indent,
            matte_total_height_mm=matte_total,
            cluster_gap_mm=cluster_gap,
            bed_width_in=bed_width,
            bed_height_in=bed_height,
            outline_color=outline_color,
            outline_thickness_mm=outline_thickness,
            cluster_count=cluster_count,
            cluster_rows=cluster_rows,
            frame_quantity=frame_quantity,
            flip_in_app=self.flip_var.get(),
            allow_raster_fallback=self.raster_fallback_var.get(),
            rotate_bottom=self.rotate_bottom_var.get(),
            row_gap_mm=row_gap,
            output_directory=output_directory,
            outline_filename=outline_filename,
            artwork_filename=artwork_filename,
        )

    def _generate_pdfs(self) -> None:
        export_outlines = self.export_outlines_var.get()
        export_artwork = self.export_artwork_var.get()
        if not export_outlines and not export_artwork:
            messagebox.showwarning(
                "No outputs selected",
                "Select at least one PDF to export.",
            )
            return
        if export_artwork and self.logo_asset is None:
            messagebox.showwarning(
                "Missing logo",
                "Please load an EPS logo before exporting artwork.",
            )
            return
        try:
            params = self._validate_inputs()
        except ValidationError as exc:
            messagebox.showerror("Invalid input", str(exc))
            return
        self.set_status("Building PDFs…")
        self.set_progress(0)
        self._set_busy(True)
        try:
            logo_asset = self.logo_asset or LogoAsset(
                name="placeholder",
                pdf_bytes=b"",
                width_pt=0.0,
                height_pt=0.0,
            )
            builder = PDFBuilder(params, logo_asset)
            result = builder.generate_pdfs(
                self._progress_callback,
                export_outlines=export_outlines,
                export_artwork=export_artwork,
            )
            exports = []
            if result.outline_path:
                exports.append(os.path.basename(result.outline_path))
            if result.artwork_path:
                exports.append(os.path.basename(result.artwork_path))
            if exports:
                self.set_status("Export complete: " + " • ".join(exports))
            else:
                self.set_status("No files generated")
            self.set_progress(100)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Generation failed", str(exc))
            self.set_status("Generation failed")
        finally:
            self._set_busy(False)

    def _progress_callback(self, value: int, message: str) -> None:
        self.set_progress(value)
        if message:
            self.set_status(message)

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        for child in self.winfo_children():
            self._set_state_recursive(child, state)
        self.update_idletasks()

    def _set_state_recursive(self, widget: tk.Widget, state: str) -> None:
        if isinstance(widget, (ttk.Button, ttk.Entry, ttk.Combobox, ttk.Checkbutton, ttk.Spinbox)):
            widget.state([state]) if state == "disabled" else widget.state(["!disabled"])
        for child in widget.winfo_children():
            self._set_state_recursive(child, state)


def run_app() -> None:
    root = tk.Tk()
    root.withdraw()
    splash = SplashScreen(root)

    def show_main() -> None:
        splash.close()
        root.deiconify()

    root.after(SPLASH_DURATION_MS, show_main)
    app = FrameSetupApp(root)
    root.mainloop()


__all__ = ["FrameSetupApp", "run_app"]

