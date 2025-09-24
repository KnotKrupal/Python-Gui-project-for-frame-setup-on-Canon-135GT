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
        self.outline_thickness_var = tk.StringVar(value="0.5")
        self.cluster_var = tk.IntVar(value=1)
        self.color_var = tk.StringVar(value=DEFAULT_COLOR_NAME)
        self.flip_var = tk.BooleanVar(value=False)
        self.raster_fallback_var = tk.BooleanVar(value=False)
        self.rotate_bottom_var = tk.BooleanVar(value=True)
        self.output_dir_var = tk.StringVar(value=os.getcwd())
        self.cluster_info_var = tk.StringVar(value="")
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
        ttk.Label(frame, text="Output directory:").grid(row=row, column=0, sticky="w", pady=4)
        output_entry = ttk.Entry(frame, textvariable=self.output_dir_var)
        output_entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=4)
        ttk.Button(frame, text="Browse", command=self._choose_output_directory).grid(
            row=row, column=3, sticky="ew", pady=4
        )

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

        ttk.Label(frame, text="Guidance", style="Header.TLabel").grid(row=7, column=0, sticky="w", pady=(12, 4))
        guidance_text = (
            "Clusters are placed from the origin (0,0) in the lower-left corner.\n"
            "Artwork and outlines share the same sheet size for guaranteed alignment."
        )
        ttk.Label(frame, text=guidance_text, wraplength=280, justify="left").grid(row=8, column=0, sticky="w")

    def _build_footer(self) -> None:
        footer = ttk.Frame(self)
        footer.pack(fill="x", padx=16, pady=(0, 12))
        progress = ttk.Progressbar(footer, variable=self.progress_var, maximum=100)
        progress.pack(fill="x", side="top")
        status = ttk.Label(footer, textvariable=self.status_var, style="Status.TLabel")
        status.pack(fill="x", side="top", pady=(4, 0))

    def _bind_variable_updates(self) -> None:
        for var in [self.glass_width_var, self.bed_width_var, self.cluster_gap_var]:
            var.trace_add("write", lambda *_: self.update_cluster_capacity())
        self.cluster_var.trace_add("write", lambda *_: self.update_cluster_capacity())
        preview_vars = [
            self.glass_width_var,
            self.glass_height_var,
            self.indent_var,
            self.matte_total_var,
            self.cluster_gap_var,
            self.bed_width_var,
            self.cluster_var,
            self.rotate_bottom_var,
            self.flip_var,
        ]
        for var in preview_vars:
            var.trace_add("write", lambda *_: self.schedule_preview_update())

    def _on_cluster_spin(self) -> None:
        self.update_cluster_capacity()
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
        except ValueError:
            self.cluster_info_var.set("Enter numeric values to compute capacity")
            self.schedule_preview_update()
            return
        capacity = calculate_capacity_from_values(glass_width, bed_width, cluster_gap)
        if capacity < 1:
            self.cluster_spin.configure(to=1)
            self.cluster_var.set(1)
            self.cluster_info_var.set("No clusters fit current bed width")
            self.schedule_preview_update()
            return
        self.cluster_spin.configure(to=capacity)
        current = self.cluster_var.get()
        if current < 1:
            current = 1
        if current > capacity:
            current = capacity
        if current != self.cluster_var.get():
            self.cluster_var.set(current)
        frames_selected = current * 4
        info = f"Clusters: {current}/{capacity}  •  Frames ready: {frames_selected}"
        self.cluster_info_var.set(info)
        self.schedule_preview_update()

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

        if glass_width <= 0 or glass_height <= 0 or bed_width_in <= 0:
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
        cluster_height = glass_height * 2
        if cluster_height <= 0:
            message = "Preview unavailable: check glass height and matte values."
            self._render_preview_message(
                layout_canvas,
                layout_width,
                layout_height,
                "Invalid height values",
            )
            self._render_preview_message(
                reference_canvas,
                reference_width,
                reference_height,
                "Invalid height values",
            )
            self.preview_message_var.set(message)
            return

        matte = calculate_matte_geometry(glass_width, glass_height, indent, matte_total)
        bed_width_mm = bed_width_in * 25.4
        cluster_width = glass_width * 2
        gap = max(cluster_gap, 0.0)
        total_clusters_width = cluster_count * cluster_width + max(cluster_count - 1, 0) * gap

        logo_width_mm = logo_height_mm = 0.0
        has_logo = False
        if self.logo_asset is not None:
            logo_width_mm = pt_to_mm(self.logo_asset.width_pt)
            logo_height_mm = pt_to_mm(self.logo_asset.height_pt)
            has_logo = logo_width_mm > 0 and logo_height_mm > 0

        arrangement_data = {
            "glass_width_mm": glass_width,
            "glass_height_mm": glass_height,
            "indent_mm": max(indent, 0.0),
            "cluster_count": cluster_count,
            "cluster_gap_mm": gap,
            "bed_width_mm": bed_width_mm,
            "total_clusters_width_mm": total_clusters_width,
            "cluster_height_mm": cluster_height,
            "logo_width_mm": logo_width_mm,
            "logo_height_mm": logo_height_mm,
            "has_logo": has_logo,
            "rotate_bottom": self.rotate_bottom_var.get(),
            "flip_in_app": self.flip_var.get(),
            "visible_band_mm": matte.visible_band_mm,
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

        if bed_width_mm > 0:
            usage_ratio = total_clusters_width / bed_width_mm
            usage_pct = usage_ratio * 100.0
            prefix = "⚠️ " if usage_ratio > 1.001 else ""
            usage_text = (
                f"{prefix}Width usage: {total_clusters_width:.1f} / {bed_width_mm:.1f} mm"
                f" ({usage_pct:.0f}%)"
            )
        else:
            usage_text = f"Total width: {total_clusters_width:.1f} mm"

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
        self.preview_message_var.set(
            f"{usage_text}\n{matte_text}\n{borders_text}\n{rotation_text} • {flip_text}\n{logo_text}"
        )

    def _draw_arrangement_preview(
        self,
        canvas: tk.Canvas,
        width: int,
        height: int,
        data: dict[str, float | bool],
        matte: MatteGeometry,
    ) -> None:
        margin = 16
        usable_width = max(width - margin * 2, 1)
        usable_height = max(height - margin * 2, 1)
        content_width = max(float(data["bed_width_mm"]), float(data["total_clusters_width_mm"]), 1.0)
        scale_x = usable_width / content_width
        scale_y = usable_height / float(data["cluster_height_mm"])
        scale = min(scale_x, scale_y)

        bed_left = margin
        bed_bottom = height - margin
        bed_right = bed_left + float(data["bed_width_mm"]) * scale
        bed_top = bed_bottom - float(data["cluster_height_mm"]) * scale

        canvas.create_rectangle(
            bed_left,
            bed_top,
            bed_right,
            bed_bottom,
            outline="#4c5c78",
            width=1.4,
        )

        seam_y = bed_bottom - float(data["glass_height_mm"]) * scale
        arrangement_right = bed_left + float(data["total_clusters_width_mm"]) * scale
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

        for cluster_index in range(int(data["cluster_count"])):
            base_x = cluster_index * (float(data["glass_width_mm"]) * 2 + float(data["cluster_gap_mm"]))
            for row in range(2):
                for column in range(2):
                    frame_x = base_x + column * float(data["glass_width_mm"])
                    frame_y = row * float(data["glass_height_mm"])
                    left = bed_left + frame_x * scale
                    right = left + float(data["glass_width_mm"]) * scale
                    top = bed_bottom - (frame_y + float(data["glass_height_mm"])) * scale
                    bottom = bed_bottom - frame_y * scale
                    canvas.create_rectangle(
                        left,
                        top,
                        right,
                        bottom,
                        fill=frame_fills[row],
                        outline="#7a879a",
                        width=1,
                    )

                    opening_left_mm = frame_x + matte.side_margin_mm
                    opening_right_mm = frame_x + float(data["glass_width_mm"]) - matte.side_margin_mm
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
                            width=1,
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
                        top_seam = bed_bottom - (
                            frame_y + float(data["glass_height_mm"]) - matte.top_margin_mm
                        ) * scale
                        canvas.create_line(
                            opening_left,
                            top_seam,
                            opening_right,
                            top_seam,
                            fill="#c6812f",
                            dash=(4, 2),
                        )

                    if has_logo and visible_band_mm > 0 and logo_width_mm > 0 and logo_height_mm > 0:
                        center_x_mm = frame_x + float(data["glass_width_mm"]) / 2
                        if row == 1:
                            center_y_mm = frame_y + visible_band_mm / 2
                        else:
                            center_y_mm = frame_y + float(data["glass_height_mm"]) - visible_band_mm / 2
                        min_center = frame_y + indent_mm + logo_height_mm / 2
                        max_center = frame_y + float(data["glass_height_mm"]) - indent_mm - logo_height_mm / 2
                        if min_center > max_center:
                            min_center = frame_y + logo_height_mm / 2
                            max_center = frame_y + float(data["glass_height_mm"]) - logo_height_mm / 2
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
        outline_thickness = parse_positive_float(
            self.outline_thickness_var.get(), "Outline thickness"
        )
        cluster_count = self.cluster_var.get()
        if cluster_count < 1:
            raise ValidationError("Cluster count", "must be at least 1")
        capacity = calculate_capacity_from_values(glass_width, bed_width, cluster_gap)
        if capacity < 1:
            raise ValidationError("Cluster count", "does not fit on the bed with current settings")
        if cluster_count > capacity:
            raise ValidationError(
                "Cluster count",
                f"cannot exceed {capacity} for the current bed width",
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
            outline_color=outline_color,
            outline_thickness_mm=outline_thickness,
            cluster_count=cluster_count,
            flip_in_app=self.flip_var.get(),
            allow_raster_fallback=self.raster_fallback_var.get(),
            rotate_bottom=self.rotate_bottom_var.get(),
            output_directory=output_directory,
            outline_filename=outline_filename,
            artwork_filename=artwork_filename,
        )

    def _generate_pdfs(self) -> None:
        if self.logo_asset is None:
            messagebox.showwarning("Missing logo", "Please load an EPS logo before generating PDFs.")
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
            builder = PDFBuilder(params, self.logo_asset)
            builder.generate_pdfs(self._progress_callback)
            self.set_status(
                f"Export complete: {params.outline_filename} and {params.artwork_filename}"
            )
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

