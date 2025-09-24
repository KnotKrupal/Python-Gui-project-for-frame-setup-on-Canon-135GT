from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

import fitz

from .colors import DEFAULT_COLOR_NAME, PRESET_COLORS
from .ghostscript import GhostscriptError, convert_eps_to_pdf_bytes
from .geometry import calculate_capacity_from_values
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
        self.status_var = tk.StringVar()
        self.progress_var = tk.DoubleVar(value=0.0)

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

    def _on_cluster_spin(self) -> None:
        self.update_cluster_capacity()

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
            return
        capacity = calculate_capacity_from_values(glass_width, bed_width, cluster_gap)
        if capacity < 1:
            self.cluster_spin.configure(to=1)
            self.cluster_var.set(1)
            self.cluster_info_var.set("No clusters fit current bed width")
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

