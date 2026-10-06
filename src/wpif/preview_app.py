from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

from PIL import Image, ImageTk

from .config import AppConfig, FitStyle, load_config, save_config, update_active_preset
from .geometry import (
    align_offset,
    fit_image,
    page_metrics_from_preview,
    pixels_to_points,
)
from .i18n import t
from .word_paste import clipboard_has_image, paste_and_fit
from .hotkeys import set_request_handler, start_hotkeys_background


class PreviewApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.cfg = load_config()
        self.sample_path = Path(__file__).resolve().parents[2] / "assets" / "sample_screenshot.png"
        if not self.sample_path.exists():
            raise FileNotFoundError(
                f"Missing sample image: {self.sample_path}. Run scripts/generate_sample_image.py"
            )
        self.sample = Image.open(self.sample_path).convert("RGBA")
        self._photo: ImageTk.PhotoImage | None = None
        self._building = False

        self.title(t(self.cfg.language, "app_title"))
        self.geometry("1180x720")
        self.minsize(980, 640)

        self._build_vars()
        self._build_ui()
        self._load_style_into_vars(self.cfg.active_style())
        self._refresh_preview()
        self._set_status(t(self.cfg.language, "status_ready"))
        # Hotkeys request paste; execute on the Tk main thread for reliable Word COM.
        set_request_handler(lambda source: self.after(0, lambda s=source: self._handle_hotkey(s)))
        start_hotkeys_background(on_status=lambda msg: self.after(0, lambda: self._set_status(msg)))

    def _build_vars(self) -> None:
        style = self.cfg.active_style()
        self.var_lang = tk.StringVar(value=self.cfg.language)
        self.var_preset = tk.StringVar(value=style.id)
        self.var_width_mode = tk.StringVar(value=style.max_width_mode)
        self.var_width_value = tk.DoubleVar(value=style.max_width_value)
        self.var_height_value = tk.DoubleVar(value=style.max_height_value)
        self.var_allow_upscale = tk.BooleanVar(value=style.allow_upscale)
        self.var_fit_cell = tk.BooleanVar(value=style.fit_table_cell)
        self.var_align = tk.StringVar(value=style.align)
        self.var_hijack = tk.BooleanVar(value=self.cfg.hijack_ctrl_v)
        self.var_demo_cell = tk.BooleanVar(value=False)
        self.var_status = tk.StringVar(value="")

    def _build_ui(self) -> None:
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        left = ttk.Frame(self, padding=12)
        left.grid(row=0, column=0, sticky="nsw")
        right = ttk.Frame(self, padding=12)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        self.lbl_lang = ttk.Label(left, text="")
        self.lbl_lang.grid(row=0, column=0, sticky="w")
        lang_box = ttk.Combobox(
            left,
            textvariable=self.var_lang,
            values=["zh", "en"],
            state="readonly",
            width=18,
        )
        lang_box.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        lang_box.bind("<<ComboboxSelected>>", self._on_language)

        self.lbl_presets = ttk.Label(left, text="")
        self.lbl_presets.grid(row=2, column=0, sticky="w")
        self.preset_box = ttk.Combobox(left, textvariable=self.var_preset, state="readonly", width=18)
        self.preset_box.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        self.preset_box.bind("<<ComboboxSelected>>", self._on_preset)

        self.lbl_width = ttk.Label(left, text="")
        self.lbl_width.grid(row=4, column=0, sticky="w")
        self.width_mode_box = ttk.Combobox(
            left,
            textvariable=self.var_width_mode,
            values=["percent_of_content", "fixed_cm"],
            state="readonly",
            width=18,
        )
        self.width_mode_box.grid(row=5, column=0, sticky="ew")
        self.width_mode_box.bind("<<ComboboxSelected>>", self._on_change)

        self.width_scale = ttk.Scale(
            left,
            from_=20,
            to=100,
            variable=self.var_width_value,
            command=lambda _v: self._on_change(),
        )
        self.width_scale.grid(row=6, column=0, sticky="ew", pady=(4, 0))
        self.lbl_width_value = ttk.Label(left, text="")
        self.lbl_width_value.grid(row=7, column=0, sticky="w", pady=(0, 10))

        self.lbl_height = ttk.Label(left, text="")
        self.lbl_height.grid(row=8, column=0, sticky="w")
        self.height_scale = ttk.Scale(
            left,
            from_=20,
            to=100,
            variable=self.var_height_value,
            command=lambda _v: self._on_change(),
        )
        self.height_scale.grid(row=9, column=0, sticky="ew", pady=(4, 0))
        self.lbl_height_value = ttk.Label(left, text="")
        self.lbl_height_value.grid(row=10, column=0, sticky="w", pady=(0, 10))

        self.chk_upscale = ttk.Checkbutton(
            left, variable=self.var_allow_upscale, command=self._on_change
        )
        self.chk_upscale.grid(row=11, column=0, sticky="w")
        self.chk_cell = ttk.Checkbutton(left, variable=self.var_fit_cell, command=self._on_change)
        self.chk_cell.grid(row=12, column=0, sticky="w")
        self.chk_demo_cell = ttk.Checkbutton(
            left, variable=self.var_demo_cell, command=self._on_change
        )
        self.chk_demo_cell.grid(row=13, column=0, sticky="w", pady=(0, 8))

        self.lbl_align = ttk.Label(left, text="")
        self.lbl_align.grid(row=14, column=0, sticky="w")
        align_row = ttk.Frame(left)
        align_row.grid(row=15, column=0, sticky="ew", pady=(0, 10))
        for i, key in enumerate(("left", "center", "right")):
            ttk.Radiobutton(
                align_row,
                text=key,
                value=key,
                variable=self.var_align,
                command=self._on_change,
            ).grid(row=0, column=i, padx=2)

        self.chk_hijack = ttk.Checkbutton(left, variable=self.var_hijack, command=self._on_change)
        self.chk_hijack.grid(row=16, column=0, sticky="w", pady=(0, 12))

        self.btn_save = ttk.Button(left, command=self._save)
        self.btn_save.grid(row=17, column=0, sticky="ew", pady=2)
        self.btn_paste = ttk.Button(left, command=self._paste_sample_or_clipboard)
        self.btn_paste.grid(row=18, column=0, sticky="ew", pady=2)

        self.lbl_hotkey = ttk.Label(left, wraplength=220, foreground="#555")
        self.lbl_hotkey.grid(row=19, column=0, sticky="w", pady=(12, 0))

        self.lbl_preview_hint = ttk.Label(right, text="")
        self.lbl_preview_hint.grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.canvas = tk.Canvas(right, background="#6b7280", highlightthickness=0)
        self.canvas.grid(row=1, column=0, sticky="nsew")
        self.canvas.bind("<Configure>", lambda _e: self._refresh_preview())

        status = ttk.Label(right, textvariable=self.var_status, anchor="w")
        status.grid(row=2, column=0, sticky="ew", pady=(8, 0))

        self._retranslate()

    def _retranslate(self) -> None:
        lang = self.var_lang.get()
        self.title(t(lang, "app_title"))
        self.lbl_lang.configure(text=t(lang, "language"))
        self.lbl_presets.configure(text=t(lang, "presets"))
        self.lbl_width.configure(text=t(lang, "max_width"))
        self.lbl_height.configure(text=t(lang, "max_height"))
        self.chk_upscale.configure(text=t(lang, "allow_upscale"))
        self.chk_cell.configure(text=t(lang, "fit_table_cell"))
        self.chk_demo_cell.configure(
            text=("预览：模拟粘贴进表格" if lang == "zh" else "Preview: simulate table cell")
        )
        self.lbl_align.configure(text=t(lang, "align"))
        self.chk_hijack.configure(text=t(lang, "hijack_ctrl_v"))
        self.btn_save.configure(text=t(lang, "save"))
        self.btn_paste.configure(text=t(lang, "smart_paste"))
        self.lbl_preview_hint.configure(text=t(lang, "preview_hint"))
        self.lbl_hotkey.configure(text=t(lang, "hotkey_hint"))

        self.preset_box.configure(values=[p.display_name(lang) for p in self.cfg.presets])
        # Keep combobox showing current preset label
        active = self.cfg.active_style()
        self.preset_box.set(active.display_name(lang))

        mode_labels = {
            "percent_of_content": t(lang, "width_mode_percent"),
            "fixed_cm": t(lang, "width_mode_fixed"),
        }
        # Keep machine values in the combobox; show friendly status elsewhere.
        self.width_mode_box.configure(values=["percent_of_content", "fixed_cm"])
        self._update_value_labels()

    def _update_value_labels(self) -> None:
        lang = self.var_lang.get()
        if self.var_width_mode.get() == "fixed_cm":
            self.width_scale.configure(from_=4, to=20)
            self.lbl_width_value.configure(
                text=f"{self.var_width_value.get():.1f} cm ({t(lang, 'width_mode_fixed')})"
            )
        else:
            self.width_scale.configure(from_=20, to=100)
            self.lbl_width_value.configure(
                text=f"{self.var_width_value.get():.0f}% ({t(lang, 'width_mode_percent')})"
            )
        self.lbl_height_value.configure(
            text=f"{self.var_height_value.get():.0f}% ({t(lang, 'height_mode_percent')})"
        )

    def _current_style_from_vars(self) -> FitStyle:
        style = self.cfg.active_style()
        return FitStyle(
            id=style.id,
            name_zh=style.name_zh,
            name_en=style.name_en,
            max_width_mode=self.var_width_mode.get(),
            max_width_value=float(self.var_width_value.get()),
            max_height_mode="percent_of_page",
            max_height_value=float(self.var_height_value.get()),
            allow_upscale=bool(self.var_allow_upscale.get()),
            wrap="inline",
            align=self.var_align.get(),
            fit_table_cell=bool(self.var_fit_cell.get()),
        )

    def _sync_cfg_from_vars(self) -> None:
        style = self._current_style_from_vars()
        self.cfg = update_active_preset(
            self.cfg,
            max_width_mode=style.max_width_mode,
            max_width_value=style.max_width_value,
            max_height_value=style.max_height_value,
            allow_upscale=style.allow_upscale,
            align=style.align,
            fit_table_cell=style.fit_table_cell,
        )
        self.cfg.language = self.var_lang.get()
        self.cfg.hijack_ctrl_v = bool(self.var_hijack.get())

    def _load_style_into_vars(self, style: FitStyle) -> None:
        self._building = True
        self.var_preset.set(style.id)
        self.var_width_mode.set(style.max_width_mode)
        self.var_width_value.set(style.max_width_value)
        self.var_height_value.set(style.max_height_value)
        self.var_allow_upscale.set(style.allow_upscale)
        self.var_fit_cell.set(style.fit_table_cell)
        self.var_align.set(style.align)
        self._building = False
        self._update_value_labels()

    def _on_language(self, _event=None) -> None:
        self.cfg.language = self.var_lang.get()
        self._retranslate()
        self._refresh_preview()

    def _on_preset(self, _event=None) -> None:
        lang = self.var_lang.get()
        label = self.preset_box.get()
        for preset in self.cfg.presets:
            if preset.display_name(lang) == label:
                self.cfg.active_preset = preset.id
                self._load_style_into_vars(preset)
                break
        self._refresh_preview()

    def _on_change(self, _event=None) -> None:
        if self._building:
            return
        self._sync_cfg_from_vars()
        # Persist hijack toggle immediately so the hotkey thread picks it up.
        save_config(self.cfg)
        self._update_value_labels()
        self._refresh_preview()

    def _set_status(self, text: str) -> None:
        self.var_status.set(text)

    def _refresh_preview(self) -> None:
        lang = self.var_lang.get()
        style = self._current_style_from_vars()
        metrics = page_metrics_from_preview(self.cfg.page_preview)

        canvas_w = max(self.canvas.winfo_width(), 400)
        canvas_h = max(self.canvas.winfo_height(), 400)
        scale = min((canvas_w - 40) / metrics.page_width, (canvas_h - 40) / metrics.page_height)
        page_w = metrics.page_width * scale
        page_h = metrics.page_height * scale
        origin_x = (canvas_w - page_w) / 2
        origin_y = (canvas_h - page_h) / 2

        cell_w = cell_h = None
        if self.var_demo_cell.get() and style.fit_table_cell:
            cell_w = metrics.content_width * 0.45
            cell_h = metrics.content_height * 0.35

        natural = pixels_to_points(self.sample.width, self.sample.height, dpi=96.0)
        fitted = fit_image(
            natural,
            style,
            content_width=metrics.content_width,
            content_height=metrics.content_height,
            page_height=metrics.page_height,
            cell_width=cell_w,
            cell_height=cell_h,
        )

        container_w = cell_w or metrics.content_width
        x_off = align_offset(style.align, container_w, fitted.width)

        self.canvas.delete("all")
        # Page
        self.canvas.create_rectangle(
            origin_x,
            origin_y,
            origin_x + page_w,
            origin_y + page_h,
            fill="#ffffff",
            outline="#111827",
            width=1,
        )
        # Content area
        cx = origin_x + metrics.margin_left * scale
        cy = origin_y + metrics.margin_top * scale
        cw = metrics.content_width * scale
        ch = metrics.content_height * scale
        self.canvas.create_rectangle(cx, cy, cx + cw, cy + ch, outline="#93c5fd", dash=(4, 3))
        self.canvas.create_text(
            cx + 4,
            cy + 4,
            anchor="nw",
            fill="#2563eb",
            text=t(lang, "content_box"),
            font=("Segoe UI", 9),
        )

        if cell_w and cell_h:
            cell_x = cx + 8 * scale
            cell_y = cy + 24 * scale
            self.canvas.create_rectangle(
                cell_x,
                cell_y,
                cell_x + cell_w * scale,
                cell_y + cell_h * scale,
                outline="#f59e0b",
                dash=(3, 2),
            )
            img_x = cell_x + x_off * scale
            img_y = cell_y + 4 * scale
        else:
            img_x = cx + x_off * scale
            img_y = cy + 18 * scale

        img_w = max(1, int(fitted.width * scale))
        img_h = max(1, int(fitted.height * scale))
        preview_img = self.sample.copy()
        preview_img.thumbnail((img_w, img_h))
        self._photo = ImageTk.PhotoImage(preview_img)
        self.canvas.create_image(img_x, img_y, anchor="nw", image=self._photo)
        self.canvas.create_rectangle(
            img_x,
            img_y,
            img_x + self._photo.width(),
            img_y + self._photo.height(),
            outline="#dc2626",
        )
        self.canvas.create_text(
            img_x,
            img_y + self._photo.height() + 4,
            anchor="nw",
            fill="#b91c1c",
            text=f"{t(lang, 'image_box')} · {fitted.scale * 100:.0f}% · {fitted.limited_by}",
            font=("Segoe UI", 9),
        )

        limit_zh = {"width": "受宽度限制", "height": "受高度限制", "both": "宽高同时限制", "none": "未缩放"}
        limit_en = {
            "width": "limited by width",
            "height": "limited by height",
            "both": "limited by both",
            "none": "no scaling",
        }
        limit_txt = limit_zh.get(fitted.limited_by, fitted.limited_by) if lang == "zh" else limit_en.get(
            fitted.limited_by, fitted.limited_by
        )
        self._set_status(
            f"{limit_txt} · {fitted.width / 72 * 2.54:.1f}×{fitted.height / 72 * 2.54:.1f} cm"
        )

    def _save(self) -> None:
        self._sync_cfg_from_vars()
        path = save_config(self.cfg)
        self._set_status(f"{t(self.cfg.language, 'saved')}: {path}")

    def _handle_hotkey(self, source: str) -> None:
        lang = self.var_lang.get()
        self._sync_cfg_from_vars()
        self._set_status(
            f"收到快捷键 {source}…" if lang == "zh" else f"Hotkey received: {source}…"
        )
        if not clipboard_has_image():
            self._set_status(
                f"{source}: 剪贴板没有图片" if lang == "zh" else f"{source}: clipboard has no image"
            )
            return
        outcome = paste_and_fit(self.cfg)
        if outcome.ok:
            self._set_status(
                f"{source} → {t(lang, 'paste_ok')} · "
                f"{outcome.width_pt / 72 * 2.54:.1f}×{outcome.height_pt / 72 * 2.54:.1f} cm"
            )
        else:
            self._set_status(f"{source} → {t(lang, 'paste_fail')}: {outcome.message}")
            messagebox.showerror(t(lang, "app_title"), f"{t(lang, 'paste_fail')}\n{outcome.message}")

    def _paste_sample_or_clipboard(self) -> None:
        self._sync_cfg_from_vars()
        # Prefer real clipboard image; fall back to sample so users can demo immediately.
        outcome = paste_and_fit(self.cfg)
        if not outcome.ok and outcome.message == "no_image":
            outcome = paste_and_fit(self.cfg, image=self.sample)
        lang = self.cfg.language
        if outcome.ok:
            self._set_status(
                f"{t(lang, 'paste_ok')} · {outcome.width_pt / 72 * 2.54:.1f}×{outcome.height_pt / 72 * 2.54:.1f} cm"
            )
        else:
            messagebox.showerror(t(lang, "app_title"), f"{t(lang, 'paste_fail')}\n{outcome.message}")


def run() -> None:
    app = PreviewApp()
    app.mainloop()
