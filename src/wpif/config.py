from __future__ import annotations

import json
import shutil
from copy import deepcopy
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


APP_DIR_NAME = "word-paste-image-fit"
DEFAULT_CONFIG_NAME = "default_presets.json"


def project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def bundled_default_config_path() -> Path:
    return project_root() / "config" / DEFAULT_CONFIG_NAME


def user_config_dir() -> Path:
    base = Path.home() / "AppData" / "Roaming" / APP_DIR_NAME
    base.mkdir(parents=True, exist_ok=True)
    return base


def user_config_path() -> Path:
    return user_config_dir() / "settings.json"


@dataclass
class FitStyle:
    id: str
    name_zh: str
    name_en: str
    max_width_mode: str = "percent_of_content"  # percent_of_content | fixed_cm
    max_width_value: float = 92.0
    max_height_mode: str = "percent_of_page"  # percent_of_page | fixed_cm
    max_height_value: float = 75.0
    allow_upscale: bool = False
    wrap: str = "inline"
    align: str = "center"  # left | center | right
    fit_table_cell: bool = True

    def display_name(self, lang: str) -> str:
        return self.name_zh if lang == "zh" else self.name_en


@dataclass
class PagePreview:
    paper: str = "A4"
    margin_cm: dict[str, float] = field(
        default_factory=lambda: {
            "top": 2.54,
            "bottom": 2.54,
            "left": 3.18,
            "right": 3.18,
        }
    )


@dataclass
class AppConfig:
    version: int = 1
    language: str = "zh"
    hijack_ctrl_v: bool = False
    smart_paste_hotkey: str = "ctrl+alt+v"
    active_preset: str = "notes"
    presets: list[FitStyle] = field(default_factory=list)
    page_preview: PagePreview = field(default_factory=PagePreview)

    def active_style(self) -> FitStyle:
        for preset in self.presets:
            if preset.id == self.active_preset:
                return preset
        if self.presets:
            return self.presets[0]
        return FitStyle(id="notes", name_zh="笔记", name_en="Notes")


def _style_from_dict(data: dict[str, Any]) -> FitStyle:
    return FitStyle(
        id=data["id"],
        name_zh=data.get("name_zh", data["id"]),
        name_en=data.get("name_en", data["id"]),
        max_width_mode=data.get("max_width_mode", "percent_of_content"),
        max_width_value=float(data.get("max_width_value", 92.0)),
        max_height_mode=data.get("max_height_mode", "percent_of_page"),
        max_height_value=float(data.get("max_height_value", 75.0)),
        allow_upscale=bool(data.get("allow_upscale", False)),
        wrap=data.get("wrap", "inline"),
        align=data.get("align", "center"),
        fit_table_cell=bool(data.get("fit_table_cell", True)),
    )


def config_from_dict(data: dict[str, Any]) -> AppConfig:
    presets = [_style_from_dict(item) for item in data.get("presets", [])]
    page = data.get("page_preview", {})
    return AppConfig(
        version=int(data.get("version", 1)),
        language=data.get("language", "zh"),
        hijack_ctrl_v=bool(data.get("hijack_ctrl_v", False)),
        smart_paste_hotkey=data.get("smart_paste_hotkey", "ctrl+alt+v"),
        active_preset=data.get("active_preset", presets[0].id if presets else "notes"),
        presets=presets,
        page_preview=PagePreview(
            paper=page.get("paper", "A4"),
            margin_cm=page.get(
                "margin_cm",
                {"top": 2.54, "bottom": 2.54, "left": 3.18, "right": 3.18},
            ),
        ),
    )


def config_to_dict(cfg: AppConfig) -> dict[str, Any]:
    return {
        "version": cfg.version,
        "language": cfg.language,
        "hijack_ctrl_v": cfg.hijack_ctrl_v,
        "smart_paste_hotkey": cfg.smart_paste_hotkey,
        "active_preset": cfg.active_preset,
        "presets": [asdict(p) for p in cfg.presets],
        "page_preview": {
            "paper": cfg.page_preview.paper,
            "margin_cm": cfg.page_preview.margin_cm,
        },
    }


def load_config() -> AppConfig:
    path = user_config_path()
    if not path.exists():
        shutil.copy2(bundled_default_config_path(), path)
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    return config_from_dict(data)


def save_config(cfg: AppConfig) -> Path:
    path = user_config_path()
    with path.open("w", encoding="utf-8") as fh:
        json.dump(config_to_dict(cfg), fh, ensure_ascii=False, indent=2)
    return path


def update_active_preset(cfg: AppConfig, **changes: Any) -> AppConfig:
    cfg = deepcopy(cfg)
    style = cfg.active_style()
    for key, value in changes.items():
        if hasattr(style, key):
            setattr(style, key, value)
    for i, preset in enumerate(cfg.presets):
        if preset.id == style.id:
            cfg.presets[i] = style
            break
    return cfg
