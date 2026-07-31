from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
import platform
import tkinter as tk
from tkinter import font as tkfont, ttk


class ThemeMode(StrEnum):
    LIGHT = "light"
    DARK = "dark"


@dataclass(frozen=True, slots=True)
class Palette:
    background: str
    surface: str
    surface_alt: str
    text: str
    muted_text: str
    border: str
    accent: str
    accent_text: str
    danger: str
    selection: str
    tooltip_background: str
    tooltip_text: str


LIGHT_PALETTE = Palette(
    background="#f5f6f8",
    surface="#ffffff",
    surface_alt="#eef1f5",
    text="#20242a",
    muted_text="#5f6874",
    border="#c8ced6",
    accent="#0b63ce",
    accent_text="#ffffff",
    danger="#b42318",
    selection="#d9eaff",
    tooltip_background="#fffbe6",
    tooltip_text="#20242a",
)

DARK_PALETTE = Palette(
    background="#20242a",
    surface="#2a2f36",
    surface_alt="#343a43",
    text="#f3f5f7",
    muted_text="#bdc5cf",
    border="#505965",
    accent="#64a9ff",
    accent_text="#101820",
    danger="#ff8a80",
    selection="#244f7a",
    tooltip_background="#343028",
    tooltip_text="#f3f5f7",
)


@dataclass(frozen=True, slots=True)
class Typography:
    family: str
    body_size: int = 10
    small_size: int = 9
    title_size: int = 20


@dataclass(frozen=True, slots=True)
class Spacing:
    xxs: int = 2
    xs: int = 4
    sm: int = 8
    md: int = 12
    lg: int = 18
    xl: int = 24


@dataclass(frozen=True, slots=True)
class AppearanceProfile:
    system_name: str
    theme_mode: ThemeMode
    palette: Palette
    typography: Typography
    spacing: Spacing
    scale: float
    minimum_target_px: int = 28


def font_family_for(system_name: str) -> str:
    if system_name == "Darwin":
        return "Helvetica Neue"
    if system_name == "Windows":
        return "Segoe UI"
    return "DejaVu Sans"


def clamp_scale(value: float) -> float:
    return max(1.0, min(2.0, round(float(value), 2)))


def scale_from_dpi(dpi: float) -> float:
    return clamp_scale(float(dpi) / 96.0)


def build_profile(
    *,
    system_name: str | None = None,
    theme_mode: ThemeMode = ThemeMode.LIGHT,
    dpi: float = 96.0,
) -> AppearanceProfile:
    system = system_name or platform.system()
    palette = DARK_PALETTE if theme_mode == ThemeMode.DARK else LIGHT_PALETTE
    return AppearanceProfile(
        system_name=system,
        theme_mode=theme_mode,
        palette=palette,
        typography=Typography(font_family_for(system)),
        spacing=Spacing(),
        scale=scale_from_dpi(dpi),
    )


def enable_windows_dpi_awareness(system_name: str | None = None) -> bool:
    """Best-effort DPI awareness. Safe no-op outside Windows and under tests."""
    if (system_name or platform.system()) != "Windows":
        return False
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
        return True
    except Exception:
        return False


def detected_dpi(root: tk.Misc) -> float:
    try:
        return float(root.winfo_fpixels("1i"))
    except (tk.TclError, TypeError, ValueError):
        return 96.0


def apply_tk_scaling(root: tk.Misc, scale: float) -> None:
    try:
        root.tk.call("tk", "scaling", clamp_scale(scale))
    except tk.TclError:
        pass


def _configure_named_fonts(root: tk.Misc, profile: AppearanceProfile) -> None:
    family = profile.typography.family
    size = profile.typography.body_size
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        try:
            named = tkfont.nametofont(name, root=root)
            named.configure(family=family, size=size)
        except tk.TclError:
            continue
    try:
        tkfont.nametofont("TkFixedFont", root=root).configure(size=size)
    except tk.TclError:
        pass


def install_appearance(root: tk.Misc, profile: AppearanceProfile | None = None) -> AppearanceProfile:
    actual = profile or build_profile(dpi=detected_dpi(root))
    apply_tk_scaling(root, actual.scale)
    _configure_named_fonts(root, actual)

    palette = actual.palette
    family = actual.typography.family
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure(".", font=(family, actual.typography.body_size))
    style.configure("TFrame", background=palette.background)
    style.configure("TLabel", background=palette.background, foreground=palette.text)
    style.configure("TLabelframe", background=palette.background, bordercolor=palette.border)
    style.configure("TLabelframe.Label", background=palette.background, foreground=palette.text)
    style.configure("TButton", padding=(10, 6))
    style.configure("Accent.TButton", font=(family, actual.typography.body_size, "bold"), padding=(12, 7))
    style.map(
        "Accent.TButton",
        background=[("active", palette.accent), ("!disabled", palette.accent)],
        foreground=[("!disabled", palette.accent_text)],
    )
    style.configure("Danger.TButton", foreground=palette.danger)
    style.configure("Title.TLabel", font=(family, actual.typography.title_size, "bold"))
    style.configure("Subtitle.TLabel", font=(family, actual.typography.body_size), foreground=palette.muted_text)
    style.configure("Hint.TLabel", font=(family, actual.typography.small_size), foreground=palette.muted_text)
    style.configure("StatusPrimary.TLabel", font=(family, actual.typography.body_size, "bold"))
    style.configure("Step.TLabel", font=(family, actual.typography.body_size, "bold"))
    style.configure("Treeview", rowheight=max(actual.minimum_target_px, int(24 * actual.scale)))
    style.configure("Treeview.Heading", font=(family, actual.typography.body_size, "bold"))
    style.map(
        "Treeview",
        background=[("selected", palette.selection)],
        foreground=[("selected", palette.text)],
    )
    try:
        root.configure(background=palette.background)
    except (tk.TclError, AttributeError):
        pass
    return actual


def accessible_widget_contract(*, minimum_target_px: int = 28) -> dict[str, object]:
    return {
        "minimum_target_px": int(minimum_target_px),
        "keyboard_focus_visible": True,
        "text_labels_required": True,
        "color_only_state_forbidden": True,
        "tooltip_delay_ms": 450,
    }
