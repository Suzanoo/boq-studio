from __future__ import annotations

import tkinter as tk

from .appearance import AppearanceProfile, build_profile


class TooltipManager:
    """Lazy Tk tooltip manager: creates a popup only after a hover delay."""

    def __init__(self, owner: tk.Misc, *, delay_ms: int = 450, profile: AppearanceProfile | None = None) -> None:
        self.owner = owner
        self.delay_ms = delay_ms
        self.profile = profile or build_profile()
        self._after_id: str | None = None
        self._popup: tk.Toplevel | None = None

    def register(self, widget: tk.Misc, text: str) -> None:
        widget.bind("<Enter>", lambda _e, w=widget, t=text: self._schedule(w, t), add="+")
        widget.bind("<Leave>", lambda _e: self.hide(), add="+")
        widget.bind("<ButtonPress>", lambda _e: self.hide(), add="+")

    def _schedule(self, widget: tk.Misc, text: str) -> None:
        self.hide()
        self._after_id = self.owner.after(self.delay_ms, lambda: self._show(widget, text))

    def _show(self, widget: tk.Misc, text: str) -> None:
        self._after_id = None
        if not text or not widget.winfo_exists():
            return
        popup = tk.Toplevel(self.owner)
        popup.wm_overrideredirect(True)
        try:
            popup.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        x = widget.winfo_rootx() + 8
        y = widget.winfo_rooty() + widget.winfo_height() + 6
        popup.wm_geometry(f"+{x}+{y}")
        tk.Label(
            popup,
            text=text,
            justify="left",
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=5,
            background=self.profile.palette.tooltip_background,
            foreground=self.profile.palette.tooltip_text,
            font=(self.profile.typography.family, self.profile.typography.small_size),
        ).pack()
        self._popup = popup

    def hide(self) -> None:
        if self._after_id is not None:
            try:
                self.owner.after_cancel(self._after_id)
            except tk.TclError:
                pass
            self._after_id = None
        if self._popup is not None:
            try:
                self._popup.destroy()
            except tk.TclError:
                pass
            self._popup = None
