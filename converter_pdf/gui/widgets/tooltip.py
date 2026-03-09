"""Tooltip au survol pour customtkinter."""

from __future__ import annotations

import tkinter as tk
import customtkinter as ctk


class Tooltip:
    """
    Tooltip qui apparaît au survol d'un widget.

    Usage:
        btn = ctk.CTkButton(parent, text="OK")
        Tooltip(btn, "Valide l'action en cours")
    """

    DELAY_MS = 400
    BG_LIGHT = "#333333"
    BG_DARK = "#2b2b2b"
    FG = "#f0f0f0"

    def __init__(self, widget: tk.Widget | ctk.CTkBaseClass, text: str):
        self.widget = widget
        self.text = text
        self._tip_window: tk.Toplevel | None = None
        self._after_id: str | None = None

        self.widget.bind("<Enter>", self._on_enter, add="+")
        self.widget.bind("<Leave>", self._on_leave, add="+")
        self.widget.bind("<ButtonPress>", self._on_leave, add="+")

    def _on_enter(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self.DELAY_MS, self._show)

    def _on_leave(self, _event=None):
        self._cancel()
        self._hide()

    def _cancel(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self._tip_window:
            return

        # Position : juste en dessous du widget
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4

        self._tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)

        # Frame avec bordure arrondie simulée
        frame = tk.Frame(tw, bg=self.BG_LIGHT, bd=1, relief="solid")
        frame.pack()

        label = tk.Label(
            frame,
            text=self.text,
            bg=self.BG_LIGHT,
            fg=self.FG,
            font=("Segoe UI", 9),
            padx=8,
            pady=4,
            wraplength=320,
            justify="left",
        )
        label.pack()

    def _hide(self):
        if self._tip_window:
            self._tip_window.destroy()
            self._tip_window = None
