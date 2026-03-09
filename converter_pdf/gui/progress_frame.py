"""Frame de progression et logs en temps réel."""

import customtkinter as ctk
from .widgets.log_viewer import LogViewer

_SECTION_FG = ("gray92", "gray14")


class ProgressFrame(ctk.CTkFrame):
    """Barre de progression, fichier courant et logs."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, corner_radius=8, fg_color=_SECTION_FG, **kwargs)

        # ── En-tête progression ──
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(10, 4))

        self.status_label = ctk.CTkLabel(
            header, text="En attente...", anchor="w",
            font=ctk.CTkFont(size=13),
        )
        self.status_label.pack(side="left", fill="x", expand=True)

        self.counter_label = ctk.CTkLabel(
            header, text="0 / 0", width=110, anchor="e",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.counter_label.pack(side="right")

        # ── Barre de progression ──
        self.progress_bar = ctk.CTkProgressBar(
            self, height=14, corner_radius=6,
            progress_color=("#2ea043", "#2ea043"),
        )
        self.progress_bar.pack(fill="x", padx=12, pady=(0, 8))
        self.progress_bar.set(0)

        # ── Séparateur visuel ──
        ctk.CTkLabel(
            self, text="  Logs", anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("gray50", "gray60"),
        ).pack(anchor="w", padx=10, pady=(2, 2))

        # ── Zone de logs ──
        self.log_viewer = LogViewer(self)
        self.log_viewer.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def update_progress(self, current: int, total: int) -> None:
        if total > 0:
            progress = min(current / total, 1.0)
            self.progress_bar.set(progress)
            pct = int(progress * 100)
            self.counter_label.configure(text=f"{current} / {total}  ({pct}%)")
        else:
            self.progress_bar.set(0)
            self.counter_label.configure(text="0 / 0")

    def set_current_file(self, filename: str) -> None:
        name = filename.split("\\")[-1].split("/")[-1] if filename else ""
        self.status_label.configure(text=f"Conversion : {name}")

    def append_log(self, level: str, message: str) -> None:
        self.log_viewer.append_log(level, message)

    def reset(self) -> None:
        self.progress_bar.set(0)
        self.counter_label.configure(text="0 / 0")
        self.status_label.configure(text="En attente...")
        self.log_viewer.clear()
