"""Widget de sélection de fichier/dossier."""

import customtkinter as ctk
from tkinter import filedialog
from pathlib import Path


class FileBrowserWidget(ctk.CTkFrame):
    """Champ de saisie de chemin avec bouton Parcourir."""

    def __init__(
        self,
        parent,
        label: str = "Dossier:",
        mode: str = "directory",
        **kwargs,
    ):
        kwargs.setdefault("fg_color", "transparent")
        super().__init__(parent, **kwargs)
        self.mode = mode

        self.columnconfigure(1, weight=1)

        self.label = ctk.CTkLabel(
            self, text=label, anchor="e", width=130,
            font=ctk.CTkFont(size=13),
        )
        self.label.grid(row=0, column=0, padx=(0, 8), sticky="e")

        self.entry = ctk.CTkEntry(
            self, height=32,
            placeholder_text="Cliquez sur Parcourir...",
            font=ctk.CTkFont(size=13),
        )
        self.entry.grid(row=0, column=1, sticky="ew")

        self.browse_btn = ctk.CTkButton(
            self, text="Parcourir", width=90, height=32,
            font=ctk.CTkFont(size=12),
            command=self._browse,
        )
        self.browse_btn.grid(row=0, column=2, padx=(6, 0))

    def _browse(self):
        if self.mode == "directory":
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename()
        if path:
            self.set_path(path)

    def get_path(self) -> Path | None:
        text = self.entry.get().strip()
        return Path(text) if text else None

    def set_path(self, path: Path | str) -> None:
        self.entry.delete(0, "end")
        self.entry.insert(0, str(path))

    def set_state(self, state: str) -> None:
        self.entry.configure(state=state)
        self.browse_btn.configure(state=state)
