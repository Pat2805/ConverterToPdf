"""Frame d'affichage du rapport post-conversion."""

import os
import sys
import customtkinter as ctk
from pathlib import Path
from typing import Any

_SECTION_FG = ("gray92", "gray14")


class ReportFrame(ctk.CTkFrame):
    """Barre de résumé avec statistiques et accès au rapport."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, corner_radius=8, fg_color=_SECTION_FG, **kwargs)
        self._report = None
        self._report_path: Path | None = None

        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="x", padx=12, pady=8)
        inner.columnconfigure(0, weight=1)

        # Résumé textuel
        self.summary_label = ctk.CTkLabel(
            inner,
            text="Aucune conversion effectuée",
            font=ctk.CTkFont(size=13),
            anchor="w",
        )
        self.summary_label.grid(row=0, column=0, sticky="ew")

        # Boutons
        btn_frame = ctk.CTkFrame(inner, fg_color="transparent")
        btn_frame.grid(row=0, column=1, padx=(10, 0))

        self.open_btn = ctk.CTkButton(
            btn_frame, text="Ouvrir fichier", width=110, height=28,
            font=ctk.CTkFont(size=12), state="disabled",
            command=self._open_report_file,
        )
        self.open_btn.pack(side="left", padx=(0, 4))

        self.view_btn = ctk.CTkButton(
            btn_frame, text="Voir rapport", width=110, height=28,
            font=ctk.CTkFont(size=12), state="disabled",
            command=self._view_report,
        )
        self.view_btn.pack(side="left")

    def display(self, report: Any, stats: dict) -> None:
        """Affiche le résumé de conversion."""
        self._report = report
        s = stats.get("success", 0)
        f = stats.get("failed", 0)
        k = stats.get("skipped", 0)
        t = stats.get("total", 0)

        self.summary_label.configure(
            text=f"Fichiers : {t}   |   Succès : {s}   |   Échecs : {f}   |   Ignorés : {k}"
        )
        self.view_btn.configure(state="normal")

        if report and hasattr(report, "output_directory") and report.output_directory:
            self._find_report_file(report.output_directory)
        elif report and hasattr(report, "source_directory") and report.source_directory:
            self._find_report_file(report.source_directory)

    def _find_report_file(self, directory: Path) -> None:
        try:
            files = sorted(directory.glob("conversion_report_*.txt"), reverse=True)
            if files:
                self._report_path = files[0]
                self.open_btn.configure(state="normal")
        except Exception:
            pass

    def _view_report(self) -> None:
        if not self._report:
            return
        win = ctk.CTkToplevel(self)
        win.title("Rapport de conversion")
        win.geometry("820x620")
        win.transient(self.winfo_toplevel())

        tb = ctk.CTkTextbox(
            win, wrap="none",
            font=ctk.CTkFont(family="Consolas", size=12),
        )
        tb.pack(fill="both", expand=True, padx=12, pady=12)
        text = self._report.generate() if hasattr(self._report, "generate") else str(self._report)
        tb.insert("1.0", text)
        tb.configure(state="disabled")

    def _open_report_file(self) -> None:
        if self._report_path and self._report_path.exists():
            if sys.platform == "win32":
                os.startfile(self._report_path)
            elif sys.platform == "darwin":
                os.system(f'open "{self._report_path}"')
            else:
                os.system(f'xdg-open "{self._report_path}"')

    def reset(self) -> None:
        self._report = None
        self._report_path = None
        self.summary_label.configure(text="Aucune conversion effectuée")
        self.view_btn.configure(state="disabled")
        self.open_btn.configure(state="disabled")
