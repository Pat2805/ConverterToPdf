"""Panneau principal de configuration (compact, sans scroll)."""

import customtkinter as ctk
from tkinter import filedialog
from pathlib import Path
from typing import Callable

from ..config import Config
from .widgets.file_browser import FileBrowserWidget
from .widgets.tooltip import Tooltip


# Mapping des filtres de format
FILTER_MAP = {
    "Tous les formats": None,
    "Images seulement": [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"],
    "Word seulement": [".doc", ".docx", ".rtf", ".odt"],
    "Excel seulement": [".xls", ".xlsx", ".xlsm", ".xlsb"],
    "XML seulement": [".xml"],
}

# Couleurs
_SEC = ("gray92", "gray14")
_GREEN = "#2ea043"
_GREEN_H = "#238636"
_RED = "#da3633"
_RED_H = "#b62324"
_SUBTLE = ("gray78", "gray35")

_FONT = lambda sz=13, **kw: ctk.CTkFont(size=sz, **kw)


class MainFrame(ctk.CTkFrame):
    """Panneau de configuration compact — tout visible sans scroller."""

    def __init__(self, parent, config: Config, **kwargs):
        kwargs.setdefault("fg_color", "transparent")
        super().__init__(parent, **kwargs)
        self.config = config

        # Callbacks (branchés par app.py)
        self.on_convert: Callable | None = None
        self.on_cancel: Callable | None = None
        self.on_check: Callable | None = None
        self.on_load_config: Callable | None = None
        self.on_save_config: Callable | None = None

        self._build_ui()
        self.load_from_config(config)

    # ─── Construction ─────────────────────────────

    def _build_ui(self):
        # ── Section Chemins ──
        sec1 = ctk.CTkFrame(self, corner_radius=8, fg_color=_SEC)
        sec1.pack(fill="x", pady=(0, 4))
        self._sec_title(sec1, "Chemins")
        inner1 = ctk.CTkFrame(sec1, fg_color="transparent")
        inner1.pack(fill="x", padx=10, pady=(0, 8))

        self.source_browser = FileBrowserWidget(inner1, label="Source :", mode="directory")
        self.source_browser.pack(fill="x", pady=1)
        Tooltip(self.source_browser, "Dossier ou fichier à convertir en PDF")

        self.output_browser = FileBrowserWidget(inner1, label="Sortie (opt.) :", mode="directory")
        self.output_browser.pack(fill="x", pady=1)
        Tooltip(self.output_browser, "Dossier de destination des PDF (par défaut : même dossier que la source)")

        # ── Section Options ──
        sec2 = ctk.CTkFrame(self, corner_radius=8, fg_color=_SEC)
        sec2.pack(fill="x", pady=4)
        self._sec_title(sec2, "Options")
        grid2 = ctk.CTkFrame(sec2, fg_color="transparent")
        grid2.pack(fill="x", padx=12, pady=(0, 8))
        for c in range(4):
            grid2.columnconfigure(c, weight=1)

        self.recursive_var = ctk.BooleanVar()
        self.force_var = ctk.BooleanVar()
        self.dry_run_var = ctk.BooleanVar()
        self.report_var = ctk.BooleanVar(value=True)
        self.delete_var = ctk.BooleanVar()
        self.hide_var = ctk.BooleanVar()
        self.keep_ext_var = ctk.BooleanVar(value=True)

        OPTS = [
            ("Récursif",       self.recursive_var, None,
             "Parcourir aussi les sous-dossiers"),
            ("Forcer",         self.force_var, None,
             "Reconvertir même si le PDF existe déjà"),
            ("Dry-run",        self.dry_run_var, None,
             "Simuler les conversions sans créer de fichiers"),
            ("Rapport",        self.report_var, None,
             "Générer un rapport de session (conversion_report_*.txt)"),
            ("Suppr. source",  self.delete_var, self._on_delete_changed,
             "Supprimer le fichier source après conversion réussie"),
            ("Cacher source",  self.hide_var, self._on_hide_changed,
             "Rendre le fichier source caché (Windows uniquement)"),
            ("Garder ext.",    self.keep_ext_var, None,
             "Nommage : doc.docx → doc.docx.pdf (sinon doc.pdf)"),
        ]
        for i, (text, var, cmd, tip) in enumerate(OPTS):
            row, col = divmod(i, 4)
            cb = ctk.CTkCheckBox(grid2, text=text, variable=var,
                                 font=_FONT(12), command=cmd)
            cb.grid(row=row, column=col, sticky="w", padx=3, pady=3)
            Tooltip(cb, tip)

        # ── Section Paramètres (tout sur 2 lignes) ──
        sec3 = ctk.CTkFrame(self, corner_radius=8, fg_color=_SEC)
        sec3.pack(fill="x", pady=4)
        self._sec_title(sec3, "Paramètres")
        grid3 = ctk.CTkFrame(sec3, fg_color="transparent")
        grid3.pack(fill="x", padx=12, pady=(0, 8))
        grid3.columnconfigure(1, weight=1)
        grid3.columnconfigure(3, weight=1)
        grid3.columnconfigure(5, weight=1)

        lf = _FONT(12)

        # Ligne 1 : Méthode | Log level | Filtre
        r = 0
        self._lbl(grid3, "Méthode :", r, 0)
        self.method_combo = self._combo(
            grid3, ["auto", "office", "libreoffice", "reportlab"], 130, r, 1,
        )
        self.method_combo.set("auto")
        Tooltip(self.method_combo, "auto = essaie Office puis LibreOffice puis ReportLab")

        self._lbl(grid3, "Log :", r, 2)
        self.log_level_combo = self._combo(
            grid3, ["DEBUG", "INFO", "WARNING", "ERROR"], 100, r, 3,
        )
        self.log_level_combo.set("INFO")
        Tooltip(self.log_level_combo, "Niveau de verbosité des logs affichés")

        self._lbl(grid3, "Formats :", r, 4)
        self.filter_combo = self._combo(
            grid3, list(FILTER_MAP.keys()), 160, r, 5,
        )
        self.filter_combo.set("Tous les formats")
        Tooltip(self.filter_combo, "Filtrer les types de fichiers à convertir")

        # Ligne 2 : OCR
        r = 1
        self.ocr_var = ctk.BooleanVar()
        ocr_cb = ctk.CTkCheckBox(grid3, text="OCR", variable=self.ocr_var,
                                  font=lf, command=self._on_ocr_changed)
        ocr_cb.grid(row=r, column=0, columnspan=2, sticky="w", padx=3, pady=4)
        Tooltip(ocr_cb, "Activer la reconnaissance de texte (OCR) sur les images")

        self._lbl(grid3, "Moteur :", r, 2)
        self.ocr_engine_combo = self._combo(
            grid3, ["auto", "tesseract", "easyocr", "paddleocr"], 130, r, 3,
        )
        self.ocr_engine_combo.set("auto")
        self.ocr_engine_combo.configure(state="disabled")
        Tooltip(self.ocr_engine_combo, "Moteur OCR : auto détecte le meilleur disponible")

        # ── Boutons config + actions (une seule ligne) ──
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", pady=(6, 2))
        btn_row.columnconfigure(3, weight=1)  # espace flexible avant Convertir

        bf = _FONT(12)
        b1 = ctk.CTkButton(btn_row, text="Vérifier config", width=130, height=30,
                            font=bf, fg_color=_SUBTLE, hover_color=("gray68", "gray45"),
                            command=lambda: self.on_check and self.on_check())
        b1.grid(row=0, column=0, padx=(0, 4))
        Tooltip(b1, "Affiche les outils et dépendances détectés sur ce système")

        b2 = ctk.CTkButton(btn_row, text="Charger config", width=120, height=30,
                            font=bf, fg_color=_SUBTLE, hover_color=("gray68", "gray45"),
                            command=lambda: self.on_load_config and self.on_load_config())
        b2.grid(row=0, column=1, padx=(0, 4))
        Tooltip(b2, "Charger les options depuis un fichier .converterrc (YAML)")

        b3 = ctk.CTkButton(btn_row, text="Sauver config", width=120, height=30,
                            font=bf, fg_color=_SUBTLE, hover_color=("gray68", "gray45"),
                            command=lambda: self.on_save_config and self.on_save_config())
        b3.grid(row=0, column=2)
        Tooltip(b3, "Sauvegarder les options actuelles dans un fichier .converterrc")

        # Convertir + Annuler (à droite)
        self.convert_btn = ctk.CTkButton(
            btn_row, text="Convertir", width=160, height=36,
            font=_FONT(14, weight="bold"),
            fg_color=_GREEN, hover_color=_GREEN_H, corner_radius=8,
            command=lambda: self.on_convert and self.on_convert(),
        )
        self.convert_btn.grid(row=0, column=4, padx=(10, 4))
        Tooltip(self.convert_btn, "Lancer la conversion de tous les fichiers du dossier source")

        self.cancel_btn = ctk.CTkButton(
            btn_row, text="Annuler", width=100, height=36,
            font=_FONT(13), fg_color=_RED, hover_color=_RED_H,
            corner_radius=8, state="disabled",
            command=lambda: self.on_cancel and self.on_cancel(),
        )
        self.cancel_btn.grid(row=0, column=5)
        Tooltip(self.cancel_btn, "Interrompre la conversion en cours")

    # ─── Helpers layout ──────────────────────────

    @staticmethod
    def _sec_title(parent, text: str):
        ctk.CTkLabel(
            parent, text=f"  {text}",
            font=_FONT(12, weight="bold"), anchor="w",
        ).pack(anchor="w", padx=6, pady=(6, 2))

    @staticmethod
    def _lbl(parent, text, row, col):
        ctk.CTkLabel(parent, text=text, font=_FONT(12), anchor="e").grid(
            row=row, column=col, sticky="e", padx=(8, 4), pady=4,
        )

    @staticmethod
    def _combo(parent, values, width, row, col):
        cb = ctk.CTkComboBox(parent, values=values, width=width, height=28,
                              state="readonly", font=_FONT(12))
        cb.grid(row=row, column=col, sticky="w", pady=4)
        return cb

    # ─── Logique ─────────────────────────────────

    def _on_delete_changed(self):
        if self.delete_var.get():
            self.hide_var.set(False)

    def _on_hide_changed(self):
        if self.hide_var.get():
            self.delete_var.set(False)

    def _on_ocr_changed(self):
        self.ocr_engine_combo.configure(
            state="readonly" if self.ocr_var.get() else "disabled",
        )

    def build_config(self) -> Config:
        """Construit un Config depuis l'état des widgets."""
        return Config(
            method=self.method_combo.get(),
            keep_extension=self.keep_ext_var.get(),
            log_level=self.log_level_combo.get(),
            report_enabled=self.report_var.get(),
            ocr_enabled=self.ocr_var.get(),
            ocr_engine=self.ocr_engine_combo.get(),
            recursive=self.recursive_var.get(),
            force=self.force_var.get(),
            delete_source=self.delete_var.get(),
            hide_source=self.hide_var.get(),
            dry_run=self.dry_run_var.get(),
            extensions=FILTER_MAP.get(self.filter_combo.get()),
        )

    def load_from_config(self, config: Config) -> None:
        """Charge un Config dans les widgets."""
        self.recursive_var.set(config.recursive)
        self.force_var.set(config.force)
        self.delete_var.set(config.delete_source)
        self.hide_var.set(config.hide_source)
        self.dry_run_var.set(config.dry_run)
        self.keep_ext_var.set(config.keep_extension)
        self.report_var.set(config.report_enabled)
        self.method_combo.set(config.method)
        self.log_level_combo.set(config.log_level)
        self.ocr_var.set(config.ocr_enabled)
        self.ocr_engine_combo.set(config.ocr_engine)
        self._on_ocr_changed()
        if config.extensions:
            for label, exts in FILTER_MAP.items():
                if exts == config.extensions:
                    self.filter_combo.set(label)
                    break
        else:
            self.filter_combo.set("Tous les formats")

    def get_source_path(self) -> Path | None:
        return self.source_browser.get_path()

    def get_output_path(self) -> Path | None:
        return self.output_browser.get_path()

    def set_running(self, running: bool) -> None:
        """Bascule l'état de l'interface pendant la conversion."""
        if running:
            self.convert_btn.configure(state="disabled")
            self.cancel_btn.configure(state="normal")
            self.source_browser.set_state("disabled")
            self.output_browser.set_state("disabled")
        else:
            self.convert_btn.configure(state="normal")
            self.cancel_btn.configure(state="disabled")
            self.source_browser.set_state("normal")
            self.output_browser.set_state("normal")
