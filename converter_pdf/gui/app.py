"""Application principale GUI pour ConverterToPdf."""

from __future__ import annotations

import io
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path

from .. import __version__
from ..config import Config
from ..signals import SignalHub
from ..gui_bridge import ConversionBridge
from .main_frame import MainFrame
from .progress_frame import ProgressFrame
from .report_frame import ReportFrame


class ConverterApp(ctk.CTk):
    """Fenêtre principale de l'application ConverterToPdf."""

    def __init__(self):
        super().__init__()

        # ── Thème ──
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # ── Fenêtre ──
        self.title(f"ConverterToPdf  v{__version__}")
        self.geometry("960x780")
        self.minsize(780, 620)

        # ── État ──
        self.config = Config.load()
        self.signals = SignalHub()
        self.bridge: ConversionBridge | None = None
        self._poll_id: str | None = None

        # ── Construction ──
        self._build_ui()

    # ──────────────────────────────────────────
    #  Layout
    # ──────────────────────────────────────────

    def _build_ui(self) -> None:
        # Titre
        title_frame = ctk.CTkFrame(self, fg_color="transparent", height=36)
        title_frame.pack(fill="x", padx=16, pady=(12, 0))
        ctk.CTkLabel(
            title_frame,
            text=f"ConverterToPdf",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            title_frame,
            text=f"v{__version__}",
            font=ctk.CTkFont(size=13),
            text_color=("gray50", "gray60"),
        ).pack(side="left", padx=(6, 0), pady=(4, 0))

        # Panel config (scrollable)
        self.main_frame = MainFrame(self, self.config)
        self.main_frame.pack(fill="x", padx=14, pady=(8, 4))

        # Progression + logs (prend tout l'espace restant)
        self.progress_frame = ProgressFrame(self)
        self.progress_frame.pack(fill="both", expand=True, padx=14, pady=4)

        # Rapport
        self.report_frame = ReportFrame(self)
        self.report_frame.pack(fill="x", padx=14, pady=(4, 12))

        # ── Callbacks ──
        self.main_frame.on_convert = self._start_conversion
        self.main_frame.on_cancel = self._cancel_conversion
        self.main_frame.on_check = self._run_check
        self.main_frame.on_load_config = self._load_config
        self.main_frame.on_save_config = self._save_config

    # ──────────────────────────────────────────
    #  Conversion
    # ──────────────────────────────────────────

    def _start_conversion(self) -> None:
        source = self.main_frame.get_source_path()
        if not source or not source.exists():
            messagebox.showerror(
                "Erreur",
                "Veuillez sélectionner un dossier ou fichier source valide.",
            )
            return

        config = self.main_frame.build_config()
        dest = self.main_frame.get_output_path()

        self.progress_frame.reset()
        self.report_frame.reset()
        self.main_frame.set_running(True)

        self.bridge = ConversionBridge(config, self.signals)
        self.bridge.start_conversion(source, dest)
        self._poll_queue()

    def _poll_queue(self) -> None:
        if not self.bridge:
            return
        for msg in self.bridge.poll_queue():
            self._handle_message(msg)
        if self.bridge.is_running:
            self._poll_id = self.after(100, self._poll_queue)
        else:
            self.main_frame.set_running(False)
            self._poll_id = None

    def _handle_message(self, msg: tuple) -> None:
        kind = msg[0]
        if kind == "log":
            self.progress_frame.append_log(msg[1], msg[2])
        elif kind == "progress":
            self.progress_frame.update_progress(msg[1], msg[2])
        elif kind == "file_start":
            self.progress_frame.set_current_file(msg[1])
        elif kind == "session_complete":
            self.report_frame.display(msg[1], msg[2])
            self.progress_frame.status_label.configure(text="Conversion terminée !")
        elif kind == "error":
            self.progress_frame.append_log("ERROR", f"Erreur : {msg[1]}")
            messagebox.showerror("Erreur", str(msg[1]))
        elif kind == "done":
            self.main_frame.set_running(False)

    def _cancel_conversion(self) -> None:
        if self.bridge:
            self.bridge.cancel()
            self.progress_frame.append_log("WARNING", "Annulation demandée...")

    # ──────────────────────────────────────────
    #  Check / Config
    # ──────────────────────────────────────────

    def _run_check(self) -> None:
        from ..cli import print_check_info
        old = sys.stdout
        sys.stdout = buf = io.StringIO()
        try:
            print_check_info()
        finally:
            sys.stdout = old

        win = ctk.CTkToplevel(self)
        win.title("Vérification de la configuration")
        win.geometry("680x520")
        win.transient(self)
        tb = ctk.CTkTextbox(win, wrap="none", font=ctk.CTkFont(family="Consolas", size=12))
        tb.pack(fill="both", expand=True, padx=12, pady=12)
        tb.insert("1.0", buf.getvalue())
        tb.configure(state="disabled")

    def _load_config(self) -> None:
        path = filedialog.askopenfilename(
            title="Charger une configuration",
            filetypes=[("Config", "*.converterrc"), ("YAML", "*.yaml *.yml"), ("Tous", "*.*")],
        )
        if path:
            try:
                cfg = Config.load(path)
                self.config = cfg
                self.main_frame.load_from_config(cfg)
                self.progress_frame.append_log("INFO", f"Configuration chargée : {path}")
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de charger la config :\n{e}")

    def _save_config(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Sauvegarder la configuration",
            defaultextension=".converterrc",
            filetypes=[("Config", "*.converterrc"), ("YAML", "*.yaml *.yml")],
        )
        if path:
            try:
                cfg = self.main_frame.build_config()
                cfg.save(path)
                self.progress_frame.append_log("INFO", f"Configuration sauvegardée : {path}")
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de sauvegarder :\n{e}")
