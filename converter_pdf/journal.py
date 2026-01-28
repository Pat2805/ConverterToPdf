"""
Journal CSV d'audit pour les conversions.

Enregistre toutes les conversions (succès, échecs, skips)
dans un fichier CSV pour traçabilité.
"""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .config import Config
    from .logger import ConverterLogger
    from .converters.base import ConversionResult


class Journal:
    """
    Journal CSV pour tracer les conversions.

    Le journal est créé dans le répertoire de traitement avec
    un nom horodaté: conversion_log_YYYYMMDD_HHMMSS.csv

    Colonnes:
    - timestamp: Date/heure de l'opération
    - status: success, failed, skipped_password, skipped_exists, etc.
    - filetype: Extension du fichier source
    - source: Chemin du fichier source
    - dest_pdf: Chemin du PDF généré (si succès)
    - duration_s: Durée de la conversion en secondes
    - method: Convertisseur utilisé
    - source_size_mb: Taille du fichier source en MB
    - dest_size_mb: Taille du PDF en MB
    - message: Message d'erreur ou d'info
    - exception: Détails de l'exception si erreur
    """

    COLUMNS = [
        "timestamp",
        "status",
        "filetype",
        "source",
        "dest_pdf",
        "duration_s",
        "method",
        "source_size_mb",
        "dest_size_mb",
        "message",
        "exception",
    ]

    def __init__(
        self,
        config: "Config",
        logger: "ConverterLogger",
        output_dir: Path | None = None,
    ):
        """
        Initialise le journal.

        Args:
            config: Configuration
            logger: Logger
            output_dir: Répertoire de sortie (par défaut: répertoire courant)
        """
        self.config = config
        self.logger = logger
        self.output_dir = output_dir or Path.cwd()
        self._file_handle = None
        self._writer = None
        self._path: Path | None = None

    def open(self) -> None:
        """Ouvre le fichier journal."""
        if not self.config.journal_enabled:
            return

        if self._file_handle is not None:
            return  # Déjà ouvert

        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self._path = self.output_dir / f"conversion_log_{timestamp}.csv"

            self._file_handle = open(
                self._path,
                "w",
                newline="",
                encoding="utf-8",
            )
            self._writer = csv.writer(self._file_handle)
            self._writer.writerow(self.COLUMNS)

            self.logger.info(f"Journal ouvert: {self._path}")

        except Exception as e:
            self.logger.error(f"Impossible de créer le journal: {e}", exc=e)
            self._file_handle = None
            self._writer = None
            self._path = None

    def close(self) -> None:
        """Ferme le fichier journal."""
        if self._file_handle is not None:
            try:
                self._file_handle.flush()
                self._file_handle.close()
                self.logger.debug(f"Journal fermé: {self._path}")
            except Exception as e:
                self.logger.error(f"Erreur fermeture journal: {e}")
            finally:
                self._file_handle = None
                self._writer = None

    def log(self, result: "ConversionResult") -> None:
        """
        Enregistre un résultat de conversion.

        Args:
            result: Résultat de la conversion
        """
        if not self.config.journal_enabled:
            return

        # Ouvrir si pas encore fait
        if self._writer is None:
            self.open()
            if self._writer is None:
                return  # Impossible d'ouvrir

        # Filtrer selon la config (erreurs seulement ou tout)
        if self.config.journal_errors_only:
            if result.is_success or result.status.value == "skipped_exists":
                return

        try:
            # Formater l'exception
            exception_str = ""
            if result.exception:
                try:
                    import traceback
                    exception_str = "".join(
                        traceback.format_exception(
                            type(result.exception),
                            result.exception,
                            result.exception.__traceback__,
                        )
                    ).strip()
                except Exception:
                    exception_str = str(result.exception)

            row = [
                datetime.now().isoformat(timespec="seconds"),
                result.status.value,
                result.source.suffix.lower().lstrip("."),
                str(result.source),
                str(result.dest) if result.dest else "",
                f"{result.duration:.3f}",
                result.method,
                f"{result.source_size_mb:.3f}",
                f"{result.dest_size_mb:.3f}",
                result.message,
                exception_str,
            ]

            self._writer.writerow(row)
            self._file_handle.flush()  # Flush immédiat pour éviter les pertes

        except Exception as e:
            self.logger.error(f"Erreur écriture journal: {e}")

    @property
    def path(self) -> Path | None:
        """Chemin du fichier journal."""
        return self._path

    def __enter__(self) -> "Journal":
        """Context manager: ouvre le journal."""
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager: ferme le journal."""
        self.close()
