"""
Convertisseur LibreOffice.

Utilise LibreOffice en mode headless pour convertir
les documents Office en PDF.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import time
from pathlib import Path
from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus

if TYPE_CHECKING:
    from ..config import Config
    from ..logger import ConverterLogger


class LibreOfficeConverter(BaseConverter):
    """
    Convertisseur via LibreOffice en mode headless.

    Supporte tous les formats Office (Word, Excel, PowerPoint)
    ainsi que d'autres formats comme ODT, ODS, etc.
    """

    name = "libreoffice"
    supported_extensions = [
        # Word
        ".doc", ".docx", ".rtf", ".odt",
        # Excel
        ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods",
        # PowerPoint
        ".ppt", ".pptx", ".odp",
    ]

    def __init__(self, config: "Config", logger: "ConverterLogger"):
        super().__init__(config, logger)
        self._libreoffice_path: Path | None = None

    def _detect_libreoffice(self) -> Path | None:
        """Détecte l'installation de LibreOffice."""
        # Utiliser le chemin configuré si disponible
        if self.config.libreoffice_path and self.config.libreoffice_path.exists():
            return self.config.libreoffice_path

        # Chercher dans PATH
        soffice = shutil.which("soffice")
        if soffice:
            return Path(soffice)

        # Chemins Windows courants
        windows_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            r"C:\Program Files\LibreOffice 7\program\soffice.exe",
            r"C:\Program Files\LibreOffice 24\program\soffice.exe",
        ]

        for path_str in windows_paths:
            path = Path(path_str)
            if path.exists():
                return path

        return None

    @property
    def libreoffice_path(self) -> Path | None:
        """Chemin vers soffice.exe (détecté au premier accès)."""
        if self._libreoffice_path is None:
            self._libreoffice_path = self._detect_libreoffice()
        return self._libreoffice_path

    def is_available(self) -> bool:
        """Vérifie que LibreOffice est installé."""
        return self.libreoffice_path is not None

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """Convertit un document via LibreOffice headless."""
        start = time.time()

        if not self.is_available():
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="LibreOffice non installé",
            )

        try:
            # LibreOffice crée le PDF dans le répertoire de sortie
            # avec le même nom que le source
            output_dir = dest.parent
            output_dir.mkdir(parents=True, exist_ok=True)

            # Commande LibreOffice
            cmd = [
                str(self.libreoffice_path),
                "--headless",
                "--convert-to", "pdf:writer_pdf_Export",
                "--outdir", str(output_dir),
                str(source.absolute()),
            ]

            self.logger.debug(f"Commande: {' '.join(cmd)}")

            # Exécuter avec timeout
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=self.config.libreoffice_timeout,
                env={**os.environ, "PYTHONIOENCODING": "utf-8"},
            )

            if result.returncode != 0:
                self.logger.error(f"LibreOffice stderr: {result.stderr}")
                return ConversionResult(
                    status=ConversionStatus.FAILED,
                    source=source,
                    dest=None,
                    duration=time.time() - start,
                    method=self.name,
                    message=f"LibreOffice erreur: {result.stderr[:200]}",
                )

            # LibreOffice crée le PDF avec le nom du source
            pdf_generated = output_dir / (source.stem + ".pdf")

            # Renommer si nécessaire
            if pdf_generated != dest:
                if dest.exists():
                    dest.unlink()
                pdf_generated.rename(dest)

            if not dest.exists():
                return ConversionResult(
                    status=ConversionStatus.FAILED,
                    source=source,
                    dest=None,
                    duration=time.time() - start,
                    method=self.name,
                    message="PDF non créé par LibreOffice",
                )

            self.logger.debug("Conversion LibreOffice réussie")
            return ConversionResult(
                status=ConversionStatus.SUCCESS,
                source=source,
                dest=dest,
                duration=time.time() - start,
                method=self.name,
            )

        except subprocess.TimeoutExpired:
            self.logger.error(f"Timeout LibreOffice ({self.config.libreoffice_timeout}s)")
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message=f"Timeout après {self.config.libreoffice_timeout}s",
            )

        except Exception as e:
            self.logger.error(f"Erreur LibreOffice: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )
