"""
Convertisseur HTML en PDF.

Utilise Chrome ou Edge en mode headless pour un rendu fidèle
des pages HTML avec CSS, images, etc.
"""

from __future__ import annotations

import shutil
import subprocess
import time
from pathlib import Path
from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus

if TYPE_CHECKING:
    from ..config import Config
    from ..logger import ConverterLogger


class HtmlConverter(BaseConverter):
    """
    Convertisseur HTML via navigateur headless (Chrome/Edge).

    Utilise la fonctionnalité print-to-pdf des navigateurs Chromium
    pour un rendu fidèle des pages HTML.
    """

    name = "html_browser"
    supported_extensions = [".htm", ".html"]

    def __init__(self, config: "Config", logger: "ConverterLogger"):
        super().__init__(config, logger)
        self._browser_path: Path | None = None

    def _detect_browser(self) -> Path | None:
        """Détecte Chrome ou Edge."""
        # Utiliser le chemin configuré si disponible
        if self.config.browser_path and self.config.browser_path.exists():
            return self.config.browser_path

        # Chercher dans PATH
        for exe in ("chrome", "chrome.exe", "msedge", "msedge.exe", "google-chrome"):
            path = shutil.which(exe)
            if path:
                return Path(path)

        # Chemins Windows courants
        windows_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        ]

        for path_str in windows_paths:
            path = Path(path_str)
            if path.exists():
                return path

        return None

    @property
    def browser_path(self) -> Path | None:
        """Chemin vers le navigateur (détecté au premier accès)."""
        if self._browser_path is None:
            self._browser_path = self._detect_browser()
        return self._browser_path

    def is_available(self) -> bool:
        """Vérifie qu'un navigateur est disponible."""
        return self.browser_path is not None

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """Convertit un fichier HTML en PDF via navigateur headless."""
        start = time.time()

        if not self.is_available():
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="Aucun navigateur (Chrome/Edge) détecté",
            )

        try:
            # Convertir en URI file://
            source_uri = source.absolute().as_uri()

            # Créer un profil temporaire pour éviter les conflits
            tmp_profile = dest.parent / f".tmp_browser_{int(time.time() * 1000)}"
            tmp_profile.mkdir(parents=True, exist_ok=True)

            # Commande navigateur headless
            cmd = [
                str(self.browser_path),
                "--headless=new",
                "--disable-gpu",
                "--no-first-run",
                "--no-default-browser-check",
                f"--user-data-dir={tmp_profile}",
                f"--print-to-pdf={dest.absolute()}",
                "--print-to-pdf-no-header",
                source_uri,
            ]

            self.logger.debug(f"Commande navigateur: {cmd[0]} --headless ...")

            # Exécuter avec timeout
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=self.config.browser_timeout,
            )

            # Nettoyer le profil temporaire
            try:
                shutil.rmtree(tmp_profile, ignore_errors=True)
            except Exception:
                pass

            if result.returncode != 0:
                self.logger.error(f"Navigateur stderr: {result.stderr[:300]}")
                return ConversionResult(
                    status=ConversionStatus.FAILED,
                    source=source,
                    dest=None,
                    duration=time.time() - start,
                    method=self.name,
                    message=f"Erreur navigateur: {result.stderr[:200]}",
                )

            if not dest.exists() or dest.stat().st_size == 0:
                return ConversionResult(
                    status=ConversionStatus.FAILED,
                    source=source,
                    dest=None,
                    duration=time.time() - start,
                    method=self.name,
                    message="PDF non créé ou vide",
                )

            self.logger.debug("Conversion HTML réussie")
            return ConversionResult(
                status=ConversionStatus.SUCCESS,
                source=source,
                dest=dest,
                duration=time.time() - start,
                method=self.name,
            )

        except subprocess.TimeoutExpired:
            self.logger.error(f"Timeout navigateur ({self.config.browser_timeout}s)")
            # Nettoyer le profil temporaire
            try:
                shutil.rmtree(tmp_profile, ignore_errors=True)
            except Exception:
                pass
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message=f"Timeout après {self.config.browser_timeout}s",
            )

        except Exception as e:
            self.logger.error(f"Erreur HTML->PDF: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )
