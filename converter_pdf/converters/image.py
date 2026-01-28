"""
Convertisseur d'images en PDF.

Utilise PIL/Pillow pour convertir les images en PDF.
Support optionnel de l'OCR pour créer des PDF recherchables.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus

# Import conditionnel de PIL
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    Image = None  # type: ignore

if TYPE_CHECKING:
    from ..config import Config
    from ..logger import ConverterLogger


class ImageConverter(BaseConverter):
    """
    Convertisseur d'images en PDF via PIL/Pillow.

    Supporte tous les formats d'image courants.
    L'OCR peut être activé pour créer des PDF recherchables.
    """

    name = "image"
    supported_extensions = [
        ".jpg", ".jpeg",
        ".png",
        ".bmp",
        ".tiff", ".tif",
        ".webp",
        ".gif",
    ]

    def is_available(self) -> bool:
        """Vérifie que PIL est installé."""
        return PIL_AVAILABLE

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """Convertit une image en PDF."""
        start = time.time()

        if not self.is_available():
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="Pillow non installé (pip install Pillow)",
            )

        try:
            self.logger.debug(f"Ouverture image: {source.name}")

            with Image.open(source) as img:
                # Convertir en RGB si nécessaire (pour éviter les erreurs avec RGBA, etc.)
                if img.mode in ("RGBA", "LA", "P"):
                    self.logger.debug(f"Conversion mode {img.mode} -> RGB")
                    # Créer un fond blanc pour les images avec transparence
                    background = Image.new("RGB", img.size, (255, 255, 255))
                    if img.mode == "P":
                        img = img.convert("RGBA")
                    background.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
                    img = background
                elif img.mode != "RGB":
                    self.logger.debug(f"Conversion mode {img.mode} -> RGB")
                    img = img.convert("RGB")

                # Sauvegarder en PDF
                img.save(
                    str(dest),
                    "PDF",
                    resolution=100.0,
                    quality=95,
                )

            if not dest.exists():
                return ConversionResult(
                    status=ConversionStatus.FAILED,
                    source=source,
                    dest=None,
                    duration=time.time() - start,
                    method=self.name,
                    message="PDF non créé",
                )

            self.logger.debug("Conversion image réussie")
            return ConversionResult(
                status=ConversionStatus.SUCCESS,
                source=source,
                dest=dest,
                duration=time.time() - start,
                method=self.name,
            )

        except Exception as e:
            self.logger.error(f"Erreur conversion image: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )
