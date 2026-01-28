"""
Convertisseur de fichiers texte en PDF.

Utilise ReportLab pour créer des PDF propres avec
police monospace et pagination.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus

# Import conditionnel de ReportLab
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Preformatted
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

if TYPE_CHECKING:
    from ..config import Config
    from ..logger import ConverterLogger


class TextConverter(BaseConverter):
    """
    Convertisseur de fichiers texte (.txt, .log) en PDF.

    Utilise ReportLab avec une police monospace pour
    préserver le formatage du texte.
    """

    name = "text"
    supported_extensions = [".txt", ".log"]

    def is_available(self) -> bool:
        """Vérifie que ReportLab est installé."""
        return REPORTLAB_AVAILABLE

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """Convertit un fichier texte en PDF."""
        start = time.time()

        if not self.is_available():
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="ReportLab non installé (pip install reportlab)",
            )

        try:
            # Lire le contenu
            self.logger.debug(f"Lecture du fichier texte: {source.name}")
            try:
                content = source.read_text(encoding="utf-8", errors="replace")
            except Exception:
                # Fallback encodage
                content = source.read_text(encoding="latin-1", errors="replace")

            # Créer les styles
            styles = getSampleStyleSheet()
            mono_style = ParagraphStyle(
                "MonoStyle",
                parent=styles["Normal"],
                fontName="Courier",
                fontSize=9,
                leading=11,
            )

            # Créer le document
            doc = SimpleDocTemplate(
                str(dest),
                pagesize=A4,
                rightMargin=36,
                leftMargin=36,
                topMargin=36,
                bottomMargin=36,
            )

            story = []

            # Titre (nom du fichier)
            title_style = styles["Heading2"]
            story.append(Paragraph(source.name, title_style))
            story.append(Spacer(1, 12))

            # Contenu avec Preformatted (préserve les espaces et retours à la ligne)
            story.append(Preformatted(content, mono_style))

            # Générer le PDF
            doc.build(story)

            if not dest.exists():
                return ConversionResult(
                    status=ConversionStatus.FAILED,
                    source=source,
                    dest=None,
                    duration=time.time() - start,
                    method=self.name,
                    message="PDF non créé",
                )

            self.logger.debug("Conversion texte réussie")
            return ConversionResult(
                status=ConversionStatus.SUCCESS,
                source=source,
                dest=dest,
                duration=time.time() - start,
                method=self.name,
            )

        except Exception as e:
            self.logger.error(f"Erreur conversion texte: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )
