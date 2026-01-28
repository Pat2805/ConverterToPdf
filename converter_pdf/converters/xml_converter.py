"""
Convertisseur de fichiers XML en PDF.

Utilise ReportLab pour créer des PDF avec
mise en forme du XML (coloration syntaxique basique).
"""

from __future__ import annotations

import time
import xml.dom.minidom
from pathlib import Path
from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus

# Import conditionnel de ReportLab
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Preformatted
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

if TYPE_CHECKING:
    from ..config import Config
    from ..logger import ConverterLogger


class XmlConverter(BaseConverter):
    """
    Convertisseur de fichiers XML en PDF.

    Utilise ReportLab avec formatage du XML pour
    une meilleure lisibilité.
    """

    name = "xml"
    supported_extensions = [".xml"]

    def is_available(self) -> bool:
        """Vérifie que ReportLab est installé."""
        return REPORTLAB_AVAILABLE

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """Convertit un fichier XML en PDF."""
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
            # Lire le contenu XML
            self.logger.debug(f"Lecture du fichier XML: {source.name}")
            try:
                xml_content = source.read_text(encoding="utf-8")
            except Exception:
                xml_content = source.read_text(encoding="latin-1", errors="replace")

            # Essayer de formater le XML pour une meilleure lisibilité
            try:
                dom = xml.dom.minidom.parseString(xml_content)
                xml_formatted = dom.toprettyxml(indent="  ")
                # Supprimer les lignes vides
                lines = [line for line in xml_formatted.split("\n") if line.strip()]
                xml_formatted = "\n".join(lines)
            except Exception:
                # Si le parsing échoue, garder le contenu original
                xml_formatted = xml_content

            # Créer les styles
            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                "XmlTitle",
                parent=styles["Heading1"],
                fontSize=16,
                spaceAfter=30,
                textColor=colors.darkblue,
            )

            code_style = ParagraphStyle(
                "XmlCode",
                parent=styles["Code"],
                fontSize=8,
                leftIndent=20,
                fontName="Courier",
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

            # Titre
            title = f"Fichier XML: {source.name}"
            story.append(Paragraph(title, title_style))
            story.append(Spacer(1, 12))

            # Contenu XML
            # Échapper les caractères spéciaux pour ReportLab
            for line in xml_formatted.split("\n"):
                if line.strip():
                    line_escaped = (
                        line.replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                    )
                    story.append(Preformatted(line_escaped, code_style))

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

            self.logger.debug("Conversion XML réussie")
            return ConversionResult(
                status=ConversionStatus.SUCCESS,
                source=source,
                dest=dest,
                duration=time.time() - start,
                method=self.name,
            )

        except Exception as e:
            self.logger.error(f"Erreur conversion XML: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )
