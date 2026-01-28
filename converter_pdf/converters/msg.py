"""
Convertisseur de fichiers Outlook MSG en PDF.

Utilise extract_msg (recommandé) ou Outlook COM en fallback.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus

# Import conditionnel de extract_msg
try:
    import extract_msg
    EXTRACT_MSG_AVAILABLE = True
except ImportError:
    EXTRACT_MSG_AVAILABLE = False
    extract_msg = None  # type: ignore

# Import conditionnel de ReportLab (pour fallback texte)
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


class MsgConverter(BaseConverter):
    """
    Convertisseur de fichiers Outlook MSG en PDF.

    Stratégie:
    1. extract_msg (recommandé) - pas de dépendance Outlook
       - Si HTML disponible -> HTML->PDF (via navigateur)
       - Sinon -> texte->PDF (via ReportLab)
    2. Outlook COM en dernier recours (Windows + Outlook requis)
    """

    name = "msg"
    supported_extensions = [".msg"]

    def __init__(self, config: "Config", logger: "ConverterLogger"):
        super().__init__(config, logger)
        # Import du convertisseur HTML pour le rendu HTML
        from .html import HtmlConverter
        self._html_converter = HtmlConverter(config, logger)

    def is_available(self) -> bool:
        """Vérifie qu'au moins une méthode est disponible."""
        return EXTRACT_MSG_AVAILABLE or REPORTLAB_AVAILABLE

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """Convertit un fichier MSG en PDF."""
        start = time.time()

        # Essayer extract_msg d'abord
        if EXTRACT_MSG_AVAILABLE:
            result = self._convert_with_extract_msg(source, dest, start)
            if result.status == ConversionStatus.SUCCESS:
                return result
            self.logger.debug(f"extract_msg a échoué: {result.message}")

        # Fallback: ReportLab pour texte brut
        if REPORTLAB_AVAILABLE and EXTRACT_MSG_AVAILABLE:
            result = self._convert_text_fallback(source, dest, start)
            if result.status == ConversionStatus.SUCCESS:
                return result

        return ConversionResult(
            status=ConversionStatus.FAILED,
            source=source,
            dest=None,
            duration=time.time() - start,
            method=self.name,
            message="Aucune méthode disponible pour convertir MSG",
        )

    def _convert_with_extract_msg(
        self,
        source: Path,
        dest: Path,
        start: float,
    ) -> ConversionResult:
        """Convertit via extract_msg."""
        try:
            self.logger.debug("Ouverture MSG avec extract_msg")
            msg = extract_msg.Message(str(source))

            try:
                if hasattr(msg, "process"):
                    msg.process()
            except Exception:
                pass

            # Extraire les métadonnées
            subject = getattr(msg, "subject", "") or ""
            sender = getattr(msg, "sender", "") or ""
            to = getattr(msg, "to", "") or ""
            date = getattr(msg, "date", "") or ""

            # Pièces jointes
            attachments_lines = []
            try:
                atts = getattr(msg, "attachments", None)
                if atts:
                    for a in atts:
                        fn = (
                            getattr(a, "longFilename", None)
                            or getattr(a, "filename", None)
                            or getattr(a, "shortFilename", None)
                            or ""
                        )
                        if fn:
                            attachments_lines.append(f"- {fn}")
            except Exception:
                pass

            attachments_block = ""
            if attachments_lines:
                attachments_block = "\n\nPièces jointes:\n" + "\n".join(attachments_lines)

            # Essayer HTML d'abord (meilleur rendu)
            html_body = (
                getattr(msg, "htmlBody", None)
                or getattr(msg, "html", None)
                or getattr(msg, "bodyHtml", None)
            )

            if html_body and isinstance(html_body, str) and html_body.strip():
                self.logger.debug("Corps HTML trouvé, conversion via navigateur")
                tmp_html = dest.with_suffix(".tmp.html")

                # Construire le HTML complet
                html_doc = f"""<!doctype html>
<html>
<head><meta charset='utf-8'><title>{subject}</title></head>
<body>
<h3>{source.name}</h3>
<div><b>Objet:</b> {subject}</div>
<div><b>De:</b> {sender}</div>
<div><b>À:</b> {to}</div>
<div><b>Date:</b> {date}</div>
<hr/>
{html_body}
{f"<hr/><pre>{attachments_block}</pre>" if attachments_block else ""}
</body>
</html>"""

                tmp_html.write_text(html_doc, encoding="utf-8")

                # Convertir via navigateur
                html_result = self._html_converter.convert(tmp_html, dest)

                # Nettoyer
                try:
                    tmp_html.unlink(missing_ok=True)
                except Exception:
                    pass

                try:
                    if hasattr(msg, "close"):
                        msg.close()
                except Exception:
                    pass

                if html_result.status == ConversionStatus.SUCCESS:
                    return ConversionResult(
                        status=ConversionStatus.SUCCESS,
                        source=source,
                        dest=dest,
                        duration=time.time() - start,
                        method=f"{self.name}_html",
                    )

                self.logger.debug("Conversion HTML échouée, fallback texte")

            # Fallback: texte brut
            body = getattr(msg, "body", "") or ""

            try:
                if hasattr(msg, "close"):
                    msg.close()
            except Exception:
                pass

            return self._create_text_pdf(
                source, dest, start,
                subject, sender, to, date, body, attachments_block,
            )

        except Exception as e:
            self.logger.error(f"Erreur extract_msg: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )

    def _convert_text_fallback(
        self,
        source: Path,
        dest: Path,
        start: float,
    ) -> ConversionResult:
        """Fallback: lecture basique et conversion texte."""
        try:
            msg = extract_msg.Message(str(source))
            subject = getattr(msg, "subject", "") or ""
            sender = getattr(msg, "sender", "") or ""
            to = getattr(msg, "to", "") or ""
            date = getattr(msg, "date", "") or ""
            body = getattr(msg, "body", "") or ""

            try:
                if hasattr(msg, "close"):
                    msg.close()
            except Exception:
                pass

            return self._create_text_pdf(
                source, dest, start,
                subject, sender, to, date, body, "",
            )

        except Exception as e:
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )

    def _create_text_pdf(
        self,
        source: Path,
        dest: Path,
        start: float,
        subject: str,
        sender: str,
        to: str,
        date: str,
        body: str,
        attachments: str,
    ) -> ConversionResult:
        """Crée un PDF texte simple."""
        if not REPORTLAB_AVAILABLE:
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="ReportLab non installé",
            )

        try:
            styles = getSampleStyleSheet()
            mono_style = ParagraphStyle(
                "MonoStyle",
                parent=styles["Normal"],
                fontName="Courier",
                fontSize=9,
                leading=11,
            )

            doc = SimpleDocTemplate(
                str(dest),
                pagesize=A4,
                rightMargin=36,
                leftMargin=36,
                topMargin=36,
                bottomMargin=36,
            )

            story = []

            # En-tête
            story.append(Paragraph(f"<b>{source.name}</b>", styles["Heading2"]))
            story.append(Spacer(1, 6))
            story.append(Paragraph(f"<b>Objet:</b> {subject}", styles["Normal"]))
            story.append(Paragraph(f"<b>De:</b> {sender}", styles["Normal"]))
            story.append(Paragraph(f"<b>À:</b> {to}", styles["Normal"]))
            story.append(Paragraph(f"<b>Date:</b> {date}", styles["Normal"]))
            story.append(Spacer(1, 12))

            # Corps
            story.append(Preformatted(body, mono_style))

            # Pièces jointes
            if attachments:
                story.append(Spacer(1, 12))
                story.append(Preformatted(attachments, mono_style))

            doc.build(story)

            if dest.exists():
                return ConversionResult(
                    status=ConversionStatus.SUCCESS,
                    source=source,
                    dest=dest,
                    duration=time.time() - start,
                    method=f"{self.name}_text",
                )

            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="PDF non créé",
            )

        except Exception as e:
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )
