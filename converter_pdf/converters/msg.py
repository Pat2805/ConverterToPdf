"""
Convertisseur de fichiers Outlook MSG en PDF.

Gère les pièces jointes :
- Crée un dossier nom_message/ contenant le PDF du message + les pièces jointes
- Convertit les pièces jointes en PDF quand c'est possible
- Conserve les fichiers originaux non convertibles (images, etc.)
- Filtre les petites images insignifiantes (logos, signatures, pixels de tracking)
"""

from __future__ import annotations

import io
import re
import shutil
import time
from pathlib import Path
from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus

# Import conditionnel de PIL pour analyse des images
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    Image = None  # type: ignore

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
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
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
    1. Créer un dossier pour le message si pièces jointes
    2. Extraire et convertir les pièces jointes en PDF
    3. Convertir le message en PDF (HTML ou texte)

    Structure de sortie (avec pièces jointes):
        message.msg-open/
        ├── _message.pdf          # Le corps du message
        ├── document.docx.pdf     # Pièce jointe convertie
        ├── image.jpg             # Pièce jointe non convertible (conservée)
        └── ...

    Sans pièces jointes: message.msg.pdf (fichier simple)
    """

    name = "msg"
    supported_extensions = [".msg"]

    # Extensions convertibles en PDF
    CONVERTIBLE_EXTENSIONS = {
        ".doc", ".docx", ".rtf", ".odt",
        ".xls", ".xlsx", ".xlsm", ".xlsb",
        ".ppt", ".pptx",
        ".txt", ".log",
        ".htm", ".html",
        ".xml",
        ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp",
        # Archives (pièces jointes compressées)
        ".zip", ".rar", ".7z",
        ".tar", ".tar.gz", ".tgz", ".tar.bz2", ".tbz2",
    }

    # Extensions d'images (pour filtrage des petites images)
    IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp"}

    # Seuils pour détecter les images insignifiantes
    # Images en dessous de ces seuils avec nom suspect sont filtrées
    MIN_IMAGE_SIZE_BYTES = 30 * 1024  # 30 KB
    MIN_IMAGE_DIMENSION = 200  # 200 pixels (largeur ET hauteur)

    # Seuils absolus : images TOUJOURS filtrées (trop petites pour être utiles)
    # Même sans nom suspect, ces images sont des icônes/logos/trackers
    ALWAYS_FILTER_SIZE_BYTES = 15 * 1024  # 15 KB
    ALWAYS_FILTER_DIMENSION = 150  # 150 pixels
    ALWAYS_FILTER_SURFACE = 25000  # ~158x158 pixels (permet 250x100 mais pas 100x100)

    # Patterns de noms de fichiers à ignorer (logos, signatures, tracking pixels)
    # NOTE: Ces patterns ne filtrent que si l'image est PETITE
    INSIGNIFICANT_IMAGE_PATTERNS = [
        r"logo",                 # logo, company_logo, logo.png
        r"signature",            # signature, email_signature
        r"spacer",               # spacer.gif
        r"pixel",                # pixel, tracking_pixel
        r"tracking",             # tracking.gif
        r"^blank$",              # blank.gif
        r"^dot$",                # dot.gif
        r"^clear$",              # clear.gif
        r"^trans(parent)?$",     # trans.gif, transparent.png
        r"^1x1$",                # 1x1.gif
        r"^icon",                # icon, icon_email
        r"footer",               # footer_logo
        r"header",               # header_image (petits)
    ]

    def __init__(self, config: "Config", logger: "ConverterLogger"):
        super().__init__(config, logger)
        # Import du convertisseur HTML pour le rendu HTML
        from .html import HtmlConverter
        self._html_converter = HtmlConverter(config, logger)
        # Les autres convertisseurs seront importés à la demande
        self._converters_cache = None

    def _get_converters(self):
        """Retourne la chaîne de convertisseurs (lazy loading)."""
        if self._converters_cache is None:
            from . import get_converter_chain
            self._converters_cache = get_converter_chain(self.config, self.logger)
        return self._converters_cache

    def is_available(self) -> bool:
        """Vérifie qu'au moins une méthode est disponible."""
        return EXTRACT_MSG_AVAILABLE or REPORTLAB_AVAILABLE

    def _sanitize_filename(self, filename: str) -> str:
        """Nettoie un nom de fichier pour le système de fichiers."""
        # Remplacer les caractères interdits
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
        # Limiter la longueur
        if len(sanitized) > 200:
            sanitized = sanitized[:200]
        return sanitized.strip()

    def _get_extension_from_mime(self, mime_type: str) -> str:
        """Retourne l'extension de fichier correspondant à un type MIME."""
        if not mime_type:
            return ""

        mime_map = {
            # Images
            "image/jpeg": ".jpg",
            "image/png": ".png",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/tiff": ".tif",
            "image/webp": ".webp",
            # Documents
            "application/pdf": ".pdf",
            "application/msword": ".doc",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
            "application/vnd.ms-excel": ".xls",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
            "application/vnd.ms-powerpoint": ".ppt",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
            # Texte
            "text/plain": ".txt",
            "text/html": ".html",
            "text/xml": ".xml",
            # Archives
            "application/zip": ".zip",
            "application/x-rar-compressed": ".rar",
            "application/x-7z-compressed": ".7z",
            # Autres
            "application/octet-stream": "",  # Binaire générique, pas d'extension
            "message/rfc822": ".eml",
        }

        mime_lower = mime_type.lower().split(";")[0].strip()
        return mime_map.get(mime_lower, "")

    def _is_insignificant_image(
        self,
        filename: str,
        data: bytes | None,
        attachment: object,
    ) -> tuple[bool, str]:
        """
        Détecte si une image est insignifiante (logo, signature, pixel de tracking).

        Une image n'est filtrée que si elle est PETITE (taille ou dimensions).
        Le nom seul ne suffit pas à filtrer une image de taille normale.

        Args:
            filename: Nom du fichier
            data: Contenu binaire de l'image
            attachment: Objet pièce jointe extract_msg

        Returns:
            Tuple (est_insignifiant, raison)
        """
        ext = Path(filename).suffix.lower()

        # Vérifier que c'est bien une image
        if ext not in self.IMAGE_EXTENSIONS:
            return False, ""

        # Si pas de données, on ne peut pas analyser
        if not data:
            return False, ""

        name_without_ext = Path(filename).stem.lower()
        file_size = len(data)
        width, height, surface = 0, 0, 0

        # Vérifier les dimensions de l'image (si PIL disponible)
        if PIL_AVAILABLE:
            try:
                img = Image.open(io.BytesIO(data))
                width, height = img.size
                surface = width * height
                img.close()
            except Exception:
                pass

        def size_info() -> str:
            if width and height:
                return f"{width}x{height}, {file_size // 1024}KB"
            return f"{file_size // 1024}KB"

        # 1. Filtrer les images de forme séparateur (ligne de 1-20px de haut/large)
        if width and height:
            aspect_ratio = max(width, height) / max(min(width, height), 1)
            if aspect_ratio > 10 and min(width, height) < 20:
                return True, f"séparateur ({size_info()})"

        # 2. Filtrer les images TOUJOURS trop petites pour être utiles
        #    Critères : taille fichier < 15KB ET (dimensions < 150x150 OU surface < 25000px²)
        is_tiny_file = file_size < self.ALWAYS_FILTER_SIZE_BYTES
        is_tiny_dimensions = width > 0 and height > 0 and (
            (width < self.ALWAYS_FILTER_DIMENSION and height < self.ALWAYS_FILTER_DIMENSION)
            or surface < self.ALWAYS_FILTER_SURFACE
        )

        if is_tiny_file and is_tiny_dimensions:
            return True, f"image trop petite ({size_info()})"

        # Si pas de dimensions disponibles mais fichier très petit, filtrer aussi
        if not width and not height and file_size < self.ALWAYS_FILTER_SIZE_BYTES:
            return True, f"fichier trop petit ({size_info()})"

        # 3. Pour les noms suspects, seuils plus permissifs (< 30KB ou < 200x200)
        is_small_file = file_size < self.MIN_IMAGE_SIZE_BYTES
        is_small_dimensions = width > 0 and height > 0 and (
            width < self.MIN_IMAGE_DIMENSION and height < self.MIN_IMAGE_DIMENSION
        )
        is_small = is_small_file or is_small_dimensions

        if is_small:
            for pattern in self.INSIGNIFICANT_IMAGE_PATTERNS:
                if re.search(pattern, name_without_ext, re.IGNORECASE):
                    return True, f"nom suspect + petite ({size_info()})"

        # Les images plus grandes sont conservées
        return False, ""

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """
        Convertit un fichier MSG en PDF.

        Si le message a des pièces jointes, crée un dossier contenant:
        - Le PDF du message
        - Les pièces jointes (converties en PDF si possible)
        """
        start = time.time()

        if not EXTRACT_MSG_AVAILABLE:
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="extract_msg non installé",
            )

        msg = None
        try:
            self.logger.debug("Ouverture MSG avec extract_msg")
            try:
                msg = extract_msg.Message(str(source))
            except TypeError as e:
                # Erreur courante avec certains MSG malformés
                self.logger.debug(f"Erreur ouverture MSG (TypeError): {e}")
                return ConversionResult(
                    status=ConversionStatus.FAILED,
                    source=source,
                    dest=None,
                    duration=time.time() - start,
                    method=self.name,
                    message=f"Fichier MSG malformé: {e}",
                    exception=e,
                )

            try:
                if hasattr(msg, "process"):
                    msg.process()
            except Exception:
                pass

            # Extraire les métadonnées
            subject = getattr(msg, "subject", "") or ""
            sender = getattr(msg, "sender", "") or ""
            to = getattr(msg, "to", "") or ""

            # Date peut être un datetime, un int (timestamp) ou une string
            raw_date = getattr(msg, "date", None)
            if raw_date is None:
                date = ""
            elif isinstance(raw_date, str):
                date = raw_date
            elif hasattr(raw_date, "strftime"):
                # C'est un datetime
                date = raw_date.strftime("%Y-%m-%d %H:%M:%S")
            elif isinstance(raw_date, (int, float)):
                # C'est un timestamp
                from datetime import datetime
                try:
                    date = datetime.fromtimestamp(raw_date).strftime("%Y-%m-%d %H:%M:%S")
                except (ValueError, OSError):
                    date = str(raw_date)
            else:
                date = str(raw_date)

            body = getattr(msg, "body", "") or ""

            # Corps HTML
            html_body = (
                getattr(msg, "htmlBody", None)
                or getattr(msg, "html", None)
                or getattr(msg, "bodyHtml", None)
            )

            # Pièces jointes
            attachments = []
            try:
                atts = getattr(msg, "attachments", None)
                self.logger.debug(f"  msg.attachments = {atts}, type = {type(atts)}")

                if atts:
                    self.logger.debug(f"  Nombre de pièces jointes brutes: {len(atts) if hasattr(atts, '__len__') else 'inconnu'}")

                    for idx, a in enumerate(atts):
                        # Log détaillé de chaque pièce jointe
                        self.logger.debug(f"  Attachment #{idx}: type={type(a).__name__}")

                        # Essayer plusieurs attributs pour le nom de fichier
                        long_fn = getattr(a, "longFilename", None)
                        fn = getattr(a, "filename", None)
                        short_fn = getattr(a, "shortFilename", None)
                        display_fn = getattr(a, "displayName", None)

                        self.logger.debug(f"    longFilename={long_fn}, filename={fn}, shortFilename={short_fn}, displayName={display_fn}")

                        final_fn = long_fn or fn or short_fn or display_fn or ""

                        # Si pas de nom, générer un nom basé sur le type MIME ou l'index
                        if not final_fn:
                            # Essayer de déterminer l'extension depuis le type MIME
                            mime_type = getattr(a, "mimetype", None) or getattr(a, "mimeType", None) or ""
                            ext = self._get_extension_from_mime(mime_type)
                            final_fn = f"attachment_{idx + 1}{ext}"
                            self.logger.debug(f"    Nom généré: {final_fn} (mime={mime_type})")

                        attachments.append((final_fn, a))
                        self.logger.debug(f"    -> Ajouté: {final_fn}")

            except Exception as e:
                self.logger.debug(f"Erreur lecture pièces jointes: {e}")
                import traceback
                self.logger.debug(traceback.format_exc())

            # Si pas de pièces jointes, conversion simple
            if not attachments:
                return self._convert_message_only(
                    source, dest, start,
                    subject, sender, to, date, body, html_body, []
                )

            # Avec pièces jointes: créer un dossier
            self.logger.debug(f"  {len(attachments)} pièce(s) jointe(s) brute(s)")

            # Le dossier porte le nom du fichier source + "-open"
            # Ex: message.msg -> message.msg-open/ (évite le conflit avec le fichier source)
            output_folder = dest.parent / f"{source.name}-open"

            # Vérifier si le dossier existe déjà (skip sauf si force)
            if output_folder.exists() and not self.config.force:
                self.logger.info(f"  Dossier déjà existant: {output_folder.name}/ (utiliser --force)")
                return ConversionResult(
                    status=ConversionStatus.SKIPPED_EXISTS,
                    source=source,
                    dest=output_folder,
                    duration=time.time() - start,
                    method=self.name,
                    message="Dossier de sortie déjà existant",
                )

            output_folder.mkdir(parents=True, exist_ok=True)

            # Extraire et convertir les pièces jointes (avec filtrage des petites images)
            attachment_results = self._process_attachments(
                attachments, output_folder, source
            )

            # Log du nombre de pièces jointes retenues
            if attachment_results:
                self.logger.info(f"  {len(attachment_results)} pièce(s) jointe(s) retenue(s)")

            # Créer la liste des pièces jointes pour le PDF du message
            attachments_info = []
            for att_name, att_dest, att_converted in attachment_results:
                if att_converted:
                    attachments_info.append(f"- {att_name} -> {att_dest.name}")
                else:
                    attachments_info.append(f"- {att_name} (non converti)")

            # Convertir le message lui-même
            message_pdf = output_folder / "_message.pdf"
            result = self._convert_message_only(
                source, message_pdf, start,
                subject, sender, to, date, body, html_body, attachments_info
            )

            if result.status == ConversionStatus.SUCCESS:
                # Succès: le dossier est la "destination"
                return ConversionResult(
                    status=ConversionStatus.SUCCESS,
                    source=source,
                    dest=output_folder,
                    duration=time.time() - start,
                    method=f"{self.name}_folder",
                    message=f"Dossier créé avec {len(attachments)} pièce(s) jointe(s)",
                )
            else:
                # Échec de conversion du message, nettoyer le dossier
                try:
                    shutil.rmtree(output_folder)
                except Exception:
                    pass
                return result

        except Exception as e:
            self.logger.error(f"Erreur MSG: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )

        finally:
            # TOUJOURS fermer le fichier MSG pour libérer les handles
            if msg is not None:
                try:
                    if hasattr(msg, "close"):
                        msg.close()
                except Exception:
                    pass

    def _process_attachments(
        self,
        attachments: list,
        output_folder: Path,
        source: Path,
    ) -> list[tuple[str, Path | None, bool]]:
        """
        Extrait et convertit les pièces jointes.

        Filtre automatiquement les petites images insignifiantes
        (logos, signatures, pixels de tracking).

        Returns:
            Liste de tuples (nom_original, chemin_destination, converti_en_pdf)
        """
        results = []
        converters = self._get_converters()
        filtered_count = 0

        # Compteur pour gérer les noms de fichiers en double
        name_counter: dict[str, int] = {}

        for filename, attachment in attachments:
            safe_name = self._sanitize_filename(filename)
            ext = Path(filename).suffix.lower()

            # Gérer les noms de fichiers en double (image.jpg, image (1).jpg, etc.)
            if safe_name in name_counter:
                name_counter[safe_name] += 1
                stem = Path(safe_name).stem
                suffix = Path(safe_name).suffix
                safe_name = f"{stem} ({name_counter[safe_name]}){suffix}"
            else:
                name_counter[safe_name] = 0

            # Extraire la pièce jointe
            try:
                att_data = None
                if hasattr(attachment, "data"):
                    att_data = attachment.data
                elif hasattr(attachment, "getStream"):
                    att_data = attachment.getStream()

                # Convertir en bytes si nécessaire
                if att_data is not None and not isinstance(att_data, bytes):
                    if hasattr(att_data, 'read'):
                        att_data = att_data.read()
                    else:
                        att_data = bytes(att_data)

                if att_data is None:
                    self.logger.debug(f"  Pièce jointe vide: {filename}")
                    results.append((filename, None, False))
                    continue

                # Filtrer les images insignifiantes AVANT de les sauvegarder
                is_insignificant, reason = self._is_insignificant_image(
                    filename, att_data, attachment
                )
                if is_insignificant:
                    self.logger.debug(f"  Image filtrée: {filename} ({reason})")
                    filtered_count += 1
                    continue

                # Sauvegarder la pièce jointe
                temp_path = output_folder / safe_name
                with open(temp_path, "wb") as f:
                    f.write(att_data)

                # Tenter de convertir en PDF si extension supportée
                if ext in self.CONVERTIBLE_EXTENSIONS:
                    pdf_dest = output_folder / (safe_name + ".pdf")
                    converted = False

                    for converter in converters:
                        if converter.name == self.name:
                            continue  # Éviter récursion infinie sur MSG imbriqués
                        if not converter.can_convert(ext):
                            continue
                        if not converter.is_available():
                            continue

                        self.logger.info(f"    [CONV] {filename}")
                        conv_result = converter.convert(temp_path, pdf_dest)

                        if conv_result.status == ConversionStatus.SUCCESS:
                            self.logger.info(f"      -> OK [{converter.name}]")
                            # Supprimer le fichier original uniquement si delete_source est activé
                            if self.config.delete_source:
                                try:
                                    temp_path.unlink()
                                    self.logger.debug(f"      -> Original supprimé")
                                except Exception:
                                    pass
                            results.append((filename, pdf_dest, True))
                            converted = True
                            break
                        else:
                            self.logger.debug(f"      -> Échec [{converter.name}]")

                    if not converted:
                        # Garder le fichier original
                        self.logger.warning(f"    [ÉCHEC] {filename} (conservé)")
                        results.append((filename, temp_path, False))
                else:
                    # Extension non convertible, garder tel quel
                    self.logger.info(f"    [KEEP] {filename}")
                    results.append((filename, temp_path, False))

            except Exception as e:
                self.logger.debug(f"  Erreur extraction {filename}: {e}")
                results.append((filename, None, False))

        if filtered_count > 0:
            self.logger.info(f"  {filtered_count} petite(s) image(s) filtrée(s)")

        return results

    def _convert_message_only(
        self,
        source: Path,
        dest: Path,
        start: float,
        subject: str,
        sender: str,
        to: str,
        date: str,
        body: str,
        html_body: str | bytes | None,
        attachments_info: list[str],
    ) -> ConversionResult:
        """Convertit uniquement le corps du message en PDF."""

        attachments_block = ""
        if attachments_info:
            attachments_block = "\n\nPièces jointes:\n" + "\n".join(attachments_info)

        # Essayer HTML d'abord
        if html_body and isinstance(html_body, (str, bytes)):
            if isinstance(html_body, bytes):
                try:
                    html_body = html_body.decode('utf-8')
                except UnicodeDecodeError:
                    html_body = html_body.decode('latin-1', errors='replace')

            if html_body.strip():
                result = self._convert_html_message(
                    source, dest, start,
                    subject, sender, to, date, html_body, attachments_block
                )
                if result.status == ConversionStatus.SUCCESS:
                    return result
                self.logger.debug("Conversion HTML échouée, fallback texte")

        # Fallback texte
        return self._create_text_pdf(
            source, dest, start,
            subject, sender, to, date, body, attachments_block
        )

    def _convert_html_message(
        self,
        source: Path,
        dest: Path,
        start: float,
        subject: str,
        sender: str,
        to: str,
        date: str,
        html_body: str,
        attachments_block: str,
    ) -> ConversionResult:
        """Convertit le message HTML en PDF via navigateur."""
        try:
            tmp_html = dest.with_suffix(".tmp.html")

            # Construire le HTML complet avec CSS pour word-wrap
            html_doc = f"""<!doctype html>
<html>
<head>
<meta charset='utf-8'>
<title>{self._escape_html(subject)}</title>
<style>
  body {{
    font-family: Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.4;
    word-wrap: break-word;
    overflow-wrap: break-word;
    max-width: 100%;
    padding: 20px;
  }}
  pre {{
    white-space: pre-wrap;
    word-wrap: break-word;
    overflow-wrap: break-word;
  }}
  a {{
    word-break: break-all;
  }}
  table {{
    max-width: 100%;
    table-layout: fixed;
  }}
  td, th {{
    word-wrap: break-word;
    overflow-wrap: break-word;
  }}
  img {{
    max-width: 100%;
    height: auto;
  }}
  .header {{
    background: #f5f5f5;
    padding: 10px;
    margin-bottom: 20px;
    border-radius: 5px;
  }}
  .attachments {{
    background: #fff3cd;
    padding: 10px;
    margin-top: 20px;
    border-radius: 5px;
  }}
</style>
</head>
<body>
<div class="header">
<h3>{self._escape_html(source.name)}</h3>
<div><b>Objet:</b> {self._escape_html(subject)}</div>
<div><b>De:</b> {self._escape_html(sender)}</div>
<div><b>À:</b> {self._escape_html(to)}</div>
<div><b>Date:</b> {self._escape_html(date)}</div>
</div>
<hr/>
{html_body}
{f'<div class="attachments"><pre>{self._escape_html(attachments_block)}</pre></div>' if attachments_block else ''}
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

            if html_result.status == ConversionStatus.SUCCESS:
                return ConversionResult(
                    status=ConversionStatus.SUCCESS,
                    source=source,
                    dest=dest,
                    duration=time.time() - start,
                    method=f"{self.name}_html",
                )

            return html_result

        except Exception as e:
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )

    def _escape_html(self, text: str) -> str:
        """Échappe les caractères HTML."""
        if not text:
            return ""
        return (
            str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    def _wrap_long_lines(self, text: str, max_chars: int = 100) -> str:
        """Découpe les lignes trop longues."""
        import textwrap

        lines = text.split('\n')
        wrapped_lines = []

        for line in lines:
            if len(line) > max_chars:
                wrapped = textwrap.fill(
                    line,
                    width=max_chars,
                    break_long_words=True,
                    break_on_hyphens=False,
                )
                wrapped_lines.append(wrapped)
            else:
                wrapped_lines.append(line)

        return '\n'.join(wrapped_lines)

    def _escape_xml(self, text) -> str:
        """Échappe les caractères spéciaux XML pour ReportLab."""
        if text is None:
            return ""
        # Convertir en string si nécessaire (date peut être un datetime ou int)
        if not isinstance(text, str):
            text = str(text)
        if not text:
            return ""
        return (
            text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
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
        """Crée un PDF texte avec word-wrap automatique."""
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

            # Style pour le corps avec word-wrap
            body_style = ParagraphStyle(
                "BodyStyle",
                parent=styles["Normal"],
                fontName="Courier",
                fontSize=9,
                leading=11,
                wordWrap='CJK',
            )

            # Style pour les en-têtes
            header_style = ParagraphStyle(
                "HeaderStyle",
                parent=styles["Normal"],
                fontSize=10,
                leading=12,
                wordWrap='CJK',
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
            story.append(Paragraph(f"<b>{self._escape_xml(source.name)}</b>", styles["Heading2"]))
            story.append(Spacer(1, 6))
            story.append(Paragraph(f"<b>Objet:</b> {self._escape_xml(subject)}", header_style))
            story.append(Paragraph(f"<b>De:</b> {self._escape_xml(sender)}", header_style))
            story.append(Paragraph(f"<b>À:</b> {self._escape_xml(self._wrap_long_lines(to, 80))}", header_style))
            story.append(Paragraph(f"<b>Date:</b> {self._escape_xml(date)}", header_style))
            story.append(Spacer(1, 12))

            # Corps
            body_wrapped = self._wrap_long_lines(body, 95)
            body_escaped = self._escape_xml(body_wrapped)
            body_html = body_escaped.replace('\n', '<br/>')
            story.append(Paragraph(body_html, body_style))

            # Pièces jointes
            if attachments:
                story.append(Spacer(1, 12))
                att_escaped = self._escape_xml(attachments)
                att_html = att_escaped.replace('\n', '<br/>')
                story.append(Paragraph(att_html, body_style))

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
