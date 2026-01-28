"""
Interface en ligne de commande pour ConverterToPdf.

Parse les arguments et configure l'application.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__


def create_parser() -> argparse.ArgumentParser:
    """Crée le parser d'arguments."""
    parser = argparse.ArgumentParser(
        prog="converter_pdf",
        description="Convertisseur de documents en PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Formats supportés:
  Documents : .doc .docx .rtf .odt .xls .xlsx .xlsm .xlsb .ppt .pptx
  Images    : .jpg .jpeg .png .bmp .tiff .tif .webp
  Web       : .htm .html
  Texte     : .txt .log
  Données   : .xml
  Email     : .msg

Exemples:
  python -m converter_pdf ./documents
  python -m converter_pdf ./documents -r -o ./pdf_output
  python -m converter_pdf ./documents --method office --log-level DEBUG
  python -m converter_pdf ./scans --images-only --ocr
""",
    )

    # Argument positionnel: chemin (optionnel si --check)
    parser.add_argument(
        "path",
        type=Path,
        nargs="?",  # Optionnel
        default=None,
        help="Fichier ou répertoire à traiter",
    )

    # Options générales
    general = parser.add_argument_group("Options générales")
    general.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="Traiter les sous-répertoires",
    )
    general.add_argument(
        "-o", "--output",
        type=Path,
        metavar="DIR",
        help="Répertoire de sortie",
    )
    general.add_argument(
        "-f", "--force",
        action="store_true",
        help="Forcer la reconversion même si le PDF existe",
    )
    general.add_argument(
        "-d", "--delete",
        action="store_true",
        help="Supprimer les fichiers sources après conversion",
    )

    # Méthode de conversion
    method = parser.add_argument_group("Méthode de conversion")
    method.add_argument(
        "--method",
        choices=["auto", "office", "libreoffice", "reportlab"],
        default="auto",
        help="Méthode de conversion (défaut: auto)",
    )

    # Logging
    logging = parser.add_argument_group("Logging")
    logging.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="Niveau de log console (défaut: INFO)",
    )
    logging.add_argument(
        "--log-file",
        type=Path,
        metavar="FILE",
        help="Fichier de log (debug complet)",
    )

    # Journal CSV
    journal = parser.add_argument_group("Journal CSV")
    journal.add_argument(
        "--no-journal",
        action="store_true",
        help="Désactiver le journal CSV",
    )
    journal.add_argument(
        "--log-all",
        action="store_true",
        help="Journaliser tous les fichiers (pas seulement les erreurs)",
    )

    # Nommage
    naming = parser.add_argument_group("Nommage des PDF")
    naming.add_argument(
        "--no-keep-ext",
        action="store_true",
        help="Nommer les PDF sans l'extension d'origine (x.pdf au lieu de x.docx.pdf)",
    )

    # Filtres de formats
    filters = parser.add_argument_group("Filtres de formats")
    filters.add_argument(
        "--images-only",
        action="store_true",
        help="Traiter uniquement les images",
    )
    filters.add_argument(
        "--word-only",
        action="store_true",
        help="Traiter uniquement les documents Word",
    )
    filters.add_argument(
        "--excel-only",
        action="store_true",
        help="Traiter uniquement les fichiers Excel",
    )
    filters.add_argument(
        "--xml-only",
        action="store_true",
        help="Traiter uniquement les fichiers XML",
    )

    # OCR
    ocr = parser.add_argument_group("OCR (images)")
    ocr.add_argument(
        "--ocr",
        action="store_true",
        help="Activer l'OCR pour les images",
    )
    ocr.add_argument(
        "--ocr-engine",
        choices=["auto", "tesseract", "easyocr", "paddleocr"],
        default="auto",
        help="Moteur OCR (défaut: auto)",
    )

    # Configuration
    config = parser.add_argument_group("Configuration")
    config.add_argument(
        "--config",
        type=Path,
        metavar="FILE",
        help="Fichier de configuration .converterrc",
    )
    config.add_argument(
        "--check",
        action="store_true",
        help="Vérifier la configuration et les outils disponibles",
    )

    # Version
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    return parser


def parse_args(args: list[str] | None = None) -> argparse.Namespace:
    """
    Parse les arguments de la ligne de commande.

    Args:
        args: Arguments (None = sys.argv)

    Returns:
        Namespace avec les arguments parsés
    """
    parser = create_parser()
    return parser.parse_args(args)


def get_extensions_filter(args: argparse.Namespace) -> list[str] | None:
    """
    Détermine le filtre d'extensions selon les arguments.

    Args:
        args: Arguments parsés

    Returns:
        Liste d'extensions ou None (toutes)
    """
    if args.images_only:
        return [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"]
    elif args.word_only:
        return [".doc", ".docx", ".rtf", ".odt"]
    elif args.excel_only:
        return [".xls", ".xlsx", ".xlsm", ".xlsb"]
    elif args.xml_only:
        return [".xml"]
    return None


def print_check_info() -> None:
    """Affiche les informations de configuration."""
    import platform

    print("\n" + "=" * 60)
    print("VÉRIFICATION DE LA CONFIGURATION")
    print("=" * 60)

    print(f"\nSystème: {platform.system()} {platform.release()}")
    print(f"Python: {platform.python_version()}")

    # pywin32
    try:
        import win32com.client
        print("\n[OK] pywin32 installé")

        # Tester Office
        from .com_utils import detect_office_installation
        from .logger import ConverterLogger
        logger = ConverterLogger()
        logger.setup(level="ERROR")  # Silencieux
        office = detect_office_installation(logger)

        for app, available in office.items():
            status = "[OK]" if available else "[--]"
            print(f"  {status} {app.capitalize()}")

    except ImportError:
        print("\n[--] pywin32 non installé (pip install pywin32)")

    # LibreOffice
    from .converters.libreoffice import LibreOfficeConverter
    from .config import Config
    from .logger import ConverterLogger
    config = Config()
    logger = ConverterLogger()
    logger.setup(level="ERROR")
    lo = LibreOfficeConverter(config, logger)
    if lo.is_available():
        print(f"\n[OK] LibreOffice: {lo.libreoffice_path}")
    else:
        print("\n[--] LibreOffice non détecté")

    # Navigateur
    from .converters.html import HtmlConverter
    html = HtmlConverter(config, logger)
    if html.is_available():
        print(f"\n[OK] Navigateur: {html.browser_path}")
    else:
        print("\n[--] Navigateur (Chrome/Edge) non détecté")

    # ReportLab
    try:
        import reportlab
        print(f"\n[OK] ReportLab: {reportlab.Version}")
    except ImportError:
        print("\n[--] ReportLab non installé (pip install reportlab)")

    # PIL
    try:
        from PIL import Image
        print(f"\n[OK] Pillow installé")
    except ImportError:
        print("\n[--] Pillow non installé (pip install Pillow)")

    # pandas
    try:
        import pandas as pd
        print(f"\n[OK] pandas: {pd.__version__}")
    except ImportError:
        print("\n[--] pandas non installé (pip install pandas)")

    # python-docx
    try:
        import docx
        print(f"\n[OK] python-docx installé")
    except ImportError:
        print("\n[--] python-docx non installé (pip install python-docx)")

    # extract_msg
    try:
        import extract_msg
        print(f"\n[OK] extract_msg installé")
    except ImportError:
        print("\n[--] extract_msg non installé (pip install extract-msg)")

    # PyYAML
    try:
        import yaml
        print(f"\n[OK] PyYAML installé")
    except ImportError:
        print("\n[--] PyYAML non installé (pip install pyyaml)")

    print("\n" + "=" * 60 + "\n")
