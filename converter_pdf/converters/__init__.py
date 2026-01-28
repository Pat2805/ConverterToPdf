"""
Module des convertisseurs.

Chaque convertisseur hérite de BaseConverter et implémente
la conversion d'un type de fichier spécifique vers PDF.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus

if TYPE_CHECKING:
    from ..config import Config
    from ..logger import ConverterLogger


def get_converter_chain(
    config: "Config",
    logger: "ConverterLogger",
) -> list[BaseConverter]:
    """
    Retourne la chaîne de convertisseurs selon la configuration.

    L'ordre est important : les convertisseurs sont essayés dans l'ordre
    jusqu'à ce qu'un réussisse.

    Args:
        config: Configuration
        logger: Logger

    Returns:
        Liste ordonnée de convertisseurs
    """
    from .office import OfficeWordConverter, OfficeExcelConverter, OfficePowerPointConverter
    from .libreoffice import LibreOfficeConverter
    from .image import ImageConverter
    from .html import HtmlConverter
    from .text import TextConverter
    from .xml_converter import XmlConverter
    from .msg import MsgConverter
    from .reportlab_fallback import ReportLabWordConverter, ReportLabExcelConverter

    converters: list[BaseConverter] = []

    # Selon la méthode configurée
    method = config.method

    if method == "auto":
        # Office en premier (meilleure qualité)
        converters.extend([
            OfficeWordConverter(config, logger),
            OfficeExcelConverter(config, logger),
            OfficePowerPointConverter(config, logger),
        ])
        # LibreOffice en fallback
        converters.append(LibreOfficeConverter(config, logger))
        # ReportLab en dernier recours
        converters.extend([
            ReportLabWordConverter(config, logger),
            ReportLabExcelConverter(config, logger),
        ])

    elif method == "office":
        converters.extend([
            OfficeWordConverter(config, logger),
            OfficeExcelConverter(config, logger),
            OfficePowerPointConverter(config, logger),
        ])

    elif method == "libreoffice":
        converters.append(LibreOfficeConverter(config, logger))

    elif method == "reportlab":
        converters.extend([
            ReportLabWordConverter(config, logger),
            ReportLabExcelConverter(config, logger),
        ])

    # Convertisseurs toujours disponibles (indépendants de la méthode)
    converters.extend([
        ImageConverter(config, logger),
        HtmlConverter(config, logger),
        TextConverter(config, logger),
        XmlConverter(config, logger),
        MsgConverter(config, logger),
    ])

    return converters


__all__ = [
    "BaseConverter",
    "ConversionResult",
    "ConversionStatus",
    "get_converter_chain",
]
