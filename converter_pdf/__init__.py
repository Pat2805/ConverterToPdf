"""
ConverterToPdf - Convertisseur de documents en PDF

Un utilitaire Python pour convertir des documents Office, images, HTML, etc. en PDF
avec journalisation complète et gestion robuste de Microsoft Office.
"""

__version__ = "2.0.0"
__author__ = "Pat2805"

from .config import Config
from .logger import ConverterLogger
from .processor import FileProcessor

__all__ = ["Config", "ConverterLogger", "FileProcessor", "__version__"]
