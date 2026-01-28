#!/usr/bin/env python3
"""
ConverterToPdf - Wrapper de compatibilité.

Ce fichier maintient la compatibilité avec l'ancienne interface.
Pour la nouvelle version modulaire, utilisez:

    python -m converter_pdf <répertoire> [options]

Ou importez directement:

    from converter_pdf import Config, ConverterLogger, FileProcessor
"""

import sys
import warnings

# Avertissement de dépréciation
warnings.warn(
    "L'utilisation de 'python converter_pdf.py' est dépréciée. "
    "Utilisez 'python -m converter_pdf' pour la nouvelle version modulaire.",
    DeprecationWarning,
    stacklevel=2,
)

# Importer et exécuter le nouveau module
from converter_pdf.__main__ import main

if __name__ == "__main__":
    sys.exit(main())
