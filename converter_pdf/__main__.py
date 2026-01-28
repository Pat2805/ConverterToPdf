"""
Point d'entrée principal pour python -m converter_pdf.
"""

from __future__ import annotations

import sys
from pathlib import Path

from .cli import parse_args, get_extensions_filter, print_check_info
from .config import Config
from .logger import ConverterLogger
from .processor import FileProcessor


def main() -> int:
    """
    Fonction principale.

    Returns:
        Code de sortie (0 = succès)
    """
    args = parse_args()

    # Mode vérification
    if args.check:
        print_check_info()
        return 0

    # Charger la configuration
    try:
        config = Config.load(args.config)
    except Exception as e:
        print(f"Erreur chargement config: {e}", file=sys.stderr)
        config = Config()

    # Mettre à jour depuis les arguments CLI
    config.update_from_args(args)

    # Filtre d'extensions
    extensions = get_extensions_filter(args)
    if extensions:
        config.extensions = extensions

    # Setup logging
    logger = ConverterLogger()
    logger.setup(
        level=config.log_level,
        log_file=config.log_file,
    )

    # Vérifier le chemin
    path = args.path
    if path is None:
        print("Erreur: chemin requis. Utilisez --help pour l'aide.", file=sys.stderr)
        return 1
    if not path.exists():
        logger.error(f"Chemin introuvable: {path}")
        return 1

    # Créer le processeur
    processor = FileProcessor(config, logger)

    # Traiter
    if path.is_file():
        result = processor.process_file(path, args.output)
        return 0 if result.is_success or result.is_skipped else 1
    else:
        stats = processor.process_directory(path, args.output)
        # Retourner 0 si au moins un succès ou aucun échec
        if stats["success"] > 0 or stats["failed"] == 0:
            return 0
        return 1


if __name__ == "__main__":
    sys.exit(main())
