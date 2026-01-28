"""
Orchestrateur de traitement des fichiers.

Gère le parcours des fichiers, la sélection des convertisseurs,
et la coordination avec le journal et le logger.
"""

from __future__ import annotations

import signal
import sys
import time
from pathlib import Path
from typing import TYPE_CHECKING

from .config import Config
from .logger import ConverterLogger
from .journal import Journal
from .converters import get_converter_chain, ConversionResult, ConversionStatus

if TYPE_CHECKING:
    from .converters.base import BaseConverter


class FileProcessor:
    """
    Orchestrateur principal pour le traitement des fichiers.

    Gère:
    - Le parcours des fichiers (récursif ou non)
    - La sélection du bon convertisseur
    - La chaîne de fallback si un convertisseur échoue
    - Le journal d'audit
    - L'interruption propre (Ctrl+C)

    Usage:
        processor = FileProcessor(config, logger)
        processor.process_directory(Path("./documents"), recursive=True)
    """

    def __init__(self, config: Config, logger: ConverterLogger):
        """
        Initialise le processeur.

        Args:
            config: Configuration
            logger: Logger configuré
        """
        self.config = config
        self.logger = logger
        self.converters = get_converter_chain(config, logger)
        self.journal: Journal | None = None
        self._interrupted = False

        # Statistiques
        self.stats = {
            "total": 0,
            "success": 0,
            "failed": 0,
            "skipped": 0,
        }

    def _setup_signal_handler(self) -> None:
        """Configure la gestion de Ctrl+C."""
        def signal_handler(signum, frame):
            self._interrupted = True
            self.logger.warning("Interruption demandée (Ctrl+C), arrêt propre...")

        signal.signal(signal.SIGINT, signal_handler)

    def _get_dest_path(self, source: Path, dest_dir: Path | None) -> Path:
        """
        Calcule le chemin de destination du PDF.

        Args:
            source: Fichier source
            dest_dir: Répertoire de destination (None = même répertoire)

        Returns:
            Chemin du PDF de destination
        """
        if dest_dir is None:
            dest_dir = source.parent
        else:
            dest_dir.mkdir(parents=True, exist_ok=True)

        if self.config.keep_extension:
            # document.docx -> document.docx.pdf
            pdf_name = source.name + ".pdf"
        else:
            # document.docx -> document.pdf
            pdf_name = source.stem + ".pdf"

        return dest_dir / pdf_name

    def process_file(
        self,
        source: Path,
        dest_dir: Path | None = None,
    ) -> ConversionResult:
        """
        Traite un fichier unique.

        Args:
            source: Fichier source
            dest_dir: Répertoire de destination (optionnel)

        Returns:
            Résultat de la conversion
        """
        dest = self._get_dest_path(source, dest_dir)

        with self.logger.file_context(source):
            # Vérifier si le fichier est déjà un PDF
            if source.suffix.lower() == ".pdf":
                if dest_dir is None or dest_dir == source.parent:
                    self.logger.skip(source.name, "déjà PDF")
                    return ConversionResult(
                        status=ConversionStatus.SKIPPED_PDF,
                        source=source,
                        dest=None,
                        duration=0,
                        method="skip",
                        message="Fichier déjà au format PDF",
                    )

            # Vérifier si le PDF existe déjà
            if dest.exists() and not self.config.force:
                self.logger.skip(source.name, "PDF existant")
                return ConversionResult(
                    status=ConversionStatus.SKIPPED_EXISTS,
                    source=source,
                    dest=dest,
                    duration=0,
                    method="skip",
                    message="PDF déjà existant",
                )

            # Gérer les conflits de noms
            if dest.exists() and self.config.force:
                self.logger.debug(f"Remplacement: {dest.name}")
            else:
                counter = 1
                original_dest = dest
                while dest.exists():
                    if self.config.keep_extension:
                        pdf_name = f"{source.name}_{counter}.pdf"
                    else:
                        pdf_name = f"{source.stem}_{counter}.pdf"
                    dest = (dest_dir or source.parent) / pdf_name
                    counter += 1

            self.logger.info(f"Conversion: {source.name} -> {dest.name}")

            # Essayer les convertisseurs en chaîne
            start = time.time()
            result: ConversionResult | None = None

            for converter in self.converters:
                if not converter.can_convert(source.suffix):
                    continue

                if not converter.is_available():
                    self.logger.debug(f"{converter.name} non disponible, skip")
                    continue

                self.logger.debug(f"Tentative avec {converter.name}")
                result = converter.convert(source, dest)

                if result.status == ConversionStatus.SUCCESS:
                    self.logger.success(
                        f"Converti en {result.duration:.1f}s",
                        method=result.method,
                        size=f"{result.source_size_mb:.1f}MB -> {result.dest_size_mb:.1f}MB",
                    )
                    break

                elif result.status == ConversionStatus.SKIPPED_PASSWORD:
                    self.logger.skip(source.name, "mot de passe requis")
                    break

                else:
                    self.logger.debug(f"{converter.name} a échoué, essai suivant...")
                    # Continuer avec le prochain convertisseur

            # Aucun convertisseur disponible ou tous ont échoué
            if result is None:
                result = ConversionResult(
                    status=ConversionStatus.SKIPPED_UNSUPPORTED,
                    source=source,
                    dest=None,
                    duration=time.time() - start,
                    method="none",
                    message=f"Aucun convertisseur pour {source.suffix}",
                )
                self.logger.skip(source.name, f"format non supporté ({source.suffix})")

            elif result.status == ConversionStatus.FAILED:
                self.logger.fail(
                    f"Échec conversion",
                    exc=result.exception,
                    method=result.method,
                )

            # Journaliser
            if self.journal:
                self.journal.log(result)

            # Supprimer le source si demandé et succès
            if result.is_success and self.config.delete_source:
                try:
                    source.unlink()
                    self.logger.debug("Fichier source supprimé")
                except Exception as e:
                    self.logger.warning(f"Impossible de supprimer le source: {e}")

            return result

    def process_directory(
        self,
        directory: Path,
        dest_dir: Path | None = None,
    ) -> dict[str, int]:
        """
        Traite tous les fichiers d'un répertoire.

        Args:
            directory: Répertoire à traiter
            dest_dir: Répertoire de destination (optionnel)

        Returns:
            Statistiques de traitement
        """
        directory = Path(directory)

        if not directory.exists() or not directory.is_dir():
            self.logger.error(f"Répertoire invalide: {directory}")
            return self.stats

        # Setup
        self._setup_signal_handler()
        self._interrupted = False

        # Initialiser le journal
        journal_dir = dest_dir or directory
        if self.config.journal_enabled:
            self.journal = Journal(self.config, self.logger, journal_dir)
            self.journal.open()

        # Réinitialiser les stats
        self.stats = {"total": 0, "success": 0, "failed": 0, "skipped": 0}

        # Afficher la configuration
        self._print_config(directory)

        # Pattern de recherche
        pattern = "**/*" if self.config.recursive else "*"
        extensions = self.config.get_all_extensions()

        self.logger.info(f"Traitement: {directory}")
        start_total = time.time()

        try:
            for file_path in directory.glob(pattern):
                if self._interrupted:
                    break

                if not file_path.is_file():
                    continue

                if file_path.suffix.lower() not in extensions:
                    continue

                self.stats["total"] += 1
                result = self.process_file(file_path, dest_dir)

                if result.is_success:
                    self.stats["success"] += 1
                elif result.is_skipped:
                    self.stats["skipped"] += 1
                else:
                    self.stats["failed"] += 1

        except KeyboardInterrupt:
            self.logger.warning("Interruption clavier")
        finally:
            # Fermer le journal
            if self.journal:
                self.journal.close()

        duration_total = time.time() - start_total

        # Afficher le résumé
        self._print_summary(duration_total)

        return self.stats

    def _print_config(self, directory: Path) -> None:
        """Affiche la configuration de traitement."""
        print("\n" + "=" * 60)
        print("CONFIGURATION")
        print("=" * 60)
        print(f"Répertoire: {directory}")
        print(f"Récursif: {'Oui' if self.config.recursive else 'Non'}")
        print(f"Méthode: {self.config.method}")
        print(f"Forcer: {'Oui' if self.config.force else 'Non'}")
        print(f"Nommage: {'x.ext.pdf' if self.config.keep_extension else 'x.pdf'}")
        print(f"Journal: {'Oui' if self.config.journal_enabled else 'Non'}")
        print(f"Log level: {self.config.log_level}")

        # Convertisseurs disponibles
        print("\nConvertisseurs disponibles:")
        for conv in self.converters:
            status = "OK" if conv.is_available() else "Non disponible"
            print(f"  - {conv.name}: {status}")

        print("=" * 60 + "\n")

    def _print_summary(self, duration: float) -> None:
        """Affiche le résumé du traitement."""
        print("\n" + "=" * 60)
        print("RÉSUMÉ")
        print("=" * 60)
        print(f"Durée totale: {duration:.1f}s")
        print(f"Fichiers traités: {self.stats['total']}")
        print(f"  - Succès: {self.stats['success']}")
        print(f"  - Ignorés: {self.stats['skipped']}")
        print(f"  - Échecs: {self.stats['failed']}")

        if self.journal and self.journal.path:
            print(f"\nJournal: {self.journal.path}")

        if self.stats["skipped"] > 0 and not self.config.force:
            print("\nAstuce: Utilisez --force pour reconvertir les fichiers existants")

        print("=" * 60 + "\n")
