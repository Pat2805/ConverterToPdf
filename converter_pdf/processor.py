"""
Orchestrateur de traitement des fichiers.

Gère le parcours des fichiers, la sélection des convertisseurs,
et la coordination avec le rapport et le logger.
"""

from __future__ import annotations

import signal
import sys
import time
from pathlib import Path
from typing import TYPE_CHECKING

from .config import Config
from .logger import ConverterLogger
from .report import SessionReport
from .converters import get_converter_chain, ConversionResult, ConversionStatus

if TYPE_CHECKING:
    from .converters.base import BaseConverter


def format_size(size_bytes: float) -> str:
    """Formate une taille en format lisible."""
    if size_bytes < 1024:
        return f"{size_bytes:.0f} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


class FileProcessor:
    """
    Orchestrateur principal pour le traitement des fichiers.

    Gère:
    - Le parcours des fichiers (récursif ou non)
    - La sélection du bon convertisseur
    - La chaîne de fallback si un convertisseur échoue
    - Le rapport de session
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
        self.report: SessionReport | None = None
        self._interrupted = False

        # Statistiques
        self.stats = {
            "total": 0,
            "success": 0,
            "failed": 0,
            "skipped": 0,
        }

        # Dossiers créés par extraction (archives/MSG) à traiter ensuite
        self._extracted_folders: list[Path] = []
        # Dossiers déjà traités (pour éviter les boucles infinies)
        self._processed_folders: set[Path] = set()

    def _setup_signal_handler(self) -> None:
        """Configure la gestion de Ctrl+C."""
        def signal_handler(signum, frame):
            self._interrupted = True
            self.logger.warning("Interruption demandée (Ctrl+C), arrêt propre...")

        signal.signal(signal.SIGINT, signal_handler)

    def _hide_file(self, file_path: Path) -> None:
        """
        Rend un fichier caché (Windows uniquement).

        Utilise l'attribut FILE_ATTRIBUTE_HIDDEN de Windows.

        Args:
            file_path: Chemin du fichier à cacher
        """
        if sys.platform != "win32":
            self.logger.debug("hide_source ignoré (non-Windows)")
            return

        import ctypes
        FILE_ATTRIBUTE_HIDDEN = 0x02

        # Obtenir les attributs actuels
        attrs = ctypes.windll.kernel32.GetFileAttributesW(str(file_path))
        if attrs == -1:
            raise OSError(f"Impossible de lire les attributs de {file_path}")

        # Ajouter l'attribut caché
        new_attrs = attrs | FILE_ATTRIBUTE_HIDDEN
        result = ctypes.windll.kernel32.SetFileAttributesW(str(file_path), new_attrs)
        if not result:
            raise OSError(f"Impossible de cacher {file_path}")

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

        # Taille du fichier source
        try:
            source_size = source.stat().st_size
            source_size_str = format_size(source_size)
        except OSError:
            source_size = 0
            source_size_str = "?"

        with self.logger.file_context(source):
            # Vérifier si le fichier est déjà un PDF
            if source.suffix.lower() == ".pdf":
                if dest_dir is None or dest_dir == source.parent:
                    self.logger.info(f"Ignoré (déjà PDF) : {source.name}")
                    result = ConversionResult(
                        status=ConversionStatus.SKIPPED_PDF,
                        source=source,
                        dest=None,
                        duration=0,
                        method="skip",
                        message="Fichier déjà au format PDF",
                    )
                    if self.report:
                        self.report.add_result(result)
                    return result

            # Vérifier si le PDF existe déjà
            if dest.exists() and not self.config.force:
                self.logger.info(f"Ignoré (existe) : {source.name} ({source_size_str})")
                result = ConversionResult(
                    status=ConversionStatus.SKIPPED_EXISTS,
                    source=source,
                    dest=dest,
                    duration=0,
                    method="skip",
                    message="PDF déjà existant",
                )
                if self.report:
                    self.report.add_result(result)
                return result

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

            # Mode dry-run: simuler sans convertir
            if self.config.dry_run:
                # Trouver le convertisseur qui serait utilisé
                converter_name = "none"
                for converter in self.converters:
                    if converter.can_convert(source.suffix) and converter.is_available():
                        converter_name = converter.name
                        break

                self.logger.info(f"[DRY-RUN] {source.name} ({source_size_str}) -> {dest.name} [{converter_name}]")
                result = ConversionResult(
                    status=ConversionStatus.SUCCESS,
                    source=source,
                    dest=dest,
                    duration=0,
                    method=f"dry_run:{converter_name}",
                    message="Simulation (dry-run)",
                )
                if self.report:
                    self.report.add_result(result)
                return result

            # Log de début de conversion avec infos
            self.logger.info(f"Conversion : {source.name} ({source_size_str})")

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
                    # Log de succès détaillé
                    dest_size_str = format_size(result.dest_size_mb * 1024 * 1024)
                    ratio = (result.dest_size_mb / result.source_size_mb * 100) if result.source_size_mb > 0 else 0
                    self.logger.info(
                        f"  -> OK : {dest.name} ({dest_size_str}, {ratio:.0f}%) "
                        f"[{result.method}, {result.duration:.1f}s]"
                    )
                    break

                elif result.status == ConversionStatus.SKIPPED_PASSWORD:
                    self.logger.warning(f"  -> Ignoré : mot de passe requis")
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
                self.logger.warning(f"  -> Ignoré : format non supporté ({source.suffix})")

            elif result.status == ConversionStatus.FAILED:
                error_msg = result.message or str(result.exception) or "Erreur inconnue"
                self.logger.error(
                    f"  -> ÉCHEC : {error_msg} [{result.method}]"
                )
                if result.exception:
                    self.logger.debug(f"Exception détaillée: {result.exception}")

            # Ajouter au rapport
            if self.report:
                self.report.add_result(result)

            # Si la conversion a créé un dossier (archive/MSG), le mémoriser pour traitement ultérieur
            if result.is_success and result.dest and result.dest.is_dir():
                resolved_dest = result.dest.resolve()
                if resolved_dest not in self._processed_folders:
                    self._extracted_folders.append(resolved_dest)
                    self.logger.debug(f"Dossier extrait à traiter: {result.dest}")

            # Supprimer le source si demandé et succès
            if result.is_success and self.config.delete_source:
                try:
                    source.unlink()
                    self.logger.debug("Fichier source supprimé")
                except Exception as e:
                    self.logger.warning(f"Impossible de supprimer le source: {e}")

            # Cacher le source si demandé et succès (Windows uniquement)
            if result.is_success and self.config.hide_source:
                try:
                    self._hide_file(source)
                    self.logger.debug("Fichier source caché")
                except Exception as e:
                    self.logger.warning(f"Impossible de cacher le source: {e}")

            return result

    def _process_single_directory(
        self,
        directory: Path,
        pattern: str,
        extensions: set[str],
        dest_dir: Path | None,
    ) -> None:
        """
        Traite les fichiers d'un seul répertoire.

        Args:
            directory: Répertoire à traiter
            pattern: Pattern glob (* ou **/*)
            extensions: Extensions à traiter
            dest_dir: Répertoire de destination (optionnel)
        """
        for file_path in sorted(directory.glob(pattern)):
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

        # Initialiser le rapport de session
        output_dir = dest_dir or directory
        if self.config.report_enabled:
            self.report = SessionReport(
                source_directory=directory,
                output_directory=output_dir,
                recursive=self.config.recursive,
            )
        else:
            self.report = None

        # Réinitialiser les stats et les listes de dossiers
        self.stats = {"total": 0, "success": 0, "failed": 0, "skipped": 0}
        self._extracted_folders = []
        self._processed_folders = set()

        # Afficher la configuration
        self._print_config(directory)

        # Pattern de recherche
        pattern = "**/*" if self.config.recursive else "*"
        extensions = self.config.get_all_extensions()

        self.logger.info(f"Démarrage du traitement : {directory}")
        self.logger.info(f"Mode : {'récursif' if self.config.recursive else 'non récursif'}")
        print("")  # Ligne vide pour aérer

        start_total = time.time()

        try:
            # Traiter le répertoire initial
            self._process_single_directory(directory, pattern, extensions, dest_dir)

            # Traiter les dossiers extraits (archives/MSG) récursivement
            extraction_pass = 0
            while self._extracted_folders and not self._interrupted:
                extraction_pass += 1
                folders_to_process = self._extracted_folders.copy()
                self._extracted_folders = []

                self.logger.info(f"\n--- Traitement des contenus extraits (passe {extraction_pass}) ---")

                for folder in folders_to_process:
                    if self._interrupted:
                        break

                    # Éviter les boucles infinies
                    if folder in self._processed_folders:
                        self.logger.debug(f"Dossier déjà traité, ignoré: {folder}")
                        continue

                    self._processed_folders.add(folder)
                    self.logger.info(f"Traitement du contenu extrait: {folder.name}/")

                    # Traiter ce dossier (toujours récursif pour les contenus extraits)
                    self._process_single_directory(folder, "**/*", extensions, None)

        except KeyboardInterrupt:
            self.logger.warning("Interruption clavier")
        finally:
            pass

        duration_total = time.time() - start_total

        # Finaliser et sauvegarder le rapport
        if self.report:
            self.report.finalize()
            report_path = self.report.save(output_dir)
            if report_path:
                self.logger.info(f"Rapport sauvegardé : {report_path}")

        # Afficher le résumé
        self._print_summary(duration_total)

        return self.stats

    def _print_config(self, directory: Path) -> None:
        """Affiche la configuration de traitement."""
        print("\n" + "=" * 80)
        if self.config.dry_run:
            print("CONFIGURATION (MODE DRY-RUN - SIMULATION)")
        else:
            print("CONFIGURATION")
        print("=" * 80)
        print(f"  Répertoire      : {directory}")
        print(f"  Récursif        : {'Oui' if self.config.recursive else 'Non'}")
        print(f"  Méthode         : {self.config.method}")
        print(f"  Forcer          : {'Oui' if self.config.force else 'Non'}")
        print(f"  Nommage PDF     : {'fichier.ext.pdf' if self.config.keep_extension else 'fichier.pdf'}")
        print(f"  Suppr. source   : {'Oui' if self.config.delete_source else 'Non'}")
        print(f"  Cacher source   : {'Oui' if self.config.hide_source else 'Non'}")
        if self.config.dry_run:
            print(f"  Mode            : DRY-RUN (simulation)")
        print(f"  Rapport         : Oui")
        print(f"  Niveau log      : {self.config.log_level}")

        # Convertisseurs disponibles
        print("\n  Convertisseurs :")
        for conv in self.converters:
            status = "OK" if conv.is_available() else "Non disponible"
            print(f"    - {conv.name}: {status}")

        print("=" * 80 + "\n")

    def _print_summary(self, duration: float) -> None:
        """Affiche le résumé du traitement."""
        print("\n" + "=" * 80)
        if self.config.dry_run:
            print("RÉSUMÉ (DRY-RUN - AUCUN FICHIER MODIFIÉ)")
        else:
            print("RÉSUMÉ")
        print("=" * 80)

        # Statistiques de base
        print(f"  Durée totale      : {duration:.1f}s")
        print(f"  Fichiers analysés : {self.stats['total']}")

        if self.stats['total'] > 0:
            success_pct = (self.stats['success'] / self.stats['total']) * 100
            print(f"    - Convertis     : {self.stats['success']} ({success_pct:.0f}%)")
        else:
            print(f"    - Convertis     : 0")

        print(f"    - Ignorés       : {self.stats['skipped']}")
        print(f"    - Échecs        : {self.stats['failed']}")

        # Statistiques de volume depuis le rapport
        if self.report and self.report.total_source_bytes > 0:
            print("")
            print(f"  Volume source     : {self.report._format_size(self.report.total_source_bytes)}")
            print(f"  Volume PDF        : {self.report._format_size(self.report.total_dest_bytes)}")
            if self.report.total_dest_bytes > 0:
                ratio = (self.report.total_dest_bytes / self.report.total_source_bytes) * 100
                print(f"  Ratio             : {ratio:.0f}%")

        # Rapport
        if self.report:
            report_dir = self.report.output_directory or self.report.source_directory
            if report_dir:
                print(f"\n  Rapport           : {report_dir}/conversion_report_*.txt")

        # Conseils
        if self.stats["failed"] > 0:
            print("\n  [!] Des fichiers ont échoué. Consultez le rapport pour les détails.")

        if self.stats["skipped"] > 0 and not self.config.force:
            print("\n  [i] Utilisez --force pour reconvertir les fichiers existants")

        print("=" * 80 + "\n")
