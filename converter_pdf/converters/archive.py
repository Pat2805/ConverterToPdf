"""
Convertisseur de fichiers compressés (ZIP, RAR, 7Z, TAR, GZ).

Décompresse les archives et convertit leur contenu en PDF.
"""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import time
import zipfile
import tarfile
from pathlib import Path
from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus

# Import conditionnel de rarfile (optionnel)
try:
    import rarfile
    RARFILE_AVAILABLE = True
except ImportError:
    RARFILE_AVAILABLE = False
    rarfile = None  # type: ignore

# Import conditionnel de py7zr (optionnel)
try:
    import py7zr
    PY7ZR_AVAILABLE = True
except ImportError:
    PY7ZR_AVAILABLE = False
    py7zr = None  # type: ignore

if TYPE_CHECKING:
    from ..config import Config
    from ..logger import ConverterLogger


class ArchiveConverter(BaseConverter):
    """
    Convertisseur de fichiers compressés en PDF.

    Stratégie:
    1. Extraire l'archive dans un dossier temporaire
    2. Convertir chaque fichier en PDF si possible
    3. Créer un dossier de sortie avec les PDF et fichiers non convertibles

    Structure de sortie:
        archive.zip.pdf/
        ├── document.docx.pdf     # Fichier converti
        ├── image.jpg.pdf         # Image convertie
        ├── subfolder/            # Structure préservée
        │   └── data.xlsx.pdf
        └── binary.exe            # Fichier non convertible (conservé)

    Formats supportés:
    - ZIP (natif Python)
    - TAR, TAR.GZ, TGZ, TAR.BZ2 (natif Python)
    - RAR (nécessite rarfile + unrar)
    - 7Z (nécessite py7zr)
    """

    name = "archive"
    supported_extensions = [
        ".zip",
        ".tar", ".tar.gz", ".tgz", ".tar.bz2", ".tbz2",
        ".rar",
        ".7z",
    ]

    # Extensions convertibles en PDF
    CONVERTIBLE_EXTENSIONS = {
        ".doc", ".docx", ".rtf", ".odt",
        ".xls", ".xlsx", ".xlsm", ".xlsb",
        ".ppt", ".pptx",
        ".txt", ".log",
        ".htm", ".html",
        ".xml",
        ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp",
        ".msg",
        ".pdf",  # Copier tel quel
    }

    # Fichiers à ignorer (système, cache, etc.)
    IGNORE_PATTERNS = {
        "__MACOSX",
        ".DS_Store",
        "Thumbs.db",
        "desktop.ini",
        ".git",
        ".svn",
        "__pycache__",
    }

    def __init__(self, config: "Config", logger: "ConverterLogger"):
        super().__init__(config, logger)
        self._converters_cache = None

    def _get_converters(self):
        """Retourne la chaîne de convertisseurs (lazy loading)."""
        if self._converters_cache is None:
            from . import get_converter_chain
            self._converters_cache = get_converter_chain(self.config, self.logger)
        return self._converters_cache

    def is_available(self) -> bool:
        """Vérifie qu'au moins ZIP est disponible (toujours vrai)."""
        return True

    def can_convert(self, extension: str) -> bool:
        """Vérifie si l'extension est supportée."""
        ext = extension.lower()
        # Gérer les extensions composées
        if ext in self.supported_extensions:
            return True
        # Vérifier .tar.gz, .tar.bz2
        for supported_ext in self.supported_extensions:
            if ext.endswith(supported_ext):
                return True
        return False

    def _sanitize_filename(self, filename: str) -> str:
        """Nettoie un nom de fichier."""
        sanitized = re.sub(r'[<>:"|?*]', '_', filename)
        if len(sanitized) > 200:
            sanitized = sanitized[:200]
        return sanitized.strip()

    def _should_ignore(self, path: Path) -> bool:
        """Vérifie si un fichier/dossier doit être ignoré."""
        for part in path.parts:
            if part in self.IGNORE_PATTERNS:
                return True
            if part.startswith('.'):
                return True
        return False

    def _get_archive_type(self, source: Path) -> str:
        """Détermine le type d'archive."""
        name = source.name.lower()
        if name.endswith('.zip'):
            return 'zip'
        elif name.endswith(('.tar.gz', '.tgz')):
            return 'tar.gz'
        elif name.endswith(('.tar.bz2', '.tbz2')):
            return 'tar.bz2'
        elif name.endswith('.tar'):
            return 'tar'
        elif name.endswith('.rar'):
            return 'rar'
        elif name.endswith('.7z'):
            return '7z'
        return 'unknown'

    def _get_effective_source_dir(self, temp_dir: Path, archive_stem: str) -> Path:
        """
        Détermine le répertoire source effectif après extraction.

        Si l'archive contient uniquement un dossier racine du même nom
        (ou très similaire), on utilise ce dossier pour éviter la duplication.
        Ex: test.zip contenant uniquement test/ -> on utilise temp_dir/test

        Args:
            temp_dir: Dossier d'extraction temporaire
            archive_stem: Nom de l'archive sans extension

        Returns:
            Le répertoire source à utiliser pour le traitement
        """
        # Lister le contenu du dossier temporaire (sans fichiers ignorés)
        items = [
            item for item in temp_dir.iterdir()
            if not self._should_ignore(item)
        ]

        # Si un seul élément et c'est un dossier
        if len(items) == 1 and items[0].is_dir():
            single_folder = items[0]
            folder_name = single_folder.name.lower()
            archive_name = archive_stem.lower()

            # Vérifier si le nom est identique ou très similaire
            # (ignorer la casse, tirets/underscores)
            def normalize(s: str) -> str:
                return s.replace('-', '').replace('_', '').replace(' ', '')

            if normalize(folder_name) == normalize(archive_name):
                self.logger.debug(
                    f"  Archive contient un seul dossier '{single_folder.name}' "
                    f"(même nom que l'archive) -> évite la duplication"
                )
                return single_folder

        return temp_dir

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """
        Convertit une archive en dossier de PDF.

        Args:
            source: Fichier archive
            dest: Chemin de destination (sera un dossier)

        Returns:
            Résultat de la conversion
        """
        start = time.time()
        archive_type = self._get_archive_type(source)

        # Vérifier la disponibilité du décompresseur
        if archive_type == 'rar' and not RARFILE_AVAILABLE:
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="rarfile non installé (pip install rarfile)",
            )

        if archive_type == '7z' and not PY7ZR_AVAILABLE:
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="py7zr non installé (pip install py7zr)",
            )

        if archive_type == 'unknown':
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message=f"Type d'archive non reconnu: {source.suffix}",
            )

        # Créer un dossier temporaire pour l'extraction
        temp_dir = None
        try:
            temp_dir = Path(tempfile.mkdtemp(prefix="converter_archive_"))

            # Extraire l'archive
            self.logger.debug(f"Extraction de {source.name} ({archive_type})")
            extracted_count = self._extract_archive(source, temp_dir, archive_type)

            if extracted_count == 0:
                return ConversionResult(
                    status=ConversionStatus.FAILED,
                    source=source,
                    dest=None,
                    duration=time.time() - start,
                    method=self.name,
                    message="Archive vide ou erreur d'extraction",
                )

            self.logger.info(f"  {extracted_count} fichier(s) extrait(s)")

            # Déterminer le nom du dossier de sortie (sans extension d'archive)
            # Ex: archive.zip -> archive/, data.tar.gz -> data/
            archive_stem = source.stem
            # Gérer les doubles extensions (.tar.gz, .tar.bz2)
            if Path(archive_stem).suffix in ('.tar',):
                archive_stem = Path(archive_stem).stem

            # Vérifier si l'archive contient uniquement un dossier du même nom
            # Ex: test.zip contenant uniquement test/ -> utiliser temp_dir/test comme source
            source_dir = self._get_effective_source_dir(temp_dir, archive_stem)

            # Créer le dossier de sortie
            output_folder = dest.parent / archive_stem
            output_folder.mkdir(parents=True, exist_ok=True)

            # Convertir les fichiers extraits
            converted, failed, kept = self._process_extracted_files(
                source_dir, output_folder
            )

            self.logger.info(
                f"  Résultat: {converted} converti(s), {kept} conservé(s), {failed} échec(s)"
            )

            return ConversionResult(
                status=ConversionStatus.SUCCESS,
                source=source,
                dest=output_folder,
                duration=time.time() - start,
                method=f"{self.name}_{archive_type}",
                message=f"{converted} PDF, {kept} fichiers conservés",
            )

        except Exception as e:
            self.logger.error(f"Erreur archive: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                exception=e,
            )
        finally:
            # Nettoyer le dossier temporaire
            if temp_dir and temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except Exception:
                    pass

    def _extract_archive(self, source: Path, dest_dir: Path, archive_type: str) -> int:
        """
        Extrait une archive.

        Returns:
            Nombre de fichiers extraits
        """
        count = 0

        if archive_type == 'zip':
            with zipfile.ZipFile(source, 'r') as zf:
                for member in zf.namelist():
                    if not self._should_ignore(Path(member)):
                        zf.extract(member, dest_dir)
                        count += 1

        elif archive_type in ('tar', 'tar.gz', 'tar.bz2'):
            mode = 'r'
            if archive_type == 'tar.gz':
                mode = 'r:gz'
            elif archive_type == 'tar.bz2':
                mode = 'r:bz2'

            with tarfile.open(source, mode) as tf:
                for member in tf.getmembers():
                    if not self._should_ignore(Path(member.name)):
                        tf.extract(member, dest_dir)
                        count += 1

        elif archive_type == 'rar' and RARFILE_AVAILABLE:
            with rarfile.RarFile(source, 'r') as rf:
                for member in rf.namelist():
                    if not self._should_ignore(Path(member)):
                        rf.extract(member, dest_dir)
                        count += 1

        elif archive_type == '7z' and PY7ZR_AVAILABLE:
            with py7zr.SevenZipFile(source, 'r') as szf:
                # py7zr extrait tout d'un coup
                szf.extractall(dest_dir)
                # Compter les fichiers
                for root, dirs, files in os.walk(dest_dir):
                    for f in files:
                        if not self._should_ignore(Path(root) / f):
                            count += 1

        return count

    def _process_extracted_files(
        self,
        source_dir: Path,
        output_dir: Path,
    ) -> tuple[int, int, int]:
        """
        Traite les fichiers extraits.

        Returns:
            Tuple (convertis, échecs, conservés)
        """
        converters = self._get_converters()
        converted = 0
        failed = 0
        kept = 0

        # Parcourir tous les fichiers
        for root, dirs, files in os.walk(source_dir):
            # Filtrer les dossiers à ignorer
            dirs[:] = [d for d in dirs if not self._should_ignore(Path(d))]

            rel_root = Path(root).relative_to(source_dir)

            for filename in files:
                if self._should_ignore(Path(filename)):
                    continue

                source_file = Path(root) / filename
                if not source_file.is_file():
                    continue

                # Créer le chemin de destination en préservant la structure
                if rel_root == Path('.'):
                    dest_subdir = output_dir
                else:
                    dest_subdir = output_dir / rel_root
                    dest_subdir.mkdir(parents=True, exist_ok=True)

                ext = source_file.suffix.lower()
                safe_name = self._sanitize_filename(filename)

                # Si c'est déjà un PDF, copier tel quel
                if ext == '.pdf':
                    dest_file = dest_subdir / safe_name
                    shutil.copy2(source_file, dest_file)
                    kept += 1
                    continue

                # Copier d'abord l'original dans le dossier de sortie
                dest_file = dest_subdir / safe_name
                shutil.copy2(source_file, dest_file)

                # Tenter de convertir si extension supportée
                if ext in self.CONVERTIBLE_EXTENSIONS:
                    pdf_dest = dest_subdir / (safe_name + ".pdf")
                    conversion_success = False

                    for converter in converters:
                        # Éviter récursion sur archives
                        if converter.name == self.name:
                            continue
                        if not converter.can_convert(ext):
                            continue
                        if not converter.is_available():
                            continue

                        self.logger.debug(f"  Conversion: {filename}")
                        result = converter.convert(dest_file, pdf_dest)

                        if result.status == ConversionStatus.SUCCESS:
                            converted += 1
                            conversion_success = True
                            # Supprimer l'original uniquement si delete_source est activé
                            if self.config.delete_source:
                                try:
                                    dest_file.unlink()
                                    self.logger.debug(f"  -> Original supprimé: {safe_name}")
                                except Exception:
                                    pass
                            break

                    if not conversion_success:
                        # Échec de conversion, l'original est déjà copié
                        failed += 1
                else:
                    # Extension non convertible, l'original est déjà copié
                    kept += 1

        return converted, failed, kept
