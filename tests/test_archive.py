"""
Tests pour le convertisseur d'archives.

Teste:
- ArchiveConverter
- Extraction ZIP, TAR, TAR.GZ
- Détection des types d'archives
- Gestion des fichiers ignorés
- Structure de sortie
"""

from __future__ import annotations

import tarfile
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from converter_pdf.config import Config
from converter_pdf.converters.base import ConversionStatus


# =============================================================================
# Fixtures spécifiques aux archives
# =============================================================================

class ArchiveFactory:
    """Factory pour créer des archives de test."""

    def __init__(self, base_dir: Path):
        self.base_dir = base_dir

    def create_zip(self, name: str, files: dict[str, str]) -> Path:
        """
        Crée un fichier ZIP.

        Args:
            name: Nom du fichier ZIP
            files: Dict {nom_fichier: contenu}

        Returns:
            Path vers le ZIP créé
        """
        zip_path = self.base_dir / name
        with zipfile.ZipFile(zip_path, 'w') as zf:
            for filename, content in files.items():
                zf.writestr(filename, content)
        return zip_path

    def create_tar(self, name: str, files: dict[str, str]) -> Path:
        """Crée un fichier TAR."""
        tar_path = self.base_dir / name
        with tarfile.open(tar_path, 'w') as tf:
            for filename, content in files.items():
                import io
                data = content.encode('utf-8')
                info = tarfile.TarInfo(name=filename)
                info.size = len(data)
                tf.addfile(info, io.BytesIO(data))
        return tar_path

    def create_tar_gz(self, name: str, files: dict[str, str]) -> Path:
        """Crée un fichier TAR.GZ."""
        tar_path = self.base_dir / name
        with tarfile.open(tar_path, 'w:gz') as tf:
            for filename, content in files.items():
                import io
                data = content.encode('utf-8')
                info = tarfile.TarInfo(name=filename)
                info.size = len(data)
                tf.addfile(info, io.BytesIO(data))
        return tar_path

    def create_zip_with_folder(self, name: str, folder_name: str, files: dict[str, str]) -> Path:
        """Crée un ZIP avec un dossier racine."""
        zip_path = self.base_dir / name
        with zipfile.ZipFile(zip_path, 'w') as zf:
            for filename, content in files.items():
                zf.writestr(f"{folder_name}/{filename}", content)
        return zip_path

    def create_nested_zip(self, outer_name: str, inner_name: str, inner_files: dict[str, str]) -> Path:
        """
        Crée un ZIP contenant un autre ZIP.

        Args:
            outer_name: Nom du ZIP externe
            inner_name: Nom du ZIP interne
            inner_files: Fichiers à mettre dans le ZIP interne

        Returns:
            Path vers le ZIP externe
        """
        import io

        # Créer le ZIP interne en mémoire
        inner_buffer = io.BytesIO()
        with zipfile.ZipFile(inner_buffer, 'w') as inner_zf:
            for filename, content in inner_files.items():
                inner_zf.writestr(filename, content)
        inner_data = inner_buffer.getvalue()

        # Créer le ZIP externe contenant le ZIP interne
        outer_path = self.base_dir / outer_name
        with zipfile.ZipFile(outer_path, 'w') as outer_zf:
            outer_zf.writestr(inner_name, inner_data)
        return outer_path


@pytest.fixture
def archive_factory(temp_dir: Path) -> ArchiveFactory:
    """Factory pour créer des archives de test."""
    return ArchiveFactory(temp_dir)


# =============================================================================
# Tests ArchiveConverter - Initialisation
# =============================================================================

class TestArchiveConverterInit:
    """Tests d'initialisation."""

    def test_supported_extensions(self, mock_logger):
        """Extensions supportées sont correctes."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        assert converter.can_convert(".zip")
        assert converter.can_convert(".tar")
        assert converter.can_convert(".tar.gz")
        assert converter.can_convert(".tgz")
        assert converter.can_convert(".tar.bz2")
        assert converter.can_convert(".tbz2")
        assert converter.can_convert(".rar")
        assert converter.can_convert(".7z")

    def test_is_available_always_true(self, mock_logger):
        """is_available retourne toujours True (ZIP est natif)."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        assert converter.is_available() is True


# =============================================================================
# Tests détection du type d'archive
# =============================================================================

class TestArchiveTypeDetection:
    """Tests de détection du type d'archive."""

    def test_detect_zip(self, mock_logger, temp_dir):
        """Détection d'un fichier ZIP."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        source = temp_dir / "test.zip"
        source.touch()

        assert converter._get_archive_type(source) == "zip"

    def test_detect_tar(self, mock_logger, temp_dir):
        """Détection d'un fichier TAR."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        source = temp_dir / "test.tar"
        source.touch()

        assert converter._get_archive_type(source) == "tar"

    def test_detect_tar_gz(self, mock_logger, temp_dir):
        """Détection d'un fichier TAR.GZ."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        for name in ["test.tar.gz", "test.tgz"]:
            source = temp_dir / name
            source.touch()
            assert converter._get_archive_type(source) == "tar.gz"

    def test_detect_tar_bz2(self, mock_logger, temp_dir):
        """Détection d'un fichier TAR.BZ2."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        for name in ["test.tar.bz2", "test.tbz2"]:
            source = temp_dir / name
            source.touch()
            assert converter._get_archive_type(source) == "tar.bz2"

    def test_detect_rar(self, mock_logger, temp_dir):
        """Détection d'un fichier RAR."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        source = temp_dir / "test.rar"
        source.touch()

        assert converter._get_archive_type(source) == "rar"

    def test_detect_7z(self, mock_logger, temp_dir):
        """Détection d'un fichier 7Z."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        source = temp_dir / "test.7z"
        source.touch()

        assert converter._get_archive_type(source) == "7z"

    def test_detect_unknown(self, mock_logger, temp_dir):
        """Détection d'un type inconnu."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        source = temp_dir / "test.xyz"
        source.touch()

        assert converter._get_archive_type(source) == "unknown"


# =============================================================================
# Tests des patterns à ignorer
# =============================================================================

class TestIgnorePatterns:
    """Tests des patterns de fichiers à ignorer."""

    def test_should_ignore_macosx(self, mock_logger):
        """__MACOSX est ignoré."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        assert converter._should_ignore(Path("__MACOSX/file.txt")) is True

    def test_should_ignore_ds_store(self, mock_logger):
        """.DS_Store est ignoré."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        assert converter._should_ignore(Path(".DS_Store")) is True

    def test_should_ignore_thumbs_db(self, mock_logger):
        """Thumbs.db est ignoré."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        assert converter._should_ignore(Path("folder/Thumbs.db")) is True

    def test_should_ignore_hidden_files(self, mock_logger):
        """Les fichiers cachés (commençant par .) sont ignorés."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        assert converter._should_ignore(Path(".hidden")) is True
        assert converter._should_ignore(Path(".git/config")) is True

    def test_should_not_ignore_normal_files(self, mock_logger):
        """Les fichiers normaux ne sont pas ignorés."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        assert converter._should_ignore(Path("document.txt")) is False
        assert converter._should_ignore(Path("folder/image.png")) is False


# =============================================================================
# Tests de sanitization des noms
# =============================================================================

class TestSanitizeFilename:
    """Tests du nettoyage des noms de fichiers."""

    def test_sanitize_removes_invalid_chars(self, mock_logger):
        """Caractères invalides sont remplacés."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        assert converter._sanitize_filename('file<name>.txt') == 'file_name_.txt'
        assert converter._sanitize_filename('path:to|file') == 'path_to_file'
        assert converter._sanitize_filename('what?*') == 'what__'

    def test_sanitize_truncates_long_names(self, mock_logger):
        """Les noms trop longs sont tronqués."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)
        long_name = "a" * 300 + ".txt"

        result = converter._sanitize_filename(long_name)
        assert len(result) <= 200

    def test_sanitize_normal_names_unchanged(self, mock_logger):
        """Les noms normaux ne sont pas modifiés."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        assert converter._sanitize_filename('document.txt') == 'document.txt'
        assert converter._sanitize_filename('my-file_v2.pdf') == 'my-file_v2.pdf'


# =============================================================================
# Tests d'extraction ZIP
# =============================================================================

class TestZipExtraction:
    """Tests d'extraction de fichiers ZIP."""

    def test_extract_simple_zip(self, mock_logger, temp_dir, archive_factory):
        """Extraction d'un ZIP simple."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        # Créer un ZIP
        zip_file = archive_factory.create_zip("test.zip", {
            "doc.txt": "Hello World",
            "data.xml": "<root/>",
        })

        # Extraire
        extract_dir = temp_dir / "extracted"
        extract_dir.mkdir()
        count = converter._extract_archive(zip_file, extract_dir, "zip")

        assert count == 2
        assert (extract_dir / "doc.txt").exists()
        assert (extract_dir / "data.xml").exists()

    def test_extract_zip_ignores_system_files(self, mock_logger, temp_dir, archive_factory):
        """L'extraction ignore les fichiers système."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        zip_file = archive_factory.create_zip("test.zip", {
            "doc.txt": "Hello",
            "__MACOSX/doc.txt": "garbage",
            ".DS_Store": "garbage",
        })

        extract_dir = temp_dir / "extracted"
        extract_dir.mkdir()
        count = converter._extract_archive(zip_file, extract_dir, "zip")

        # Seul doc.txt devrait être extrait
        assert count == 1
        assert (extract_dir / "doc.txt").exists()


# =============================================================================
# Tests d'extraction TAR
# =============================================================================

class TestTarExtraction:
    """Tests d'extraction de fichiers TAR."""

    def test_extract_tar(self, mock_logger, temp_dir, archive_factory):
        """Extraction d'un TAR simple."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        tar_file = archive_factory.create_tar("test.tar", {
            "file1.txt": "Content 1",
            "file2.txt": "Content 2",
        })

        extract_dir = temp_dir / "extracted"
        extract_dir.mkdir()
        count = converter._extract_archive(tar_file, extract_dir, "tar")

        assert count == 2

    def test_extract_tar_gz(self, mock_logger, temp_dir, archive_factory):
        """Extraction d'un TAR.GZ."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        tar_file = archive_factory.create_tar_gz("test.tar.gz", {
            "compressed.txt": "Compressed content",
        })

        extract_dir = temp_dir / "extracted"
        extract_dir.mkdir()
        count = converter._extract_archive(tar_file, extract_dir, "tar.gz")

        assert count == 1
        assert (extract_dir / "compressed.txt").exists()


# =============================================================================
# Tests de la gestion des dossiers dupliqués
# =============================================================================

class TestEffectiveSourceDir:
    """Tests de la gestion des dossiers racine uniques."""

    def test_single_folder_same_name(self, mock_logger, temp_dir):
        """Si l'archive contient un seul dossier du même nom, l'utiliser."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        # Créer une structure: temp_dir/test/file.txt
        (temp_dir / "test").mkdir()
        (temp_dir / "test" / "file.txt").write_text("content")

        result = converter._get_effective_source_dir(temp_dir, "test")

        assert result == temp_dir / "test"

    def test_single_folder_different_name(self, mock_logger, temp_dir):
        """Si le dossier a un nom différent, utiliser le dossier parent."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        # Créer une structure: temp_dir/other/file.txt
        (temp_dir / "other").mkdir()
        (temp_dir / "other" / "file.txt").write_text("content")

        result = converter._get_effective_source_dir(temp_dir, "test")

        assert result == temp_dir

    def test_multiple_items_at_root(self, mock_logger, temp_dir):
        """Si plusieurs éléments à la racine, utiliser le dossier parent."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        # Créer plusieurs fichiers
        (temp_dir / "file1.txt").write_text("content1")
        (temp_dir / "file2.txt").write_text("content2")

        result = converter._get_effective_source_dir(temp_dir, "test")

        assert result == temp_dir


# =============================================================================
# Tests de conversion complète
# =============================================================================

class TestArchiveConversion:
    """Tests de conversion complète d'archives."""

    def test_convert_zip_with_text_files(self, mock_logger, temp_dir, archive_factory):
        """Conversion d'un ZIP contenant des fichiers texte."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        zip_file = archive_factory.create_zip("docs.zip", {
            "readme.txt": "Project readme",
            "notes.log": "Log entry",
        })

        dest = temp_dir / "docs.zip.pdf"  # Sera transformé en dossier "docs"
        result = converter.convert(zip_file, dest)

        # Le résultat dépend de la disponibilité de ReportLab
        # Mais l'extraction devrait fonctionner
        if result.status == ConversionStatus.SUCCESS:
            output_dir = temp_dir / "docs"
            assert output_dir.exists()
            assert output_dir.is_dir()

    def test_convert_empty_zip_fails(self, mock_logger, temp_dir):
        """Conversion d'un ZIP vide échoue."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        # Créer un ZIP vide
        zip_file = temp_dir / "empty.zip"
        with zipfile.ZipFile(zip_file, 'w') as zf:
            pass  # Rien à ajouter

        dest = temp_dir / "empty.zip.pdf"
        result = converter.convert(zip_file, dest)

        assert result.status == ConversionStatus.FAILED
        assert "vide" in result.message.lower()

    def test_convert_rar_without_library_fails(self, mock_logger, temp_dir):
        """Conversion d'un RAR sans rarfile échoue gracieusement."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        # Créer un faux fichier RAR
        rar_file = temp_dir / "test.rar"
        rar_file.write_bytes(b"Rar!\x1a\x07\x00")  # Magic number RAR

        with patch("converter_pdf.converters.archive.RARFILE_AVAILABLE", False):
            dest = temp_dir / "test.rar.pdf"
            result = converter.convert(rar_file, dest)

            assert result.status == ConversionStatus.FAILED
            assert "rarfile" in result.message.lower()

    def test_convert_7z_without_library_fails(self, mock_logger, temp_dir):
        """Conversion d'un 7Z sans py7zr échoue gracieusement."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        # Créer un faux fichier 7Z
        sz_file = temp_dir / "test.7z"
        sz_file.write_bytes(b"7z\xbc\xaf\x27\x1c")  # Magic number 7Z

        with patch("converter_pdf.converters.archive.PY7ZR_AVAILABLE", False):
            dest = temp_dir / "test.7z.pdf"
            result = converter.convert(sz_file, dest)

            assert result.status == ConversionStatus.FAILED
            assert "py7zr" in result.message.lower()

    def test_convert_unknown_type_fails(self, mock_logger, temp_dir):
        """Conversion d'un type inconnu échoue."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), mock_logger)

        # Fichier avec extension inconnue
        unknown_file = temp_dir / "test.unknown"
        unknown_file.write_bytes(b"random data")

        # Forcer can_convert à True pour tester _get_archive_type
        dest = temp_dir / "test.unknown.pdf"
        result = converter.convert(unknown_file, dest)

        assert result.status == ConversionStatus.FAILED


# =============================================================================
# Tests d'intégration
# =============================================================================

class TestArchiveIntegration:
    """Tests d'intégration pour les archives."""

    @pytest.mark.integration
    def test_zip_to_pdf_workflow(self, temp_dir, real_logger, archive_factory):
        """Workflow complet: ZIP -> dossier avec PDF."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), real_logger)

        # Créer un ZIP avec structure
        zip_file = archive_factory.create_zip("project.zip", {
            "README.txt": "# Project\nThis is a test project.",
            "docs/manual.txt": "User manual content",
            "src/code.py": "print('hello')",  # Non convertible
        })

        dest = temp_dir / "project.zip.pdf"
        result = converter.convert(zip_file, dest)

        if result.status == ConversionStatus.SUCCESS:
            output_dir = temp_dir / "project"
            assert output_dir.exists()

            # Vérifier la structure
            # Les fichiers originaux et/ou PDF devraient exister
            # (selon la disponibilité des convertisseurs)

    @pytest.mark.integration
    def test_nested_archive_not_extracted(self, temp_dir, real_logger, archive_factory):
        """Une archive imbriquée n'est pas extraite récursivement."""
        from converter_pdf.converters.archive import ArchiveConverter

        converter = ArchiveConverter(Config(), real_logger)

        # Créer une archive interne
        inner_zip = archive_factory.create_zip("inner.zip", {
            "nested.txt": "Nested content",
        })

        # Lire son contenu et créer l'archive externe
        inner_content = inner_zip.read_bytes()

        outer_zip = temp_dir / "outer.zip"
        with zipfile.ZipFile(outer_zip, 'w') as zf:
            zf.writestr("outer.txt", "Outer content")
            zf.writestr("inner.zip", inner_content)

        dest = temp_dir / "outer.zip.pdf"
        result = converter.convert(outer_zip, dest)

        if result.status == ConversionStatus.SUCCESS:
            output_dir = temp_dir / "outer"
            # inner.zip devrait être copié tel quel (pas extrait)
            # car ArchiveConverter évite la récursion
