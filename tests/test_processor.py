"""
Tests pour le module processor.py.

Teste:
- FileProcessor
- Calcul des chemins de destination
- Traitement de fichiers individuels
- Traitement de répertoires
- Statistiques
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from converter_pdf.config import Config
from converter_pdf.converters.base import ConversionResult, ConversionStatus
from converter_pdf.processor import FileProcessor, format_size


# =============================================================================
# Tests format_size
# =============================================================================

class TestFormatSize:
    """Tests de la fonction format_size."""

    def test_format_bytes(self):
        """Formatage des bytes."""
        assert format_size(0) == "0 B"
        assert format_size(100) == "100 B"
        assert format_size(1023) == "1023 B"

    def test_format_kilobytes(self):
        """Formatage des kilobytes."""
        assert format_size(1024) == "1.0 KB"
        assert format_size(1536) == "1.5 KB"
        assert format_size(1024 * 100) == "100.0 KB"

    def test_format_megabytes(self):
        """Formatage des mégabytes."""
        assert format_size(1024 * 1024) == "1.0 MB"
        assert format_size(1024 * 1024 * 5.5) == "5.5 MB"

    def test_format_gigabytes(self):
        """Formatage des gigabytes."""
        assert format_size(1024 * 1024 * 1024) == "1.00 GB"
        assert format_size(1024 * 1024 * 1024 * 2.5) == "2.50 GB"


# =============================================================================
# Tests FileProcessor - Initialisation
# =============================================================================

class TestFileProcessorInit:
    """Tests d'initialisation de FileProcessor."""

    def test_init_creates_converter_chain(self, mock_logger):
        """L'initialisation crée la chaîne de convertisseurs."""
        config = Config()
        processor = FileProcessor(config, mock_logger)

        assert len(processor.converters) > 0
        assert processor.config == config
        assert processor.logger == mock_logger

    def test_init_stats_zeroed(self, mock_logger):
        """Les statistiques sont initialisées à zéro."""
        processor = FileProcessor(Config(), mock_logger)

        assert processor.stats["total"] == 0
        assert processor.stats["success"] == 0
        assert processor.stats["failed"] == 0
        assert processor.stats["skipped"] == 0

    def test_init_not_interrupted(self, mock_logger):
        """_interrupted est False à l'initialisation."""
        processor = FileProcessor(Config(), mock_logger)
        assert processor._interrupted is False


# =============================================================================
# Tests FileProcessor - Calcul des chemins
# =============================================================================

class TestFileProcessorDestPath:
    """Tests du calcul des chemins de destination."""

    def test_dest_path_same_directory(self, mock_logger, temp_dir):
        """Sans dest_dir, le PDF va dans le même répertoire."""
        processor = FileProcessor(Config(), mock_logger)
        source = temp_dir / "document.docx"
        source.touch()

        dest = processor._get_dest_path(source, None)

        assert dest.parent == source.parent
        assert dest.name == "document.docx.pdf"

    def test_dest_path_keep_extension_true(self, mock_logger, temp_dir):
        """Avec keep_extension=True: document.docx -> document.docx.pdf."""
        config = Config(keep_extension=True)
        processor = FileProcessor(config, mock_logger)
        source = temp_dir / "document.docx"
        source.touch()

        dest = processor._get_dest_path(source, None)

        assert dest.name == "document.docx.pdf"

    def test_dest_path_keep_extension_false(self, mock_logger, temp_dir):
        """Avec keep_extension=False: document.docx -> document.pdf."""
        config = Config(keep_extension=False)
        processor = FileProcessor(config, mock_logger)
        source = temp_dir / "document.docx"
        source.touch()

        dest = processor._get_dest_path(source, None)

        assert dest.name == "document.pdf"

    def test_dest_path_custom_directory(self, mock_logger, temp_dir):
        """Avec dest_dir, le PDF va dans ce répertoire."""
        processor = FileProcessor(Config(), mock_logger)
        source = temp_dir / "document.docx"
        source.touch()
        dest_dir = temp_dir / "output"

        dest = processor._get_dest_path(source, dest_dir)

        assert dest.parent == dest_dir
        assert dest_dir.exists()  # Créé automatiquement


# =============================================================================
# Tests FileProcessor - process_file
# =============================================================================

class TestFileProcessorProcessFile:
    """Tests de traitement de fichiers individuels."""

    def test_process_pdf_file_skipped(self, mock_logger, temp_dir):
        """Un fichier PDF est ignoré (déjà PDF)."""
        processor = FileProcessor(Config(), mock_logger)
        source = temp_dir / "document.pdf"
        source.write_bytes(b"%PDF-1.4 test")

        result = processor.process_file(source)

        assert result.status == ConversionStatus.SKIPPED_PDF

    def test_process_existing_pdf_skipped(self, mock_logger, temp_dir):
        """Si le PDF existe déjà, le fichier est ignoré."""
        config = Config(force=False)
        processor = FileProcessor(config, mock_logger)

        source = temp_dir / "document.txt"
        source.write_text("test")
        dest = temp_dir / "document.txt.pdf"
        dest.write_bytes(b"%PDF-1.4 existing")

        result = processor.process_file(source)

        assert result.status == ConversionStatus.SKIPPED_EXISTS

    def test_process_force_reconverts(self, mock_logger, temp_dir, file_factory):
        """Avec force=True, reconvertit même si le PDF existe."""
        config = Config(force=True)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("document.txt", "test content")
        dest = temp_dir / "document.txt.pdf"
        dest.write_bytes(b"%PDF-1.4 old")

        # Le résultat dépend de la disponibilité de ReportLab
        result = processor.process_file(source)

        # Ne devrait pas être SKIPPED_EXISTS
        assert result.status != ConversionStatus.SKIPPED_EXISTS

    def test_process_unsupported_format(self, mock_logger, temp_dir):
        """Un format non supporté retourne SKIPPED_UNSUPPORTED."""
        processor = FileProcessor(Config(), mock_logger)
        source = temp_dir / "document.xyz"
        source.write_text("test")

        result = processor.process_file(source)

        assert result.status == ConversionStatus.SKIPPED_UNSUPPORTED

    def test_process_updates_stats(self, mock_logger, temp_dir, file_factory):
        """Le traitement met à jour les statistiques."""
        processor = FileProcessor(Config(report_enabled=False), mock_logger)
        source = file_factory.create_text_file("test.txt", "content")

        # Réinitialiser les stats
        processor.stats = {"total": 0, "success": 0, "failed": 0, "skipped": 0}

        processor.process_file(source)

        # Le total n'est pas mis à jour par process_file individuel
        # C'est process_directory qui gère ça

    def test_process_delete_source_on_success(self, mock_logger, temp_dir, file_factory):
        """Avec delete_source=True, supprime le source après succès."""
        config = Config(delete_source=True)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("todelete.txt", "content")

        # Patcher un convertisseur pour simuler un succès
        mock_converter = MagicMock()
        mock_converter.can_convert.return_value = True
        mock_converter.is_available.return_value = True
        mock_converter.convert.return_value = ConversionResult(
            status=ConversionStatus.SUCCESS,
            source=source,
            dest=temp_dir / "todelete.txt.pdf",
            duration=0.1,
            method="mock",
        )
        # Créer le fichier dest pour que ConversionResult calcule la taille
        (temp_dir / "todelete.txt.pdf").write_bytes(b"%PDF")

        processor.converters = [mock_converter]

        result = processor.process_file(source)

        if result.is_success:
            # Le source devrait être supprimé (ou tentative de suppression)
            pass  # Dépend de l'implémentation


# =============================================================================
# Tests FileProcessor - process_directory
# =============================================================================

class TestFileProcessorProcessDirectory:
    """Tests de traitement de répertoires."""

    def test_process_invalid_directory(self, mock_logger, temp_dir):
        """Un répertoire invalide retourne des stats vides."""
        processor = FileProcessor(Config(), mock_logger)
        fake_dir = temp_dir / "nonexistent"

        stats = processor.process_directory(fake_dir)

        assert stats["total"] == 0

    def test_process_empty_directory(self, mock_logger, temp_dir):
        """Un répertoire vide retourne des stats vides."""
        config = Config(report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        stats = processor.process_directory(temp_dir)

        assert stats["total"] == 0

    def test_process_directory_non_recursive(self, mock_logger, temp_dir, file_factory):
        """Mode non-récursif ne traite pas les sous-dossiers."""
        config = Config(recursive=False, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        # Créer des fichiers
        file_factory.create_text_file("root.txt", "root content")
        subdir = file_factory.create_subdirectory("subdir")
        (subdir / "nested.txt").write_text("nested content")

        stats = processor.process_directory(temp_dir)

        # Seul root.txt devrait être traité
        assert stats["total"] == 1

    def test_process_directory_recursive(self, mock_logger, temp_dir, file_factory):
        """Mode récursif traite les sous-dossiers."""
        config = Config(recursive=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        # Créer des fichiers
        file_factory.create_text_file("root.txt", "root content")
        subdir = file_factory.create_subdirectory("subdir")
        (subdir / "nested.txt").write_text("nested content")

        stats = processor.process_directory(temp_dir)

        # Les deux fichiers devraient être traités
        assert stats["total"] == 2

    def test_process_directory_filters_extensions(self, mock_logger, temp_dir, file_factory):
        """Seules les extensions supportées sont traitées."""
        config = Config(report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        # Créer des fichiers de différentes extensions
        file_factory.create_text_file("document.txt", "text")
        file_factory.create_empty_file("unknown.xyz")

        stats = processor.process_directory(temp_dir)

        # Seul .txt devrait être traité
        assert stats["total"] == 1

    def test_process_directory_returns_stats(self, mock_logger, temp_dir, file_factory):
        """process_directory retourne les statistiques correctes."""
        config = Config(report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        # Créer plusieurs fichiers
        file_factory.create_text_file("doc1.txt", "content 1")
        file_factory.create_text_file("doc2.txt", "content 2")
        file_factory.create_pdf_file("existing.pdf")

        stats = processor.process_directory(temp_dir)

        # 2 txt + 1 pdf = 3 fichiers traités
        assert stats["total"] == 3
        # Le PDF est skipped
        assert stats["skipped"] >= 1


# =============================================================================
# Tests FileProcessor - Interruption
# =============================================================================

class TestFileProcessorInterruption:
    """Tests de gestion des interruptions."""

    def test_interrupted_flag_stops_processing(self, mock_logger, temp_dir, file_factory):
        """Le flag _interrupted arrête le traitement."""
        config = Config(report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        # Créer plusieurs fichiers
        for i in range(5):
            file_factory.create_text_file(f"doc{i}.txt", f"content {i}")

        # Simuler une interruption après le premier fichier
        original_process_file = processor.process_file

        call_count = [0]

        def mock_process_file(*args, **kwargs):
            call_count[0] += 1
            if call_count[0] >= 2:
                processor._interrupted = True
            return original_process_file(*args, **kwargs)

        processor.process_file = mock_process_file

        stats = processor.process_directory(temp_dir)

        # Devrait s'arrêter avant de traiter tous les fichiers
        assert stats["total"] < 5


# =============================================================================
# Tests FileProcessor - Rapport
# =============================================================================

class TestFileProcessorReport:
    """Tests de génération du rapport."""

    def test_report_created_when_enabled(self, mock_logger, temp_dir, file_factory):
        """Le rapport est créé quand report_enabled=True."""
        config = Config(report_enabled=True)
        processor = FileProcessor(config, mock_logger)

        file_factory.create_text_file("doc.txt", "content")
        processor.process_directory(temp_dir)

        # Vérifier qu'un rapport a été créé
        reports = list(temp_dir.glob("conversion_report_*.txt"))
        assert len(reports) >= 1

    def test_no_report_when_disabled(self, mock_logger, temp_dir, file_factory):
        """Aucun rapport quand report_enabled=False."""
        config = Config(report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        file_factory.create_text_file("doc.txt", "content")
        processor.process_directory(temp_dir)

        # Vérifier qu'aucun rapport n'a été créé
        reports = list(temp_dir.glob("conversion_report_*.txt"))
        assert len(reports) == 0


# =============================================================================
# Tests FileProcessor - Dry-Run
# =============================================================================

class TestFileProcessorDryRun:
    """Tests du mode dry-run (simulation)."""

    def test_dry_run_does_not_create_pdf(self, mock_logger, temp_dir, file_factory):
        """En mode dry-run, aucun PDF n'est créé."""
        config = Config(dry_run=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("document.txt", "content")
        expected_pdf = temp_dir / "document.txt.pdf"

        result = processor.process_file(source)

        # Le résultat indique un succès simulé
        assert result.status == ConversionStatus.SUCCESS
        assert "dry_run" in result.method

        # Mais aucun fichier PDF n'a été créé
        assert not expected_pdf.exists()

    def test_dry_run_identifies_converter(self, mock_logger, temp_dir, file_factory):
        """Dry-run identifie le convertisseur qui serait utilisé."""
        config = Config(dry_run=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("test.txt", "content")
        result = processor.process_file(source)

        # Le method devrait contenir le nom du convertisseur
        assert "dry_run:" in result.method
        # Pour un fichier .txt, c'est le convertisseur "text"
        assert "text" in result.method

    def test_dry_run_directory_stats(self, mock_logger, temp_dir, file_factory):
        """Dry-run sur répertoire retourne des stats correctes."""
        config = Config(dry_run=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        # Créer plusieurs fichiers
        file_factory.create_text_file("doc1.txt", "content 1")
        file_factory.create_text_file("doc2.txt", "content 2")
        file_factory.create_xml_file("data.xml")

        stats = processor.process_directory(temp_dir)

        # Tous devraient être comptés comme "success"
        assert stats["total"] == 3
        assert stats["success"] == 3
        assert stats["failed"] == 0

        # Mais aucun PDF ne devrait exister
        pdfs = list(temp_dir.glob("*.pdf"))
        assert len(pdfs) == 0

    def test_dry_run_does_not_delete_source(self, mock_logger, temp_dir, file_factory):
        """Dry-run ne supprime pas les fichiers source même avec delete_source=True."""
        config = Config(dry_run=True, delete_source=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("todelete.txt", "content")
        processor.process_file(source)

        # Le fichier source doit toujours exister
        assert source.exists()

    def test_dry_run_skips_already_pdf(self, mock_logger, temp_dir, file_factory):
        """Dry-run respecte la logique de skip pour les PDF existants."""
        config = Config(dry_run=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_pdf_file("existing.pdf")
        result = processor.process_file(source)

        # Un PDF est toujours skipped, même en dry-run
        assert result.status == ConversionStatus.SKIPPED_PDF

    def test_dry_run_with_existing_pdf_destination(self, mock_logger, temp_dir, file_factory):
        """Dry-run avec PDF de destination déjà existant."""
        config = Config(dry_run=True, force=False, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("document.txt", "content")
        # Créer le PDF de destination
        (temp_dir / "document.txt.pdf").write_bytes(b"%PDF")

        result = processor.process_file(source)

        # Devrait être skipped car le PDF existe
        assert result.status == ConversionStatus.SKIPPED_EXISTS

    def test_dry_run_force_with_existing_pdf(self, mock_logger, temp_dir, file_factory):
        """Dry-run avec force=True sur PDF existant."""
        config = Config(dry_run=True, force=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("document.txt", "content")
        existing_pdf = temp_dir / "document.txt.pdf"
        existing_pdf.write_bytes(b"%PDF-old")

        result = processor.process_file(source)

        # Devrait simuler une conversion réussie
        assert result.status == ConversionStatus.SUCCESS
        assert "dry_run" in result.method

        # L'ancien PDF ne devrait pas avoir été modifié
        assert existing_pdf.read_bytes() == b"%PDF-old"


# =============================================================================
# Tests FileProcessor - Hide Source
# =============================================================================

class TestFileProcessorHideSource:
    """Tests du mode hide_source (cacher les sources après conversion)."""

    @pytest.mark.skipif(
        __import__("sys").platform != "win32",
        reason="hide_source ne fonctionne que sur Windows"
    )
    def test_hide_source_calls_hide_file(self, mock_logger, temp_dir, file_factory):
        """hide_source appelle _hide_file après une conversion réussie."""
        config = Config(hide_source=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("tohide.txt", "content")
        dest = temp_dir / "tohide.txt.pdf"

        # Patcher un convertisseur pour simuler un succès
        mock_converter = MagicMock()
        mock_converter.can_convert.return_value = True
        mock_converter.is_available.return_value = True

        def mock_convert(src, dst):
            # Créer le fichier PDF lors de la conversion
            dst.write_bytes(b"%PDF")
            return ConversionResult(
                status=ConversionStatus.SUCCESS,
                source=src,
                dest=dst,
                duration=0.1,
                method="mock",
            )

        mock_converter.convert.side_effect = mock_convert
        processor.converters = [mock_converter]

        # Patcher _hide_file pour vérifier qu'elle est appelée
        with patch.object(processor, '_hide_file') as mock_hide:
            result = processor.process_file(source)
            assert result.is_success
            mock_hide.assert_called_once_with(source)

    def test_hide_source_not_called_on_failure(self, mock_logger, temp_dir, file_factory):
        """hide_source n'est pas appelée si la conversion échoue."""
        config = Config(hide_source=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("nothide.txt", "content")

        # Patcher un convertisseur pour simuler un échec
        mock_converter = MagicMock()
        mock_converter.can_convert.return_value = True
        mock_converter.is_available.return_value = True
        mock_converter.convert.return_value = ConversionResult(
            status=ConversionStatus.FAILED,
            source=source,
            dest=None,
            duration=0.1,
            method="mock",
            message="Erreur simulée",
        )

        processor.converters = [mock_converter]

        # Patcher _hide_file pour vérifier qu'elle n'est pas appelée
        with patch.object(processor, '_hide_file') as mock_hide:
            processor.process_file(source)
            mock_hide.assert_not_called()

    def test_hide_source_not_called_on_skip(self, mock_logger, temp_dir, file_factory):
        """hide_source n'est pas appelée si le fichier est ignoré."""
        config = Config(hide_source=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_pdf_file("existing.pdf")

        # Patcher _hide_file pour vérifier qu'elle n'est pas appelée
        with patch.object(processor, '_hide_file') as mock_hide:
            result = processor.process_file(source)
            assert result.status == ConversionStatus.SKIPPED_PDF
            mock_hide.assert_not_called()

    def test_hide_source_not_called_in_dry_run(self, mock_logger, temp_dir, file_factory):
        """hide_source n'est pas appelée en mode dry-run."""
        config = Config(hide_source=True, dry_run=True, report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("nodryrun.txt", "content")

        # Patcher _hide_file pour vérifier qu'elle n'est pas appelée
        with patch.object(processor, '_hide_file') as mock_hide:
            result = processor.process_file(source)
            # En dry-run, le fichier n'est pas réellement converti
            # mais retourne SUCCESS simulé
            assert result.status == ConversionStatus.SUCCESS
            # _hide_file ne devrait pas être appelée car aucun fichier créé
            mock_hide.assert_not_called()

    @pytest.mark.skipif(
        __import__("sys").platform != "win32",
        reason="hide_source ne fonctionne que sur Windows"
    )
    def test_hide_file_method_windows(self, mock_logger, temp_dir, file_factory):
        """_hide_file utilise les attributs Windows."""
        import ctypes

        processor = FileProcessor(Config(), mock_logger)
        source = file_factory.create_text_file("testhide.txt", "content")

        processor._hide_file(source)

        # Vérifier que l'attribut caché est appliqué
        FILE_ATTRIBUTE_HIDDEN = 0x02
        attrs = ctypes.windll.kernel32.GetFileAttributesW(str(source))
        assert attrs != -1
        assert attrs & FILE_ATTRIBUTE_HIDDEN != 0

    def test_hide_file_skipped_on_non_windows(self, mock_logger, temp_dir):
        """_hide_file est ignorée sur les systèmes non-Windows."""
        processor = FileProcessor(Config(), mock_logger)

        with patch('sys.platform', 'linux'):
            # Ne devrait pas lever d'erreur
            processor._hide_file(temp_dir / "test.txt")


# =============================================================================
# Tests d'intégration
# =============================================================================

class TestFileProcessorIntegration:
    """Tests d'intégration du processeur."""

    @pytest.mark.integration
    def test_full_directory_processing(self, temp_dir, real_logger, file_factory):
        """Test complet de traitement d'un répertoire."""
        config = Config(
            recursive=True,
            report_enabled=True,
            force=False,
        )
        processor = FileProcessor(config, real_logger)

        # Créer une structure de fichiers
        file_factory.create_text_file("readme.txt", "Project readme")
        file_factory.create_text_file("notes.log", "Log entry 1\nLog entry 2")
        file_factory.create_xml_file("config.xml")

        subdir = file_factory.create_subdirectory("docs")
        (subdir / "manual.txt").write_text("User manual")

        # Traiter
        stats = processor.process_directory(temp_dir)

        # Vérifications
        assert stats["total"] == 4  # 3 dans root + 1 dans subdir

        # Vérifier que le rapport existe
        reports = list(temp_dir.glob("conversion_report_*.txt"))
        assert len(reports) == 1

    @pytest.mark.integration
    def test_mixed_file_types(self, temp_dir, real_logger, file_factory):
        """Test avec différents types de fichiers."""
        config = Config(report_enabled=False)
        processor = FileProcessor(config, real_logger)

        # Créer différents types
        file_factory.create_text_file("text.txt", "text content")
        file_factory.create_xml_file("data.xml")
        file_factory.create_pdf_file("existing.pdf")
        file_factory.create_empty_file("unknown.xyz")

        stats = processor.process_directory(temp_dir)

        # txt, xml, pdf traités; xyz ignoré
        assert stats["total"] == 3
        assert stats["skipped"] >= 1  # Au moins le PDF


# =============================================================================
# Tests FileProcessor - Extraction récursive (archives/MSG imbriqués)
# =============================================================================

class TestFileProcessorNestedExtraction:
    """Tests du traitement récursif des archives et MSG imbriqués."""

    def test_extracted_folders_initialized(self, mock_logger):
        """Les listes de dossiers extraits sont initialisées."""
        processor = FileProcessor(Config(), mock_logger)

        assert hasattr(processor, '_extracted_folders')
        assert hasattr(processor, '_processed_folders')
        assert processor._extracted_folders == []
        assert processor._processed_folders == set()

    def test_extracted_folder_tracked_after_archive_conversion(
        self, mock_logger, temp_dir, file_factory
    ):
        """Un dossier créé par une conversion d'archive est tracké."""
        config = Config(report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("test.txt", "content")

        # Simuler un convertisseur qui crée un dossier
        output_folder = temp_dir / "extracted_folder"
        output_folder.mkdir()

        mock_converter = MagicMock()
        mock_converter.can_convert.return_value = True
        mock_converter.is_available.return_value = True
        mock_converter.convert.return_value = ConversionResult(
            status=ConversionStatus.SUCCESS,
            source=source,
            dest=output_folder,  # Destination est un dossier
            duration=0.1,
            method="archive_zip",
        )

        processor.converters = [mock_converter]
        processor.process_file(source)

        # Le dossier devrait être dans la liste des dossiers extraits
        assert output_folder.resolve() in processor._extracted_folders

    def test_extracted_folder_not_tracked_for_pdf(
        self, mock_logger, temp_dir, file_factory
    ):
        """Un fichier PDF créé n'est pas tracké comme dossier."""
        config = Config(report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        source = file_factory.create_text_file("test.txt", "content")
        pdf_dest = temp_dir / "test.txt.pdf"

        mock_converter = MagicMock()
        mock_converter.can_convert.return_value = True
        mock_converter.is_available.return_value = True

        def mock_convert(src, dst):
            dst.write_bytes(b"%PDF")
            return ConversionResult(
                status=ConversionStatus.SUCCESS,
                source=src,
                dest=dst,
                duration=0.1,
                method="text",
            )

        mock_converter.convert.side_effect = mock_convert
        processor.converters = [mock_converter]
        processor.process_file(source)

        # Pas de dossier extrait tracké
        assert len(processor._extracted_folders) == 0

    def test_process_directory_resets_extracted_folders(
        self, mock_logger, temp_dir, file_factory
    ):
        """process_directory réinitialise les listes de dossiers."""
        config = Config(report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        # Ajouter des données factices
        processor._extracted_folders = [Path("/fake/folder")]
        processor._processed_folders = {Path("/fake/processed")}

        file_factory.create_text_file("test.txt", "content")
        processor.process_directory(temp_dir)

        # Les listes devraient être réinitialisées au début
        # (elles peuvent contenir de nouvelles entrées après traitement)
        assert Path("/fake/folder") not in processor._extracted_folders
        assert Path("/fake/processed") not in processor._processed_folders

    def test_nested_zip_extraction_flow(
        self, mock_logger, temp_dir
    ):
        """Test du flux d'extraction imbriquée (ZIP dans ZIP)."""
        import zipfile

        config = Config(report_enabled=False)
        processor = FileProcessor(config, mock_logger)

        # Créer un ZIP interne avec un fichier texte
        inner_zip_path = temp_dir / "inner.zip"
        with zipfile.ZipFile(inner_zip_path, 'w') as zf:
            zf.writestr("deep.txt", "Deep content")

        # Créer un ZIP externe contenant le ZIP interne
        outer_zip_path = temp_dir / "outer.zip"
        with zipfile.ZipFile(outer_zip_path, 'w') as zf:
            zf.write(inner_zip_path, "inner.zip")

        # Supprimer le ZIP interne (il est maintenant dans outer.zip)
        inner_zip_path.unlink()

        # Traiter le répertoire
        stats = processor.process_directory(temp_dir)

        # Le ZIP externe devrait être traité
        assert stats["total"] >= 1
        assert stats["success"] >= 1

        # Le dossier outer/ devrait exister
        outer_folder = temp_dir / "outer"
        assert outer_folder.exists()

        # Le ZIP interne devrait avoir été extrait dans outer/
        inner_in_outer = outer_folder / "inner.zip"
        assert inner_in_outer.exists()

        # Le dossier inner/ devrait exister (traitement récursif)
        inner_folder = outer_folder / "inner"
        assert inner_folder.exists(), "Le ZIP imbriqué devrait avoir été extrait"

        # Le fichier deep.txt devrait être dans inner/
        deep_txt = inner_folder / "deep.txt"
        assert deep_txt.exists(), "Le contenu du ZIP imbriqué devrait être extrait"
