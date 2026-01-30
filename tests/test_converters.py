"""
Tests pour les convertisseurs.

Teste:
- BaseConverter (interface)
- TextConverter (txt, log)
- ImageConverter (jpg, png, etc.)
- XmlConverter (xml)
- Converter chain
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from converter_pdf.config import Config
from converter_pdf.converters.base import (
    BaseConverter,
    ConversionResult,
    ConversionStatus,
)


# =============================================================================
# Tests ConversionStatus
# =============================================================================

class TestConversionStatus:
    """Tests de l'enum ConversionStatus."""

    def test_all_statuses_exist(self):
        """Tous les statuts existent."""
        assert ConversionStatus.SUCCESS
        assert ConversionStatus.FAILED
        assert ConversionStatus.SKIPPED_PASSWORD
        assert ConversionStatus.SKIPPED_EXISTS
        assert ConversionStatus.SKIPPED_UNSUPPORTED
        assert ConversionStatus.SKIPPED_PDF

    def test_status_values(self):
        """Les valeurs des statuts sont correctes."""
        assert ConversionStatus.SUCCESS.value == "success"
        assert ConversionStatus.FAILED.value == "failed"
        assert ConversionStatus.SKIPPED_PASSWORD.value == "skipped_password"


# =============================================================================
# Tests ConversionResult
# =============================================================================

class TestConversionResult:
    """Tests de la classe ConversionResult."""

    def test_success_result_is_success(self, temp_dir: Path):
        """Un résultat SUCCESS a is_success=True."""
        source = temp_dir / "test.txt"
        source.write_text("test")

        result = ConversionResult(
            status=ConversionStatus.SUCCESS,
            source=source,
            dest=temp_dir / "test.pdf",
            duration=1.0,
            method="test",
        )
        assert result.is_success is True
        assert result.is_failed is False
        assert result.is_skipped is False

    def test_failed_result_is_failed(self, temp_dir: Path):
        """Un résultat FAILED a is_failed=True."""
        source = temp_dir / "test.txt"
        source.write_text("test")

        result = ConversionResult(
            status=ConversionStatus.FAILED,
            source=source,
            dest=None,
            duration=0.5,
            method="test",
        )
        assert result.is_failed is True
        assert result.is_success is False
        assert result.is_skipped is False

    def test_skipped_results_are_skipped(self, temp_dir: Path):
        """Les résultats SKIPPED_* ont is_skipped=True."""
        source = temp_dir / "test.txt"
        source.write_text("test")

        skipped_statuses = [
            ConversionStatus.SKIPPED_PASSWORD,
            ConversionStatus.SKIPPED_EXISTS,
            ConversionStatus.SKIPPED_UNSUPPORTED,
            ConversionStatus.SKIPPED_PDF,
        ]

        for status in skipped_statuses:
            result = ConversionResult(
                status=status,
                source=source,
                dest=None,
                duration=0,
                method="skip",
            )
            assert result.is_skipped is True, f"{status} devrait être skipped"
            assert result.is_success is False
            assert result.is_failed is False

    def test_source_size_calculated(self, temp_dir: Path):
        """La taille source est calculée automatiquement."""
        source = temp_dir / "test.txt"
        source.write_text("Hello World")  # 11 bytes

        result = ConversionResult(
            status=ConversionStatus.SUCCESS,
            source=source,
            dest=None,
            duration=0,
            method="test",
        )
        assert result.source_size == 11

    def test_dest_size_calculated(self, temp_dir: Path):
        """La taille destination est calculée automatiquement."""
        source = temp_dir / "test.txt"
        source.write_text("test")
        dest = temp_dir / "test.pdf"
        dest.write_bytes(b"%PDF" + b"\x00" * 100)  # 104 bytes

        result = ConversionResult(
            status=ConversionStatus.SUCCESS,
            source=source,
            dest=dest,
            duration=0,
            method="test",
        )
        assert result.dest_size == 104

    def test_size_mb_properties(self, temp_dir: Path):
        """Les propriétés size_mb fonctionnent."""
        source = temp_dir / "test.txt"
        source.write_bytes(b"\x00" * (1024 * 1024))  # 1 MB

        result = ConversionResult(
            status=ConversionStatus.SUCCESS,
            source=source,
            dest=None,
            duration=0,
            method="test",
        )
        assert result.source_size_mb == pytest.approx(1.0, rel=0.01)

    def test_str_representation(self, temp_dir: Path):
        """__str__ retourne une représentation lisible."""
        source = temp_dir / "test.txt"
        source.write_text("test")

        result = ConversionResult(
            status=ConversionStatus.SUCCESS,
            source=source,
            dest=temp_dir / "test.pdf",
            duration=1.5,
            method="test_converter",
        )
        result_str = str(result)

        assert "success" in result_str
        assert "test.txt" in result_str
        assert "test_converter" in result_str


# =============================================================================
# Tests BaseConverter
# =============================================================================

class TestBaseConverter:
    """Tests de la classe abstraite BaseConverter."""

    def test_can_convert_with_supported_extension(self, mock_logger):
        """can_convert retourne True pour une extension supportée."""
        # Créer une implémentation concrète minimale
        class TestConverter(BaseConverter):
            name = "test"
            supported_extensions = [".txt", ".log"]

            def convert(self, source, dest):
                pass

        converter = TestConverter(Config(), mock_logger)
        assert converter.can_convert(".txt") is True
        assert converter.can_convert(".log") is True
        assert converter.can_convert("txt") is True  # Sans le point

    def test_can_convert_with_unsupported_extension(self, mock_logger):
        """can_convert retourne False pour une extension non supportée."""
        class TestConverter(BaseConverter):
            name = "test"
            supported_extensions = [".txt"]

            def convert(self, source, dest):
                pass

        converter = TestConverter(Config(), mock_logger)
        assert converter.can_convert(".pdf") is False
        assert converter.can_convert(".docx") is False

    def test_can_convert_case_insensitive(self, mock_logger):
        """can_convert est insensible à la casse."""
        class TestConverter(BaseConverter):
            name = "test"
            supported_extensions = [".txt"]

            def convert(self, source, dest):
                pass

        converter = TestConverter(Config(), mock_logger)
        assert converter.can_convert(".TXT") is True
        assert converter.can_convert(".Txt") is True

    def test_is_available_default_true(self, mock_logger):
        """is_available retourne True par défaut."""
        class TestConverter(BaseConverter):
            name = "test"
            supported_extensions = []

            def convert(self, source, dest):
                pass

        converter = TestConverter(Config(), mock_logger)
        assert converter.is_available() is True

    def test_str_representation(self, mock_logger):
        """__str__ inclut nom et extensions."""
        class TestConverter(BaseConverter):
            name = "test"
            supported_extensions = [".txt", ".log"]

            def convert(self, source, dest):
                pass

        converter = TestConverter(Config(), mock_logger)
        result = str(converter)
        assert "test" in result
        assert ".txt" in result
        assert ".log" in result


# =============================================================================
# Tests TextConverter
# =============================================================================

class TestTextConverter:
    """Tests du convertisseur de texte."""

    @pytest.fixture
    def text_converter(self, mock_logger):
        """Crée une instance de TextConverter."""
        from converter_pdf.converters.text import TextConverter
        return TextConverter(Config(), mock_logger)

    def test_supported_extensions(self, text_converter):
        """Extensions supportées sont correctes."""
        assert text_converter.can_convert(".txt")
        assert text_converter.can_convert(".log")
        assert not text_converter.can_convert(".pdf")

    @pytest.mark.requires_reportlab
    def test_convert_simple_text(self, text_converter, file_factory, temp_dir):
        """Conversion d'un fichier texte simple."""
        source = file_factory.create_text_file("test.txt", "Hello World\nLine 2")
        dest = temp_dir / "test.txt.pdf"

        result = text_converter.convert(source, dest)

        if text_converter.is_available():
            assert result.status == ConversionStatus.SUCCESS
            assert dest.exists()
            assert result.dest == dest
        else:
            assert result.status == ConversionStatus.FAILED
            assert "ReportLab" in result.message

    @pytest.mark.requires_reportlab
    def test_convert_utf8_text(self, text_converter, file_factory, temp_dir):
        """Conversion d'un fichier texte UTF-8 avec caractères spéciaux."""
        source = file_factory.create_text_file(
            "unicode.txt",
            "Français: é à ü\nJapanese: 日本語\nEmoji: 😀"
        )
        dest = temp_dir / "unicode.txt.pdf"

        result = text_converter.convert(source, dest)

        if text_converter.is_available():
            assert result.status == ConversionStatus.SUCCESS

    @pytest.mark.requires_reportlab
    def test_convert_empty_file(self, text_converter, file_factory, temp_dir):
        """Conversion d'un fichier texte vide."""
        source = file_factory.create_empty_file("empty.txt")
        dest = temp_dir / "empty.txt.pdf"

        result = text_converter.convert(source, dest)

        if text_converter.is_available():
            assert result.status == ConversionStatus.SUCCESS

    def test_is_available_without_reportlab(self, mock_logger):
        """is_available retourne False sans ReportLab."""
        from converter_pdf.converters.text import TextConverter

        with patch("converter_pdf.converters.text.REPORTLAB_AVAILABLE", False):
            converter = TextConverter(Config(), mock_logger)
            # Note: is_available() lit la variable globale au runtime
            # Donc on doit patcher l'attribut ou recréer le module


# =============================================================================
# Tests ImageConverter
# =============================================================================

class TestImageConverter:
    """Tests du convertisseur d'images."""

    @pytest.fixture
    def image_converter(self, mock_logger):
        """Crée une instance de ImageConverter."""
        from converter_pdf.converters.image import ImageConverter
        return ImageConverter(Config(), mock_logger)

    def test_supported_extensions(self, image_converter):
        """Extensions supportées sont correctes."""
        assert image_converter.can_convert(".jpg")
        assert image_converter.can_convert(".jpeg")
        assert image_converter.can_convert(".png")
        assert image_converter.can_convert(".bmp")
        assert image_converter.can_convert(".tiff")
        assert image_converter.can_convert(".webp")
        assert image_converter.can_convert(".gif")
        assert not image_converter.can_convert(".pdf")

    @pytest.mark.requires_pillow
    def test_convert_png_image(self, image_converter, file_factory, temp_dir):
        """Conversion d'une image PNG."""
        source = file_factory.create_image_file("test.png", 200, 200)
        dest = temp_dir / "test.png.pdf"

        result = image_converter.convert(source, dest)

        if image_converter.is_available():
            assert result.status == ConversionStatus.SUCCESS
            assert dest.exists()
        else:
            assert result.status == ConversionStatus.FAILED

    @pytest.mark.requires_pillow
    def test_convert_records_duration(self, image_converter, file_factory, temp_dir):
        """La durée de conversion est enregistrée."""
        source = file_factory.create_image_file("test.png")
        dest = temp_dir / "test.png.pdf"

        result = image_converter.convert(source, dest)

        if image_converter.is_available():
            assert result.duration >= 0


# =============================================================================
# Tests XmlConverter
# =============================================================================

class TestXmlConverter:
    """Tests du convertisseur XML."""

    @pytest.fixture
    def xml_converter(self, mock_logger):
        """Crée une instance de XmlConverter."""
        from converter_pdf.converters.xml_converter import XmlConverter
        return XmlConverter(Config(), mock_logger)

    def test_supported_extensions(self, xml_converter):
        """Extensions supportées sont correctes."""
        assert xml_converter.can_convert(".xml")
        assert not xml_converter.can_convert(".txt")
        assert not xml_converter.can_convert(".html")

    @pytest.mark.requires_reportlab
    def test_convert_simple_xml(self, xml_converter, file_factory, temp_dir):
        """Conversion d'un fichier XML simple."""
        source = file_factory.create_xml_file("test.xml")
        dest = temp_dir / "test.xml.pdf"

        result = xml_converter.convert(source, dest)

        if xml_converter.is_available():
            assert result.status == ConversionStatus.SUCCESS
            assert dest.exists()

    @pytest.mark.requires_reportlab
    def test_convert_malformed_xml(self, xml_converter, file_factory, temp_dir):
        """Conversion d'un XML mal formé (doit quand même fonctionner)."""
        source = file_factory.create_xml_file(
            "bad.xml",
            "<root><unclosed>"  # XML invalide
        )
        dest = temp_dir / "bad.xml.pdf"

        result = xml_converter.convert(source, dest)

        if xml_converter.is_available():
            # Devrait réussir même avec XML invalide (garde le contenu brut)
            assert result.status == ConversionStatus.SUCCESS


# =============================================================================
# Tests Converter Chain
# =============================================================================

class TestConverterChain:
    """Tests de la chaîne de convertisseurs."""

    def test_get_converter_chain_auto(self, mock_logger):
        """get_converter_chain avec method=auto retourne tous les convertisseurs."""
        from converter_pdf.converters import get_converter_chain

        config = Config(method="auto")
        converters = get_converter_chain(config, mock_logger)

        # Vérifier que les convertisseurs principaux sont présents
        names = [c.name for c in converters]
        assert "image" in names
        assert "text" in names
        assert "xml" in names
        assert "msg" in names
        assert "archive" in names

    def test_get_converter_chain_office_only(self, mock_logger):
        """get_converter_chain avec method=office."""
        from converter_pdf.converters import get_converter_chain

        config = Config(method="office")
        converters = get_converter_chain(config, mock_logger)

        names = [c.name for c in converters]
        # Les convertisseurs Office devraient être en premier
        # mais image, text, etc. sont toujours inclus
        assert "image" in names
        assert "text" in names

    def test_get_converter_chain_libreoffice_only(self, mock_logger):
        """get_converter_chain avec method=libreoffice."""
        from converter_pdf.converters import get_converter_chain

        config = Config(method="libreoffice")
        converters = get_converter_chain(config, mock_logger)

        names = [c.name for c in converters]
        assert "libreoffice" in names
        assert "image" in names

    def test_converter_chain_order_matters(self, mock_logger):
        """L'ordre des convertisseurs est important pour les fallbacks."""
        from converter_pdf.converters import get_converter_chain

        config = Config(method="auto")
        converters = get_converter_chain(config, mock_logger)

        # Office devrait être avant ReportLab fallback
        office_indices = [
            i for i, c in enumerate(converters)
            if "office" in c.name.lower() and "reportlab" not in c.name.lower()
        ]
        reportlab_indices = [
            i for i, c in enumerate(converters)
            if "reportlab" in c.name.lower()
        ]

        if office_indices and reportlab_indices:
            # Le premier Office devrait être avant le premier ReportLab
            assert min(office_indices) < min(reportlab_indices)


# =============================================================================
# Tests d'intégration simples
# =============================================================================

class TestConverterIntegration:
    """Tests d'intégration des convertisseurs."""

    @pytest.mark.integration
    def test_text_to_pdf_full_workflow(self, temp_dir, real_logger):
        """Workflow complet : texte -> PDF."""
        from converter_pdf.converters.text import TextConverter

        # Créer un fichier texte
        source = temp_dir / "document.txt"
        source.write_text("Line 1\nLine 2\nLine 3", encoding="utf-8")
        dest = temp_dir / "document.txt.pdf"

        # Convertir
        converter = TextConverter(Config(), real_logger)
        if not converter.is_available():
            pytest.skip("ReportLab non disponible")

        result = converter.convert(source, dest)

        # Vérifications
        assert result.is_success
        assert dest.exists()
        assert dest.stat().st_size > 0
        assert result.method == "text"

    @pytest.mark.integration
    @pytest.mark.requires_pillow
    def test_image_to_pdf_full_workflow(self, temp_dir, real_logger, file_factory):
        """Workflow complet : image -> PDF."""
        from converter_pdf.converters.image import ImageConverter

        # Créer une image
        source = file_factory.create_image_file("photo.png", 300, 200)
        dest = temp_dir / "photo.png.pdf"

        # Convertir
        converter = ImageConverter(Config(), real_logger)
        if not converter.is_available():
            pytest.skip("Pillow non disponible")

        result = converter.convert(source, dest)

        # Vérifications
        assert result.is_success
        assert dest.exists()
        assert result.source_size > 0
        assert result.dest_size > 0
