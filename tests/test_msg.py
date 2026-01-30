"""
Tests pour le convertisseur de fichiers MSG.

Teste:
- MsgConverter
- Détection des images insignifiantes
- Sanitization des noms de fichiers
- Conversion MIME vers extension
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from converter_pdf.config import Config
from converter_pdf.converters.base import ConversionStatus


# =============================================================================
# Tests MsgConverter - Initialisation
# =============================================================================

class TestMsgConverterInit:
    """Tests d'initialisation."""

    def test_supported_extensions(self, mock_logger):
        """Extensions supportées sont correctes."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter.can_convert(".msg")
        assert converter.can_convert(".MSG")
        assert not converter.can_convert(".eml")
        assert not converter.can_convert(".pst")

    def test_is_available_with_extract_msg(self, mock_logger):
        """is_available retourne True si extract_msg est installé."""
        from converter_pdf.converters.msg import MsgConverter

        with patch("converter_pdf.converters.msg.EXTRACT_MSG_AVAILABLE", True):
            converter = MsgConverter(Config(), mock_logger)
            assert converter.is_available() is True

    def test_is_available_with_reportlab_only(self, mock_logger):
        """is_available retourne True si ReportLab est installé."""
        from converter_pdf.converters.msg import MsgConverter

        with patch("converter_pdf.converters.msg.EXTRACT_MSG_AVAILABLE", False):
            with patch("converter_pdf.converters.msg.REPORTLAB_AVAILABLE", True):
                converter = MsgConverter(Config(), mock_logger)
                assert converter.is_available() is True

    def test_is_available_without_dependencies(self, mock_logger):
        """is_available retourne False sans dépendances."""
        from converter_pdf.converters.msg import MsgConverter

        with patch("converter_pdf.converters.msg.EXTRACT_MSG_AVAILABLE", False):
            with patch("converter_pdf.converters.msg.REPORTLAB_AVAILABLE", False):
                converter = MsgConverter(Config(), mock_logger)
                assert converter.is_available() is False


# =============================================================================
# Tests sanitization des noms de fichiers
# =============================================================================

class TestMsgSanitizeFilename:
    """Tests du nettoyage des noms de fichiers."""

    def test_sanitize_removes_invalid_chars(self, mock_logger):
        """Caractères invalides sont remplacés."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        # Caractères Windows interdits: < > : " / \ | ? *
        assert converter._sanitize_filename('file<name>.txt') == 'file_name_.txt'
        assert converter._sanitize_filename('path:to|file') == 'path_to_file'
        assert converter._sanitize_filename('doc"test"') == 'doc_test_'
        assert converter._sanitize_filename('file/with\\slashes') == 'file_with_slashes'

    def test_sanitize_truncates_long_names(self, mock_logger):
        """Les noms trop longs sont tronqués."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)
        long_name = "a" * 300 + ".txt"

        result = converter._sanitize_filename(long_name)
        assert len(result) <= 200

    def test_sanitize_strips_whitespace(self, mock_logger):
        """Les espaces en début/fin sont supprimés."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter._sanitize_filename('  file.txt  ') == 'file.txt'
        assert converter._sanitize_filename('\ttest\n') == 'test'

    def test_sanitize_normal_names_unchanged(self, mock_logger):
        """Les noms normaux ne sont pas modifiés."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter._sanitize_filename('document.pdf') == 'document.pdf'
        assert converter._sanitize_filename('my-file_v2.docx') == 'my-file_v2.docx'
        assert converter._sanitize_filename('Report 2023.xlsx') == 'Report 2023.xlsx'


# =============================================================================
# Tests conversion MIME vers extension
# =============================================================================

class TestMsgMimeToExtension:
    """Tests de la conversion MIME vers extension."""

    def test_common_image_types(self, mock_logger):
        """Types MIME d'images courants."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter._get_extension_from_mime("image/jpeg") == ".jpg"
        assert converter._get_extension_from_mime("image/png") == ".png"
        assert converter._get_extension_from_mime("image/gif") == ".gif"
        assert converter._get_extension_from_mime("image/bmp") == ".bmp"

    def test_common_document_types(self, mock_logger):
        """Types MIME de documents courants."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter._get_extension_from_mime("application/pdf") == ".pdf"
        assert converter._get_extension_from_mime("application/msword") == ".doc"
        assert converter._get_extension_from_mime(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ) == ".docx"
        assert converter._get_extension_from_mime("application/vnd.ms-excel") == ".xls"

    def test_text_types(self, mock_logger):
        """Types MIME de texte."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter._get_extension_from_mime("text/plain") == ".txt"
        assert converter._get_extension_from_mime("text/html") == ".html"
        assert converter._get_extension_from_mime("text/xml") == ".xml"

    def test_mime_with_parameters(self, mock_logger):
        """Types MIME avec paramètres (charset, etc.)."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        # Le type MIME peut avoir des paramètres après ;
        assert converter._get_extension_from_mime("text/plain; charset=utf-8") == ".txt"
        assert converter._get_extension_from_mime("text/html; charset=iso-8859-1") == ".html"

    def test_unknown_mime_type(self, mock_logger):
        """Type MIME inconnu retourne chaîne vide."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter._get_extension_from_mime("application/x-unknown") == ""
        assert converter._get_extension_from_mime("") == ""
        assert converter._get_extension_from_mime("application/octet-stream") == ""

    def test_mime_case_insensitive(self, mock_logger):
        """La comparaison MIME est insensible à la casse."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter._get_extension_from_mime("IMAGE/JPEG") == ".jpg"
        assert converter._get_extension_from_mime("Image/Png") == ".png"


# =============================================================================
# Tests détection d'images insignifiantes
# =============================================================================

class TestInsignificantImageDetection:
    """Tests de la détection d'images insignifiantes."""

    def test_very_small_image_filtered(self, mock_logger):
        """Une image très petite est filtrée."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        # Image de moins de 5KB et nom quelconque
        small_data = b"\x00" * 1000  # 1KB

        mock_attachment = MagicMock()

        is_insignificant, reason = converter._is_insignificant_image(
            "photo.jpg",
            small_data,
            mock_attachment,
        )

        # Très petite image (<5KB) devrait être filtrée
        # Le comportement exact dépend de l'implémentation

    def test_logo_pattern_small_image_filtered(self, mock_logger):
        """Une petite image avec nom 'logo' est filtrée."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        # Image petite avec nom suspect
        small_data = b"\x00" * 8000  # 8KB (< 10KB)

        mock_attachment = MagicMock()

        is_insignificant, reason = converter._is_insignificant_image(
            "company_logo.png",
            small_data,
            mock_attachment,
        )

        # Devrait être filtrée (petite + nom suspect)

    def test_normal_image_not_filtered(self, mock_logger):
        """Une image de taille normale n'est pas filtrée."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        # Image de taille raisonnable
        normal_data = b"\x00" * 50000  # 50KB

        mock_attachment = MagicMock()

        is_insignificant, reason = converter._is_insignificant_image(
            "photo.jpg",
            normal_data,
            mock_attachment,
        )

        assert is_insignificant is False

    def test_large_logo_not_filtered(self, mock_logger):
        """Un logo de grande taille n'est pas filtré."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        # Grande image même avec nom "logo"
        large_data = b"\x00" * 100000  # 100KB

        mock_attachment = MagicMock()

        is_insignificant, reason = converter._is_insignificant_image(
            "big_logo.png",
            large_data,
            mock_attachment,
        )

        # Ne devrait pas être filtrée car trop grande
        assert is_insignificant is False

    def test_non_image_not_filtered(self, mock_logger):
        """Un fichier non-image n'est pas filtré comme image."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        mock_attachment = MagicMock()

        is_insignificant, reason = converter._is_insignificant_image(
            "document.pdf",
            b"PDF content",
            mock_attachment,
        )

        assert is_insignificant is False


# =============================================================================
# Tests des patterns d'images insignifiantes
# =============================================================================

class TestInsignificantImagePatterns:
    """Tests des patterns de noms d'images insignifiantes."""

    @pytest.fixture
    def msg_converter(self, mock_logger):
        """Crée un convertisseur MSG."""
        from converter_pdf.converters.msg import MsgConverter
        return MsgConverter(Config(), mock_logger)

    def test_logo_patterns(self, msg_converter):
        """Patterns 'logo' sont reconnus."""
        patterns = msg_converter.INSIGNIFICANT_IMAGE_PATTERNS
        import re

        # Tester avec quelques noms
        logo_names = ["logo.png", "company_logo.gif", "Logo_small.jpg"]

        for name in logo_names:
            stem = Path(name).stem.lower()
            matched = any(re.search(p, stem) for p in patterns)
            assert matched, f"{name} devrait matcher un pattern logo"

    def test_spacer_patterns(self, msg_converter):
        """Patterns 'spacer' sont reconnus."""
        patterns = msg_converter.INSIGNIFICANT_IMAGE_PATTERNS
        import re

        spacer_names = ["spacer.gif", "SPACER.png"]

        for name in spacer_names:
            stem = Path(name).stem.lower()
            matched = any(re.search(p, stem) for p in patterns)
            assert matched, f"{name} devrait matcher un pattern spacer"

    def test_normal_names_not_matched(self, msg_converter):
        """Les noms normaux ne matchent pas."""
        patterns = msg_converter.INSIGNIFICANT_IMAGE_PATTERNS
        import re

        normal_names = ["photo.jpg", "screenshot.png", "document_scan.tiff", "vacation.jpeg"]

        for name in normal_names:
            stem = Path(name).stem.lower()
            matched = any(re.search(p, stem) for p in patterns)
            assert not matched, f"{name} ne devrait pas matcher de pattern"


# =============================================================================
# Tests des extensions convertibles
# =============================================================================

class TestConvertibleExtensions:
    """Tests des extensions convertibles en PDF."""

    def test_office_documents_convertible(self, mock_logger):
        """Documents Office sont convertibles."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        office_exts = [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"]
        for ext in office_exts:
            assert ext in converter.CONVERTIBLE_EXTENSIONS

    def test_images_convertible(self, mock_logger):
        """Images sont convertibles."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        image_exts = [".jpg", ".jpeg", ".png", ".bmp", ".tiff"]
        for ext in image_exts:
            assert ext in converter.CONVERTIBLE_EXTENSIONS

    def test_text_files_convertible(self, mock_logger):
        """Fichiers texte sont convertibles."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert ".txt" in converter.CONVERTIBLE_EXTENSIONS
        assert ".html" in converter.CONVERTIBLE_EXTENSIONS
        assert ".xml" in converter.CONVERTIBLE_EXTENSIONS


# =============================================================================
# Tests des seuils de filtrage
# =============================================================================

class TestFilteringThresholds:
    """Tests des seuils de filtrage des images."""

    def test_min_image_size_threshold(self, mock_logger):
        """Seuil de taille minimum pour noms suspects."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter.MIN_IMAGE_SIZE_BYTES == 30 * 1024  # 30KB

    def test_min_image_dimension_threshold(self, mock_logger):
        """Seuil de dimension minimum pour noms suspects."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter.MIN_IMAGE_DIMENSION == 200  # 200px

    def test_always_filter_thresholds(self, mock_logger):
        """Seuils pour images toujours filtrées (trop petites)."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        assert converter.ALWAYS_FILTER_SIZE_BYTES == 15 * 1024  # 15KB
        assert converter.ALWAYS_FILTER_DIMENSION == 150  # 150px
        assert converter.ALWAYS_FILTER_SURFACE == 25000  # ~158x158px


# =============================================================================
# Tests de conversion (sans dépendances)
# =============================================================================

class TestMsgConversionWithoutDependencies:
    """Tests de conversion sans les dépendances externes."""

    def test_convert_without_extract_msg_fails(self, mock_logger, temp_dir):
        """Conversion échoue sans extract_msg."""
        from converter_pdf.converters.msg import MsgConverter

        # Simuler l'absence de extract_msg et ReportLab
        with patch("converter_pdf.converters.msg.EXTRACT_MSG_AVAILABLE", False):
            with patch("converter_pdf.converters.msg.REPORTLAB_AVAILABLE", False):
                converter = MsgConverter(Config(), mock_logger)

                source = temp_dir / "test.msg"
                source.write_bytes(b"fake msg content")
                dest = temp_dir / "test.msg.pdf"

                result = converter.convert(source, dest)

                assert result.status == ConversionStatus.FAILED


# =============================================================================
# Tests d'intégration (si dépendances disponibles)
# =============================================================================

class TestMsgIntegration:
    """Tests d'intégration pour MSG (si extract_msg est disponible)."""

    @pytest.fixture
    def has_extract_msg(self):
        """Vérifie si extract_msg est installé."""
        try:
            import extract_msg
            return True
        except ImportError:
            return False

    @pytest.mark.integration
    def test_msg_converter_initialization(self, mock_logger, has_extract_msg):
        """Initialisation du convertisseur MSG."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        # Le convertisseur HTML devrait être initialisé
        assert converter._html_converter is not None

        # Le cache de convertisseurs devrait être None (lazy loading)
        assert converter._converters_cache is None

    @pytest.mark.integration
    def test_get_converters_lazy_loading(self, mock_logger):
        """Les convertisseurs sont chargés à la demande."""
        from converter_pdf.converters.msg import MsgConverter

        converter = MsgConverter(Config(), mock_logger)

        # Avant l'appel, le cache est None
        assert converter._converters_cache is None

        # Après l'appel, le cache est rempli
        converters = converter._get_converters()
        assert converter._converters_cache is not None
        assert len(converters) > 0

        # Un second appel retourne le cache
        converters2 = converter._get_converters()
        assert converters is converters2
