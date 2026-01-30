"""
Fixtures pytest pour ConverterToPdf.

Fournit:
- Configurations de test
- Logger mock
- Fichiers de test temporaires
- Helpers pour créer des fichiers de différents formats
"""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Generator
from unittest.mock import MagicMock

import pytest

from converter_pdf.config import Config
from converter_pdf.logger import ConverterLogger
from converter_pdf.converters.base import ConversionResult, ConversionStatus


# =============================================================================
# Fixtures de configuration
# =============================================================================

@pytest.fixture
def default_config() -> Config:
    """Configuration par défaut pour les tests."""
    return Config()


@pytest.fixture
def config_force() -> Config:
    """Configuration avec force=True."""
    return Config(force=True)


@pytest.fixture
def config_recursive() -> Config:
    """Configuration avec récursivité activée."""
    return Config(recursive=True)


@pytest.fixture
def config_delete_source() -> Config:
    """Configuration avec suppression des sources."""
    return Config(delete_source=True)


@pytest.fixture
def config_no_report() -> Config:
    """Configuration sans rapport."""
    return Config(report_enabled=False)


@pytest.fixture
def config_office_only() -> Config:
    """Configuration méthode Office uniquement."""
    return Config(method="office")


@pytest.fixture
def config_libreoffice_only() -> Config:
    """Configuration méthode LibreOffice uniquement."""
    return Config(method="libreoffice")


@pytest.fixture
def config_reportlab_only() -> Config:
    """Configuration méthode ReportLab uniquement."""
    return Config(method="reportlab")


# =============================================================================
# Fixtures de logger
# =============================================================================

@pytest.fixture
def mock_logger() -> MagicMock:
    """Logger mock pour éviter les sorties pendant les tests."""
    logger = MagicMock(spec=ConverterLogger)
    logger.file_context = MagicMock()
    # Créer un context manager mock
    logger.file_context.return_value.__enter__ = MagicMock(return_value=None)
    logger.file_context.return_value.__exit__ = MagicMock(return_value=False)
    return logger


@pytest.fixture
def real_logger() -> ConverterLogger:
    """Logger réel mais silencieux (niveau CRITICAL)."""
    logger = ConverterLogger("test_logger")
    logger.setup(level="CRITICAL", console_colors=False)
    return logger


# =============================================================================
# Fixtures de répertoires temporaires
# =============================================================================

@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """Crée un répertoire temporaire, nettoyé après le test."""
    tmp = Path(tempfile.mkdtemp(prefix="converter_test_"))
    yield tmp
    shutil.rmtree(tmp, ignore_errors=True)


@pytest.fixture
def temp_file(temp_dir: Path) -> Generator[Path, None, None]:
    """Crée un fichier temporaire vide."""
    file = temp_dir / "test_file.txt"
    file.touch()
    yield file


# =============================================================================
# Helpers pour créer des fichiers de test
# =============================================================================

class FileFactory:
    """Factory pour créer des fichiers de test."""

    def __init__(self, base_dir: Path):
        self.base_dir = base_dir

    def create_text_file(
        self,
        name: str = "test.txt",
        content: str = "Test content\nLine 2\nLine 3"
    ) -> Path:
        """Crée un fichier texte."""
        file = self.base_dir / name
        file.write_text(content, encoding="utf-8")
        return file

    def create_xml_file(
        self,
        name: str = "test.xml",
        content: str | None = None
    ) -> Path:
        """Crée un fichier XML."""
        if content is None:
            content = """<?xml version="1.0" encoding="UTF-8"?>
<root>
    <item id="1">Test item</item>
    <item id="2">Another item</item>
</root>"""
        file = self.base_dir / name
        file.write_text(content, encoding="utf-8")
        return file

    def create_html_file(
        self,
        name: str = "test.html",
        content: str | None = None
    ) -> Path:
        """Crée un fichier HTML."""
        if content is None:
            content = """<!DOCTYPE html>
<html>
<head><title>Test</title></head>
<body>
    <h1>Test Document</h1>
    <p>This is a test paragraph.</p>
</body>
</html>"""
        file = self.base_dir / name
        file.write_text(content, encoding="utf-8")
        return file

    def create_empty_file(self, name: str) -> Path:
        """Crée un fichier vide."""
        file = self.base_dir / name
        file.touch()
        return file

    def create_binary_file(self, name: str, size: int = 1024) -> Path:
        """Crée un fichier binaire de taille spécifiée."""
        file = self.base_dir / name
        file.write_bytes(b"\x00" * size)
        return file

    def create_pdf_file(self, name: str = "test.pdf") -> Path:
        """Crée un fichier PDF minimal valide."""
        # PDF minimal valide
        pdf_content = b"""%PDF-1.4
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >> endobj
xref
0 4
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
trailer << /Size 4 /Root 1 0 R >>
startxref
196
%%EOF"""
        file = self.base_dir / name
        file.write_bytes(pdf_content)
        return file

    def create_image_file(
        self,
        name: str = "test.png",
        width: int = 100,
        height: int = 100
    ) -> Path:
        """Crée un fichier image PNG simple."""
        try:
            from PIL import Image
            img = Image.new("RGB", (width, height), color="white")
            file = self.base_dir / name
            img.save(file)
            return file
        except ImportError:
            # Fallback: créer un fichier PNG minimal (1x1 pixel)
            # PNG minimal valide
            png_data = bytes([
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,  # PNG signature
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,  # IHDR chunk
                0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
                0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,  # IDAT chunk
                0x54, 0x08, 0xD7, 0x63, 0xF8, 0xFF, 0xFF, 0x3F,
                0x00, 0x05, 0xFE, 0x02, 0xFE, 0xDC, 0xCC, 0x59,
                0xE7, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,  # IEND chunk
                0x44, 0xAE, 0x42, 0x60, 0x82
            ])
            file = self.base_dir / name
            file.write_bytes(png_data)
            return file

    def create_subdirectory(self, name: str) -> Path:
        """Crée un sous-répertoire."""
        subdir = self.base_dir / name
        subdir.mkdir(parents=True, exist_ok=True)
        return subdir


@pytest.fixture
def file_factory(temp_dir: Path) -> FileFactory:
    """Factory pour créer des fichiers de test."""
    return FileFactory(temp_dir)


# =============================================================================
# Fixtures pour résultats de conversion
# =============================================================================

@pytest.fixture
def success_result(temp_dir: Path) -> ConversionResult:
    """Résultat de conversion réussie."""
    source = temp_dir / "test.txt"
    source.write_text("test")
    dest = temp_dir / "test.txt.pdf"
    dest.write_bytes(b"%PDF-1.4 test")

    return ConversionResult(
        status=ConversionStatus.SUCCESS,
        source=source,
        dest=dest,
        duration=1.5,
        method="test_converter",
        message="Conversion réussie",
    )


@pytest.fixture
def failed_result(temp_dir: Path) -> ConversionResult:
    """Résultat de conversion échouée."""
    source = temp_dir / "test.txt"
    source.write_text("test")

    return ConversionResult(
        status=ConversionStatus.FAILED,
        source=source,
        dest=None,
        duration=0.5,
        method="test_converter",
        message="Erreur de conversion",
        exception=ValueError("Test error"),
    )


@pytest.fixture
def skipped_result(temp_dir: Path) -> ConversionResult:
    """Résultat de conversion ignorée."""
    source = temp_dir / "test.txt"
    source.write_text("test")

    return ConversionResult(
        status=ConversionStatus.SKIPPED_EXISTS,
        source=source,
        dest=None,
        duration=0,
        method="skip",
        message="PDF existe déjà",
    )


# =============================================================================
# Markers personnalisés
# =============================================================================

def pytest_configure(config):
    """Configure les markers personnalisés."""
    config.addinivalue_line(
        "markers", "slow: marks tests as slow (deselect with '-m \"not slow\"')"
    )
    config.addinivalue_line(
        "markers", "integration: marks tests as integration tests"
    )
    config.addinivalue_line(
        "markers", "requires_office: marks tests that require Microsoft Office"
    )
    config.addinivalue_line(
        "markers", "requires_libreoffice: marks tests that require LibreOffice"
    )
    config.addinivalue_line(
        "markers", "requires_pillow: marks tests that require Pillow"
    )
    config.addinivalue_line(
        "markers", "requires_reportlab: marks tests that require ReportLab"
    )
