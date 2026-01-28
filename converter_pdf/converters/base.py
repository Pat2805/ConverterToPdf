"""
Classes de base pour les convertisseurs.

Définit l'interface commune et les structures de données
pour tous les convertisseurs.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..config import Config
    from ..logger import ConverterLogger


class ConversionStatus(Enum):
    """Statut de conversion."""

    SUCCESS = "success"
    """Conversion réussie"""

    FAILED = "failed"
    """Conversion échouée"""

    SKIPPED_PASSWORD = "skipped_password"
    """Ignoré: fichier protégé par mot de passe"""

    SKIPPED_EXISTS = "skipped_exists"
    """Ignoré: PDF existe déjà"""

    SKIPPED_UNSUPPORTED = "skipped_unsupported"
    """Ignoré: format non supporté"""

    SKIPPED_PDF = "skipped_pdf"
    """Ignoré: déjà un PDF"""


@dataclass
class ConversionResult:
    """
    Résultat d'une conversion.

    Contient toutes les informations sur le résultat d'une tentative
    de conversion, qu'elle ait réussi ou échoué.
    """

    status: ConversionStatus
    """Statut de la conversion"""

    source: Path
    """Fichier source"""

    dest: Path | None
    """Fichier de destination (None si échec)"""

    duration: float
    """Durée de la conversion en secondes"""

    method: str
    """Nom du convertisseur utilisé"""

    message: str = ""
    """Message explicatif (optionnel)"""

    exception: Exception | None = None
    """Exception si erreur (optionnel)"""

    source_size: int = 0
    """Taille du fichier source en bytes"""

    dest_size: int = 0
    """Taille du PDF généré en bytes"""

    def __post_init__(self):
        """Calcule les tailles si possible."""
        if self.source and self.source.exists():
            self.source_size = self.source.stat().st_size
        if self.dest and self.dest.exists():
            self.dest_size = self.dest.stat().st_size

    @property
    def is_success(self) -> bool:
        """True si la conversion a réussi."""
        return self.status == ConversionStatus.SUCCESS

    @property
    def is_skipped(self) -> bool:
        """True si le fichier a été ignoré."""
        return self.status in (
            ConversionStatus.SKIPPED_PASSWORD,
            ConversionStatus.SKIPPED_EXISTS,
            ConversionStatus.SKIPPED_UNSUPPORTED,
            ConversionStatus.SKIPPED_PDF,
        )

    @property
    def is_failed(self) -> bool:
        """True si la conversion a échoué."""
        return self.status == ConversionStatus.FAILED

    @property
    def source_size_mb(self) -> float:
        """Taille source en MB."""
        return self.source_size / (1024 * 1024)

    @property
    def dest_size_mb(self) -> float:
        """Taille destination en MB."""
        return self.dest_size / (1024 * 1024)

    def __str__(self) -> str:
        """Représentation lisible."""
        parts = [f"{self.status.value}: {self.source.name}"]
        if self.dest:
            parts.append(f"-> {self.dest.name}")
        parts.append(f"({self.duration:.1f}s, {self.method})")
        if self.message:
            parts.append(f"- {self.message}")
        return " ".join(parts)


class BaseConverter(ABC):
    """
    Classe de base abstraite pour tous les convertisseurs.

    Chaque convertisseur doit:
    - Définir `name` et `supported_extensions`
    - Implémenter `convert()`
    - Ne jamais lever d'exception (retourner ConversionResult avec status FAILED)

    Usage:
        class MyConverter(BaseConverter):
            name = "my_converter"
            supported_extensions = [".xyz"]

            def convert(self, source: Path, dest: Path) -> ConversionResult:
                # ... conversion ...
                return ConversionResult(...)
    """

    name: str = "base"
    """Nom du convertisseur (pour les logs)"""

    supported_extensions: list[str] = []
    """Extensions supportées (avec le point, ex: [".docx", ".doc"])"""

    def __init__(self, config: "Config", logger: "ConverterLogger"):
        """
        Initialise le convertisseur.

        Args:
            config: Configuration globale
            logger: Logger pour ce convertisseur
        """
        self.config = config
        self.logger = logger

    def can_convert(self, extension: str) -> bool:
        """
        Vérifie si ce convertisseur supporte l'extension.

        Args:
            extension: Extension du fichier (avec ou sans point)

        Returns:
            True si l'extension est supportée
        """
        ext = extension.lower()
        if not ext.startswith("."):
            ext = f".{ext}"
        return ext in self.supported_extensions

    @abstractmethod
    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """
        Convertit un fichier en PDF.

        Cette méthode NE DOIT JAMAIS lever d'exception.
        Toutes les erreurs doivent être encapsulées dans le ConversionResult.

        Args:
            source: Chemin du fichier source
            dest: Chemin du PDF de destination

        Returns:
            ConversionResult avec le statut et les détails
        """
        pass

    def is_available(self) -> bool:
        """
        Vérifie si le convertisseur est disponible.

        Peut être surchargé pour vérifier les dépendances
        (ex: Office installé, LibreOffice présent, etc.)

        Returns:
            True si le convertisseur peut être utilisé
        """
        return True

    def __str__(self) -> str:
        return f"{self.name} ({', '.join(self.supported_extensions)})"

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__} name={self.name}>"
