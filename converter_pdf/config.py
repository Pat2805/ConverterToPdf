"""
Configuration pour ConverterToPdf.

Supporte:
- Configuration par défaut
- Fichier .converterrc (YAML)
- Arguments CLI (priorité maximale)
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

# Import optionnel de YAML
try:
    import yaml
    YAML_AVAILABLE = True
except ImportError:
    YAML_AVAILABLE = False
    yaml = None  # type: ignore


@dataclass
class Config:
    """
    Configuration du convertisseur PDF.

    Priorité de chargement:
    1. Arguments CLI (priorité max)
    2. Fichier .converterrc
    3. Valeurs par défaut

    Usage:
        # Charger la config
        config = Config.load(Path(".converterrc"))

        # Mettre à jour depuis CLI
        config.update_from_args(args)

        # Utiliser
        if config.method == "auto":
            ...
    """

    # === Méthode de conversion ===
    method: str = "auto"
    """Méthode de conversion: auto, office, libreoffice, reportlab"""

    # === Nommage des PDF ===
    keep_extension: bool = True
    """Conserver l'extension d'origine: doc.docx -> doc.docx.pdf"""

    # === Logging ===
    log_level: str = "INFO"
    """Niveau de log console: DEBUG, INFO, WARNING, ERROR"""

    log_file: Path | None = None
    """Fichier de log (optionnel)"""

    log_file_level: str = "DEBUG"
    """Niveau de log fichier"""

    # === Journal CSV ===
    journal_enabled: bool = True
    """Activer le journal CSV d'audit"""

    journal_errors_only: bool = True
    """Journaliser uniquement les erreurs (pas les succès)"""

    # === OCR ===
    ocr_enabled: bool = False
    """Activer l'OCR pour les images"""

    ocr_engine: str = "auto"
    """Moteur OCR: auto, tesseract, easyocr, paddleocr"""

    # === Timeouts (secondes) ===
    office_timeout: int = 60
    """Timeout pour les opérations Microsoft Office"""

    libreoffice_timeout: int = 60
    """Timeout pour LibreOffice"""

    browser_timeout: int = 30
    """Timeout pour le navigateur headless (HTML)"""

    # === Chemins externes (détectés automatiquement) ===
    libreoffice_path: Path | None = None
    """Chemin vers soffice.exe (détecté automatiquement si None)"""

    browser_path: Path | None = None
    """Chemin vers Chrome/Edge (détecté automatiquement si None)"""

    # === Options de traitement ===
    recursive: bool = False
    """Traiter les sous-répertoires"""

    force: bool = False
    """Forcer la reconversion même si le PDF existe"""

    delete_source: bool = False
    """Supprimer le fichier source après conversion réussie"""

    # === Filtres de formats ===
    extensions: list[str] | None = None
    """Extensions à traiter (None = toutes)"""

    def __post_init__(self):
        """Validation et conversion des types après initialisation."""
        # Convertir les chemins en Path
        if self.log_file and not isinstance(self.log_file, Path):
            self.log_file = Path(self.log_file)
        if self.libreoffice_path and not isinstance(self.libreoffice_path, Path):
            self.libreoffice_path = Path(self.libreoffice_path)
        if self.browser_path and not isinstance(self.browser_path, Path):
            self.browser_path = Path(self.browser_path)

        # Valider les valeurs
        valid_methods = {"auto", "office", "libreoffice", "reportlab"}
        if self.method not in valid_methods:
            raise ValueError(f"method doit être dans {valid_methods}, pas '{self.method}'")

        valid_levels = {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}
        if self.log_level.upper() not in valid_levels:
            raise ValueError(f"log_level doit être dans {valid_levels}")
        self.log_level = self.log_level.upper()

        valid_ocr = {"auto", "tesseract", "easyocr", "paddleocr"}
        if self.ocr_engine not in valid_ocr:
            raise ValueError(f"ocr_engine doit être dans {valid_ocr}")

    @classmethod
    def load(cls, config_path: Path | str | None = None) -> "Config":
        """
        Charge la configuration depuis un fichier YAML.

        Args:
            config_path: Chemin vers le fichier de config.
                        Si None, cherche .converterrc dans le répertoire courant.

        Returns:
            Instance de Config
        """
        # Chercher le fichier de config
        if config_path is None:
            config_path = Path.cwd() / ".converterrc"
        else:
            config_path = Path(config_path)

        # Si le fichier n'existe pas, retourner la config par défaut
        if not config_path.exists():
            return cls()

        # Vérifier que YAML est disponible
        if not YAML_AVAILABLE:
            raise ImportError(
                "PyYAML n'est pas installé. "
                "Installez-le avec: pip install pyyaml"
            )

        # Charger le fichier
        with open(config_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}

        # Convertir les chemins
        for key in ["log_file", "libreoffice_path", "browser_path"]:
            if key in data and data[key]:
                data[key] = Path(data[key])

        return cls(**data)

    def save(self, config_path: Path | str) -> None:
        """
        Sauvegarde la configuration dans un fichier YAML.

        Args:
            config_path: Chemin de destination
        """
        if not YAML_AVAILABLE:
            raise ImportError(
                "PyYAML n'est pas installé. "
                "Installez-le avec: pip install pyyaml"
            )

        config_path = Path(config_path)
        config_path.parent.mkdir(parents=True, exist_ok=True)

        # Convertir en dict avec les chemins en strings
        data = asdict(self)
        for key in ["log_file", "libreoffice_path", "browser_path"]:
            if data[key]:
                data[key] = str(data[key])

        with open(config_path, "w", encoding="utf-8") as f:
            yaml.dump(data, f, default_flow_style=False, allow_unicode=True)

    def update(self, **kwargs: Any) -> None:
        """
        Met à jour la configuration avec les valeurs fournies.

        Args:
            **kwargs: Valeurs à mettre à jour
        """
        for key, value in kwargs.items():
            if hasattr(self, key) and value is not None:
                setattr(self, key, value)

        # Re-valider
        self.__post_init__()

    def update_from_args(self, args: Any) -> None:
        """
        Met à jour depuis les arguments CLI (argparse namespace).

        Args:
            args: Namespace d'argparse
        """
        updates = {}

        # Mapping des arguments CLI vers les attributs de config
        arg_mapping = {
            "method": "method",
            "recursive": "recursive",
            "force": "force",
            "delete": "delete_source",
            "log_level": "log_level",
            "log_file": "log_file",
            "output": None,  # Géré séparément
            "ocr": "ocr_enabled",
            "ocr_engine": "ocr_engine",
            "no_keep_ext": None,  # Géré spécialement
            "no_journal": None,  # Géré spécialement
            "log_all": None,  # Géré spécialement
        }

        for arg_name, config_name in arg_mapping.items():
            if config_name and hasattr(args, arg_name):
                value = getattr(args, arg_name)
                if value is not None:
                    updates[config_name] = value

        # Gestion des flags inversés
        if hasattr(args, "no_keep_ext") and args.no_keep_ext:
            updates["keep_extension"] = False

        if hasattr(args, "no_journal") and args.no_journal:
            updates["journal_enabled"] = False

        if hasattr(args, "log_all") and args.log_all:
            updates["journal_errors_only"] = False

        self.update(**updates)

    def get_all_extensions(self) -> list[str]:
        """
        Retourne la liste de toutes les extensions supportées.

        Returns:
            Liste des extensions (avec le point)
        """
        if self.extensions:
            return self.extensions

        return [
            # Documents Office
            ".doc", ".docx", ".rtf", ".odt",
            ".xls", ".xlsx", ".xlsm", ".xlsb",
            ".ppt", ".pptx",
            # Images
            ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp",
            # Web
            ".htm", ".html",
            # Texte
            ".txt", ".log",
            # Données
            ".xml",
            # Email
            ".msg",
            # PDF (copie/skip)
            ".pdf",
        ]

    def to_dict(self) -> dict[str, Any]:
        """Convertit la config en dictionnaire."""
        data = asdict(self)
        # Convertir les Path en str pour sérialisation
        for key in ["log_file", "libreoffice_path", "browser_path"]:
            if data[key]:
                data[key] = str(data[key])
        return data

    def __str__(self) -> str:
        """Représentation lisible de la config."""
        lines = ["Configuration:"]
        for key, value in self.to_dict().items():
            if value is not None:
                lines.append(f"  {key}: {value}")
        return "\n".join(lines)


def create_default_config(path: Path | str = ".converterrc") -> Config:
    """
    Crée un fichier de configuration par défaut.

    Args:
        path: Chemin du fichier à créer

    Returns:
        L'instance de Config créée
    """
    config = Config()
    config.save(path)
    return config
