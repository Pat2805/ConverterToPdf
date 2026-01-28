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

    # === Rapport de session ===
    report_enabled: bool = True
    """Activer la génération du rapport de session (fichier texte)"""

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
                        Si None, cherche .converterrc dans plusieurs emplacements.

        Returns:
            Instance de Config
        """
        # Chercher le fichier de config dans plusieurs emplacements
        if config_path is None:
            # Emplacements de recherche (par ordre de priorité)
            search_paths = [
                Path.cwd() / ".converterrc",  # Répertoire de travail
                Path(__file__).parent.parent / ".converterrc",  # Répertoire du package
                Path.home() / ".converterrc",  # Répertoire utilisateur
            ]
            config_path = None
            for path in search_paths:
                if path.exists():
                    config_path = path
                    break
            if config_path is None:
                return cls()  # Aucun fichier trouvé, config par défaut
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

        Les arguments CLI ont priorité sur le fichier de config,
        SAUF pour les flags booléens (store_true) qui ne sont mis à jour
        que s'ils sont explicitement activés sur la ligne de commande.

        Args:
            args: Namespace d'argparse
        """
        updates = {}

        # Mapping des arguments non-booléens (valeur None si non spécifié)
        non_bool_mapping = {
            "method": "method",
            "log_level": "log_level",
            "log_file": "log_file",
            "ocr_engine": "ocr_engine",
        }

        for arg_name, config_name in non_bool_mapping.items():
            if hasattr(args, arg_name):
                value = getattr(args, arg_name)
                if value is not None:
                    updates[config_name] = value

        # Flags booléens (store_true): ne mettre à jour que si True
        # Car False signifie "non spécifié sur la CLI", pas "désactivé"
        bool_flags = {
            "recursive": "recursive",
            "force": "force",
            "delete": "delete_source",
            "ocr": "ocr_enabled",
        }

        for arg_name, config_name in bool_flags.items():
            if hasattr(args, arg_name) and getattr(args, arg_name):
                updates[config_name] = True

        # Gestion des flags inversés (--no-xxx)
        if hasattr(args, "no_keep_ext") and args.no_keep_ext:
            updates["keep_extension"] = False

        if hasattr(args, "no_report") and args.no_report:
            updates["report_enabled"] = False

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
            # Archives
            ".zip", ".rar", ".7z",
            ".tar", ".tar.gz", ".tgz", ".tar.bz2", ".tbz2",
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
