"""
Tests pour le module config.py.

Teste:
- Valeurs par défaut
- Validation des paramètres
- Chargement depuis fichier YAML
- Mise à jour depuis CLI
- Sauvegarde de configuration
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from converter_pdf.config import Config, create_default_config


class TestConfigDefaults:
    """Tests des valeurs par défaut."""

    def test_default_method(self):
        """La méthode par défaut est 'auto'."""
        config = Config()
        assert config.method == "auto"

    def test_default_keep_extension(self):
        """keep_extension est True par défaut."""
        config = Config()
        assert config.keep_extension is True

    def test_default_log_level(self):
        """log_level par défaut est INFO."""
        config = Config()
        assert config.log_level == "INFO"

    def test_default_recursive(self):
        """recursive est False par défaut."""
        config = Config()
        assert config.recursive is False

    def test_default_force(self):
        """force est False par défaut."""
        config = Config()
        assert config.force is False

    def test_default_delete_source(self):
        """delete_source est False par défaut."""
        config = Config()
        assert config.delete_source is False

    def test_default_report_enabled(self):
        """report_enabled est True par défaut."""
        config = Config()
        assert config.report_enabled is True

    def test_default_dry_run(self):
        """dry_run est False par défaut."""
        config = Config()
        assert config.dry_run is False

    def test_default_hide_source(self):
        """hide_source est False par défaut."""
        config = Config()
        assert config.hide_source is False

    def test_default_ocr_disabled(self):
        """OCR est désactivé par défaut."""
        config = Config()
        assert config.ocr_enabled is False

    def test_default_timeouts(self):
        """Timeouts par défaut."""
        config = Config()
        assert config.office_timeout == 60
        assert config.libreoffice_timeout == 60
        assert config.browser_timeout == 30


class TestConfigValidation:
    """Tests de validation des paramètres."""

    def test_invalid_method_raises(self):
        """Une méthode invalide lève ValueError."""
        with pytest.raises(ValueError, match="method doit être dans"):
            Config(method="invalid")

    def test_valid_methods(self):
        """Les méthodes valides sont acceptées."""
        for method in ["auto", "office", "libreoffice", "reportlab"]:
            config = Config(method=method)
            assert config.method == method

    def test_invalid_log_level_raises(self):
        """Un niveau de log invalide lève ValueError."""
        with pytest.raises(ValueError, match="log_level doit être dans"):
            Config(log_level="INVALID")

    def test_valid_log_levels(self):
        """Les niveaux de log valides sont acceptés."""
        for level in ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]:
            config = Config(log_level=level)
            assert config.log_level == level

    def test_log_level_case_insensitive(self):
        """Le niveau de log est normalisé en majuscules."""
        config = Config(log_level="debug")
        assert config.log_level == "DEBUG"

        config = Config(log_level="Info")
        assert config.log_level == "INFO"

    def test_invalid_ocr_engine_raises(self):
        """Un moteur OCR invalide lève ValueError."""
        with pytest.raises(ValueError, match="ocr_engine doit être dans"):
            Config(ocr_engine="invalid")

    def test_valid_ocr_engines(self):
        """Les moteurs OCR valides sont acceptés."""
        for engine in ["auto", "tesseract", "easyocr", "paddleocr"]:
            config = Config(ocr_engine=engine)
            assert config.ocr_engine == engine

    def test_hide_source_delete_source_incompatible(self):
        """hide_source et delete_source sont incompatibles."""
        with pytest.raises(ValueError, match="delete_source et hide_source sont incompatibles"):
            Config(hide_source=True, delete_source=True)

    def test_hide_source_alone_valid(self):
        """hide_source seul est valide."""
        config = Config(hide_source=True)
        assert config.hide_source is True
        assert config.delete_source is False

    def test_delete_source_alone_valid(self):
        """delete_source seul est valide."""
        config = Config(delete_source=True)
        assert config.delete_source is True
        assert config.hide_source is False


class TestConfigPathConversion:
    """Tests de conversion des chemins."""

    def test_log_file_converted_to_path(self):
        """log_file string est converti en Path."""
        config = Config(log_file="test.log")
        assert isinstance(config.log_file, Path)
        assert config.log_file == Path("test.log")

    def test_libreoffice_path_converted(self):
        """libreoffice_path string est converti en Path."""
        config = Config(libreoffice_path="C:\\Program Files\\LibreOffice\\soffice.exe")
        assert isinstance(config.libreoffice_path, Path)

    def test_browser_path_converted(self):
        """browser_path string est converti en Path."""
        config = Config(browser_path="C:\\Chrome\\chrome.exe")
        assert isinstance(config.browser_path, Path)

    def test_none_paths_remain_none(self):
        """Les chemins None restent None."""
        config = Config()
        assert config.log_file is None
        assert config.libreoffice_path is None
        assert config.browser_path is None


class TestConfigUpdate:
    """Tests de mise à jour de configuration."""

    def test_update_single_value(self):
        """update() met à jour une seule valeur."""
        config = Config()
        config.update(method="office")
        assert config.method == "office"

    def test_update_multiple_values(self):
        """update() met à jour plusieurs valeurs."""
        config = Config()
        config.update(method="office", recursive=True, force=True)
        assert config.method == "office"
        assert config.recursive is True
        assert config.force is True

    def test_update_ignores_none(self):
        """update() ignore les valeurs None."""
        config = Config(method="office")
        config.update(method=None)
        assert config.method == "office"  # Inchangé

    def test_update_validates_values(self):
        """update() valide les nouvelles valeurs."""
        config = Config()
        with pytest.raises(ValueError):
            config.update(method="invalid")

    def test_update_ignores_unknown_keys(self):
        """update() ignore les clés inconnues."""
        config = Config()
        config.update(unknown_key="value")  # Ne devrait pas lever d'erreur
        assert not hasattr(config, "unknown_key")


class TestConfigFromArgs:
    """Tests de mise à jour depuis CLI (argparse)."""

    def test_update_from_args_method(self):
        """update_from_args met à jour la méthode."""
        config = Config()

        class Args:
            method = "office"
            log_level = None
            log_file = None
            ocr_engine = None
            recursive = False
            force = False
            delete = False
            hide = False
            ocr = False
            dry_run = False
            no_keep_ext = False
            no_report = False

        config.update_from_args(Args())
        assert config.method == "office"

    def test_update_from_args_bool_flags(self):
        """update_from_args traite les flags booléens."""
        config = Config()

        class Args:
            method = None
            log_level = None
            log_file = None
            ocr_engine = None
            recursive = True
            force = True
            delete = True
            hide = False
            ocr = False
            dry_run = False
            no_keep_ext = False
            no_report = False

        config.update_from_args(Args())
        assert config.recursive is True
        assert config.force is True
        assert config.delete_source is True

    def test_update_from_args_no_keep_ext(self):
        """update_from_args traite --no-keep-ext."""
        config = Config()
        assert config.keep_extension is True

        class Args:
            method = None
            log_level = None
            log_file = None
            ocr_engine = None
            recursive = False
            force = False
            delete = False
            hide = False
            ocr = False
            dry_run = False
            no_keep_ext = True
            no_report = False

        config.update_from_args(Args())
        assert config.keep_extension is False

    def test_update_from_args_no_report(self):
        """update_from_args traite --no-report."""
        config = Config()
        assert config.report_enabled is True

        class Args:
            method = None
            log_level = None
            log_file = None
            ocr_engine = None
            recursive = False
            force = False
            delete = False
            hide = False
            ocr = False
            dry_run = False
            no_keep_ext = False
            no_report = True

        config.update_from_args(Args())
        assert config.report_enabled is False

    def test_update_from_args_preserves_config_values(self):
        """Les flags False ne modifient pas les valeurs du fichier de config."""
        config = Config(recursive=True, force=True)

        class Args:
            method = None
            log_level = None
            log_file = None
            ocr_engine = None
            recursive = False  # False = non spécifié sur CLI
            force = False
            delete = False
            hide = False
            ocr = False
            dry_run = False
            no_keep_ext = False
            no_report = False

        config.update_from_args(Args())
        # Les valeurs du fichier de config sont préservées
        assert config.recursive is True
        assert config.force is True

    def test_update_from_args_hide(self):
        """update_from_args traite --hide."""
        config = Config()
        assert config.hide_source is False

        class Args:
            method = None
            log_level = None
            log_file = None
            ocr_engine = None
            recursive = False
            force = False
            delete = False
            hide = True
            ocr = False
            dry_run = False
            no_keep_ext = False
            no_report = False

        config.update_from_args(Args())
        assert config.hide_source is True

    def test_update_from_args_hide_delete_incompatible(self):
        """update_from_args lève une erreur si hide et delete sont tous deux True."""
        config = Config()

        class Args:
            method = None
            log_level = None
            log_file = None
            ocr_engine = None
            recursive = False
            force = False
            delete = True
            hide = True
            ocr = False
            dry_run = False
            no_keep_ext = False
            no_report = False

        with pytest.raises(ValueError, match="delete_source et hide_source sont incompatibles"):
            config.update_from_args(Args())


class TestConfigLoad:
    """Tests de chargement depuis fichier."""

    def test_load_nonexistent_returns_default(self):
        """Charger un fichier inexistant retourne la config par défaut."""
        config = Config.load(Path("nonexistent.yaml"))
        assert config.method == "auto"

    def test_load_none_returns_default(self):
        """Charger None retourne la config par défaut."""
        config = Config.load(None)
        # Devrait retourner les valeurs par défaut si aucun .converterrc trouvé
        assert isinstance(config, Config)

    @pytest.mark.skipif(
        not __import__("importlib.util").util.find_spec("yaml"),
        reason="PyYAML non installé"
    )
    def test_load_from_yaml_file(self, temp_dir: Path):
        """Charger depuis un fichier YAML."""
        import yaml

        config_path = temp_dir / ".converterrc"
        config_data = {
            "method": "office",
            "recursive": True,
            "force": True,
            "log_level": "DEBUG",
        }
        with open(config_path, "w") as f:
            yaml.dump(config_data, f)

        config = Config.load(config_path)
        assert config.method == "office"
        assert config.recursive is True
        assert config.force is True
        assert config.log_level == "DEBUG"


class TestConfigSave:
    """Tests de sauvegarde de configuration."""

    @pytest.mark.skipif(
        not __import__("importlib.util").util.find_spec("yaml"),
        reason="PyYAML non installé"
    )
    def test_save_creates_file(self, temp_dir: Path):
        """save() crée le fichier."""
        config = Config(method="office", recursive=True)
        config_path = temp_dir / ".converterrc"

        config.save(config_path)
        assert config_path.exists()

    @pytest.mark.skipif(
        not __import__("importlib.util").util.find_spec("yaml"),
        reason="PyYAML non installé"
    )
    def test_save_and_load_roundtrip(self, temp_dir: Path):
        """Sauvegarde puis chargement préserve les valeurs."""
        original = Config(
            method="office",
            recursive=True,
            force=True,
            log_level="DEBUG",
            office_timeout=120,
        )
        config_path = temp_dir / ".converterrc"

        original.save(config_path)
        loaded = Config.load(config_path)

        assert loaded.method == original.method
        assert loaded.recursive == original.recursive
        assert loaded.force == original.force
        assert loaded.log_level == original.log_level
        assert loaded.office_timeout == original.office_timeout


class TestConfigExtensions:
    """Tests des extensions supportées."""

    def test_get_all_extensions_default(self):
        """get_all_extensions retourne les extensions par défaut."""
        config = Config()
        extensions = config.get_all_extensions()

        # Vérifier quelques extensions clés
        assert ".doc" in extensions
        assert ".docx" in extensions
        assert ".pdf" in extensions
        assert ".jpg" in extensions
        assert ".png" in extensions
        assert ".txt" in extensions
        assert ".zip" in extensions
        assert ".msg" in extensions

    def test_get_all_extensions_custom(self):
        """get_all_extensions retourne les extensions personnalisées si définies."""
        config = Config(extensions=[".doc", ".pdf"])
        extensions = config.get_all_extensions()

        assert extensions == [".doc", ".pdf"]


class TestConfigToDict:
    """Tests de conversion en dictionnaire."""

    def test_to_dict_contains_all_fields(self):
        """to_dict contient tous les champs."""
        config = Config()
        data = config.to_dict()

        assert "method" in data
        assert "keep_extension" in data
        assert "log_level" in data
        assert "recursive" in data
        assert "force" in data
        assert "delete_source" in data

    def test_to_dict_paths_as_strings(self):
        """to_dict convertit les Path en strings."""
        config = Config(log_file=Path("test.log"))
        data = config.to_dict()

        assert isinstance(data["log_file"], str)
        assert data["log_file"] == "test.log"


class TestConfigStr:
    """Tests de représentation string."""

    def test_str_readable_format(self):
        """__str__ retourne un format lisible."""
        config = Config(method="office", recursive=True)
        result = str(config)

        assert "Configuration:" in result
        assert "method: office" in result
        assert "recursive: True" in result


class TestCreateDefaultConfig:
    """Tests de create_default_config."""

    @pytest.mark.skipif(
        not __import__("importlib.util").util.find_spec("yaml"),
        reason="PyYAML non installé"
    )
    def test_create_default_config_creates_file(self, temp_dir: Path):
        """create_default_config crée un fichier."""
        config_path = temp_dir / ".converterrc"
        config = create_default_config(config_path)

        assert config_path.exists()
        assert isinstance(config, Config)

    @pytest.mark.skipif(
        not __import__("importlib.util").util.find_spec("yaml"),
        reason="PyYAML non installé"
    )
    def test_create_default_config_loadable(self, temp_dir: Path):
        """Le fichier créé est chargeable."""
        config_path = temp_dir / ".converterrc"
        create_default_config(config_path)

        loaded = Config.load(config_path)
        assert loaded.method == "auto"
