"""
Système de logging structuré pour ConverterToPdf.

Fournit un logging multi-niveaux avec:
- Sortie console colorée
- Fichier de log rotatif (optionnel)
- Contexte fichier en cours de traitement
- Format structuré pour debug efficace
"""

import logging
import sys
from contextlib import contextmanager
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any


class ColoredFormatter(logging.Formatter):
    """Formatter avec couleurs ANSI pour la console."""

    COLORS = {
        logging.DEBUG: "\033[36m",     # Cyan
        logging.INFO: "\033[32m",      # Vert
        logging.WARNING: "\033[33m",   # Jaune
        logging.ERROR: "\033[31m",     # Rouge
        logging.CRITICAL: "\033[35m",  # Magenta
    }
    RESET = "\033[0m"
    BOLD = "\033[1m"

    def __init__(self, fmt: str | None = None, include_colors: bool = True):
        super().__init__(fmt)
        self.include_colors = include_colors

    def format(self, record: logging.LogRecord) -> str:
        # Ajouter le contexte fichier si présent
        file_context = getattr(record, "current_file", None)
        if file_context:
            record.file_ctx = f"[{file_context}] "
        else:
            record.file_ctx = ""

        # Formater le message de base
        message = super().format(record)

        # Ajouter les couleurs si activées
        if self.include_colors and sys.stdout.isatty():
            color = self.COLORS.get(record.levelno, "")
            level_name = record.levelname
            # Colorer uniquement le niveau
            message = message.replace(
                level_name,
                f"{color}{self.BOLD}{level_name}{self.RESET}",
                1
            )

        return message


class FileContextFilter(logging.Filter):
    """Filter qui ajoute le contexte fichier aux records."""

    def __init__(self, logger_instance: "ConverterLogger"):
        super().__init__()
        self.logger_instance = logger_instance

    def filter(self, record: logging.LogRecord) -> bool:
        record.current_file = self.logger_instance._current_file
        return True


class ConverterLogger:
    """
    Logger principal pour ConverterToPdf.

    Usage:
        logger = ConverterLogger()
        logger.setup(level="DEBUG", log_file=Path("converter.log"))

        with logger.file_context(Path("document.docx")):
            logger.info("Conversion démarrée")
            logger.debug("Ouverture via COM")
            logger.info("Conversion réussie", duration=2.3, method="office")
    """

    # Format par défaut
    DEFAULT_FORMAT = "%(asctime)s | %(levelname)-7s | %(file_ctx)s%(message)s"
    DEFAULT_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

    def __init__(self, name: str = "converter_pdf"):
        self.name = name
        self.logger = logging.getLogger(name)
        self._current_file: str | None = None
        self._setup_done = False

    def setup(
        self,
        level: str = "INFO",
        log_file: Path | None = None,
        log_file_level: str = "DEBUG",
        max_file_size: int = 10 * 1024 * 1024,  # 10 MB
        backup_count: int = 5,
        console_colors: bool = True,
    ) -> None:
        """
        Configure le logger avec handlers console et fichier.

        Args:
            level: Niveau de log pour la console (DEBUG, INFO, WARNING, ERROR)
            log_file: Chemin du fichier de log (optionnel)
            log_file_level: Niveau de log pour le fichier (par défaut DEBUG)
            max_file_size: Taille max du fichier avant rotation (10 MB par défaut)
            backup_count: Nombre de fichiers de backup à conserver
            console_colors: Activer les couleurs dans la console
        """
        if self._setup_done:
            # Éviter la double configuration
            return

        self.logger.setLevel(logging.DEBUG)  # Le logger accepte tout
        self.logger.handlers.clear()  # Nettoyer les handlers existants

        # Ajouter le filter de contexte fichier
        context_filter = FileContextFilter(self)
        self.logger.addFilter(context_filter)

        # Handler console
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(getattr(logging, level.upper()))
        console_formatter = ColoredFormatter(
            fmt=self.DEFAULT_FORMAT,
            include_colors=console_colors,
        )
        console_formatter.datefmt = self.DEFAULT_DATE_FORMAT
        console_handler.setFormatter(console_formatter)
        self.logger.addHandler(console_handler)

        # Handler fichier (optionnel)
        if log_file:
            log_file = Path(log_file)
            log_file.parent.mkdir(parents=True, exist_ok=True)

            file_handler = RotatingFileHandler(
                log_file,
                maxBytes=max_file_size,
                backupCount=backup_count,
                encoding="utf-8",
            )
            file_handler.setLevel(getattr(logging, log_file_level.upper()))
            file_formatter = logging.Formatter(
                fmt=self.DEFAULT_FORMAT,
                datefmt=self.DEFAULT_DATE_FORMAT,
            )
            file_handler.setFormatter(file_formatter)
            self.logger.addHandler(file_handler)

            self.info(f"Fichier de log: {log_file}")

        self._setup_done = True

    @contextmanager
    def file_context(self, source_file: Path | str):
        """
        Context manager pour tracker le fichier en cours de traitement.

        Usage:
            with logger.file_context(Path("document.docx")):
                logger.info("Traitement en cours")
                # Tous les logs afficheront [document.docx]
        """
        previous = self._current_file
        self._current_file = Path(source_file).name if source_file else None
        try:
            yield
        finally:
            self._current_file = previous

    def _log(
        self,
        level: int,
        msg: str,
        exc: Exception | None = None,
        **extra: Any,
    ) -> None:
        """Méthode interne de logging avec extras formatés."""
        # Formater les extras dans le message
        if extra:
            extras_str = ", ".join(f"{k}={v}" for k, v in extra.items())
            msg = f"{msg} ({extras_str})"

        if exc:
            self.logger.log(level, msg, exc_info=exc)
        else:
            self.logger.log(level, msg)

    def debug(self, msg: str, **extra: Any) -> None:
        """Log niveau DEBUG - détails techniques pour debug."""
        self._log(logging.DEBUG, msg, **extra)

    def info(self, msg: str, **extra: Any) -> None:
        """Log niveau INFO - informations générales."""
        self._log(logging.INFO, msg, **extra)

    def warning(self, msg: str, **extra: Any) -> None:
        """Log niveau WARNING - avertissements non bloquants."""
        self._log(logging.WARNING, msg, **extra)

    def error(self, msg: str, exc: Exception | None = None, **extra: Any) -> None:
        """Log niveau ERROR - erreurs avec traceback optionnel."""
        self._log(logging.ERROR, msg, exc=exc, **extra)

    def critical(self, msg: str, exc: Exception | None = None, **extra: Any) -> None:
        """Log niveau CRITICAL - erreurs fatales."""
        self._log(logging.CRITICAL, msg, exc=exc, **extra)

    def success(self, msg: str, **extra: Any) -> None:
        """Log de succès (niveau INFO avec préfixe)."""
        self.info(f"[OK] {msg}", **extra)

    def fail(self, msg: str, exc: Exception | None = None, **extra: Any) -> None:
        """Log d'échec (niveau ERROR avec préfixe)."""
        self.error(f"[FAIL] {msg}", exc=exc, **extra)

    def skip(self, msg: str, reason: str = "", **extra: Any) -> None:
        """Log de skip (niveau WARNING avec préfixe)."""
        full_msg = f"[SKIP] {msg}"
        if reason:
            full_msg += f" - {reason}"
        self.warning(full_msg, **extra)


# Instance globale pour usage simplifié
_default_logger: ConverterLogger | None = None


def get_logger() -> ConverterLogger:
    """Retourne le logger par défaut (crée si nécessaire)."""
    global _default_logger
    if _default_logger is None:
        _default_logger = ConverterLogger()
    return _default_logger


def setup_logging(
    level: str = "INFO",
    log_file: Path | None = None,
    **kwargs: Any,
) -> ConverterLogger:
    """Configure et retourne le logger par défaut."""
    logger = get_logger()
    logger.setup(level=level, log_file=log_file, **kwargs)
    return logger
