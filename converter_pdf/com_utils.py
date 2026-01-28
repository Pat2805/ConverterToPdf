"""
Utilitaires COM robustes pour Microsoft Office.

Ce module résout les problèmes de conflit avec Word/Excel déjà ouverts:
- Utilisation systématique de DispatchEx (nouvelle instance)
- Context managers pour garantir le nettoyage
- Timeout sur les opérations COM
- Kill process en cas de blocage

IMPORTANT:
- DispatchEx = crée toujours une NOUVELLE instance Office
- Dispatch = peut se connecter à une instance existante (à éviter!)
"""

from __future__ import annotations

import subprocess
import threading
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError
from contextlib import contextmanager
from typing import Any, Callable, TypeVar

from .logger import ConverterLogger, get_logger

# Import conditionnel de pywin32
try:
    import pythoncom
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    pythoncom = None  # type: ignore
    win32com = None  # type: ignore


T = TypeVar("T")


class COMError(Exception):
    """Exception pour les erreurs COM."""
    pass


class COMTimeoutError(COMError):
    """Exception pour les timeouts COM."""
    pass


class COMNotAvailableError(COMError):
    """Exception quand pywin32 n'est pas installé."""
    pass


def check_com_available() -> None:
    """Vérifie que pywin32 est disponible."""
    if not WIN32COM_AVAILABLE:
        raise COMNotAvailableError(
            "pywin32 n'est pas installé. "
            "Installez-le avec: pip install pywin32"
        )


@contextmanager
def com_context():
    """
    Context manager pour CoInitialize/CoUninitialize propre.

    IMPORTANT: Chaque thread qui utilise COM doit appeler CoInitialize.
    Ce context manager garantit que CoUninitialize est appelé même en cas d'erreur.

    Usage:
        with com_context():
            app = win32com.client.DispatchEx("Word.Application")
            # ... utiliser app ...
    """
    check_com_available()
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass  # Ignorer les erreurs de cleanup


def create_office_app(
    app_name: str,
    visible: bool = False,
    display_alerts: bool = False,
    logger: ConverterLogger | None = None,
) -> Any:
    """
    Crée une NOUVELLE instance d'application Office.

    CRITIQUE: Utilise DispatchEx, pas Dispatch!
    - DispatchEx = crée une nouvelle instance (ce qu'on veut)
    - Dispatch = peut se connecter à une instance existante (problèmes!)

    Args:
        app_name: Nom de l'application ("Word.Application", "Excel.Application", etc.)
        visible: Rendre l'application visible (False par défaut)
        display_alerts: Afficher les alertes Office (False par défaut)
        logger: Logger optionnel pour debug

    Returns:
        Instance COM de l'application Office

    Raises:
        COMError: Si la création échoue
    """
    check_com_available()
    log = logger or get_logger()

    try:
        log.debug(f"Création nouvelle instance {app_name} via DispatchEx")

        # TOUJOURS utiliser DispatchEx pour créer une nouvelle instance
        app = win32com.client.DispatchEx(app_name)

        # Configuration silencieuse
        app.Visible = visible

        # DisplayAlerts: 0 = wdAlertsNone (désactive toutes les alertes)
        try:
            app.DisplayAlerts = 0 if not display_alerts else -1
        except AttributeError:
            pass  # Certaines apps n'ont pas cette propriété

        # AutomationSecurity: 3 = msoAutomationSecurityForceDisable
        # Désactive les macros pendant l'automatisation
        try:
            app.AutomationSecurity = 3
        except AttributeError:
            pass

        log.debug(f"{app_name} créé avec succès (nouvelle instance)")
        return app

    except Exception as e:
        log.error(f"Échec création {app_name}: {e}", exc=e)
        raise COMError(f"Impossible de créer {app_name}: {e}") from e


def quit_office_app(app: Any, logger: ConverterLogger | None = None) -> None:
    """
    Ferme proprement une application Office.

    Args:
        app: Instance COM de l'application
        logger: Logger optionnel
    """
    log = logger or get_logger()

    if app is None:
        return

    try:
        log.debug("Fermeture de l'application Office")
        app.Quit()
        log.debug("Application fermée avec succès")
    except Exception as e:
        log.warning(f"Erreur lors de la fermeture Office: {e}")


@contextmanager
def office_app_context(
    app_name: str,
    visible: bool = False,
    display_alerts: bool = False,
    logger: ConverterLogger | None = None,
):
    """
    Context manager complet pour une application Office.

    Garantit:
    - CoInitialize/CoUninitialize propre
    - Création d'une nouvelle instance (DispatchEx)
    - Fermeture même en cas d'erreur

    Usage:
        with office_app_context("Word.Application") as word:
            doc = word.Documents.Open(path)
            doc.ExportAsFixedFormat(...)
            doc.Close()
        # word.Quit() est appelé automatiquement

    Args:
        app_name: Nom de l'application
        visible: Rendre visible
        display_alerts: Afficher les alertes
        logger: Logger optionnel

    Yields:
        Instance COM de l'application
    """
    log = logger or get_logger()
    app = None

    with com_context():
        try:
            app = create_office_app(
                app_name,
                visible=visible,
                display_alerts=display_alerts,
                logger=log,
            )
            yield app
        finally:
            quit_office_app(app, logger=log)


def run_with_timeout(
    func: Callable[[], T],
    timeout_seconds: int = 60,
    logger: ConverterLogger | None = None,
) -> T:
    """
    Exécute une fonction avec timeout.

    Utile pour les opérations COM qui peuvent bloquer indéfiniment.

    Args:
        func: Fonction à exécuter (sans arguments)
        timeout_seconds: Timeout en secondes
        logger: Logger optionnel

    Returns:
        Résultat de la fonction

    Raises:
        COMTimeoutError: Si le timeout est dépassé
    """
    log = logger or get_logger()

    with ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(func)
        try:
            return future.result(timeout=timeout_seconds)
        except FuturesTimeoutError:
            log.error(f"Timeout après {timeout_seconds}s")
            # Tenter de tuer les processus Office bloqués
            kill_office_processes(logger=log)
            raise COMTimeoutError(
                f"Opération COM timeout après {timeout_seconds}s"
            )


def kill_office_processes(
    processes: list[str] | None = None,
    logger: ConverterLogger | None = None,
) -> None:
    """
    Tue les processus Office orphelins.

    À utiliser en dernier recours si une opération COM bloque.

    Args:
        processes: Liste des noms de processus (par défaut: Word, Excel, PowerPoint)
        logger: Logger optionnel
    """
    log = logger or get_logger()

    if processes is None:
        processes = ["WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE"]

    for proc_name in processes:
        try:
            result = subprocess.run(
                ["taskkill", "/F", "/IM", proc_name],
                capture_output=True,
                text=True,
                check=False,
            )
            if result.returncode == 0:
                log.warning(f"Processus {proc_name} tué")
            # returncode != 0 signifie généralement que le processus n'existe pas
        except Exception as e:
            log.debug(f"Impossible de tuer {proc_name}: {e}")


def is_password_error(error: Exception | str) -> bool:
    """
    Détecte si une erreur indique un fichier protégé par mot de passe.

    Args:
        error: Exception ou message d'erreur

    Returns:
        True si l'erreur indique une protection par mot de passe
    """
    try:
        msg = str(error).lower()
    except Exception:
        return False

    keywords = [
        "password",
        "mot de passe",
        "mdp",
        "protected",
        "protégé",
        "protege",
        "protection",
        "encrypt",
        "encrypted",
        "chiffré",
        "chiffre",
        "cannot be opened because it is password",
        "the password is incorrect",
        "requires a password",
        "un mot de passe est requis",
    ]

    return any(keyword in msg for keyword in keywords)


def detect_office_installation(logger: ConverterLogger | None = None) -> dict[str, bool]:
    """
    Détecte quelles applications Office sont installées.

    Returns:
        Dict avec les applications disponibles
    """
    log = logger or get_logger()
    result = {
        "word": False,
        "excel": False,
        "powerpoint": False,
        "outlook": False,
    }

    if not WIN32COM_AVAILABLE:
        log.warning("pywin32 non installé, impossible de détecter Office")
        return result

    apps = {
        "word": "Word.Application",
        "excel": "Excel.Application",
        "powerpoint": "PowerPoint.Application",
        "outlook": "Outlook.Application",
    }

    with com_context():
        for name, prog_id in apps.items():
            try:
                app = win32com.client.DispatchEx(prog_id)
                app.Quit()
                result[name] = True
                log.debug(f"{name.capitalize()} détecté")
            except Exception:
                log.debug(f"{name.capitalize()} non disponible")

    return result
