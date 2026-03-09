"""Pont entre la GUI et le backend de conversion."""

from __future__ import annotations

import logging
import queue
import threading
from pathlib import Path
from typing import Any

from .config import Config
from .logger import ConverterLogger
from .processor import FileProcessor
from .signals import SignalHub


class GUILogHandler(logging.Handler):
    """Handler de logging qui redirige les messages vers une queue pour la GUI."""

    def __init__(self, message_queue: queue.Queue):
        super().__init__(level=logging.DEBUG)
        self._queue = message_queue

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            self._queue.put(("log", record.levelname, msg))
        except Exception:
            pass


class ConversionBridge:
    """
    Pont entre le FileProcessor et la GUI.

    Lance les conversions dans un thread worker, communique via queue.Queue.
    La GUI poll la queue avec after() pour rester responsive.
    """

    def __init__(self, config: Config, signals: SignalHub | None = None):
        self.config = config
        self.signals = signals or SignalHub()
        self._queue: queue.Queue = queue.Queue()
        self._worker_thread: threading.Thread | None = None
        self._processor: FileProcessor | None = None
        self._total_files: int = 0
        self._processed_files: int = 0
        self._running = False

    def pre_scan(self, directory: Path) -> int:
        """Compte les fichiers a traiter (meme logique que FileProcessor)."""
        pattern = "**/*" if self.config.recursive else "*"
        extensions = set(self.config.get_all_extensions())
        count = 0
        for file_path in directory.glob(pattern):
            if file_path.is_file() and file_path.suffix.lower() in extensions:
                count += 1
        return count

    def start_conversion(self, path: Path, dest_dir: Path | None = None) -> None:
        """Lance la conversion dans un thread worker."""
        if self._running:
            return

        self._running = True
        self._processed_files = 0

        if path.is_dir():
            self._total_files = self.pre_scan(path)
        else:
            self._total_files = 1

        self.signals.emit("on_progress", 0, self._total_files)
        self._queue.put(("progress", 0, self._total_files))

        self._worker_thread = threading.Thread(
            target=self._run_worker,
            args=(path, dest_dir),
            daemon=True,
        )
        self._worker_thread.start()

    def _run_worker(self, path: Path, dest_dir: Path | None) -> None:
        """Thread worker : execute la conversion."""
        try:
            # Logger frais pour ce run (evite le guard _setup_done)
            logger = ConverterLogger("converter_pdf_gui")
            logger.setup(level=self.config.log_level, console_colors=False)

            # Ajouter le handler GUI pour capturer les logs
            gui_handler = GUILogHandler(self._queue)
            gui_handler.setFormatter(logging.Formatter(
                "%(asctime)s | %(levelname)-7s | %(message)s",
                datefmt="%H:%M:%S",
            ))
            logger.logger.addHandler(gui_handler)

            # Creer le processeur
            processor = FileProcessor(self.config, logger)
            self._processor = processor

            # Neutraliser le signal handler (ne fonctionne que depuis le main thread)
            processor._setup_signal_handler = lambda: None

            # Wrapper process_file pour emettre les signaux de progression
            original_process_file = processor.process_file

            def wrapped_process_file(source, dest_dir_inner=None):
                self.signals.emit("on_file_start", source)
                self._queue.put(("file_start", str(source)))
                result = original_process_file(source, dest_dir_inner)
                self._processed_files += 1
                self.signals.emit("on_file_complete", result)
                self.signals.emit("on_progress", self._processed_files, self._total_files)
                self._queue.put(("file_complete", result))
                self._queue.put(("progress", self._processed_files, self._total_files))
                return result

            processor.process_file = wrapped_process_file

            # Lancer la conversion
            if path.is_file():
                processor.process_file(path, dest_dir)
            else:
                processor.process_directory(path, dest_dir)

            # Session terminee
            self._queue.put(("session_complete", processor.report, processor.stats))
            self.signals.emit("on_session_complete", processor.report)

        except Exception as e:
            self._queue.put(("error", e))
            self.signals.emit("on_error", e)
        finally:
            self._running = False
            self._queue.put(("done",))

    def cancel(self) -> None:
        """Demande l'annulation de la conversion en cours."""
        if self._processor:
            self._processor._interrupted = True

    @property
    def is_running(self) -> bool:
        return self._running

    def poll_queue(self) -> list[tuple]:
        """Draine tous les messages en attente (appele depuis le thread GUI)."""
        messages = []
        while True:
            try:
                msg = self._queue.get_nowait()
                messages.append(msg)
            except queue.Empty:
                break
        return messages
