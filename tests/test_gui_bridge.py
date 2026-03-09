"""Tests pour le pont GUI <-> Backend."""

import logging
import queue
import time
import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from converter_pdf.config import Config
from converter_pdf.gui_bridge import ConversionBridge, GUILogHandler
from converter_pdf.signals import SignalHub


class TestGUILogHandler:
    """Tests pour GUILogHandler."""

    def test_handler_captures_log_message(self):
        q = queue.Queue()
        handler = GUILogHandler(q)
        handler.setFormatter(logging.Formatter("%(message)s"))

        record = logging.LogRecord(
            name="test", level=logging.INFO, pathname="", lineno=0,
            msg="Test message", args=(), exc_info=None,
        )
        handler.emit(record)

        msg = q.get_nowait()
        assert msg[0] == "log"
        assert msg[1] == "INFO"
        assert "Test message" in msg[2]

    def test_handler_captures_level(self):
        q = queue.Queue()
        handler = GUILogHandler(q)
        handler.setFormatter(logging.Formatter("%(message)s"))

        for level_name, level in [("WARNING", logging.WARNING), ("ERROR", logging.ERROR)]:
            record = logging.LogRecord(
                name="test", level=level, pathname="", lineno=0,
                msg="msg", args=(), exc_info=None,
            )
            handler.emit(record)
            msg = q.get_nowait()
            assert msg[1] == level_name


class TestConversionBridge:
    """Tests pour ConversionBridge."""

    def test_pre_scan_counts_files(self, temp_dir, file_factory):
        """pre_scan compte correctement les fichiers supportes."""
        file_factory.create_text_file("doc1.txt", "hello")
        file_factory.create_text_file("doc2.txt", "world")
        file_factory.create_text_file("ignore.xyz", "not supported")  # extension non supportee

        config = Config()
        bridge = ConversionBridge(config)
        count = bridge.pre_scan(temp_dir)
        assert count == 2  # Only .txt files are supported

    def test_pre_scan_recursive(self, temp_dir, file_factory):
        """pre_scan avec mode recursif."""
        file_factory.create_text_file("doc1.txt", "hello")
        sub = file_factory.create_subdirectory("sub")
        (sub / "doc2.txt").write_text("world")

        config = Config(recursive=True)
        bridge = ConversionBridge(config)
        count = bridge.pre_scan(temp_dir)
        assert count == 2

    def test_pre_scan_non_recursive(self, temp_dir, file_factory):
        """pre_scan sans mode recursif n'inclut pas les sous-dossiers."""
        file_factory.create_text_file("doc1.txt", "hello")
        sub = file_factory.create_subdirectory("sub")
        (sub / "doc2.txt").write_text("world")

        config = Config(recursive=False)
        bridge = ConversionBridge(config)
        count = bridge.pre_scan(temp_dir)
        assert count == 1

    def test_is_running_initially_false(self):
        config = Config()
        bridge = ConversionBridge(config)
        assert bridge.is_running is False

    def test_poll_queue_empty(self):
        config = Config()
        bridge = ConversionBridge(config)
        assert bridge.poll_queue() == []

    def test_cancel_without_processor(self):
        """cancel() sans processeur ne crash pas."""
        config = Config()
        bridge = ConversionBridge(config)
        bridge.cancel()  # Should not raise

    def test_signals_connected(self):
        """Les signaux sont initialises par defaut."""
        config = Config()
        signals = SignalHub()
        bridge = ConversionBridge(config, signals)
        assert bridge.signals is signals

    def test_start_conversion_sets_running(self, temp_dir, file_factory):
        """start_conversion lance le thread worker."""
        file_factory.create_text_file("test.txt", "content")

        config = Config(dry_run=True, report_enabled=False)
        bridge = ConversionBridge(config)
        bridge.start_conversion(temp_dir)

        # Wait for thread to finish
        if bridge._worker_thread:
            bridge._worker_thread.join(timeout=10)

        # Should have messages in queue
        messages = bridge.poll_queue()
        msg_types = [m[0] for m in messages]
        assert "progress" in msg_types
        assert "done" in msg_types

    def test_double_start_ignored(self, temp_dir, file_factory):
        """start_conversion ignore si deja en cours."""
        file_factory.create_text_file("test.txt", "content")

        config = Config(dry_run=True, report_enabled=False)
        bridge = ConversionBridge(config)

        # Simulate running state
        bridge._running = True
        bridge.start_conversion(temp_dir)

        # No thread should have been created
        assert bridge._worker_thread is None
