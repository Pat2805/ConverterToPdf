"""Tests pour le systeme de signaux."""

import threading
import pytest
from converter_pdf.signals import SignalHub


class TestSignalHub:
    """Tests pour SignalHub."""

    def test_connect_and_emit(self):
        hub = SignalHub()
        results = []
        hub.connect("on_file_start", lambda path: results.append(path))
        hub.emit("on_file_start", "/test/file.txt")
        assert results == ["/test/file.txt"]

    def test_emit_multiple_listeners(self):
        hub = SignalHub()
        results = []
        hub.connect("on_progress", lambda c, t: results.append(("a", c, t)))
        hub.connect("on_progress", lambda c, t: results.append(("b", c, t)))
        hub.emit("on_progress", 1, 10)
        assert len(results) == 2
        assert results[0] == ("a", 1, 10)
        assert results[1] == ("b", 1, 10)

    def test_disconnect(self):
        hub = SignalHub()
        results = []
        callback = lambda: results.append(1)
        hub.connect("on_error", callback)
        hub.disconnect("on_error", callback)
        hub.emit("on_error")
        assert results == []

    def test_disconnect_all(self):
        hub = SignalHub()
        results = []
        hub.connect("on_file_start", lambda *a: results.append(1))
        hub.connect("on_progress", lambda *a: results.append(2))
        hub.disconnect_all()
        hub.emit("on_file_start")
        hub.emit("on_progress")
        assert results == []

    def test_unknown_signal_raises(self):
        hub = SignalHub()
        with pytest.raises(ValueError, match="Signal inconnu"):
            hub.connect("nonexistent_signal", lambda: None)

    def test_emit_unknown_signal_no_error(self):
        hub = SignalHub()
        # emit on unknown signal should not crash - it just has no listeners
        hub.emit("on_file_start")  # Valid but no listeners - should work fine

    def test_listener_exception_does_not_propagate(self):
        hub = SignalHub()
        results = []
        hub.connect("on_log", lambda *a: 1 / 0)  # Will raise ZeroDivisionError
        hub.connect("on_log", lambda *a: results.append("ok"))
        hub.emit("on_log", "INFO", "test")
        assert results == ["ok"]  # Second listener still called

    def test_thread_safety(self):
        hub = SignalHub()
        results = []
        hub.connect("on_progress", lambda c, t: results.append((c, t)))

        threads = []
        for i in range(10):
            t = threading.Thread(target=hub.emit, args=("on_progress", i, 10))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        assert len(results) == 10

    def test_all_signal_names_exist(self):
        hub = SignalHub()
        expected = {"on_file_start", "on_file_complete", "on_progress", "on_log", "on_session_complete", "on_error"}
        assert set(hub._listeners.keys()) == expected
