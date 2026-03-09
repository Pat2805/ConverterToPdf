"""Systeme de signaux thread-safe pour la communication GUI <-> Backend."""

from __future__ import annotations

import threading
from dataclasses import dataclass, field
from typing import Any, Callable


@dataclass
class SignalHub:
    """Hub de signaux thread-safe pour la communication entre composants."""

    _listeners: dict[str, list[Callable[..., Any]]] = field(
        default_factory=lambda: {
            "on_file_start": [],
            "on_file_complete": [],
            "on_progress": [],
            "on_log": [],
            "on_session_complete": [],
            "on_error": [],
        }
    )
    _lock: threading.Lock = field(default_factory=threading.Lock)

    def connect(self, signal_name: str, callback: Callable[..., Any]) -> None:
        """Enregistre un callback pour un signal."""
        with self._lock:
            if signal_name not in self._listeners:
                raise ValueError(f"Signal inconnu: {signal_name}")
            self._listeners[signal_name].append(callback)

    def disconnect(self, signal_name: str, callback: Callable[..., Any]) -> None:
        """Deconnecte un callback."""
        with self._lock:
            if signal_name in self._listeners:
                self._listeners[signal_name].remove(callback)

    def emit(self, signal_name: str, *args: Any, **kwargs: Any) -> None:
        """Emet un signal (appelle tous les callbacks enregistres)."""
        with self._lock:
            listeners = list(self._listeners.get(signal_name, []))
        for listener in listeners:
            try:
                listener(*args, **kwargs)
            except Exception:
                pass  # Les listeners ne doivent pas casser l'emetteur

    def disconnect_all(self) -> None:
        """Deconnecte tous les callbacks."""
        with self._lock:
            for key in self._listeners:
                self._listeners[key] = []
