"""Widget de visualisation de logs avec coloration syntaxique."""

import customtkinter as ctk


class LogViewer(ctk.CTkFrame):
    """Zone de logs scrollable avec coloration par niveau."""

    MAX_LINES = 5000

    def __init__(self, parent, **kwargs):
        kwargs.setdefault("fg_color", "transparent")
        super().__init__(parent, **kwargs)

        self.textbox = ctk.CTkTextbox(
            self, state="disabled", wrap="word",
            font=ctk.CTkFont(family="Consolas", size=12),
            corner_radius=6,
        )
        self.textbox.pack(fill="both", expand=True)

        # Tags de couleur sur le widget Text interne
        inner = self.textbox._textbox
        inner.tag_config("ERROR", foreground="#FF6B6B")
        inner.tag_config("WARNING", foreground="#FFB347")
        inner.tag_config("DEBUG", foreground="#777777")
        inner.tag_config("INFO", foreground="#C8C8C8")
        inner.tag_config("CRITICAL", foreground="#FF4444", underline=True)

    def append_log(self, level: str, message: str) -> None:
        """Ajoute une ligne de log avec coloration."""
        self.textbox.configure(state="normal")
        tag = level.upper() if level.upper() in (
            "ERROR", "WARNING", "DEBUG", "INFO", "CRITICAL",
        ) else "INFO"
        self.textbox._textbox.insert("end", message + "\n", tag)
        self._trim_lines()
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def _trim_lines(self) -> None:
        line_count = int(self.textbox._textbox.index("end-1c").split(".")[0])
        if line_count > self.MAX_LINES:
            excess = line_count - self.MAX_LINES
            self.textbox._textbox.delete("1.0", f"{excess}.0")

    def clear(self) -> None:
        """Efface tous les logs."""
        self.textbox.configure(state="normal")
        self.textbox._textbox.delete("1.0", "end")
        self.textbox.configure(state="disabled")
