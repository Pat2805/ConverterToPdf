"""Module GUI pour ConverterToPdf utilisant customtkinter."""


def launch_gui() -> None:
    """Lance l'application GUI."""
    from .app import ConverterApp
    app = ConverterApp()
    app.mainloop()
