"""Entry point: sets up logging, gates on license activation, launches the
Tk main loop. Run with:  python app.py   (from inside this directory).
"""
import tkinter as tk

import config
import gui
import licensing


def ensure_activated() -> bool:
    if licensing.esta_activado():
        return True
    return gui.mostrar_dialogo_activacion(licensing.obtener_id_equipo())


def main() -> None:
    config.setup_logging()
    if not ensure_activated():
        return
    root = tk.Tk()
    root.geometry(gui.WINDOW_GEOMETRY)
    gui.EtiquetadoraApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
