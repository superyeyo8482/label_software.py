"""ZPL label generation and printer output.

win32print is Windows-only and is isolated behind PrinterBackend so that
ZPL generation and every other module can be imported and unit-tested on
non-Windows machines. Only WindowsZplBackend touches win32print, and only
at send()-time (lazy import), so importing this module never requires
pywin32 to be installed.
"""
from __future__ import annotations

import logging
import platform
import time
from typing import Callable

import config

logger = logging.getLogger(__name__)


class PrintError(Exception):
    """A label could not be sent to the printer/output device."""


class PrinterBackend:
    def send(self, data: bytes) -> None:
        raise NotImplementedError


class WindowsZplBackend(PrinterBackend):
    def __init__(self, printer_name: str = config.NOMBRE_IMPRESORA):
        self.printer_name = printer_name

    def send(self, data: bytes) -> None:
        import win32print

        try:
            handle = win32print.OpenPrinter(self.printer_name)
        except Exception as exc:
            # win32print raises pywintypes.error, which isn't reliably
            # importable/typed cross-platform; this is a genuine hardware/
            # driver boundary, so a broad catch here (re-raised as our own
            # typed PrintError, logged, never silently swallowed) is
            # appropriate rather than a narrower exception list.
            raise PrintError(f"No se pudo abrir la impresora '{self.printer_name}': {exc}") from exc
        try:
            win32print.StartDocPrinter(handle, 1, ("etiqueta", None, "RAW"))
            win32print.StartPagePrinter(handle)
            win32print.WritePrinter(handle, data)
            win32print.EndPagePrinter(handle)
            win32print.EndDocPrinter(handle)
        except Exception as exc:
            raise PrintError(f"No se pudo imprimir: {exc}") from exc
        finally:
            win32print.ClosePrinter(handle)


class PosixFileBackend(PrinterBackend):
    """Non-Windows fallback: writes raw ZPL to a USB printer device node."""

    def __init__(self, device_path: str = config.POSIX_PRINTER_DEVICE):
        self.device_path = device_path

    def send(self, data: bytes) -> None:
        try:
            with open(self.device_path, "wb") as lp:
                lp.write(data)
        except OSError as exc:
            raise PrintError(f"No se pudo imprimir (Linux): {exc}") from exc


def get_default_backend() -> PrinterBackend:
    if platform.system() == "Windows":
        return WindowsZplBackend()
    return PosixFileBackend()


def generar_zpl(producto: dict, ls: int, lt: int) -> str:
    precio = producto['precio']
    try:
        precio_str = f"{float(precio):,.0f}"
    except (ValueError, TypeError):
        logger.warning("Precio no numérico para producto %s: %r", producto.get('id'), precio)
        precio_str = str(precio)
    return config.PLANTILLA_ZPL.format(
        ls=ls, lt=lt,
        precio=precio_str,
        proveedor=producto['proveedor'].upper(),
        idProducto=producto['id'].upper(),
        material=producto['material'].upper(),
        tipo=producto['tipo'].upper(),
        actualizado=producto['actualizado'].upper(),
        gramos=producto['gramos'].zfill(3),
    )


def enviar_zpl(zpl: str, backend: PrinterBackend | None = None) -> bool:
    """Returns True/False (same contract as the original) instead of
    popping a messagebox itself - the caller (gui.py) owns showing errors
    to the user, so this module has no tkinter dependency.
    """
    backend = backend or get_default_backend()
    try:
        backend.send(zpl.encode())
        return True
    except PrintError as exc:
        logger.error(str(exc))
        return False


def imprimir_lote(
    productos: list[dict],
    ls: int,
    lt: int,
    debe_detenerse: Callable[[], bool],
    on_progreso: Callable[[int], None] | None = None,
    backend: PrinterBackend | None = None,
    delay: float = 0.1,
) -> int:
    """Prints each product `cantidad` times, checking `debe_detenerse()`
    before every label so a caller can implement a STOP button. Raises
    PrintError on the first failed send (matches the original: abort the
    whole batch, don't keep going after a printer error). Returns the
    count of labels actually sent.

    This is the threading/UI-agnostic core of what was gui.py's
    `_impresion_hilo` in the original - gui.py now only owns starting the
    thread and updating widgets via `on_progreso`/exception handling.
    """
    backend = backend or get_default_backend()
    total = 0
    for producto in productos:
        if debe_detenerse():
            return total
        for _ in range(producto['cantidad']):
            if debe_detenerse():
                return total
            zpl = generar_zpl(producto, ls, lt)
            if not enviar_zpl(zpl, backend=backend):
                raise PrintError(f"No se pudo imprimir el producto {producto.get('id', '?')}")
            total += 1
            if on_progreso:
                on_progreso(total)
            if delay:
                time.sleep(delay)
    return total
