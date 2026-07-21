"""Static configuration: printer name, print offsets, file paths, ZPL
template, logging setup. No secrets live here - see licensing.py (and
REFACTOR_NOTES.md) for why the activation key does NOT belong in a module
like this one.
"""
import logging

NOMBRE_IMPRESORA = "ZDesigner ZD220-203dpi ZPL"

LS_VAL_DEFAULT = 50
LT_VAL_DEFAULT = 53

PLANTILLA_ZPL = """^XA
^PW600
^LS{ls}
^LT{lt}
^MD25
^PR4
^FO15,15^A0N,35,35^FD${precio}^FS
^FO15,75^A0N,20,20^FD{proveedor}^FS
^FO70,75^A0N,18,18^FD{idProducto}^FS
^FO15,105^A0N,20,20^FD{tipo}^FS
^FO115,105^A0N,18,18^FD{gramos}^FS
^FO15,135^A0N,20,20^FD{material}^FS
^FO115,135^A0N,18,18^FD{actualizado}^FS
^XZ"""

CONFIG_FILE = "etiquetadora_config.json"
LICENSE_FILE = "license.key"
LOG_FILE = "label_software.log"

POSIX_PRINTER_DEVICE = "/dev/usb/lp0"

DEFAULT_LANGUAGE = "es"


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
