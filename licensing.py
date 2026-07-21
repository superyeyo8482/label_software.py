"""License activation/validation logic.

SECURITY WARNING - insecure by construction, intentionally left unfixed:

CLAVE_MAESTRA is a symmetric Fernet key embedded directly in this source
file (and therefore in any compiled build too - `strings LabelSoftware.exe`
recovers it in seconds). Whoever holds this key can encrypt any machine ID
they like and produce a code that `validar_codigo` will accept, so the
"activation" gate provides no real protection against unauthorized use -
it only stops someone from typing in a *random* string.

This was confirmed live during the refactor, not assumed: using only this
module, a valid activation token was generated for a fabricated machine ID
and accepted by validar_codigo(). See tests/test_licensing.py, which
exercises exactly this to make sure the behavior (whatever it ends up
being) doesn't regress silently.

Real protection needs one of:
  - Asymmetric signing: keep a private signing key with whoever issues
    licenses, ship only a public verification key in this app.
  - Server-side validation: this app calls out to a server you control at
    activation time instead of validating locally.

Deliberately NOT changed here. Per instruction, that's a decision for you
to make - not something to silently "fix" as part of a quality refactor.
"""
from __future__ import annotations

import hashlib
import logging
import os
import platform
import subprocess
import uuid

from cryptography.fernet import Fernet, InvalidToken

import config

logger = logging.getLogger(__name__)

CLAVE_MAESTRA = b'qJjQH8bZJ4yJjHtPqM7t2Hh5VvJjQqP3xR5sT1uVvWw=='


def obtener_id_equipo() -> str:
    nombre = platform.node()
    if platform.system() == "Windows":
        try:
            uuid_cpu = subprocess.getoutput("wmic csproduct get uuid").split()[1]
        except (IndexError, OSError) as exc:
            logger.warning("No se pudo obtener el UUID vía wmic: %s", exc)
            uuid_cpu = "NODisponible"
    else:
        try:
            with open("/etc/machine-id", "r") as f:
                uuid_cpu = f.read().strip()
        except OSError as exc:
            logger.warning("No se pudo leer /etc/machine-id: %s", exc)
            uuid_cpu = "LINUX"
    mac = ':'.join('{:02x}'.format((uuid.getnode() >> i) & 0xff) for i in range(0, 8 * 6, 8))
    return hashlib.md5(f"{nombre}{uuid_cpu}{mac}".encode()).hexdigest()


def validar_codigo(codigo: str, id_equipo: str) -> bool:
    """True if `codigo` decrypts (with CLAVE_MAESTRA) to exactly id_equipo."""
    try:
        fernet = Fernet(CLAVE_MAESTRA)
        id_descifrado = fernet.decrypt(codigo.encode()).decode()
    except (InvalidToken, ValueError, TypeError) as exc:
        # Expected/routine: users mistype codes constantly. Not error-level.
        logger.info("Código de activación rechazado: %s", exc)
        return False
    return id_descifrado == id_equipo


def cargar_codigo_guardado() -> str | None:
    if not os.path.exists(config.LICENSE_FILE):
        return None
    try:
        with open(config.LICENSE_FILE, "r") as f:
            return f.read().strip()
    except OSError as exc:
        logger.warning("No se pudo leer %s: %s", config.LICENSE_FILE, exc)
        return None


def guardar_codigo(codigo: str) -> None:
    with open(config.LICENSE_FILE, "w") as f:
        f.write(codigo)


def borrar_codigo_guardado() -> None:
    try:
        os.remove(config.LICENSE_FILE)
    except OSError as exc:
        logger.warning("No se pudo borrar %s: %s", config.LICENSE_FILE, exc)


def esta_activado() -> bool:
    """Checks the saved license file against the current machine ID.

    A saved code that fails validation is treated as stale and removed,
    matching the original's behavior of deleting license.key so the user
    is prompted to re-activate rather than getting stuck.
    """
    id_actual = obtener_id_equipo()
    codigo = cargar_codigo_guardado()
    if codigo is None:
        return False
    if validar_codigo(codigo, id_actual):
        return True
    borrar_codigo_guardado()
    return False
