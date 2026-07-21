from cryptography.fernet import Fernet

import config
import licensing


def test_validar_codigo_accepts_correctly_encrypted_machine_id():
    id_equipo = "abc123"
    token = Fernet(licensing.CLAVE_MAESTRA).encrypt(id_equipo.encode()).decode()
    assert licensing.validar_codigo(token, id_equipo) is True


def test_validar_codigo_rejects_token_for_a_different_machine():
    token = Fernet(licensing.CLAVE_MAESTRA).encrypt(b"other-machine").decode()
    assert licensing.validar_codigo(token, "abc123") is False


def test_validar_codigo_rejects_garbage_input():
    assert licensing.validar_codigo("not-a-real-token", "abc123") is False


def test_obtener_id_equipo_is_stable_and_hex():
    id1 = licensing.obtener_id_equipo()
    id2 = licensing.obtener_id_equipo()
    assert id1 == id2
    assert len(id1) == 32
    int(id1, 16)  # raises ValueError if not hex


def test_esta_activado_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "LICENSE_FILE", str(tmp_path / "license.key"))
    assert licensing.esta_activado() is False

    id_equipo = licensing.obtener_id_equipo()
    token = Fernet(licensing.CLAVE_MAESTRA).encrypt(id_equipo.encode()).decode()
    licensing.guardar_codigo(token)
    assert licensing.esta_activado() is True


def test_esta_activado_removes_stale_license_file(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "LICENSE_FILE", str(tmp_path / "license.key"))
    licensing.guardar_codigo("garbage-not-a-token")
    assert licensing.esta_activado() is False
    assert not (tmp_path / "license.key").exists()


def test_hardcoded_key_lets_anyone_mint_a_valid_activation_code():
    """This is not a bug in the test - it's a live demonstration of the
    flagged vulnerability (REFACTOR_NOTES.md #5 / licensing.py docstring).
    Anyone with this source file can produce a valid code for any machine
    ID, with no access to whatever system actually issues licenses.
    """
    forged_machine_id = "some-machine-i-dont-own"
    forged_token = Fernet(licensing.CLAVE_MAESTRA).encrypt(forged_machine_id.encode()).decode()
    assert licensing.validar_codigo(forged_token, forged_machine_id) is True
