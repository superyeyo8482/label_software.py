import printing


class FakeBackend(printing.PrinterBackend):
    def __init__(self, fail=False):
        self.sent = None
        self.fail = fail

    def send(self, data: bytes) -> None:
        if self.fail:
            raise printing.PrintError("boom")
        self.sent = data


def _producto(**overrides):
    base = dict(proveedor='ab', id='or001', material='oro 10k', tipo='anillo',
                actualizado='s', gramos='8', precio='150000')
    base.update(overrides)
    return base


def test_generar_zpl_uppercases_fields_and_pads_gramos():
    zpl = printing.generar_zpl(_producto(), ls=-50, lt=-43)
    assert '^LS-50' in zpl
    assert '^LT-43' in zpl
    assert 'AB' in zpl
    assert 'OR001' in zpl
    assert '008' in zpl  # gramos zfill(3)


def test_generar_zpl_formats_price_with_thousands_separator():
    zpl = printing.generar_zpl(_producto(precio='150000'), ls=0, lt=0)
    assert '$150,000' in zpl


def test_generar_zpl_non_numeric_price_falls_back_to_raw_string():
    zpl = printing.generar_zpl(_producto(precio='n/a'), ls=0, lt=0)
    assert '$n/a' in zpl


def test_enviar_zpl_success_returns_true_and_sends_bytes():
    backend = FakeBackend()
    assert printing.enviar_zpl("^XA^XZ", backend=backend) is True
    assert backend.sent == b"^XA^XZ"


def test_enviar_zpl_failure_returns_false_instead_of_raising():
    backend = FakeBackend(fail=True)
    assert printing.enviar_zpl("^XA^XZ", backend=backend) is False


def test_default_backend_selection_does_not_require_pywin32_on_non_windows(monkeypatch):
    monkeypatch.setattr(printing.platform, "system", lambda: "Darwin")
    backend = printing.get_default_backend()
    assert isinstance(backend, printing.PosixFileBackend)


def test_imprimir_lote_sends_one_label_per_unit_of_cantidad():
    backend = FakeBackend()
    productos = [
        _producto(cantidad=3),
        _producto(id='or002', cantidad=2),
    ]
    total = printing.imprimir_lote(productos, ls=0, lt=0, debe_detenerse=lambda: False, backend=backend, delay=0)
    assert total == 5


def test_imprimir_lote_stops_early_when_debe_detenerse_becomes_true():
    backend = FakeBackend()
    productos = [{**_producto(), 'cantidad': 10}]
    calls = {'n': 0}

    def debe_detenerse():
        calls['n'] += 1
        return calls['n'] > 2

    total = printing.imprimir_lote(productos, ls=0, lt=0, debe_detenerse=debe_detenerse, backend=backend, delay=0)
    assert total < 10


def test_imprimir_lote_raises_print_error_on_first_failure_and_stops():
    backend = FakeBackend(fail=True)
    productos = [{**_producto(), 'cantidad': 5}]
    try:
        printing.imprimir_lote(productos, ls=0, lt=0, debe_detenerse=lambda: False, backend=backend, delay=0)
        assert False, "expected PrintError"
    except printing.PrintError:
        pass


def test_imprimir_lote_calls_progress_callback_with_running_total():
    backend = FakeBackend()
    productos = [{**_producto(), 'cantidad': 3}]
    progreso = []
    printing.imprimir_lote(productos, ls=0, lt=0, debe_detenerse=lambda: False,
                            on_progreso=progreso.append, backend=backend, delay=0)
    assert progreso == [1, 2, 3]
