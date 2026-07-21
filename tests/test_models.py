import models


def test_generar_id_increments_per_proveedor_independently():
    productos = [{'proveedor': 'AB'}, {'proveedor': 'AB'}, {'proveedor': 'CD'}]
    assert models.generar_id(productos, 'AB') == 'AB-003'
    assert models.generar_id(productos, 'CD') == 'CD-002'
    assert models.generar_id(productos, 'EF') == 'EF-001'


def test_calcular_resumen_sums_labels_and_value():
    productos = [
        {'precio': '100000', 'cantidad': 2},
        {'precio': '50000', 'cantidad': 1},
    ]
    resumen = models.calcular_resumen(productos)
    assert resumen.num_productos == 2
    assert resumen.total_etiquetas == 3
    assert resumen.total_valor == 250000.0


def test_calcular_resumen_skips_rows_with_non_numeric_price_instead_of_raising():
    productos = [{'precio': 'garbage', 'cantidad': 1}, {'precio': '10', 'cantidad': 1}]
    resumen = models.calcular_resumen(productos)
    assert resumen.num_productos == 2
    assert resumen.total_etiquetas == 2
    assert resumen.total_valor == 10.0


def test_calcular_resumen_empty_batch():
    resumen = models.calcular_resumen([])
    assert resumen.num_productos == 0
    assert resumen.total_etiquetas == 0
    assert resumen.total_valor == 0.0
