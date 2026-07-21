import csv_io


def test_row_with_provider_is_converted():
    row = {'proveedor': 'AB', 'id': 'OR001', 'material': 'ORO 10k', 'tipo': 'ANILLO',
           'actualizado': 'S', 'gramos': '008', 'precio': '150000', 'cantidad': '2'}
    p = csv_io.fila_csv_a_producto(row)
    assert p['proveedor'] == 'AB'
    assert p['cantidad'] == 2


def test_row_without_provider_is_skipped():
    row = {'proveedor': '', 'id': 'OR001'}
    assert csv_io.fila_csv_a_producto(row) is None


def test_missing_columns_fall_back_to_defaults():
    row = {'proveedor': 'AB', 'id': 'OR001'}
    p = csv_io.fila_csv_a_producto(row)
    assert p['material'] == '-'
    assert p['gramos'] == '000'
    assert p['precio'] == '0'
    assert p['cantidad'] == 1


def test_present_but_blank_cell_stays_blank_not_defaulted():
    """Preserves an asymmetry from the original: a genuinely missing
    column falls back to '-', but a present-and-empty cell does not.
    """
    row = {'proveedor': 'AB', 'id': 'OR001', 'material': ''}
    p = csv_io.fila_csv_a_producto(row)
    assert p['material'] == ''
