import excel_paste


def test_tab_separated_with_header_row():
    raw = "Prov\tID\tMat\tTipo\tAct\tGramos\tPrecio\tCant\nAB\tOR001\tORO 10k\tANILLO\tS\t8\t150000\t2"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=True)
    assert len(r.productos) == 1
    p = r.productos[0]
    assert p['proveedor'] == 'AB'
    assert p['id'] == 'OR001'
    assert p['gramos'] == '008'
    assert p['precio'] == '150000'
    assert p['cantidad'] == 2


def test_comma_separated_without_header_row():
    raw = "AB,OR001,ORO 10k,ANILLO,S,8,150000,2"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    assert len(r.productos) == 1
    assert r.productos[0]['id'] == 'OR001'


def test_missing_id_is_kept_not_skipped_proveedor_only_required():
    """Matches the confirmed-working source's documented design: 'Solo
    Proveedor es obligatorio. ID se genera automáticamente si está vacío.'
    An empty id is preserved as '' here; gui.py's _completar_ids_vacios
    fills it in later via a user prompt.
    """
    raw = "AB\t\tORO 10k\tANILLO\tS\t8\t150000\t1"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    assert len(r.productos) == 1
    assert r.productos[0]['id'] == ''


def test_row_missing_provider_is_skipped_and_reported():
    raw = "AB\tOR001\t\t\t\t\t\t\n\tOR002\t\t\t\t\t\t"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    assert len(r.productos) == 1
    assert len(r.errores) == 1


def test_missing_trailing_columns_default_to_dash_and_zero():
    raw = "AB\tOR001"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    p = r.productos[0]
    assert p['material'] == '-'
    assert p['tipo'] == '-'
    assert p['actualizado'] == '-'
    assert p['gramos'] == '000'
    assert p['precio'] == '0'
    assert p['cantidad'] == 1


def test_price_cleaning_strips_dollar_sign_and_thousands_comma():
    raw = "AB\tOR001\tM\tT\tS\t8\t$1,500\t1"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    assert r.productos[0]['precio'] == '1500'


def test_price_cleaning_strips_non_breaking_space():
    raw = "AB\tOR001\tM\tT\tS\t8\t1\xa0500\t1"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    assert r.productos[0]['precio'] == '1500'


def test_gramos_cleaning_strips_unit_words():
    raw = "AB\tOR001\tM\tT\tS\t15gr\t100\t1"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    assert r.productos[0]['gramos'] == '015'


def test_gramos_cleaning_strips_full_word_and_padding():
    raw = "AB\tOR001\tM\tT\tS\t8 gramos\t100\t1"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    assert r.productos[0]['gramos'] == '008'


def test_cantidad_defaults_to_one_when_blank_or_below_one():
    raw = "AB\tOR001\tM\tT\tS\t1\t100\t0"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    assert r.productos[0]['cantidad'] == 1


def test_empty_clipboard_yields_no_products():
    r = excel_paste.parsear_filas_pegadas("", primera_fila_es_encabezado=False)
    assert r.productos == []
    assert r.errores == []


def test_blank_lines_between_rows_are_skipped():
    raw = "AB\tOR001\tM\tT\tS\t8\t100\t1\n\nCD\tPL002\tM\tT\tS\t8\t100\t1"
    r = excel_paste.parsear_filas_pegadas(raw, primera_fila_es_encabezado=False)
    assert len(r.productos) == 2
