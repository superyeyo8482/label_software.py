"""Pure CSV row conversion, extracted from cargar()/guardar() in the
original. File dialogs and open()/csv.DictReader/csv.DictWriter stay in
gui.py; this module only turns a CSV row dict into a product dict (or
None to signal "skip this row").
"""

CSV_FIELDNAMES = ['proveedor', 'id', 'material', 'tipo', 'actualizado', 'gramos', 'precio', 'cantidad']


def fila_csv_a_producto(row: dict):
    """None if the row has no proveedor (matches the original's silent
    `continue`). Note asymmetry preserved from the source: a genuinely
    missing column falls back to '-'/'0'/'000', but a present-but-blank
    cell stays blank - `row.get('material', '-')` only uses the default
    when the key itself is absent, not when the value is ''.
    """
    proveedor = row.get('proveedor', '').strip()
    if not proveedor:
        return None
    return {
        'proveedor': proveedor,
        'id': row.get('id', '').strip(),
        'material': row.get('material', '-').strip(),
        'tipo': row.get('tipo', '-').strip(),
        'actualizado': row.get('actualizado', '-').strip(),
        'gramos': row.get('gramos', '000').strip(),
        'precio': row.get('precio', '0').strip(),
        'cantidad': int(row.get('cantidad', 1)),
    }
