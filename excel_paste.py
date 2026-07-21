"""Pure parsing logic extracted from pegar_desde_excel: turns clipboard text
copied from Excel/Google Sheets/LibreOffice into product dicts.

Hybrid, reconciled from the two source files found during this refactor
(see REFACTOR_NOTES.md "Stage 4") - they disagreed on this one function:

- *Requirements* (only `proveedor` is mandatory; a missing `id` is filled
  in later via a one-time "generate IDs automatically?" prompt, not
  silently dropped) come from label_software_v2_backup.py, the
  confirmed-working file whose own header comment states this design
  choice explicitly: "Solo Proveedor es obligatorio. ID se genera
  automáticamente si está vacío." It also matches how agregar() already
  treats every field but proveedor as optional-with-defaults.
- *Field-cleaning robustness* (tab-or-comma separator detection, gramos
  unit-word stripping like "15gr"/"15 g", price cleanup that also
  tolerates non-breaking spaces from Excel-formatted numbers and a
  comma-as-decimal-separator fallback) comes from label_software_2.0.py's
  surviving fragment, which was more defensive about messy pasted content
  even though the file around it was corrupted.

No tkinter/clipboard access here - gui.py owns reading the clipboard and
asking whether the first row is a header.
"""
from dataclasses import dataclass, field


@dataclass
class ResultadoPegado:
    productos: list = field(default_factory=list)
    errores: list = field(default_factory=list)


def _limpiar_gramos(gramos_raw: str) -> str:
    limpio = gramos_raw.lower().replace('gramos', '').replace('gr', '').replace('g', '').replace(',', '').strip()
    try:
        g = int(float(limpio)) if limpio else 0
        return f"{g:03d}"
    except (ValueError, TypeError):
        return "000"


def _limpiar_precio(precio_raw: str) -> str:
    limpio = precio_raw.replace("$", "").replace(" ", "").replace(",", "").replace("\xa0", "").strip()
    if limpio == "":
        limpio = "0"
    try:
        return f"{float(limpio):.0f}"
    except (ValueError, TypeError):
        try:
            limpio = precio_raw.replace("$", "").replace(" ", "").replace(",", ".")
            return f"{float(limpio):.0f}"
        except (ValueError, TypeError):
            return "0"


def _limpiar_cantidad(cantidad_raw: str) -> int:
    try:
        cantidad = int(float(cantidad_raw)) if cantidad_raw else 1
        return cantidad if cantidad >= 1 else 1
    except (ValueError, TypeError):
        return 1


def parsear_filas_pegadas(datos_raw: str, primera_fila_es_encabezado: bool) -> ResultadoPegado:
    resultado = ResultadoPegado()
    lineas = datos_raw.replace('\r', '').strip().split('\n')
    if not lineas or not lineas[0]:
        return resultado

    sep = '\t' if '\t' in lineas[0] else ','
    inicio = 1 if primera_fila_es_encabezado else 0

    for idx_fila, linea in enumerate(lineas[inicio:], start=1):
        if not linea.strip():
            continue

        celdas = linea.split(sep)
        while len(celdas) < 8:
            celdas.append('')

        try:
            proveedor = celdas[0].strip()
            id_prod = celdas[1].strip()
            material = celdas[2].strip() or "-"
            tipo = celdas[3].strip() or "-"
            actualizado = celdas[4].strip() or "-"

            if not proveedor:
                resultado.errores.append(f"Fila {idx_fila}: Proveedor vacío")
                continue

            resultado.productos.append({
                'proveedor': proveedor,
                'id': id_prod,
                'material': material,
                'tipo': tipo,
                'actualizado': actualizado,
                'gramos': _limpiar_gramos(celdas[5].strip()),
                'precio': _limpiar_precio(celdas[6].strip()),
                'cantidad': _limpiar_cantidad(celdas[7].strip()),
            })
        except Exception as exc:
            # Broad on purpose: guards a whole row of unpredictable pasted
            # content; each numeric field already has its own narrower
            # fallback above.
            resultado.errores.append(f"Fila {idx_fila}: {exc}")

    return resultado
