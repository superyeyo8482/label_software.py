"""Product / batch data structures and pure calculations.

Products are plain dicts (matching the original code), with keys:
proveedor, id, material, tipo, actualizado, gramos, precio, cantidad.
"""
from dataclasses import dataclass


def generar_id(productos: list[dict], proveedor: str) -> str:
    count = sum(1 for p in productos if p['proveedor'] == proveedor) + 1
    return f"{proveedor}-{count:03d}"


def hay_ids_vacios(productos: list[dict]) -> bool:
    return any(not p['id'] for p in productos)


@dataclass
class ResumenLote:
    num_productos: int
    total_etiquetas: int
    total_valor: float


def calcular_resumen(productos: list[dict]) -> ResumenLote:
    """Batch summary: distinct product rows, total label count, total value.

    The original had this calculation written out twice (actualizar_resumen
    and pegar_desde_excel), and the two copies disagreed on error handling:
    actualizar_resumen skipped rows with a non-numeric precio, while
    pegar_desde_excel's inline copy had no guard and could raise ValueError
    on the same bad data. Consolidated into one function using the safer
    (skip-on-error) behavior, so both call sites get the same, crash-proof
    result.
    """
    num_productos = len(productos)
    total_etiquetas = sum(p['cantidad'] for p in productos)
    total_valor = 0.0
    for p in productos:
        try:
            total_valor += float(p['precio']) * p['cantidad']
        except (ValueError, TypeError):
            continue
    return ResumenLote(num_productos, total_etiquetas, total_valor)
