# label_software - Stage 1: recovered, deduplicated, syntactically valid source.
#
# This file is NOT the final modular refactor. It is the mechanical recovery
# step: one clean copy of every method, with the corruption removed, so the
# "real" codebase can be reviewed before it gets split into modules.
#
# See REFACTOR_NOTES.md in this directory for the full explanation of what
# was found and why each change below was made.
#
# Changes from label_software_2.0.py in this file are ONLY:
#   1. Deduplication: cargar_ejemplos/cambiar_idioma collapsed from 243
#      copies each down to 1 (they were byte-identical apart from corrupted
#      splice fragments injected between repeats - see notes).
#   2. Two indentation bugs fixed so the file actually parses (see
#      "RECOVERED-BUG" comments below). The file as shipped does not run.
#   3. Stubs added for methods that button `command=` callbacks reference
#      but that do not exist anywhere in the 14,000-line source (verified
#      by exhaustive search). Each is marked "RECOVERED-STUB" and shows a
#      visible "not implemented" message instead of guessing real business
#      logic. No business behavior has been invented.
#
# Everything else is byte-for-byte the same logic as the original.

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import win32print
import platform
import json
import os
import threading
import time
import hashlib
import subprocess
import uuid
import sys
from cryptography.fernet import Fernet

# ========== CONFIGURACION ==========
NOMBRE_IMPRESORA = "ZDesigner ZD220-203dpi ZPL"
LS_VAL = -50
LT_VAL = -43

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

# SECURITY WARNING (issue #2 from the requested audit): this is a symmetric
# Fernet key embedded directly in source. Anyone who runs `strings` on the
# .py file or the compiled .exe can extract it and mint their own valid
# activation codes for any machine ID. This "fixes" nothing about the
# licensing threat model - it is flagged, not fixed, per your instruction
# that this decision (asymmetric signing vs. server-side validation) is
# yours to make, not mine.
CLAVE_MAESTRA = b'qJjQH8bZJ4yJjHtPqM7t2Hh5VvJjQqP3xR5sT1uVvWw=='

# ========== TEXTOS (IDIOMAS) ==========
TEXTOS = {
    'es': {
        'titulo': 'Label Software para Zebra ZD220 - Joyería',
        'nuevo_producto': 'Nuevo producto',
        'proveedor': 'Proveedor (2 letras):',
        'id': 'ID Producto:',
        'material': 'Material (ej. ORO 10k):',
        'tipo': 'Tipo (ANILLO/PULSERA):',
        'actualizado': 'Actualizado (S/N):',
        'gramos': 'Gramos:',
        'precio': 'Precio ($):',
        'cantidad': 'Cantidad:',
        'agregar': '➕ Agregar producto',
        'ajustes': 'Ajustes de impresión',
        'ls': 'Desplazamiento horizontal (^LS):',
        'lt': 'Desplazamiento vertical (^LT):',
        'imprimir_prueba': '🖨️ Imprimir prueba',
        'acciones': 'Acciones',
        'eliminar': '🗑️ Eliminar seleccionados',
        'duplicar': '🔄 Duplicar fila marcada',
        'imprimir_sel': '✅ Imprimir seleccionados',
        'imprimir_todos': '🖨️ Imprimir todos',
        'stop': '🛑 STOP',
        'pegar_excel': '📋 Pegar Excel',
        'limpiar': '🧹 Limpiar tabla',
        'guardar_csv': '💾 Guardar CSV',
        'cargar_csv': '📂 Cargar CSV',
        'salir': '❌ Salir',
        'lista_productos': 'Lista de productos',
        'resumen': 'Resumen del lote',
        'productos_distintos': '📦 Productos:',
        'total_etiquetas': '🏷️ Etiquetas:',
        'valor_total': '💰 Total:',
        'listo': '✅ Listo',
        'idioma': 'Idioma / Language:',
        'espanol': 'Español',
        'ingles': 'English',
        'campo_obligatorio': 'Proveedor e ID son obligatorios.',
        'campos_incompletos': 'Campos incompletos',
        'id_vacio': 'IDs vacíos detectados',
        'generar_ids': '¿Generar IDs automáticamente?',
        'sin_impresion': 'Sin impresión',
        'no_hay_impresion': 'No hay ninguna impresión en curso.',
        'confirmar_impresion': 'Confirmar impresión',
        'exito': 'Éxito',
        'error': 'Error',
        'no_activado': 'Este equipo no está activado.',
        'id_equipo': 'ID del equipo (cópielo y envíelo a su proveedor):',
        'copiar_id': '📋 Copiar ID',
        'ingresar_clave': 'Ingrese la clave que recibió:',
        'activar': 'Activar',
        'copiado': 'Copiado',
        'id_copiado': 'ID copiado al portapapeles',
        'clave_invalida': 'Clave inválida',
    },
    'en': {
        'titulo': 'Label Software for Zebra ZD220 - Jewelry',
        'nuevo_producto': 'New product',
        'proveedor': 'Supplier (2 letters):',
        'id': 'Product ID:',
        'material': 'Material (e.g. GOLD 10k):',
        'tipo': 'Type (RING/BRACELET):',
        'actualizado': 'Updated (Y/N):',
        'gramos': 'Grams:',
        'precio': 'Price ($):',
        'cantidad': 'Quantity:',
        'agregar': '➕ Add product',
        'ajustes': 'Print adjustments',
        'ls': 'Horizontal adjustment (^LS):',
        'lt': 'Vertical adjustment (^LT):',
        'imprimir_prueba': '🖨️ Test print',
        'acciones': 'Actions',
        'eliminar': '🗑️ Delete selected',
        'duplicar': '🔄 Duplicate selected row',
        'imprimir_sel': '✅ Print selected',
        'imprimir_todos': '🖨️ Print all',
        'stop': '🛑 STOP',
        'pegar_excel': '📋 Paste from Excel',
        'limpiar': '🧹 Clear table',
        'guardar_csv': '💾 Save CSV',
        'cargar_csv': '📂 Load CSV',
        'salir': '❌ Exit',
        'lista_productos': 'Product list',
        'resumen': 'Batch summary',
        'productos_distintos': '📦 Products:',
        'total_etiquetas': '🏷️ Labels:',
        'valor_total': '💰 Total:',
        'listo': '✅ Ready',
        'idioma': 'Idioma / Language:',
        'espanol': 'Spanish',
        'ingles': 'English',
        'campo_obligatorio': 'Supplier and ID are required.',
        'campos_incompletos': 'Incomplete fields',
        'id_vacio': 'Empty IDs detected',
        'generar_ids': 'Generate IDs automatically?',
        'sin_impresion': 'No print job',
        'no_hay_impresion': 'There is no print job in progress.',
        'confirmar_impresion': 'Confirm print',
        'exito': 'Success',
        'error': 'Error',
        'no_activado': 'This device is not activated.',
        'id_equipo': 'Device ID (copy and send to your provider):',
        'copiar_id': '📋 Copy ID',
        'ingresar_clave': 'Enter the key you received:',
        'activar': 'Activate',
        'copiado': 'Copied',
        'id_copiado': 'ID copied to clipboard',
        'clave_invalida': 'Invalid key',
    }
}

# ========== FUNCIONES ZPL ==========
def generar_zpl(producto, ls, lt):
    precio = producto['precio']
    try:
        precio_float = float(precio)
        precio_str = f"{precio_float:,.0f}".replace(",", ",")
    except:
        precio_str = str(precio)
    return PLANTILLA_ZPL.format(
        ls=ls, lt=lt,
        precio=precio_str,
        proveedor=producto['proveedor'].upper(),
        idProducto=producto['id'].upper(),
        material=producto['material'].upper(),
        tipo=producto['tipo'].upper(),
        actualizado=producto['actualizado'].upper(),
        gramos=producto['gramos'].zfill(3)
    )

def enviar_zpl(zpl):
    if platform.system() == "Windows":
        try:
            h = win32print.OpenPrinter(NOMBRE_IMPRESORA)
            win32print.StartDocPrinter(h, 1, ("etiqueta", None, "RAW"))
            win32print.StartPagePrinter(h)
            win32print.WritePrinter(h, zpl.encode())
            win32print.EndPagePrinter(h)
            win32print.EndDocPrinter(h)
            win32print.ClosePrinter(h)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo imprimir:\n{e}")
            return False
    else:
        try:
            with open("/dev/usb/lp0", "wb") as lp:
                lp.write(zpl.encode())
            return True
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo imprimir (Linux):\n{e}")
            return False

# ========== SISTEMA DE LICENCIAS ==========
def obtener_id_equipo():
    if platform.system() == "Windows":
        nombre = platform.node()
        try:
            uuid_cpu = subprocess.getoutput("wmic csproduct get uuid").split()[1]
        except:
            uuid_cpu = "NODisponible"
        mac = ':'.join(['{:02x}'.format((uuid.getnode() >> i) & 0xff) for i in range(0, 8*6, 8)])
        return hashlib.md5(f"{nombre}{uuid_cpu}{mac}".encode()).hexdigest()
    else:
        nombre = platform.node()
        try:
            with open("/etc/machine-id", "r") as f:
                uuid_cpu = f.read().strip()
        except:
            uuid_cpu = "LINUX"
        mac = ':'.join(['{:02x}'.format((uuid.getnode() >> i) & 0xff) for i in range(0, 8*6, 8)])
        return hashlib.md5(f"{nombre}{uuid_cpu}{mac}".encode()).hexdigest()

def validar_licencia_ingresada(licencia_ingresada, id_equipo):
    try:
        fernet = Fernet(CLAVE_MAESTRA)
        id_descifrado = fernet.decrypt(licencia_ingresada.encode()).decode()
        return id_descifrado == id_equipo
    except:
        return False

def validar_licencia():
    archivo_licencia = "license.key"
    id_actual = obtener_id_equipo()

    if os.path.exists(archivo_licencia):
        with open(archivo_licencia, "r") as f:
            licencia_guardada = f.read().strip()
            if validar_licencia_ingresada(licencia_guardada, id_actual):
                return True
        os.remove(archivo_licencia)

    root_temp = tk.Tk()
    root_temp.withdraw()
    ventana = tk.Toplevel(root_temp)
    ventana.title("Activación - Label Software")
    ventana.geometry("500x300")
    ventana.resizable(False, False)

    tk.Label(ventana, text="Este equipo no está activado.", font=("Arial", 10)).pack(pady=10)
    tk.Label(ventana, text="ID del equipo (cópielo y envíelo a su proveedor):").pack()
    id_var = tk.StringVar(value=id_actual)
    entry_id = tk.Entry(ventana, textvariable=id_var, width=50, state="readonly")
    entry_id.pack(pady=5)

    def copiar_id():
        ventana.clipboard_clear()
        ventana.clipboard_append(id_actual)
        messagebox.showinfo("Copiado", "ID copiado al portapapeles")
    tk.Button(ventana, text="📋 Copiar ID", command=copiar_id).pack(pady=5)

    tk.Label(ventana, text="Ingrese la clave que recibió:").pack()
    clave_var = tk.StringVar()
    tk.Entry(ventana, textvariable=clave_var, width=50).pack(pady=5)

    def verificar():
        entrada = clave_var.get().strip()
        if validar_licencia_ingresada(entrada, id_actual):
            with open(archivo_licencia, "w") as f:
                f.write(entrada)
            messagebox.showinfo("Éxito", "Software activado correctamente")
            ventana.destroy()
            root_temp.destroy()
            return True
        else:
            messagebox.showerror("Error", "Clave inválida")
            return False

    tk.Button(ventana, text="Activar", command=verificar).pack(pady=10)
    ventana.protocol("WM_DELETE_WINDOW", lambda: (root_temp.destroy(), ventana.destroy()))
    root_temp.wait_window(ventana)
    return False

# ========== CLASE PRINCIPAL ==========
class EtiquetadoraApp:
    def __init__(self, root):
        self.root = root
        self.idioma = 'es'
        self.productos = []
        self.check_vars = []
        self.imprimiendo = False
        self.stop_impresion = False

        ls_saved, lt_saved = self.cargar_configuracion()
        self.ls_var = tk.IntVar(value=ls_saved)
        self.lt_var = tk.IntVar(value=lt_saved)

        self.crear_widgets()
        self.actualizar_tabla()
        self.actualizar_resumen()
        self.cargar_ejemplos()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def t(self, key):
        return TEXTOS[self.idioma].get(key, key)

    def cargar_configuracion(self):
        config_file = "etiquetadora_config.json"
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                return config.get('ls', LS_VAL), config.get('lt', LT_VAL)
            except:
                return LS_VAL, LT_VAL
        return LS_VAL, LT_VAL

    def guardar_configuracion(self):
        config = {'ls': self.ls_var.get(), 'lt': self.lt_var.get()}
        try:
            with open("etiquetadora_config.json", 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
        except:
            pass

    def on_closing(self):
        self.stop_impresion = True
        self.guardar_configuracion()
        self.root.destroy()

    def _focus_next(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

    def generar_id(self, proveedor):
        count = sum(1 for p in self.productos if p['proveedor'] == proveedor) + 1
        return f"{proveedor}-{count:03d}"

    # ---------- INTERFAZ GRAFICA ----------
    def crear_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Formulario
        self.frame_form = ttk.LabelFrame(main_frame, text=self.t('nuevo_producto'), padding=10)
        self.frame_form.pack(fill=tk.X, pady=(0,5))
        inner_frame = tk.Frame(self.frame_form)
        inner_frame.pack(fill=tk.X, padx=5, pady=5)
        col_izq = tk.Frame(inner_frame)
        col_izq.grid(row=0, column=0, padx=(0, 20), sticky="n")
        col_der = tk.Frame(inner_frame)
        col_der.grid(row=0, column=1, padx=(0, 0), sticky="n")
        ENTRY_WIDTH = 25
        campos_izq = [
            ("proveedor", self.t('proveedor')),
            ("id", self.t('id')),
            ("material", self.t('material')),
            ("tipo", self.t('tipo'))
        ]
        self.entries = {}
        for key, label in campos_izq:
            row_frame = tk.Frame(col_izq)
            row_frame.pack(fill=tk.X, pady=4)
            lbl = ttk.Label(row_frame, text=label, width=22, anchor="e")
            lbl.pack(side=tk.LEFT, padx=(0, 8))
            entry = ttk.Entry(row_frame, width=ENTRY_WIDTH)
            entry.pack(side=tk.LEFT)
            self.entries[key] = entry
            entry.bind('<Return>', self._focus_next)
        campos_der = [
            ("actualizado", self.t('actualizado')),
            ("gramos", self.t('gramos')),
            ("precio", self.t('precio')),
            ("cantidad", self.t('cantidad'))
        ]
        for key, label in campos_der:
            row_frame = tk.Frame(col_der)
            row_frame.pack(fill=tk.X, pady=4)
            lbl = ttk.Label(row_frame, text=label, width=22, anchor="e")
            lbl.pack(side=tk.LEFT, padx=(0, 8))
            entry = ttk.Entry(row_frame, width=ENTRY_WIDTH)
            entry.pack(side=tk.LEFT)
            self.entries[key] = entry
            if key != "cantidad":
                entry.bind('<Return>', self._focus_next)
            else:
                entry.bind('<Return>', lambda event: self.agregar())
        btn_frame = tk.Frame(self.frame_form)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        self.btn_agregar = ttk.Button(btn_frame, text=self.t('agregar'), command=self.agregar, width=35)
        self.btn_agregar.pack()

        # Ajustes
        self.frame_offset = ttk.LabelFrame(main_frame, text=self.t('ajustes'), padding=8)
        self.frame_offset.pack(fill=tk.X, pady=5)
        ajustes_frame = tk.Frame(self.frame_offset)
        ajustes_frame.pack()
        ttk.Label(ajustes_frame, text=self.t('ls')).grid(row=0, column=0, padx=5)
        ls_spin = ttk.Spinbox(ajustes_frame, from_=-200, to=200, increment=1, textvariable=self.ls_var, width=8)
        ls_spin.grid(row=0, column=1, padx=5)
        ttk.Label(ajustes_frame, text=self.t('lt')).grid(row=0, column=2, padx=5)
        lt_spin = ttk.Spinbox(ajustes_frame, from_=-200, to=200, increment=1, textvariable=self.lt_var, width=8)
        lt_spin.grid(row=0, column=3, padx=5)
        self.btn_test = ttk.Button(ajustes_frame, text=self.t('imprimir_prueba'), command=self.probar_offset)
        self.btn_test.grid(row=0, column=4, padx=10)

        # Idioma
        frame_idioma = tk.Frame(ajustes_frame)
        frame_idioma.grid(row=1, column=0, columnspan=5, pady=5)
        ttk.Label(frame_idioma, text=self.t('idioma')).pack(side=tk.LEFT, padx=5)
        self.btn_idioma = ttk.Button(frame_idioma, text="English", command=self.cambiar_idioma)
        self.btn_idioma.pack(side=tk.LEFT, padx=5)

        # Botones
        self.frame_botones = ttk.LabelFrame(main_frame, text=self.t('acciones'), padding=8)
        self.frame_botones.pack(fill=tk.X, pady=5)
        btn_eliminar = ttk.Button(self.frame_botones, text=self.t('eliminar'), command=self.eliminar, width=25)
        btn_eliminar.grid(row=0, column=0, padx=5, pady=2, sticky="ew")
        btn_duplicar = ttk.Button(self.frame_botones, text=self.t('duplicar'), command=self.duplicar, width=25)
        btn_duplicar.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.btn_imprimir_sel = ttk.Button(self.frame_botones, text=self.t('imprimir_sel'), command=self.imprimir_chequeados, width=25)
        self.btn_imprimir_sel.grid(row=0, column=2, padx=5, pady=2, sticky="ew")
        self.btn_imprimir_todos = ttk.Button(self.frame_botones, text=self.t('imprimir_todos'), command=self.imprimir_todos, width=25)
        self.btn_imprimir_todos.grid(row=1, column=0, padx=5, pady=2, sticky="ew")
        self.btn_stop = tk.Button(self.frame_botones, text=self.t('stop'), command=self.detener_impresion,
                                  bg="#f44336", fg="white", font=("Arial", 10, "bold"),
                                  width=25, height=1, relief="raised")
        self.btn_stop.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        self.btn_pegar = ttk.Button(self.frame_botones, text=self.t('pegar_excel'), command=self.pegar_desde_excel, width=25)
        self.btn_pegar.grid(row=1, column=2, padx=5, pady=2, sticky="ew")
        self.btn_limpiar = ttk.Button(self.frame_botones, text=self.t('limpiar'), command=self.limpiar, width=25)
        self.btn_limpiar.grid(row=2, column=0, padx=5, pady=2, sticky="ew")
        self.btn_cargar = ttk.Button(self.frame_botones, text=self.t('cargar_csv'), command=self.cargar, width=25)
        self.btn_cargar.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        self.btn_guardar = ttk.Button(self.frame_botones, text=self.t('guardar_csv'), command=self.guardar, width=25)
        self.btn_guardar.grid(row=2, column=2, padx=5, pady=2, sticky="ew")
        btn_salir = ttk.Button(self.frame_botones, text=self.t('salir'), command=self.root.destroy, width=25)
        btn_salir.grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        for col in range(3):
            self.frame_botones.columnconfigure(col, weight=1)

        # Tabla
        self.frame_tabla = ttk.LabelFrame(main_frame, text=self.t('lista_productos'), padding=8)
        self.frame_tabla.pack(fill=tk.BOTH, expand=True, pady=5)
        tree_frame = ttk.Frame(self.frame_tabla)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        columnas = ("No.", "✓", "Proveedor", "ID", "Material", "Tipo", "Act", "Gramos", "Precio", "Cantidad")
        self.tree = ttk.Treeview(tree_frame, columns=columnas, show="headings",
                                  yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        self.tree.heading("No.", text="#")
        self.tree.column("No.", width=40, anchor='center', stretch=False)
        self.tree.heading("✓", text="✓")
        self.tree.column("✓", width=40, anchor='center', stretch=False)
        col_widths = {"Proveedor": 80, "ID": 100, "Material": 120, "Tipo": 120,
                      "Act": 50, "Gramos": 70, "Precio": 100, "Cantidad": 80}
        for col in columnas[2:]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 100))
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.tree.bind('<ButtonRelease-1>', self.on_click_check)
        self.tree.bind('<Double-1>', self.editar_celda)

        # Resumen del lote (donde se muestra el resumen)
        self.frame_resumen = ttk.LabelFrame(main_frame, text=self.t('resumen'), padding=8)
        self.frame_resumen.pack(fill=tk.X, pady=5)
        self.label_productos = ttk.Label(self.frame_resumen, text=f"{self.t('productos_distintos')} 0",
                                         font=("Arial", 11, "bold"), foreground="#2c5f8a")
        self.label_productos.pack(side=tk.LEFT, padx=15, pady=3)
        self.label_etiquetas = ttk.Label(self.frame_resumen, text=f"{self.t('total_etiquetas')} 0",
                                         font=("Arial", 11, "bold"), foreground="#2c5f8a")
        self.label_etiquetas.pack(side=tk.LEFT, padx=15, pady=3)
        self.label_valor = ttk.Label(self.frame_resumen, text=f"{self.t('valor_total')} $0",
                                     font=("Arial", 11, "bold"), foreground="#2c5f8a")
        self.label_valor.pack(side=tk.LEFT, padx=15, pady=3)
        self.label_estado = ttk.Label(self.frame_resumen, text=self.t('listo'), foreground="green", font=("Arial", 10))
        self.label_estado.pack(side=tk.RIGHT, padx=15, pady=3)

    # ---------- ACTUALIZAR RESUMEN ----------
    # RECOVERED-BUG: in the original file this was nested 4 extra spaces
    # inside crear_widgets (a local closure, never bound to self, with a
    # body indented no deeper than its own `def` line - a plain
    # IndentationError). Dedented here to a normal method so `self.actualizar_resumen()`
    # (called from __init__ and actualizar_tabla) actually resolves.
    def actualizar_resumen(self):
        num_productos = len(self.productos)
        total_etiquetas = sum(p['cantidad'] for p in self.productos)
        total_precio = 0.0
        for p in self.productos:
            try:
                total_precio += float(p['precio']) * p['cantidad']
            except:
                pass
        self.label_productos.config(text=f"{self.t('productos_distintos')} {num_productos}")
        self.label_etiquetas.config(text=f"{self.t('total_etiquetas')} {total_etiquetas}")
        self.label_valor.config(text=f"{self.t('valor_total')} ${total_precio:,.0f}")

    def actualizar_tabla(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, prod in enumerate(self.productos, start=1):
            check = "✓" if self.check_vars[i-1].get() else ""
            self.tree.insert("", tk.END, values=(
                i,
                check,
                prod['proveedor'], prod['id'], prod['material'],
                prod['tipo'], prod['actualizado'], prod['gramos'],
                prod['precio'], prod['cantidad']
            ))
        self.actualizar_resumen()

    # ---------- PEGAR DESDE EXCEL (CON COLUMNAS VACIAS) ----------
    # RECOVERED-BUG: `def pegar_desde_excel(self):` was indented 20 spaces
    # in the original (16 more than a class method should be). Fixed to a
    # normal 4-space method def; body is otherwise untouched.
    def pegar_desde_excel(self):
        try:
            datos_raw = self.root.clipboard_get()
            if not datos_raw:
                messagebox.showwarning("Portapapeles vacío", "No hay datos para pegar.")
                return

            lineas = datos_raw.replace('\r', '').strip().split('\n')
            if not lineas:
                return

            sep = '\t' if '\t' in lineas[0] else ','
            respuesta = messagebox.askyesno("Encabezados", "¿La primera fila contiene los títulos?")
            inicio = 1 if respuesta else 0

            nuevos = []
            errores = []
            filas_ignoradas = 0

            for idx_fila, linea in enumerate(lineas[inicio:], start=1):
                if not linea.strip():
                    continue

                celdas = linea.split(sep)
                while len(celdas) < 8:
                    celdas.append('')

                proveedor = celdas[0].strip()
                idProd = celdas[1].strip()
                material = celdas[2].strip() if celdas[2].strip() else "-"
                tipo = celdas[3].strip() if celdas[3].strip() else "-"
                actualizado = celdas[4].strip() if celdas[4].strip() else "-"
                gramos_raw = celdas[5].strip()
                precio_raw = celdas[6].strip()
                cantidad_raw = celdas[7].strip()

                if not proveedor or not idProd:
                    filas_ignoradas += 1
                    errores.append(f"Fila {idx_fila}: Proveedor o ID vacío")
                    continue

                gramos_clean = gramos_raw.lower().replace('gramos', '').replace('gr', '').replace('g', '').replace(',', '').strip()
                try:
                    g = int(float(gramos_clean)) if gramos_clean else 0
                    gramos = str(g)
                except ValueError:
                    gramos = "0"

                # Limpiar precio (usamos comillas dobles para evitar errores de sintaxis)
                precio_clean = precio_raw.replace("$", "").replace(" ", "").replace(",", "").replace("\xa0", "").strip()
                if precio_clean == '':
                    precio_clean = '0'
                try:
                    p = float(precio_clean)
                    precio = f"{p:.0f}"
                except ValueError:
                    try:
                        precio_clean = precio_raw.replace("$", "").replace(" ", "").replace(",", ".")
                        p = float(precio_clean)
                        precio = f"{p:.0f}"
                    except ValueError:
                        precio = "0"

                try:
                    cantidad = int(float(cantidad_raw)) if cantidad_raw else 1
                    if cantidad < 1:
                        cantidad = 1
                except ValueError:
                    cantidad = 1

                nuevos.append({
                    'proveedor': proveedor,
                    'id': idProd,
                    'material': material,
                    'tipo': tipo,
                    'actualizado': actualizado,
                    'gramos': gramos,
                    'precio': precio,
                    'cantidad': cantidad
                })

            if nuevos:
                self.productos.extend(nuevos)
                self.check_vars.extend([tk.BooleanVar(value=False) for _ in nuevos])

            self.actualizar_tabla()

            total_productos = len(self.productos)
            total_etiquetas = sum(p['cantidad'] for p in self.productos)
            total_valor = sum(float(p['precio']) * p['cantidad'] for p in self.productos)

            msg = "📊 RESUMEN DEL LOTE\n"
            msg += f"📦 Productos en lista: {total_productos}\n"
            msg += f"🏷️ Etiquetas a imprimir: {total_etiquetas}\n"
            msg += f"💰 Valor total: ${total_valor:,.0f}\n"

            if nuevos:
                msg += f"\n✅ Se agregaron {len(nuevos)} productos nuevos."
            else:
                msg += "\n⚠️ No se agregaron productos nuevos."

            if filas_ignoradas > 0:
                msg += f"\n⚠️ {filas_ignoradas} filas ignoradas (falta proveedor o ID)."
            if errores:
                msg += f"\n⚠️ {len(errores)} errores adicionales (revisa el log)."

            messagebox.showinfo("Pegado completado", msg)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo pegar:\n{str(e)}")

    def cargar_ejemplos(self):
        ejemplos = [
            {'proveedor':'AB','id':'OR001','material':'ORO 10k','tipo':'ANILLO','actualizado':'S','gramos':'008','precio':'150000','cantidad':2},
            {'proveedor':'CD','id':'PL002','material':'PLATA 925','tipo':'PULSERA','actualizado':'N','gramos':'015','precio':'85000','cantidad':1}
        ]
        for prod in ejemplos:
            self.productos.append(prod)
            self.check_vars.append(tk.BooleanVar(value=False))
        self.actualizar_tabla()

    # ---------- CAMBIAR IDIOMA ----------
    # Confirmed dead no-op in the original (issue #1 from the audit): the
    # button is wired up and visible, but does nothing.
    def cambiar_idioma(self):
        # (Mantén la función que ya tienes para cambiar idioma)
        pass

    # ================================================================
    # RECOVERED-STUB METHODS
    #
    # Everything below is referenced by `command=self.<name>` in
    # crear_widgets() above, but has NO implementation anywhere in the
    # 14,000-line original file (confirmed by exhaustive search - these
    # names appear nowhere except inside TEXTOS strings and the button
    # wiring itself). As shipped, the app throws AttributeError the
    # instant crear_widgets() runs and never reaches mainloop().
    #
    # No business logic has been invented for these. Each stub just logs
    # and shows a visible "not implemented" message so the window can open
    # and the rest of the recovered code can be inspected/run. Real
    # behavior needs to be written by whoever has the actual logic (or it
    # needs to be re-authored from scratch against the CSV/table/printing
    # primitives that DO exist: generar_zpl, enviar_zpl, actualizar_tabla,
    # self.productos, self.check_vars).
    # ================================================================

    def _no_implementado(self, nombre):
        messagebox.showwarning(
            "No implementado",
            f"'{nombre}' no tiene una implementación en el archivo original "
            f"(label_software_2.0.py) - ver REFACTOR_NOTES.md."
        )

    def agregar(self):
        self._no_implementado('agregar')

    def probar_offset(self):
        self._no_implementado('probar_offset')

    def eliminar(self):
        self._no_implementado('eliminar')

    def duplicar(self):
        self._no_implementado('duplicar')

    def imprimir_chequeados(self):
        self._no_implementado('imprimir_chequeados')

    def imprimir_todos(self):
        self._no_implementado('imprimir_todos')

    def detener_impresion(self):
        self._no_implementado('detener_impresion')

    def limpiar(self):
        self._no_implementado('limpiar')

    def cargar(self):
        self._no_implementado('cargar')

    def guardar(self):
        self._no_implementado('guardar')

    def on_click_check(self, event):
        self._no_implementado('on_click_check')

    def editar_celda(self, event):
        self._no_implementado('editar_celda')


# ========== MAIN ==========
if __name__ == "__main__":
    if not validar_licencia():
        sys.exit()

    root = tk.Tk()
    root.geometry("1300x800")
    app = EtiquetadoraApp(root)
    root.mainloop()
