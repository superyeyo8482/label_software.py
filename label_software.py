# label_software.py
# Versión 2.0 - Corregida y funcional
# Solo Proveedor es obligatorio. ID se genera automáticamente si está vacío.

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

# ========== CONFIGURACIÓN ==========
NOMBRE_IMPRESORA = "ZDesigner ZD220-203dpi ZPL"
LS_VAL = 50
LT_VAL = 53

# Plantilla ZPL definitiva
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

# ========== CLAVE MAESTRA PARA LICENCIAS ==========
CLAVE_MAESTRA = b'qJjQH8bZJ4yJjHtPqM7t2Hh5VvJjQqP3xR5sT1uVvWw=='

# ========== TEXTOS PARA IDIOMAS ==========
TEXTOS = {
    'es': {
        'titulo': 'Label Software para Zebra ZD220 - Joyería',
        'nuevo_producto': 'Nuevo producto',
        'proveedor': 'Proveedor (2 letras):',
        'id': 'ID Producto:',
        'material': 'Material (ej. ORO 10k):',
        'tipo': 'Tipo (ANILLO/PULSERA):',
        'actualizado': 'Actualizado (S/N):',
        'gramos': 'Gramos (3 dígitos):',
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
        'productos_distintos': '📦 Productos distintos:',
        'total_etiquetas': '🏷️ Total de etiquetas:',
        'valor_total': '💰 Valor total:',
        'listo': '✅ Listo',
        'idioma': 'Idioma / Language:',
        'espanol': 'Español',
        'ingles': 'English',
        'campo_obligatorio': 'Proveedor es obligatorio.',
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
        'gramos': 'Grams (3 digits):',
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
        'productos_distintos': '📦 Distinct products:',
        'total_etiquetas': '🏷️ Total labels:',
        'valor_total': '💰 Total value:',
        'listo': '✅ Ready',
        'idioma': 'Idioma / Language:',
        'espanol': 'Spanish',
        'ingles': 'English',
        'campo_obligatorio': 'Supplier is required.',
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
    
    # Ventana de activación
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

    def _actualizar_etiquetas_recursivo(self, widget, mapeo, t):
        if isinstance(widget, ttk.Label):
            texto = widget.cget('text')
            if texto in mapeo:
                widget.config(text=t[mapeo[texto]])
        for child in widget.winfo_children():
            self._actualizar_etiquetas_recursivo(child, mapeo, t)

    def actualizar_etiquetas_formulario(self):
        mapeo_etiquetas = {
            "Proveedor (2 letras):": "proveedor",
            "ID Producto:": "id",
            "Material (ej. ORO 10k):": "material",
            "Tipo (ANILLO/PULSERA):": "tipo",
            "Actualizado (S/N):": "actualizado",
            "Gramos (3 dígitos):": "gramos",
            "Precio ($):": "precio",
            "Cantidad:": "cantidad",
        }
        t = TEXTOS[self.idioma]
        for child in self.frame_form.winfo_children():
            self._actualizar_etiquetas_recursivo(child, mapeo_etiquetas, t)

    def cambiar_idioma(self):
        self.idioma = 'en' if self.idioma == 'es' else 'es'
        t = TEXTOS[self.idioma]
        
        self.frame_form.config(text=t['nuevo_producto'])
        self.frame_offset.config(text=t['ajustes'])
        self.frame_botones.config(text=t['acciones'])
        self.frame_tabla.config(text=t['lista_productos'])
        self.frame_resumen.config(text=t['resumen'])
        
        self.btn_agregar.config(text=t['agregar'])
        self.btn_test.config(text=t['imprimir_prueba'])
        self.btn_stop.config(text=t['stop'])
        self.btn_pegar.config(text=t['pegar_excel'])
        self.btn_limpiar.config(text=t['limpiar'])
        self.btn_cargar.config(text=t['cargar_csv'])
        self.btn_guardar.config(text=t['guardar_csv'])
        self.btn_imprimir_sel.config(text=t['imprimir_sel'])
        self.btn_imprimir_todos.config(text=t['imprimir_todos'])
        self.btn_idioma.config(text=t['ingles'] if self.idioma == 'es' else t['espanol'])
        
        self.actualizar_etiquetas_formulario()
        
        for child in self.frame_offset.winfo_children():
            if isinstance(child, tk.Frame):
                for subchild in child.winfo_children():
                    if isinstance(subchild, ttk.Label):
                        texto_actual = subchild.cget('text')
                        if 'horizontal' in texto_actual.lower() or 'desplazamiento horizontal' in texto_actual.lower():
                            subchild.config(text=t['ls'])
                        elif 'vertical' in texto_actual.lower() or 'desplazamiento vertical' in texto_actual.lower():
                            subchild.config(text=t['lt'])
                        elif 'idioma' in texto_actual.lower() or 'language' in texto_actual.lower():
                            subchild.config(text=t['idioma'])
        
        for child in self.frame_botones.winfo_children():
            if isinstance(child, ttk.Button):
                texto_actual = child.cget('text')
                if 'Eliminar' in texto_actual or 'Delete' in texto_actual:
                    child.config(text=t['eliminar'])
                elif 'Duplicar' in texto_actual or 'Duplicate' in texto_actual:
                    child.config(text=t['duplicar'])
                elif 'Salir' in texto_actual or 'Exit' in texto_actual:
                    child.config(text=t['salir'])
        
        self.actualizar_resumen()
        self.label_estado.config(text=t['listo'])
        self.root.title(t['titulo'])
        
        if self.idioma == 'en':
            encabezados = ["#", "✓", "Supplier", "ID", "Material", "Type", "Updated", "Grams", "Price", "Quantity"]
        else:
            encabezados = ["No.", "✓", "Proveedor", "ID", "Material", "Tipo", "Act", "Gramos", "Precio", "Cantidad"]
        
        columnas = ("No.", "✓", "Proveedor", "ID", "Material", "Tipo", "Act", "Gramos", "Precio", "Cantidad")
        for i, col in enumerate(columnas):
            self.tree.heading(col, text=encabezados[i])

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

        # Resumen
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

    def on_click_check(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#2":
                item = self.tree.identify_row(event.y)
                if item:
                    idx = self.tree.index(item)
                    if idx < len(self.check_vars):
                        self.check_vars[idx].set(not self.check_vars[idx].get())
                        self.actualizar_tabla()

    def agregar(self):
        proveedor = self.entries['proveedor'].get().strip()
        idProd = self.entries['id'].get().strip()
        material = self.entries['material'].get().strip()
        tipo = self.entries['tipo'].get().strip()
        actualizado = self.entries['actualizado'].get().strip()
        gramos = self.entries['gramos'].get().strip()
        precio = self.entries['precio'].get().strip()
        cantidad = self.entries['cantidad'].get().strip()
        
        if not proveedor:
            messagebox.showwarning(self.t('campos_incompletos'), self.t('campo_obligatorio'))
            return
        
        if not idProd:
            idProd = self.generar_id(proveedor)
            self.entries['id'].insert(0, idProd)
        
        try:
            cant = int(cantidad) if cantidad else 1
        except ValueError:
            messagebox.showwarning("Cantidad inválida", "Debe ser un número entero.")
            return
        
        nuevo = {
            'proveedor': proveedor,
            'id': idProd,
            'material': material if material else "-",
            'tipo': tipo if tipo else "-",
            'actualizado': actualizado if actualizado else "-",
            'gramos': gramos if gramos else "000",
            'precio': precio if precio else "0",
            'cantidad': cant
        }
        self.productos.append(nuevo)
        self.check_vars.append(tk.BooleanVar(value=False))
        self.actualizar_tabla()
        for key in self.entries:
            self.entries[key].delete(0, tk.END)
        self.entries['cantidad'].insert(0, "1")
        self.entries['proveedor'].focus()

    def eliminar(self):
        indices = [i for i, var in enumerate(self.check_vars) if var.get()]
        if not indices:
            messagebox.showinfo("Sin selección", "Marca las filas a eliminar.")
            return
        if messagebox.askyesno("Confirmar", f"¿Eliminar {len(indices)} producto(s)?"):
            for i in sorted(indices, reverse=True):
                del self.productos[i]
                del self.check_vars[i]
            self.actualizar_tabla()

    def duplicar(self):
        indices = [i for i, var in enumerate(self.check_vars) if var.get()]
        if not indices:
            messagebox.showinfo("Selecciona", "Marca la fila que quieras duplicar.")
            return
        idx = indices[0]
        copia = self.productos[idx].copy()
        copia['cantidad'] = 1
        self.productos.append(copia)
        self.check_vars.append(tk.BooleanVar(value=False))
        self.actualizar_tabla()

    def limpiar(self):
        if messagebox.askyesno("Limpiar", "¿Borrar todos los productos de la lista?"):
            self.productos.clear()
            self.check_vars.clear()
            self.actualizar_tabla()

    def pegar_desde_excel(self):
        try:
            datos_raw = self.root.clipboard_get()
            if not datos_raw:
                messagebox.showwarning("Portapapeles vacío", "No hay datos para pegar.")
                return
            lineas = datos_raw.strip().split('\n')
            if not lineas:
                return
            respuesta_encabezados = messagebox.askyesno("Encabezados", "¿La primera fila contiene los títulos?")
            inicio = 1 if respuesta_encabezados else 0
            nuevos = []
            errores = []
            hay_ids_vacios = False
            
            for idx_fila, linea in enumerate(lineas[inicio:], start=1):
                if not linea.strip():
                    continue
                celdas = linea.split('\t')
                if len(celdas) < 7:
                    errores.append(f"Fila {idx_fila}: Solo {len(celdas)} columnas")
                    continue
                try:
                    proveedor = celdas[0].strip()
                    idProd = celdas[1].strip()
                    material = celdas[2].strip() if len(celdas) > 2 and celdas[2].strip() else "-"
                    tipo = celdas[3].strip() if len(celdas) > 3 and celdas[3].strip() else "-"
                    actualizado = celdas[4].strip() if len(celdas) > 4 and celdas[4].strip() else "-"
                    gramos_raw = celdas[5].strip() if len(celdas) > 5 else "0"
                    precio_raw = celdas[6].strip() if len(celdas) > 6 else "0"
                    cantidad_raw = celdas[7].strip() if len(celdas) > 7 else "1"
                    if not proveedor:
                        errores.append(f"Fila {idx_fila}: Proveedor vacío")
                        continue
                    if not idProd:
                        hay_ids_vacios = True
                    try:
                        g = int(float(gramos_raw)) if gramos_raw else 0
                        gramos = f"{g:03d}"
                    except:
                        gramos = "000"
                    precio_clean = precio_raw.replace('$', '').replace(',', '').strip()
                    try:
                        p = float(precio_clean) if precio_clean else 0
                        precio = f"{p:.0f}"
                    except:
                        precio = "0"
                    try:
                        cantidad = int(float(cantidad_raw))
                        if cantidad < 1:
                            cantidad = 1
                    except:
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
                except Exception as e:
                    errores.append(f"Fila {idx_fila}: {str(e)}")
            
            if hay_ids_vacios and nuevos:
                resposta = messagebox.askyesno(
                    self.t('id_vacio'),
                    f"{self.t('id_vacio')}\n\n{self.t('generar_ids')}"
                )
                if resposta:
                    for produto in nuevos:
                        if not produto['id']:
                            produto['id'] = self.generar_id(produto['proveedor'])
            
            if nuevos:
                self.productos.extend(nuevos)
                self.check_vars.extend([tk.BooleanVar(value=False) for _ in nuevos])
                self.actualizar_tabla()
                msg = f"✅ Se agregaron {len(nuevos)} productos."
                if errores:
                    msg += f"\n⚠️ {len(errores)} errores"
                messagebox.showinfo("Éxito", msg)
            else:
                messagebox.showwarning("Error", "No se pudo procesar ningún dato")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo pegar:\n{str(e)}")

    def detener_impresion(self):
        if self.imprimiendo:
            self.stop_impresion = True
            self.label_estado.config(text="🛑 Deteniendo...", foreground="red")
        else:
            messagebox.showinfo("Sin impresión", "No hay ninguna impresión en curso.")

    def _imprimir(self, lista):
        if not lista:
            return
        total_etiquetas = sum(p['cantidad'] for p in lista)
        respuesta = messagebox.askyesno("Confirmar impresión", 
            f"📊 Productos: {len(lista)}\n"
            f"🏷️ Etiquetas: {total_etiquetas}\n"
            f"💰 Valor total: ${sum(float(p['precio']) * p['cantidad'] for p in lista):,.0f}\n\n"
            f"¿Imprimir?"
        )
        if not respuesta:
            return
        self.stop_impresion = False
        self.imprimiendo = True
        self.label_estado.config(text="🖨️ Imprimiendo...", foreground="orange")
        hilo = threading.Thread(target=self._impresion_hilo, args=(lista, total_etiquetas))
        hilo.daemon = True
        hilo.start()

    def _impresion_hilo(self, lista, total_etiquetas):
        ls = self.ls_var.get()
        lt = self.lt_var.get()
        total = 0
        try:
            for prod in lista:
                if self.stop_impresion:
                    break
                for _ in range(prod['cantidad']):
                    if self.stop_impresion:
                        break
                    zpl = generar_zpl(prod, ls, lt)
                    if not enviar_zpl(zpl):
                        self.root.after(0, lambda: self.label_estado.config(text="❌ Error", foreground="red"))
                        self.imprimiendo = False
                        return
                    total += 1
                    self.root.after(0, lambda t=total: self.label_estado.config(text=f"🖨️ {t}/{total_etiquetas}", foreground="orange"))
                    time.sleep(0.1)
            if not self.stop_impresion:
                enviar_zpl("~FF")
                self.root.after(0, lambda: self.label_estado.config(text=self.t('listo'), foreground="green"))
                self.root.after(0, lambda: messagebox.showinfo("Éxito", f"Se imprimieron {total} etiquetas."))
            else:
                self.root.after(0, lambda: self.label_estado.config(text="⏹️ Detenido", foreground="orange"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error en impresión: {e}"))
        finally:
            self.imprimiendo = False
            self.stop_impresion = False

    def imprimir_chequeados(self):
        if self.imprimiendo:
            messagebox.showwarning("Impresión en curso", "Espera a que termine o usa STOP.")
            return
        indices = [i for i, var in enumerate(self.check_vars) if var.get()]
        if not indices:
            messagebox.showinfo("Sin selección", "Marca productos con ✓")
            return
        self._imprimir([self.productos[i] for i in indices])

    def imprimir_todos(self):
        if self.imprimiendo:
            messagebox.showwarning("Impresión en curso", "Espera a que termine o usa STOP.")
            return
        if not self.productos:
            messagebox.showinfo("Sin datos", "Lista vacía")
            return
        self._imprimir(self.productos)

    def probar_offset(self):
        self.guardar_configuracion()
        ejemplo = {
            'proveedor': 'AB',
            'id': 'TEST',
            'material': 'ORO 10k',
            'tipo': 'ANILLO',
            'actualizado': 'S',
            'gramos': '008',
            'precio': '1000',
            'cantidad': 1
        }
        zpl = generar_zpl(ejemplo, self.ls_var.get(), self.lt_var.get())
        if enviar_zpl(zpl):
            messagebox.showinfo("Prueba", "Etiqueta de prueba enviada. Verifica la posición.")

    def guardar(self):
        if not self.productos:
            messagebox.showwarning("Lista vacía", "No hay productos que guardar.")
            return
        archivo = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if archivo:
            with open(archivo, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['proveedor','id','material','tipo','actualizado','gramos','precio','cantidad'])
                writer.writeheader()
                writer.writerows(self.productos)
            messagebox.showinfo("Guardado", f"Lista guardada en {archivo}")

    def cargar(self):
        archivo = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if archivo:
            try:
                with open(archivo, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    novos = []
                    hay_ids_vacios = False
                    for row in reader:
                        proveedor = row.get('proveedor', '').strip()
                        idProd = row.get('id', '').strip()
                        if not proveedor:
                            continue
                        if not idProd:
                            hay_ids_vacios = True
                        novos.append({
                            'proveedor': proveedor,
                            'id': idProd,
                            'material': row.get('material', '-').strip(),
                            'tipo': row.get('tipo', '-').strip(),
                            'actualizado': row.get('actualizado', '-').strip(),
                            'gramos': row.get('gramos', '000').strip(),
                            'precio': row.get('precio', '0').strip(),
                            'cantidad': int(row.get('cantidad', 1))
                        })
                    if hay_ids_vacios and novos:
                        resposta = messagebox.askyesno(
                            self.t('id_vacio'),
                            f"{self.t('id_vacio')}\n\n{self.t('generar_ids')}"
                        )
                        if resposta:
                            for produto in novos:
                                if not produto['id']:
                                    produto['id'] = self.generar_id(produto['proveedor'])
                    self.productos.extend(novos)
                    self.check_vars.extend([tk.BooleanVar(value=False) for _ in novos])
                    self.actualizar_tabla()
                messagebox.showinfo("Cargado", f"Se agregaron {len(novos)} productos.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar:\n{e}")

    def editar_celda(self, event):
        if hasattr(self, 'edit_entry') and self.edit_entry:
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        column = self.tree.identify_column(event.x)
        idx = self.tree.index(item)
        if idx >= len(self.productos):
            return
        col_index = int(column[1:]) - 1
        if col_index <= 1:
            return
        self._crear_entry_edicion(item, idx, col_index)

    def _crear_entry_edicion(self, item, idx, col_index):
        valores = list(self.tree.item(item, 'values'))
        valor_actual = valores[col_index]
        x, y, w, h = self.tree.bbox(item, f"#{col_index+1}")
        entry = tk.Entry(self.tree, font=('TkDefaultFont', 10))
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, str(valor_actual))
        entry.select_range(0, tk.END)
        entry.focus()
        self.edit_entry = entry
        self.edit_item = (item, idx, col_index)

        def guardar(event=None):
            new_val = entry.get().strip()
            self._guardar_edicion(idx, col_index, new_val)
            entry.destroy()
            self.edit_entry = None
            self.edit_item = None
        entry.bind('<Return>', lambda e: guardar())
        entry.bind('<FocusOut>', lambda e: guardar())

    def _guardar_edicion(self, idx, col_index, new_val):
        if col_index == 2:
            self.productos[idx]['proveedor'] = new_val
        elif col_index == 3:
            self.productos[idx]['id'] = new_val
        elif col_index == 4:
            self.productos[idx]['material'] = new_val if new_val else "-"
        elif col_index == 5:
            self.productos[idx]['tipo'] = new_val if new_val else "-"
        elif col_index == 6:
            self.productos[idx]['actualizado'] = new_val if new_val else "-"
        elif col_index == 7:
            try:
                g = int(float(new_val)) if new_val else 0
                self.productos[idx]['gramos'] = f"{g:03d}"
            except:
                self.productos[idx]['gramos'] = "000"
        elif col_index == 8:
            self.productos[idx]['precio'] = new_val
        elif col_index == 9:
            try:
                self.productos[idx]['cantidad'] = int(new_val)
                if self.productos[idx]['cantidad'] < 1:
                    self.productos[idx]['cantidad'] = 1
            except:
                messagebox.showwarning("Error", "Cantidad debe ser número")
                return
        self.actualizar_tabla()

    def cargar_ejemplos(self):
        ejemplos = [
            {'proveedor':'AB','id':'OR001','material':'ORO 10k','tipo':'ANILLO','actualizado':'S','gramos':'008','precio':'150000','cantidad':2},
            {'proveedor':'CD','id':'PL002','material':'PLATA 925','tipo':'PULSERA','actualizado':'N','gramos':'015','precio':'85000','cantidad':1}
        ]
        for prod in ejemplos:
            self.productos.append(prod)
            self.check_vars.append(tk.BooleanVar(value=False))
        self.actualizar_tabla()

# ========== MAIN ==========
if __name__ == "__main__":
    if not validar_licencia():
        sys.exit()
    
    root = tk.Tk()
    root.geometry("1300x800")
    app = EtiquetadoraApp(root)
    root.mainloop()
