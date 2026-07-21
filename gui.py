"""Tkinter view code. Delegates ZPL/printing to printing.py, license checks
to licensing.py, translated strings to i18n.py, batch math to models.py,
CSV row conversion to csv_io.py, and clipboard-paste parsing to
excel_paste.py - this module should only contain widget wiring and calls
out to those.
"""
import csv
import json
import logging
import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import config
import csv_io
import excel_paste
import i18n
import licensing
import models
import printing

logger = logging.getLogger(__name__)

ENTRY_WIDTH = 25
WINDOW_GEOMETRY = "1300x800"
ACTIVATION_WINDOW_GEOMETRY = "500x300"
TREE_COLUMN_WIDTHS = {
    "Proveedor": 80, "ID": 100, "Material": 120, "Tipo": 120,
    "Act": 50, "Gramos": 70, "Precio": 100, "Cantidad": 80,
}


def mostrar_dialogo_activacion(id_actual: str) -> bool:
    """Blocking activation window. Returns True once a valid code has been
    entered and saved, False if the user closes the window without
    activating. Mirrors the original validar_licencia() window exactly.
    """
    activado = False

    root_temp = tk.Tk()
    root_temp.withdraw()
    ventana = tk.Toplevel(root_temp)
    ventana.title("Activación - Label Software")
    ventana.geometry(ACTIVATION_WINDOW_GEOMETRY)
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
        nonlocal activado
        entrada = clave_var.get().strip()
        if licensing.validar_codigo(entrada, id_actual):
            licensing.guardar_codigo(entrada)
            messagebox.showinfo("Éxito", "Software activado correctamente")
            activado = True
            ventana.destroy()
            root_temp.destroy()
        else:
            messagebox.showerror("Error", "Clave inválida")

    tk.Button(ventana, text="Activar", command=verificar).pack(pady=10)
    ventana.protocol("WM_DELETE_WINDOW", lambda: (root_temp.destroy(), ventana.destroy()))
    root_temp.wait_window(ventana)
    return activado


class EtiquetadoraApp:
    def __init__(self, root):
        self.root = root
        self.idioma = config.DEFAULT_LANGUAGE
        self.productos = []
        self.check_vars = []
        self.imprimiendo = False
        self.stop_impresion = False
        self.edit_entry = None
        self.edit_item = None

        ls_saved, lt_saved = self.cargar_configuracion()
        self.ls_var = tk.IntVar(value=ls_saved)
        self.lt_var = tk.IntVar(value=lt_saved)

        self.crear_widgets()
        self.actualizar_tabla()
        self.actualizar_resumen()
        self.cargar_ejemplos()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def t(self, key: str) -> str:
        return i18n.t(self.idioma, key)

    def cargar_configuracion(self):
        if not os.path.exists(config.CONFIG_FILE):
            return config.LS_VAL_DEFAULT, config.LT_VAL_DEFAULT
        try:
            with open(config.CONFIG_FILE, 'r', encoding='utf-8') as f:
                datos = json.load(f)
            return datos.get('ls', config.LS_VAL_DEFAULT), datos.get('lt', config.LT_VAL_DEFAULT)
        except (json.JSONDecodeError, OSError, KeyError) as exc:
            logger.warning("Config corrupta o ilegible (%s), usando valores por defecto", exc)
            return config.LS_VAL_DEFAULT, config.LT_VAL_DEFAULT

    def guardar_configuracion(self):
        datos = {'ls': self.ls_var.get(), 'lt': self.lt_var.get()}
        try:
            with open(config.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(datos, f, indent=4)
        except OSError as exc:
            logger.error("No se pudo guardar la configuración: %s", exc)

    def on_closing(self):
        self.stop_impresion = True
        self.guardar_configuracion()
        self.root.destroy()

    def _focus_next(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

    def generar_id(self, proveedor: str) -> str:
        return models.generar_id(self.productos, proveedor)

    def _completar_ids_vacios(self, productos: list):
        """Shared by pegar_desde_excel and cargar: if any product in the
        newly-imported batch is missing an ID, ask once whether to
        auto-generate IDs for all of them.
        """
        if not models.hay_ids_vacios(productos):
            return
        if messagebox.askyesno(self.t('id_vacio'), f"{self.t('id_vacio')}\n\n{self.t('generar_ids')}"):
            for producto in productos:
                if not producto['id']:
                    producto['id'] = self.generar_id(producto['proveedor'])

    # ---------- INTERFAZ GRAFICA ----------
    def crear_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Formulario
        self.frame_form = ttk.LabelFrame(main_frame, text=self.t('nuevo_producto'), padding=10)
        self.frame_form.pack(fill=tk.X, pady=(0, 5))
        inner_frame = tk.Frame(self.frame_form)
        inner_frame.pack(fill=tk.X, padx=5, pady=5)
        col_izq = tk.Frame(inner_frame)
        col_izq.grid(row=0, column=0, padx=(0, 20), sticky="n")
        col_der = tk.Frame(inner_frame)
        col_der.grid(row=0, column=1, padx=(0, 0), sticky="n")
        campos_izq = [
            ("proveedor", self.t('proveedor')),
            ("id", self.t('id')),
            ("material", self.t('material')),
            ("tipo", self.t('tipo')),
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
            ("cantidad", self.t('cantidad')),
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
        for col in columnas[2:]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=TREE_COLUMN_WIDTHS.get(col, 100))
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.tree.bind('<ButtonRelease-1>', self.on_click_check)
        self.tree.bind('<Double-1>', self.editar_celda)

        # Resumen del lote
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
        resumen = models.calcular_resumen(self.productos)
        self.label_productos.config(text=f"{self.t('productos_distintos')} {resumen.num_productos}")
        self.label_etiquetas.config(text=f"{self.t('total_etiquetas')} {resumen.total_etiquetas}")
        self.label_valor.config(text=f"{self.t('valor_total')} ${resumen.total_valor:,.0f}")

    def actualizar_tabla(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, prod in enumerate(self.productos, start=1):
            check = "✓" if self.check_vars[i - 1].get() else ""
            self.tree.insert("", tk.END, values=(
                i, check,
                prod['proveedor'], prod['id'], prod['material'],
                prod['tipo'], prod['actualizado'], prod['gramos'],
                prod['precio'], prod['cantidad'],
            ))
        self.actualizar_resumen()

    def on_click_check(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        if column != "#2":
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
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
            'cantidad': cant,
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

    # ---------- PEGAR DESDE EXCEL ----------
    def pegar_desde_excel(self):
        try:
            datos_raw = self.root.clipboard_get()
            if not datos_raw:
                messagebox.showwarning("Portapapeles vacío", "No hay datos para pegar.")
                return

            respuesta_encabezados = messagebox.askyesno("Encabezados", "¿La primera fila contiene los títulos?")
            resultado = excel_paste.parsear_filas_pegadas(datos_raw, primera_fila_es_encabezado=respuesta_encabezados)

            self._completar_ids_vacios(resultado.productos)

            if resultado.productos:
                self.productos.extend(resultado.productos)
                self.check_vars.extend([tk.BooleanVar(value=False) for _ in resultado.productos])
                self.actualizar_tabla()
                msg = f"✅ Se agregaron {len(resultado.productos)} productos."
                if resultado.errores:
                    msg += f"\n⚠️ {len(resultado.errores)} errores"
                messagebox.showinfo("Éxito", msg)
            else:
                messagebox.showwarning("Error", "No se pudo procesar ningún dato")
        except tk.TclError as exc:
            logger.info("Portapapeles vacío o inaccesible: %s", exc)
            messagebox.showwarning("Portapapeles vacío", "No hay datos para pegar.")
        except Exception as exc:
            logger.exception("Error inesperado al pegar desde Excel")
            messagebox.showerror("Error", f"No se pudo pegar:\n{exc}")

    # ---------- IMPRESION ----------
    def detener_impresion(self):
        if self.imprimiendo:
            self.stop_impresion = True
            self.label_estado.config(text="🛑 Deteniendo...", foreground="red")
        else:
            messagebox.showinfo("Sin impresión", "No hay ninguna impresión en curso.")

    def _imprimir(self, lista):
        if not lista:
            return
        resumen = models.calcular_resumen(lista)
        respuesta = messagebox.askyesno(
            "Confirmar impresión",
            f"📊 Productos: {len(lista)}\n"
            f"🏷️ Etiquetas: {resumen.total_etiquetas}\n"
            f"💰 Valor total: ${resumen.total_valor:,.0f}\n\n"
            f"¿Imprimir?"
        )
        if not respuesta:
            return
        self.stop_impresion = False
        self.imprimiendo = True
        self.label_estado.config(text="🖨️ Imprimiendo...", foreground="orange")
        hilo = threading.Thread(target=self._impresion_hilo, args=(lista, resumen.total_etiquetas))
        hilo.daemon = True
        hilo.start()

    def _impresion_hilo(self, lista, total_etiquetas):
        ls = self.ls_var.get()
        lt = self.lt_var.get()
        try:
            total = printing.imprimir_lote(
                lista, ls, lt,
                debe_detenerse=lambda: self.stop_impresion,
                on_progreso=lambda t: self.root.after(
                    0, lambda t=t: self.label_estado.config(text=f"🖨️ {t}/{total_etiquetas}", foreground="orange")
                ),
            )
        except printing.PrintError as exc:
            logger.error(str(exc))
            self.root.after(0, lambda: self.label_estado.config(text="❌ Error", foreground="red"))
            return
        except Exception as exc:
            logger.exception("Error inesperado durante la impresión")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error en impresión: {exc}"))
            return
        else:
            if not self.stop_impresion:
                printing.enviar_zpl("~FF")
                self.root.after(0, lambda: self.label_estado.config(text=self.t('listo'), foreground="green"))
                self.root.after(0, lambda: messagebox.showinfo("Éxito", f"Se imprimieron {total} etiquetas."))
            else:
                self.root.after(0, lambda: self.label_estado.config(text="⏹️ Detenido", foreground="orange"))
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
            'cantidad': 1,
        }
        zpl = printing.generar_zpl(ejemplo, self.ls_var.get(), self.lt_var.get())
        if printing.enviar_zpl(zpl):
            messagebox.showinfo("Prueba", "Etiqueta de prueba enviada. Verifica la posición.")

    # ---------- CSV ----------
    def guardar(self):
        if not self.productos:
            messagebox.showwarning("Lista vacía", "No hay productos que guardar.")
            return
        archivo = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not archivo:
            return
        try:
            with open(archivo, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=csv_io.CSV_FIELDNAMES)
                writer.writeheader()
                writer.writerows(self.productos)
        except OSError as exc:
            logger.error("No se pudo guardar CSV en %s: %s", archivo, exc)
            messagebox.showerror("Error", f"No se pudo guardar:\n{exc}")
            return
        messagebox.showinfo("Guardado", f"Lista guardada en {archivo}")

    def cargar(self):
        archivo = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if not archivo:
            return
        try:
            with open(archivo, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                nuevos = [p for p in (csv_io.fila_csv_a_producto(row) for row in reader) if p is not None]
        except (OSError, csv.Error, ValueError) as exc:
            logger.error("No se pudo cargar CSV %s: %s", archivo, exc)
            messagebox.showerror("Error", f"No se pudo cargar:\n{exc}")
            return

        self._completar_ids_vacios(nuevos)
        self.productos.extend(nuevos)
        self.check_vars.extend([tk.BooleanVar(value=False) for _ in nuevos])
        self.actualizar_tabla()
        messagebox.showinfo("Cargado", f"Se agregaron {len(nuevos)} productos.")

    # ---------- EDICION DE CELDAS ----------
    def editar_celda(self, event):
        if self.edit_entry:
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
        x, y, w, h = self.tree.bbox(item, f"#{col_index + 1}")
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
            except (ValueError, TypeError):
                self.productos[idx]['gramos'] = "000"
        elif col_index == 8:
            self.productos[idx]['precio'] = new_val
        elif col_index == 9:
            try:
                cantidad = int(new_val)
                self.productos[idx]['cantidad'] = cantidad if cantidad >= 1 else 1
            except (ValueError, TypeError):
                messagebox.showwarning("Error", "Cantidad debe ser número")
                return
        self.actualizar_tabla()

    def cargar_ejemplos(self):
        ejemplos = [
            {'proveedor': 'AB', 'id': 'OR001', 'material': 'ORO 10k', 'tipo': 'ANILLO', 'actualizado': 'S', 'gramos': '008', 'precio': '150000', 'cantidad': 2},
            {'proveedor': 'CD', 'id': 'PL002', 'material': 'PLATA 925', 'tipo': 'PULSERA', 'actualizado': 'N', 'gramos': '015', 'precio': '85000', 'cantidad': 1},
        ]
        for prod in ejemplos:
            self.productos.append(prod)
            self.check_vars.append(tk.BooleanVar(value=False))
        self.actualizar_tabla()

    # ---------- IDIOMA ----------
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
        t = i18n.TEXTOS[self.idioma]
        for child in self.frame_form.winfo_children():
            self._actualizar_etiquetas_recursivo(child, mapeo_etiquetas, t)

    def cambiar_idioma(self):
        self.idioma = 'en' if self.idioma == 'es' else 'es'
        t = i18n.TEXTOS[self.idioma]

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
