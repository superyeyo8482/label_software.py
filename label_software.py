import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import win32print
import platform
import json
import os
import threading
import time

# ========== CONFIGURACIÓN ==========
NOMBRE_IMPRESORA = "ZDesigner ZD220-203dpi ZPL"
LS_VAL = -50
LT_VAL = -43

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

class EtiquetadoraApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Label Software para Zebra ZD220 - Joyería")
        self.root.geometry("1250x850")
        self.root.minsize(1050, 750)
        self.root.configure(bg="#e8f0f8")
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

    # ---------- Configuración ----------
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

    # ---------- Interfaz gráfica ----------
    def crear_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Formulario
        frame_form = ttk.LabelFrame(main_frame, text="Nuevo producto", padding=10)
        frame_form.pack(fill=tk.X, pady=(0,5))
        inner_frame = tk.Frame(frame_form)
        inner_frame.pack(fill=tk.X, padx=5, pady=5)
        col_izq = tk.Frame(inner_frame)
        col_izq.grid(row=0, column=0, padx=(0, 20), sticky="n")
        col_der = tk.Frame(inner_frame)
        col_der.grid(row=0, column=1, padx=(0, 0), sticky="n")
        ENTRY_WIDTH = 25
        campos_izq = [
            ("Proveedor (2 letras):", "proveedor"),
            ("ID Producto:", "id"),
            ("Material (ej. ORO 10k):", "material"),
            ("Tipo (ANILLO/PULSERA):", "tipo")
        ]
        self.entries = {}
        for i, (label, key) in enumerate(campos_izq):
            row_frame = tk.Frame(col_izq)
            row_frame.pack(fill=tk.X, pady=4)
            lbl = ttk.Label(row_frame, text=label, width=22, anchor="e")
            lbl.pack(side=tk.LEFT, padx=(0, 8))
            entry = ttk.Entry(row_frame, width=ENTRY_WIDTH)
            entry.pack(side=tk.LEFT)
            self.entries[key] = entry
            entry.bind('<Return>', self._focus_next)
        campos_der = [
            ("Actualizado (S/N):", "actualizado"),
            ("Gramos (3 dígitos):", "gramos"),
            ("Precio ($):", "precio"),
            ("Cantidad:", "cantidad")
        ]
        for i, (label, key) in enumerate(campos_der):
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
        btn_frame = tk.Frame(frame_form)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        btn_agregar = ttk.Button(btn_frame, text="➕ Agregar producto", command=self.agregar, width=35)
        btn_agregar.pack()

        # Ajustes
        frame_offset = ttk.LabelFrame(main_frame, text="Ajustes de impresión", padding=8)
        frame_offset.pack(fill=tk.X, pady=5)
        ajustes_frame = tk.Frame(frame_offset)
        ajustes_frame.pack()
        ttk.Label(ajustes_frame, text="Desplazamiento horizontal (^LS):").grid(row=0, column=0, padx=5)
        ls_spin = ttk.Spinbox(ajustes_frame, from_=-200, to=200, increment=1, textvariable=self.ls_var, width=8)
        ls_spin.grid(row=0, column=1, padx=5)
        ttk.Label(ajustes_frame, text="Desplazamiento vertical (^LT):").grid(row=0, column=2, padx=5)
        lt_spin = ttk.Spinbox(ajustes_frame, from_=-200, to=200, increment=1, textvariable=self.lt_var, width=8)
        lt_spin.grid(row=0, column=3, padx=5)
        btn_test = ttk.Button(ajustes_frame, text="🖨️ Imprimir prueba", command=self.probar_offset)
        btn_test.grid(row=0, column=4, padx=10)

        # Botones
        frame_botones = ttk.LabelFrame(main_frame, text="Acciones", padding=8)
        frame_botones.pack(fill=tk.X, pady=5)
        btn_eliminar = ttk.Button(frame_botones, text="🗑️ Eliminar seleccionados", command=self.eliminar, width=25)
        btn_eliminar.grid(row=0, column=0, padx=5, pady=2, sticky="ew")
        btn_duplicar = ttk.Button(frame_botones, text="🔄 Duplicar fila marcada", command=self.duplicar, width=25)
        btn_duplicar.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        btn_imprimir_sel = ttk.Button(frame_botones, text="✅ Imprimir seleccionados", command=self.imprimir_chequeados, width=25)
        btn_imprimir_sel.grid(row=0, column=2, padx=5, pady=2, sticky="ew")
        btn_imprimir_todos = ttk.Button(frame_botones, text="🖨️ Imprimir todos", command=self.imprimir_todos, width=25)
        btn_imprimir_todos.grid(row=1, column=0, padx=5, pady=2, sticky="ew")
        self.btn_stop = tk.Button(frame_botones, text="🛑 STOP", command=self.detener_impresion,
                                  bg="#f44336", fg="white", font=("Arial", 10, "bold"),
                                  width=25, height=1, relief="raised")
        self.btn_stop.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        btn_pegar = ttk.Button(frame_botones, text="📋 Pegar Excel", command=self.pegar_desde_excel, width=25)
        btn_pegar.grid(row=1, column=2, padx=5, pady=2, sticky="ew")
        btn_limpiar = ttk.Button(frame_botones, text="🧹 Limpiar tabla", command=self.limpiar, width=25)
        btn_limpiar.grid(row=2, column=0, padx=5, pady=2, sticky="ew")
        btn_cargar = ttk.Button(frame_botones, text="📂 Cargar CSV", command=self.cargar, width=25)
        btn_cargar.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        btn_guardar = ttk.Button(frame_botones, text="💾 Guardar CSV", command=self.guardar, width=25)
        btn_guardar.grid(row=2, column=2, padx=5, pady=2, sticky="ew")
        btn_salir = ttk.Button(frame_botones, text="❌ Salir", command=self.root.destroy, width=25)
        btn_salir.grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        for col in range(3):
            frame_botones.columnconfigure(col, weight=1)

        # Tabla
        frame_tabla = ttk.LabelFrame(main_frame, text="Lista de productos", padding=8)
        frame_tabla.pack(fill=tk.BOTH, expand=True, pady=5)
        tree_frame = ttk.Frame(frame_tabla)
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
        frame_resumen = ttk.LabelFrame(main_frame, text="Resumen del lote", padding=8)
        frame_resumen.pack(fill=tk.X, pady=5)
        self.label_productos = ttk.Label(frame_resumen, text="📦 Productos distintos: 0", 
                                         font=("Arial", 11, "bold"), foreground="#2c5f8a")
        self.label_productos.pack(side=tk.LEFT, padx=15, pady=3)
        self.label_etiquetas = ttk.Label(frame_resumen, text="🏷️ Total de etiquetas: 0", 
                                         font=("Arial", 11, "bold"), foreground="#2c5f8a")
        self.label_etiquetas.pack(side=tk.LEFT, padx=15, pady=3)
        self.label_valor = ttk.Label(frame_resumen, text="💰 Valor total: $0", 
                                     font=("Arial", 11, "bold"), foreground="#2c5f8a")
        self.label_valor.pack(side=tk.LEFT, padx=15, pady=3)
        self.label_estado = ttk.Label(frame_resumen, text="✅ Listo", foreground="green", font=("Arial", 10))
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
        self.label_productos.config(text=f"📦 Productos distintos: {num_productos}")
        self.label_etiquetas.config(text=f"🏷️ Total de etiquetas: {total_etiquetas}")
        self.label_valor.config(text=f"💰 Valor total: ${total_precio:,.0f}")

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
        if not (proveedor and idProd):
            messagebox.showwarning("Campos incompletos", "Proveedor e ID son obligatorios.")
            return
        try:
            cant = int(cantidad) if cantidad else 1
        except:
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
            respuesta = messagebox.askyesno("Encabezados", "¿La primera fila contiene los títulos?")
            inicio = 1 if respuesta else 0
            nuevos = []
            errores = []
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
                    if not proveedor or not idProd:
                        errores.append(f"Fila {idx_fila}: Proveedor o ID vacío")
                        continue
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
                self.root.after(0, lambda: self.label_estado.config(text="✅ Completado", foreground="green"))
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
                    nuevos = []
                    for row in reader:
                        nuevos.append({
                            'proveedor': row.get('proveedor', '').strip(),
                            'id': row.get('id', '').strip(),
                            'material': row.get('material', '-').strip(),
                            'tipo': row.get('tipo', '-').strip(),
                            'actualizado': row.get('actualizado', '-').strip(),
                            'gramos': row.get('gramos', '000').strip(),
                            'precio': row.get('precio', '0').strip(),
                            'cantidad': int(row.get('cantidad', 1))
                        })
                    self.productos.extend(nuevos)
                    self.check_vars.extend([tk.BooleanVar(value=False) for _ in nuevos])
                    self.actualizar_tabla()
                messagebox.showinfo("Cargado", f"Se agregaron {len(nuevos)} productos.")
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

if __name__ == "__main__":
    root = tk.Tk()
    app = EtiquetadoraApp(root)
    root.mainloop()
