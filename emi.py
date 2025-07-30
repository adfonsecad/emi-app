# Línea 1: Versión del software
__version__ = "1.0.1"

# imports


import tkinter as tk
from tkinter import messagebox, ttk
import urllib3
import os
from datetime import datetime
import uuid

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from cryptography.hazmat.primitives.serialization.pkcs12 import load_key_and_certificates
from cryptography.hazmat.backends import default_backend

from packaging import version
import requests
from tkinter import messagebox
import sys
import os
from cryptography.hazmat.primitives.serialization.pkcs12 import load_key_and_certificates
from cryptography.hazmat.backends import default_backend

def ruta_recurso(nombre_archivo):
    """Devuelve la ruta correcta del archivo, compatible con PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, nombre_archivo)
    return os.path.join(os.path.abspath("."), nombre_archivo)

# === Validar Version  ===

def verificar_actualizacion():
    url_version = "https://raw.githubusercontent.com/adfonsecad/emi-app/main/version.txt"
    url_archivo = "https://raw.githubusercontent.com/adfonsecad/emi-app/main/emi.py"
    
    try:
        version_remota = requests.get(url_version, timeout=5).text.strip()
        version_local = __version__.strip()
        
        if version.parse(version_remota) > version.parse(version_local):
            respuesta = messagebox.askyesno(
                "Actualización disponible",
                f"Hay una nueva versión ({version_remota}). ¿Deseas descargarla?"
            )
            if respuesta:
                nuevo_codigo = requests.get(url_archivo, timeout=10).text

                # Reemplazar la línea de versión en el nuevo código
                lineas = nuevo_codigo.splitlines()
                for i, linea in enumerate(lineas):
                    if linea.startswith("__version__"):
                        lineas[i] = f'__version__ = "{version_remota}"'
                        break
                nuevo_codigo_actualizado = "\n".join(lineas)

                with open("emi_actualizado.py", "w", encoding="utf-8") as f:
                    f.write(nuevo_codigo_actualizado)

                messagebox.showinfo("Actualizado", "Se descargó la nueva versión como 'emi_actualizado.py'.")
        else:
            messagebox.showinfo("Sin actualizaciones", "Ya tienes la última versión.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo verificar la actualización:\n{e}")


# === DESACTIVA ADVERTENCIAS SSL SOLO EN DESARROLLO ===
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# === GOOGLE SHEETS CONFIGURACIÓN ===
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = '1tZmLipOF8I_QPkNbWMwhfUEtf-fq2vJ6wLZ9sEKPVqU'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Usar ruta_recurso para acceder al archivo JSON
ruta_json = ruta_recurso("credenciales.json")

# Cargar credenciales y construir el servicio
credenciales = Credentials.from_service_account_file(ruta_json, scopes=SCOPES)
servicio = build('sheets', 'v4', credentials=credenciales)
hoja = servicio.spreadsheets()


# === FUNCIÓN PARA OBTENER USUARIOS DESDE GOOGLE SHEETS ===
def obtener_usuarios():
    try:
        resultado = hoja.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Log Ins!A2:E"  # Incluye columna E (Permisos)
        ).execute()
        valores = resultado.get('values', [])
        usuarios = []
        for fila in valores:
            fila += [''] * (5 - len(fila))  # Asegura 5 columnas
            usuarios.append({
                "User": fila[0].strip().upper(),
                "Password": fila[1].strip(),
                "email": fila[2].strip(),
                "company": fila[3].strip(),
                "Permisos": [p.strip() for p in fila[4].split(",")] if fila[4] else []
            })
        return usuarios
    except Exception as e:
        print("❌ Error al obtener usuarios:", e)
        return []

# === FUNCIÓN PARA OBTENER CLIENTES DESDE GOOGLE SHEETS ===
def obtener_clientes():
    try:
        resultado = hoja.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Client ID!A2:I"
        ).execute()
        valores = resultado.get('values', [])
        clientes = []
        for fila in valores:
            fila += [''] * (8 - len(fila))
            clientes.append({
                'CustomerID': fila[0],
                'Nombre': fila[1],
                'Identificacion': fila[2],
                'email': fila[3],
                'Activity': fila[4],
                'Contact': fila[5],
                'Vendedor': fila[6],
                'Descuento': fila[7],
                'Excento de IVA': fila[8],
            })
        return clientes
    except Exception as e:
        print("Error al obtener clientes:", e)
        return []

# === FUNCIÓN PARA OBTENER PRODUCTOS (con Código Cabys y Cabys Exento) ===
def obtener_productos():
    try:
        resultado = hoja.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Producto!A2:E"  # Ampliado hasta columna E
        ).execute()
        valores = resultado.get('values', [])
        productos = []
        for fila in valores:
            # Asegurar que la fila tenga al menos 4 columnas
            if len(fila) < 4:
                fila += [''] * (4 - len(fila))
            
            especie = fila[0]
            try:
                valor_unitario = float(fila[1])
            except:
                valor_unitario = 0.0

            codigo_cabys = fila[2]
            codigo_cabys_exento = fila[3]

            productos.append({
                'especie': especie,
                'valor_unitario': valor_unitario,
                'codigo_cabys': codigo_cabys,
                'codigo_cabys_exento': codigo_cabys_exento
            })
        return productos
    except Exception as e:
        print("Error al obtener productos:", e)
        return []

# === CERTIFICADO P12 ===

def cargar_certificado_p12(nombre_archivo_p12, clave):
    ruta_completa = ruta_recurso(nombre_archivo_p12)

    try:
        with open(ruta_completa, 'rb') as archivo:
            contenido_p12 = archivo.read()

        private_key, cert, additional_certs = load_key_and_certificates(
            contenido_p12,
            password=clave.encode(),
            backend=default_backend()
        )
        print("✅ Certificado .p12 cargado correctamente.")
        return cert, private_key
    except Exception as e:
        print(f"❌ Error al cargar el certificado: {e}")
        return None, None

# Llamada a la función
cert, key = cargar_certificado_p12("030495001321.p12", "1987")


# === Clase para editar clientes ===
class EditarClientes(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Editar Clientes - Base de Datos")
        self.geometry("900x500")

        self.clientes = []
        self.clientes_filas = []

        columnas = ["CustomerID", "Nombre", "Identificacion", "email", "Activity", "Contact", "Vendedor", "Descuento", "Excento de IVA"]
        self.tree = ttk.Treeview(self, columns=columnas, show="headings")
        for col in columnas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<<TreeviewSelect>>", self.cargar_cliente_seleccionado)

        frame_edicion = tk.Frame(self)
        frame_edicion.pack(pady=10, fill='x')

        self.vars_campos = {}
        for i, col in enumerate(columnas):
            tk.Label(frame_edicion, text=col + ":").grid(row=i, column=0, sticky='e')
            var = tk.StringVar()
            entry = tk.Entry(frame_edicion, textvariable=var, width=60)
            entry.grid(row=i, column=1, padx=5, pady=2, sticky='w')
            self.vars_campos[col] = var

        btn_guardar = tk.Button(self, text="Guardar Cambios", command=self.guardar_cambios)
        btn_guardar.pack(pady=10)

        self.cargar_clientes()

    def cargar_clientes(self):
        try:
            resultado = hoja.values().get(
                spreadsheetId=SPREADSHEET_ID,
                range="Client ID!A2:I"
            ).execute()
            valores = resultado.get('values', [])
            self.clientes = valores
            self.tree.delete(*self.tree.get_children())
            self.clientes_filas = []
            for i, fila in enumerate(valores, start=2):
                fila += [''] * (9 - len(fila))
                self.tree.insert('', 'end', values=fila)
                self.clientes_filas.append(i)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar la base de datos: {e}")

    def cargar_cliente_seleccionado(self, event):
        seleccion = self.tree.selection()
        if seleccion:
            valores = self.tree.item(seleccion[0], 'values')
            for i, col in enumerate(self.vars_campos):
                self.vars_campos[col].set(valores[i])

    def guardar_cambios(self):
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Selección", "Seleccione un cliente para guardar cambios.")
            return

        fila_seleccionada = self.tree.index(seleccion[0])
        fila_hoja = self.clientes_filas[fila_seleccionada]

        datos_actualizados = [self.vars_campos[col].get() for col in self.vars_campos]

        try:
            rango = f"Client ID!A{fila_hoja}:I{fila_hoja}"
            hoja.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=rango,
                valueInputOption="RAW",
                body={"values": [datos_actualizados]}
            ).execute()

            self.tree.item(seleccion[0], values=datos_actualizados)
            messagebox.showinfo("Éxito", "Datos guardados correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar los datos: {e}")

# === Gestionar Lineas ===
class GestionLineas(tk.Toplevel):
    def __init__(self, master, tipo_documento):
        super().__init__(master)
        self.title(f"Factura Electrónica - Gestión de Líneas (Tipo: {tipo_documento})")
        self.geometry("1100x700")

        self.clientes = obtener_clientes()
        self.clientes_por_id = {c['CustomerID']: c for c in self.clientes}

        self.productos = obtener_productos()
        self.productos_por_especie = {p['especie']: p['valor_unitario'] for p in self.productos}

        frame_gen = tk.Frame(self)
        frame_gen.pack(pady=10, fill='x')

        tk.Label(frame_gen, text="Consecutivo:").grid(row=0, column=0, sticky='w')
        tk.Label(frame_gen, text="01010001001000000045").grid(row=0, column=1, sticky='w')

        tk.Label(frame_gen, text="Fecha:").grid(row=0, column=2, sticky='w', padx=(20, 0))
        tk.Label(frame_gen, text=datetime.today().strftime("%d/%m/%Y")).grid(row=0, column=3, sticky='w')

        tk.Label(frame_gen, text="Tipo de cambio (Compra):").grid(row=0, column=4, sticky='w', padx=(20, 0))
        self.tipo_cambio_var = tk.StringVar(value="No disponible")
        self.tipo_cambio_entry = tk.Entry(frame_gen, textvariable=self.tipo_cambio_var, width=15)
        self.tipo_cambio_entry.grid(row=0, column=5, sticky='w')


        # --- CLIENTE ---
        frame_cliente = tk.LabelFrame(self, text="Datos del Cliente")
        frame_cliente.pack(padx=10, pady=10, fill='x')

        tk.Label(frame_cliente, text="Customer ID:").grid(row=0, column=0, sticky='e', padx=5, pady=3)
        self.cliente_id_var = tk.StringVar()
        self.combo_cliente = ttk.Combobox(frame_cliente, textvariable=self.cliente_id_var,
                                          values=[c['CustomerID'] for c in self.clientes], state="readonly", width=30)
        self.combo_cliente.grid(row=0, column=1, sticky='w', padx=5, pady=3)
        self.combo_cliente.bind("<<ComboboxSelected>>", self.on_cliente_seleccionado)

        etiquetas = ["Nombre", "Identificación", "Email", "Activity", "Contact", "Vendedor", "Descuento", "Excento de IVA"]
        self.campos_cliente = {}
        for i, etiqueta in enumerate(etiquetas, start=1):
            tk.Label(frame_cliente, text=etiqueta + ":").grid(row=i, column=0, sticky='e', padx=5, pady=3)
            var = tk.StringVar()
            entry = tk.Entry(frame_cliente, textvariable=var, width=50)
            entry.grid(row=i, column=1, sticky='w', padx=5, pady=3)
            self.campos_cliente[etiqueta] = var

        # --- PRODUCCIÓN DE MARIPOSAS ---
        frame_mariposas = tk.LabelFrame(self, text="Producción de Mariposas")
        frame_mariposas.pack(pady=10, padx=10, fill="x")

        tk.Label(frame_mariposas, text="Especie:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.especie_var = tk.StringVar()
        self.combo_especie = ttk.Combobox(
            frame_mariposas,
            textvariable=self.especie_var,
            values=[p['especie'] for p in self.productos],
            width=40,
            state="readonly"
        )
        self.combo_especie.grid(row=0, column=1, padx=5, pady=5)
        self.combo_especie.bind("<<ComboboxSelected>>", self.on_especie_seleccionada)

        tk.Label(frame_mariposas, text="Valor Unitario:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.valor_unitario_var = tk.StringVar(value="0.0")
        self.entry_valor_unitario = tk.Entry(frame_mariposas, textvariable=self.valor_unitario_var, width=15, state='readonly')
        self.entry_valor_unitario.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(frame_mariposas, text="Cantidad de pupas:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.cantidad_var = tk.StringVar()
        self.entry_cantidad = tk.Entry(frame_mariposas, textvariable=self.cantidad_var, width=10)
        self.entry_cantidad.grid(row=0, column=5, padx=5, pady=5)

        btn_agregar = tk.Button(frame_mariposas, text="Agregar a pedido", command=self.agregar_a_lista)
        btn_agregar.grid(row=0, column=6, padx=10, pady=5)

        tk.Label(frame_mariposas, text="Código Cabys:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.codigo_cabys_var = tk.StringVar(value="No disponible")
        self.entry_codigo_cabys = tk.Entry(frame_mariposas, textvariable=self.codigo_cabys_var, width=40, state='readonly')
        self.entry_codigo_cabys.grid(row=1, column=1, padx=5, pady=5, columnspan=2)

        # Tabla con Item, Especie, Cantidad, Valor Unitario y Código Cabys
        self.tree = ttk.Treeview(self, columns=("Item", "Especie", "Cantidad", "Valor Unitario", "Código Cabys"), show="headings", height=10)
        self.tree.heading("Item", text="Item")
        self.tree.heading("Especie", text="Especie")
        self.tree.heading("Cantidad", text="Cantidad de pupas")
        self.tree.heading("Valor Unitario", text="Valor Unitario")
        self.tree.heading("Código Cabys", text="Código Cabys")
        self.tree.column("Item", width=50, anchor='center')
        self.tree.column("Especie", width=250)
        self.tree.column("Cantidad", width=150, anchor='center')
        self.tree.column("Valor Unitario", width=100, anchor='e')
        self.tree.column("Código Cabys", width=200)
        self.tree.pack(padx=10, pady=10, fill="both", expand=True)

        btn_guardar = tk.Button(self, text="Guardar pedido en Google Sheets", command=self.guardar_pedido)
        btn_guardar.pack(pady=10)

    def on_cliente_seleccionado(self, event):
        cid = self.cliente_id_var.get()
        cliente = self.clientes_por_id.get(cid)
        if cliente:
            for campo in self.campos_cliente:
                self.campos_cliente[campo].set(cliente.get(campo, ''))
        else:
            for key in self.campos_cliente:
                self.campos_cliente[key].set('')

    def on_especie_seleccionada(self, event):
        especie = self.especie_var.get()
        valor = self.productos_por_especie.get(especie, 0.0)
        self.valor_unitario_var.set(f"{valor:.2f}")

        cid = self.cliente_id_var.get()
        cliente = self.clientes_por_id.get(cid)
        es_exento = cliente.get("Excento de IVA", "").strip().lower() in ["Y", "y", "true", "1"]

        producto = next((p for p in self.productos if p['especie'] == especie), None)
        if producto:
            codigo = producto.get('Codigo Cabys Excento') if es_exento else producto.get('Codigo Cabys')
        else:
            codigo = None  # o podrías usar "No disponible" directamente

        self.codigo_cabys_var.set(codigo if codigo else "No disponible")


    def agregar_a_lista(self):
        especie = self.especie_var.get()
        cantidad = self.cantidad_var.get()
        valor_unitario = self.valor_unitario_var.get()
        codigo_cabys = self.codigo_cabys_var.get()

        if not especie:
            messagebox.showwarning("Falta especie", "Debe seleccionar una especie.")
            return
        if not cantidad.isdigit() or int(cantidad) <= 0:
            messagebox.showwarning("Cantidad inválida", "Debe ingresar una cantidad válida (número entero mayor que 0).")
            return

        item_num = len(self.tree.get_children()) + 1
        self.tree.insert('', 'end', values=(item_num, especie, cantidad, valor_unitario, codigo_cabys))
        self.especie_var.set('')
        self.cantidad_var.set('')
        self.valor_unitario_var.set('0.00')
        self.codigo_cabys_var.set('No disponible')


    def guardar_pedido(self):
        if not self.tree.get_children():
            messagebox.showwarning("Lista vacía", "No hay artículos para guardar.")
            return

        pedido = []
        for iid in self.tree.get_children():
            item = self.tree.item(iid)['values']
            pedido.append(item)

        valores_a_guardar = []
        for item in pedido:
            valores_a_guardar.append([str(v) for v in item])  # Incluye Código Cabys

        try:
            hoja.values().append(
                spreadsheetId=SPREADSHEET_ID,
                range="Pedidos!A2",
                valueInputOption="RAW",
                body={"values": valores_a_guardar}
            ).execute()
            messagebox.showinfo("Guardado", "Pedido guardado correctamente en Google Sheets.")
            self.tree.delete(*self.tree.get_children())
        except Exception as e:
            print("Error al guardar:", e)
            messagebox.showerror("Error", f"No se pudo guardar el pedido: {e}")


# === LOGIN VENTANA ===
class LoginVentana(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Login - emi")
        self.geometry("300x200")
        self.resizable(False, False)

        tk.Label(self, text="Usuario:", font=("Arial", 12)).pack(pady=(15, 5))
        self.usuario_var = tk.StringVar()
        tk.Entry(self, textvariable=self.usuario_var, font=("Arial", 12)).pack()

        tk.Label(self, text="Contraseña:", font=("Arial", 12)).pack(pady=(10, 5))
        self.password_var = tk.StringVar()
        tk.Entry(self, textvariable=self.password_var, font=("Arial", 12), show="*").pack()

        tk.Button(self, text="Ingresar", command=self.validar_usuario).pack(pady=15)

        try:
            self.usuarios = obtener_usuarios()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar la lista de usuarios: {e}")
            self.usuarios = []

    def validar_usuario(self):
        usuario = self.usuario_var.get().strip().upper()
        password = self.password_var.get().strip()

        for u in self.usuarios:
            if u.get('User', '').strip().upper() == usuario and u.get('Password', '').strip() == password:
                self.destroy()
                app = App(u)  # pasa el diccionario completo
                app.mainloop()
                return

        messagebox.showerror("Acceso denegado", "Usuario o contraseña incorrectos.")

# === APP PRINCIPAL ===

class App(tk.Tk):
    def __init__(self, usuario_info):
        super().__init__()
        self.title("Sistema de Facturación - emi")
        self.geometry("945x566")

        self.usuario = usuario_info["User"]
        self.email = usuario_info.get("email", "")
        self.empresa = usuario_info.get("company", "")
        self.permisos = usuario_info.get("Permisos", [])

        # Título general
        tk.Label(self, text="emi - Creación de Documentos", font=("Arial", 16)).pack(pady=10)

        # Frame para usuario, correo, empresa y botón actualización
        frame_usuario = tk.Frame(self)
        frame_usuario.pack(fill="x", padx=10)

        # Frame izquierdo para etiquetas
        frame_info = tk.Frame(frame_usuario)
        frame_info.pack(side="left", anchor="w")

        tk.Label(frame_info, text=f"Usuario activo: {self.usuario}", font=("Arial", 10)).pack(anchor="w")
        tk.Label(frame_info, text=f"Correo: {self.email}", font=("Arial", 10)).pack(anchor="w")
        tk.Label(frame_info, text=f"Empresa: {self.empresa}", font=("Arial", 10)).pack(anchor="w", pady=(0,5))

        # Botón de actualización a la derecha
        btn_actualizar = tk.Button(frame_usuario, text="Buscar actualizaciones", command=verificar_actualizacion)
        btn_actualizar.pack(side="right")

        # Botones de documentos
        for nombre, codigo in TIPOS_DOCUMENTO.items():
            if nombre in self.permisos:
                btn = tk.Button(self, text=nombre, width=30,
                                command=lambda c=codigo: self.abrir_gestion_lineas(c))
                btn.pack(pady=5)

        # Botón de clientes
        if "Datos Cliente y Resumen" in self.permisos:
            btn_cliente = tk.Button(self, text="Datos Cliente y Resumen", width=30,
                                    command=self.abrir_editar_clientes)
            btn_cliente.pack(pady=5)

    # ✅ Estos métodos deben estar dentro de la clase
    def abrir_gestion_lineas(self, tipo_documento):
        GestionLineas(self, tipo_documento)

    def abrir_editar_clientes(self):
        EditarClientes(self)

# === TIPOS DE DOCUMENTO ===
TIPOS_DOCUMENTO = {
    "Factura Electrónica": "01",
    "Nota de Crédito": "03",
    "Nota de Débito": "02",
    "Tiquete Electrónico": "04",
    "Factura de Compra": "05"
}

# === EJECUCIÓN ===
if __name__ == "__main__":
    login = LoginVentana()
    login.mainloop()
