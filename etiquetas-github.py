import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
import csv
import sqlite3
import json
import re
from PIL import Image, ImageDraw, ImageFont, ImageTk
import barcode
from barcode.writer import ImageWriter
import io
import win32print

# --- CONSTANTS ---
DIRECTORIO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
base_dir = DIRECTORIO_SCRIPT
RUTA_DB = os.path.join(base_dir, "control_rollos.db")
RUTA_CONFIG = os.path.join(base_dir, "config.json")
PASSWORD_ADMIN = "admin123"

COLUMNAS_CSV = [
    "Fecha entrada", "Fecha salida", "Area", "Numero de parte",
    "Lote interno", "Linea", "No. piezas", "Piezas por bolsa",
    "No. bolsas", "Lote juliano", "FIFO", "Sobrante",
    "Serial inicial", "Serial final", "Pedido ID", "Fecha pedido", "Scrap", "Destino"
]

# --- GENERAL STYLES ---
COLOR_FONDO = "#f0f0f0"
COLOR_FONDO_FRAME = "#e8e8e8"
COLOR_TITULO = "#2c3e50"
COLOR_BOTON = "#34495e"
COLOR_LABEL = "#333333"
COLOR_ENTRY = "#ffffff"
FUENTE_TITULO = ("Segoe UI", 16, "bold")
FUENTE_SUBTITULO = ("Segoe UI", 11, "bold")
FUENTE_LABEL = ("Segoe UI", 9)
FUENTE_BOTON = ("Segoe UI", 9, "bold")
FUENTE_CAMPO = ("Segoe UI", 9)

# --- CONFIG JSON ---
def config_vacia():
    return {"clientes": {}, "partes": {}, "impresora": "ZDesigner ZM400 300 dpi (ZPL)"}

def cargar_config():
    if not os.path.exists(RUTA_CONFIG):
        guardar_config(config_vacia())
    try:
        with open(RUTA_CONFIG, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return config_vacia()

def guardar_config(cfg):
    with open(RUTA_CONFIG, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=4)

def obtener_impresora():
    cfg = cargar_config()
    return cfg.get("impresora", "ZDesigner ZM400 300 dpi (ZPL)")

def guardar_impresora(nombre_impresora):
    cfg = cargar_config()
    cfg["impresora"] = nombre_impresora
    guardar_config(cfg)

# --- LOAD PARTS ---
def cargar_partes():
    cfg = cargar_config()
    partes = {}
    for parte, datos in cfg["partes"].items():
        cliente = datos.get("cliente", "")
        posicion = datos.get("posicion", "inicio")
        complemento = datos.get("complemento", "")
        
        partes[parte] = (complemento, posicion, cliente)
    return partes

partes_data = cargar_partes()

def obtener_complemento(parte_sel):
    datos = partes_data.get(parte_sel, ("", "inicio", ""))
    return datos[0], datos[1]

def obtener_cliente_de_parte(parte_sel):
    datos = partes_data.get(parte_sel, ("", "", ""))
    return datos[2] if len(datos) > 2 else ""

def obtener_partes_por_cliente(cliente):
    return [p for p, datos in partes_data.items() if len(datos) > 2 and datos[2] == cliente]

def obtener_destinos_de_cliente(cliente):
    cfg = cargar_config()
    if cliente in cfg["clientes"]:
        return cfg["clientes"][cliente].get("destinos", [])
    return []

# --- DATABASE ---
def inicializar_db():
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_entrada TEXT,
            fecha_salida TEXT,
            area TEXT,
            numero_parte TEXT,
            lote_interno TEXT,
            linea TEXT,
            no_piezas TEXT,
            piezas_por_bolsa TEXT,
            no_bolsas TEXT,
            lote_juliano TEXT,
            fifo TEXT,
            sobrante TEXT,
            serial_inicial INTEGER,
            serial_final INTEGER,
            pedido_id INTEGER,
            pedido_fecha TEXT,
            scrap TEXT,
            destino TEXT
        )
    ''')
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS pedidos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_pedido TEXT,
            cliente TEXT,
            destino TEXT,
            estado TEXT DEFAULT 'activo'
        )
    ''')
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS pedido_detalles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pedido_id INTEGER,
            numero_parte TEXT,
            cantidad_meta INTEGER,
            acumulado INTEGER DEFAULT 0,
            ultimo_serial INTEGER DEFAULT 0,
            fecha_salida_inicial TEXT,
            FOREIGN KEY(pedido_id) REFERENCES pedidos(id)
        )
    ''')
    
    conn.commit()
    conn.close()

def migrar_db():
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    
    c.execute("PRAGMA table_info(pedidos)")
    columnas_ped = [col[1] for col in c.fetchall()]
    if "destino" not in columnas_ped:
        try:
            c.execute('ALTER TABLE pedidos ADD COLUMN destino TEXT')
        except:
            pass
    
    c.execute("PRAGMA table_info(registros)")
    columnas_reg = [col[1] for col in c.fetchall()]
    if "destino" not in columnas_reg:
        try:
            c.execute('ALTER TABLE registros ADD COLUMN destino TEXT')
        except:
            pass
    
    conn.commit()
    conn.close()

inicializar_db()
migrar_db()

# --- ORDER FUNCTIONS ---
def obtener_pedidos_activos():
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('SELECT id, fecha_pedido, cliente, destino FROM pedidos WHERE estado="activo" ORDER BY id DESC')
    resultados = c.fetchall()
    conn.close()
    return resultados

def obtener_todos_pedidos():
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('SELECT id, fecha_pedido, cliente, destino, estado FROM pedidos WHERE estado != "eliminado" ORDER BY id DESC')
    resultados = c.fetchall()
    conn.close()
    return resultados

def obtener_detalle_pedido(pedido_id, numero_parte):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('''
        SELECT cantidad_meta, fecha_salida_inicial
        FROM pedido_detalles 
        WHERE pedido_id=? AND numero_parte=?
    ''', (pedido_id, numero_parte))
    fila = c.fetchone()
    conn.close()
    if fila:
        return {
            "meta": fila[0], 
            "fecha_salida_inicial": fila[1]
        }
    return None

def obtener_destino_pedido(pedido_id):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('SELECT destino FROM pedidos WHERE id=?', (pedido_id,))
    fila = c.fetchone()
    conn.close()
    return fila[0] if fila else ""

def obtener_acumulado_real(pedido_id, numero_parte):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('''
        SELECT SUM(r.no_piezas) 
        FROM registros r
        LEFT JOIN pedidos p ON r.pedido_id = p.id
        WHERE r.pedido_id=? AND r.numero_parte=?
        AND (p.estado IS NULL OR p.estado != 'eliminado')
    ''', (pedido_id, numero_parte))
    resultado = c.fetchone()[0]
    conn.close()
    return resultado if resultado else 0

def obtener_ultimo_serial(pedido_id, numero_parte):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('''
        SELECT r.serial_final 
        FROM registros r
        LEFT JOIN pedidos p ON r.pedido_id = p.id
        WHERE r.pedido_id=? AND r.numero_parte=? 
        AND (p.estado IS NULL OR p.estado != 'eliminado')
        ORDER BY CAST(r.serial_final AS INTEGER) DESC
        LIMIT 1
    ''', (pedido_id, numero_parte))
    fila = c.fetchone()
    conn.close()
    return int(fila[0]) if fila else 0

def actualizar_acumulado_pedido(pedido_id, numero_parte):
    acumulado = obtener_acumulado_real(pedido_id, numero_parte)
    ultimo_serial = obtener_ultimo_serial(pedido_id, numero_parte)
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('''
        UPDATE pedido_detalles 
        SET acumulado=?, ultimo_serial=?
        WHERE pedido_id=? AND numero_parte=?
    ''', (acumulado, ultimo_serial, pedido_id, numero_parte))
    conn.commit()
    
    c.execute('SELECT SUM(cantidad_meta - acumulado) FROM pedido_detalles WHERE pedido_id=? AND cantidad_meta > acumulado', (pedido_id,))
    resto = c.fetchone()[0]
    if resto == 0 or resto is None:
        c.execute('UPDATE pedidos SET estado="cerrado" WHERE id=?', (pedido_id,))
        conn.commit()
    conn.close()

def actualizar_acumulado_con_nuevo_registro(pedido_id, numero_parte, nuevo_acumulado, ultimo_serial, fecha_salida=None):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    if fecha_salida:
        c.execute('''
            UPDATE pedido_detalles 
            SET acumulado=?, ultimo_serial=?, fecha_salida_inicial=COALESCE(fecha_salida_inicial, ?)
            WHERE pedido_id=? AND numero_parte=?
        ''', (nuevo_acumulado, ultimo_serial, fecha_salida, pedido_id, numero_parte))
    else:
        c.execute('''
            UPDATE pedido_detalles 
            SET acumulado=?, ultimo_serial=?
            WHERE pedido_id=? AND numero_parte=?
        ''', (nuevo_acumulado, ultimo_serial, pedido_id, numero_parte))
    conn.commit()
    
    c.execute('SELECT SUM(cantidad_meta - acumulado) FROM pedido_detalles WHERE pedido_id=? AND cantidad_meta > acumulado', (pedido_id,))
    resto = c.fetchone()[0]
    if resto == 0 or resto is None:
        c.execute('UPDATE pedidos SET estado="cerrado" WHERE id=?', (pedido_id,))
        conn.commit()
    conn.close()

def obtener_sobrante_pendiente(parte, pedido_id_actual):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('''
        SELECT r.sobrante, r.lote_interno, r.fecha_entrada, r.area, r.linea, 
               r.piezas_por_bolsa, r.fifo, r.lote_juliano, r.fecha_salida
        FROM registros r
        LEFT JOIN pedidos p ON r.pedido_id = p.id
        WHERE r.numero_parte=? 
        AND (p.estado IS NULL OR p.estado != 'eliminado')
        ORDER BY r.id DESC
        LIMIT 1
    ''', (parte,))
    fila = c.fetchone()
    conn.close()
    if fila:
        try:
            sobrante_val = int(float(str(fila[0]).strip()))
        except:
            sobrante_val = 0
        if sobrante_val == 0:
            return None
        return {
            "sobrante": sobrante_val,
            "lote_interno": fila[1],
            "fecha_entrada": fila[2],
            "area": fila[3],
            "linea": fila[4],
            "piezas_por_bolsa": fila[5],
            "fifo": fila[6],
            "lote_juliano": fila[7],
            "fecha_salida": fila[8]
        }
    return None

def cancelar_pedido(pedido_id):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('UPDATE pedidos SET estado="cancelado" WHERE id=?', (pedido_id,))
    conn.commit()
    conn.close()

def reactivar_pedido(pedido_id):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('UPDATE pedidos SET estado="activo" WHERE id=?', (pedido_id,))
    conn.commit()
    conn.close()

def eliminar_pedido(pedido_id):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('UPDATE pedidos SET estado="eliminado" WHERE id=?', (pedido_id,))
    c.execute('UPDATE registros SET pedido_id=NULL WHERE pedido_id=?', (pedido_id,))
    conn.commit()
    conn.close()

def actualizar_fecha_salida_pedido(pedido_id, nueva_fecha):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('UPDATE pedido_detalles SET fecha_salida_inicial=? WHERE pedido_id=?', (nueva_fecha, pedido_id))
    c.execute('UPDATE registros SET fecha_salida=? WHERE pedido_id=?', (nueva_fecha, pedido_id))
    conn.commit()
    conn.close()

# --- FORMAT FUNCTIONS ---
def formatear_fecha(event):
    texto = event.widget.get().replace("/", "")
    if not texto.isdigit(): 
        return
    texto = texto[:8]
    if len(texto) >= 5:
        nuevo = "{0}/{1}/{2}".format(texto[:2], texto[2:4], texto[4:])
    elif len(texto) >= 3:
        nuevo = "{0}/{1}".format(texto[:2], texto[2:])
    else:
        nuevo = texto
    event.widget.delete(0, tk.END)
    event.widget.insert(0, nuevo)

def focus_siguiente(event, proximo_widget):
    proximo_widget.focus_set()
    if isinstance(proximo_widget, ttk.Combobox):
        proximo_widget.event_generate('<Down>')
    return "break"

# --- ZPL AND ZEBRA ---
def generar_codigo_zpl(data, serial):
    zpl = "^XA\n"
    zpl += "^PW1263\n"
    zpl += "^LL944\n"
    zpl += "^FX\n"
    zpl += "^FO74,59^GB178,178,178^FS\n"
    zpl += "^FO96,81^GC59,59,W^FS\n"
    zpl += "^FO111,96^GC30,30,B^FS\n"
    zpl += "^FO141,111^GC237,37,W^FS\n"
    zpl += "^CF0,37\n"
    zpl += "^FO266,81^FDA^FS\n"
    zpl += "^FO266,118^FDS^FS\n"
    zpl += "^FO266,155^FDS^FS\n"
    zpl += "^FO266,192^FDA^FS\n"
    zpl += "^FO333,126^FD{0}^FS\n".format("Company Name")
    zpl += "^FO474,178^FD{1}^FS\n".format("Subsidiary Name")
    zpl += "^FX --- DATE ---\n"
    zpl += "^FO503,281^FDDATE: {0}^FS\n".format(data['Fecha salida'])
    zpl += "^FX --- ROW 1: PART NUM AND LOT BOX ---\n"
    zpl += "^FO89,400^A0N,37,37^FDPART NUM. {0}^FS\n".format(data['Numero de parte'])
    zpl += "^FO89,459^BY3,3,89^BCN,89,Y,N,N^FDP{0}^FS\n".format(data['Numero de parte'])
    zpl += "^FO666,400^A0N,37,37^FDLOT BOX: {0}^FS\n".format(data['Lote interno'])
    zpl += "^FO666,459^BY3,3,89^BCN,89,Y,N,N^FD1T{0}^FS\n".format(data['Lote interno'])
    zpl += "^FX --- ROW 2: QUANTITY AND SERIAL ---\n"
    zpl += "^FO89,666^A0N,37,37^FDQUANTITY: {0}^FS\n".format(data['Piezas por bolsa'])
    zpl += "^FO89,725^BY3,3,89^BCN,89,Y,N,N^FDQ{0}^FS\n".format(data['Piezas por bolsa'])
    zpl += "^FO666,725^A0N,37,37^FDSERIAL: {0}^FS\n".format(serial)
    zpl += "^FO666,799^A0N,37,37^FDAJuliano LOT NUM. {0}^FS\n".format(data['Lote juliano'])
    zpl += "^XZ"
    return zpl

def enviar_a_zebra(codigo_zpl):
    nombre_impresora = obtener_impresora()
    try:
        hPrinter = win32print.OpenPrinter(nombre_impresora)
        try:
            win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta Zebra", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, codigo_zpl.encode('utf-8'))
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            return True
        finally:
            win32print.ClosePrinter(hPrinter)
    except Exception as e:
        messagebox.showerror("Error Zebra", "No se pudo imprimir en '{0}'\n\n{1}\n\nVerifique que la impresora esta conectada y encendida.".format(nombre_impresora, str(e)))
        return False

# --- LABEL IMAGE ---
def generar_imagen_etiqueta(data, serial):
    ruta_plantilla = os.path.join(base_dir, "plantilla_vacia.png")
    if os.path.exists(ruta_plantilla):
        etiqueta_visual = Image.open(ruta_plantilla).convert("RGB")
    else:
        etiqueta_visual = Image.new('RGB', (800, 650), 'white')

    draw = ImageDraw.Draw(etiqueta_visual)
    try:
        fuente = ImageFont.truetype("segoeui.ttf", 25)
    except:
        try:
            fuente = ImageFont.truetype("arial.ttf", 25)
        except:
            fuente = ImageFont.load_default()

    def pegar_barcode(contenido, x, y, ancho=0.2):
        buffer = io.BytesIO()
        COD = barcode.get_barcode_class('code128')
        writer_options = {
            'module_width': ancho,
            'module_height': 4.0,
            'font_size': 10,
            'text_distance': 3.0,
            'quiet_zone': 1.0
        }
        instancia = COD(str(contenido), writer=ImageWriter())
        instancia.write(buffer, options=writer_options)
        buffer.seek(0)
        img_bar = Image.open(buffer)
        etiqueta_visual.paste(img_bar, (x, y))

    draw.text((410, 185), str(data['Fecha salida']), fill="black", font=fuente)
    draw.text((190, 265), str(data['Numero de parte']), fill="black", font=fuente)
    pegar_barcode("P" + str(data['Numero de parte']), 45, 300, ancho=0.22)
    draw.text((560, 265), str(data['Lote interno']), fill="black", font=fuente)
    pegar_barcode("1T" + str(data['Lote interno']), 430, 300, ancho=0.15)
    draw.text((190, 445), str(data['Piezas por bolsa']), fill="black", font=fuente)
    pegar_barcode("Q" + str(data['Piezas por bolsa']), 45, 480, ancho=0.22)
    draw.text((550, 485), str(serial), fill="black", font=fuente)
    draw.text((610, 535), str(data['Lote juliano']), fill="black", font=fuente)

    return etiqueta_visual

# --- PRODUCTION FUNCTIONS ---
def detectar_duplicado(data, pedido_id):
    try:
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        c.execute('''
            SELECT id FROM registros
            WHERE numero_parte=? AND lote_juliano=? 
            AND CAST(serial_inicial AS INTEGER)=? AND CAST(serial_final AS INTEGER)=?
            AND (pedido_id=? OR (pedido_id IS NULL AND ? IS NULL))
        ''', (
            str(data["Numero de parte"]).strip(),
            str(data["Lote juliano"]).strip(),
            int(data["Serial inicial"]),
            int(data["Serial final"]),
            pedido_id if pedido_id else 0,
            pedido_id if pedido_id else 0
        ))
        resultado = c.fetchone()
        conn.close()
        return resultado is not None
    except:
        return False

def guardar_db(data):
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute('''
        INSERT INTO registros (
            fecha_entrada, fecha_salida, area, numero_parte,
            lote_interno, linea, no_piezas, piezas_por_bolsa,
            no_bolsas, lote_juliano, fifo, sobrante,
            serial_inicial, serial_final, pedido_id, pedido_fecha, scrap,
            destino
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ''', (
        data["Fecha entrada"], data["Fecha salida"], data["Area"],
        data["Numero de parte"], data["Lote interno"], data["Linea"],
        data["No. piezas"], data["Piezas por bolsa"], data["No. bolsas"],
        data["Lote juliano"], data["FIFO"], data["Sobrante"],
        data["Serial inicial"], data["Serial final"],
        data.get("Pedido ID"), data.get("Fecha pedido"), data.get("Scrap", ""),
        data.get("Destino", "")
    ))
    conn.commit()
    conn.close()

# --- EXCESS ADJUSTMENT ---
def ajustar_excedente(data, meta, acumulado_actual):
    piezas_producidas = int(data["No. piezas"])
    piezas_por_bolsa_original = int(data["Piezas por bolsa"])
    serial_inicial = int(data["Serial inicial"])
    
    faltante = meta - acumulado_actual
    piezas_para_pedido = min(piezas_producidas, faltante)
    sobrante = piezas_producidas - piezas_para_pedido
    
    if piezas_para_pedido % piezas_por_bolsa_original == 0:
        bolsas = piezas_para_pedido // piezas_por_bolsa_original
        piezas_ultima_bolsa = piezas_por_bolsa_original
    else:
        bolsas = (piezas_para_pedido // piezas_por_bolsa_original) + 1
        piezas_ultima_bolsa = piezas_para_pedido % piezas_por_bolsa_original
    
    serial_final = serial_inicial + bolsas - 1
    
    data["No. piezas"] = str(piezas_para_pedido)
    data["Piezas por bolsa"] = str(piezas_por_bolsa_original)
    data["No. bolsas"] = str(bolsas)
    data["Serial final"] = str(serial_final)
    data["Sobrante"] = str(sobrante)
    
    return data, sobrante, piezas_por_bolsa_original, piezas_ultima_bolsa, bolsas

# --- PRODUCTION FORM ---
class FormularioProduccion:
    def __init__(self, parent, modo="manual", pedido_id=None, parte=None, pedido_fecha=None, fecha_salida_inicial=None, destino=None):
        self.parent = parent
        self.modo = modo
        self.pedido_id = pedido_id
        self.parte_fija = parte
        self.pedido_fecha = pedido_fecha
        self.fecha_salida_inicial = fecha_salida_inicial
        self.destino_fijo = destino
        self.ventana = None
        self.entries = {}
        self.combo_parte = None
        self.sobrante_ya_cargado = False
        self.pedido_completado = False
        self.fecha_sobrante_original = None
        self.fecha_sobrante_usada = False
        
    def abrir(self):
        self.ventana = tk.Toplevel(self.parent)
        self.ventana.configure(bg=COLOR_FONDO)
        if self.modo == "pedido":
            self.ventana.title("PRODUCTION - Control System")
        else:
            self.ventana.title("MANUAL REGISTRATION - Control System")
        self.ventana.geometry("1350x600")
        self.ventana.attributes("-topmost", True)
        
        main_frame = tk.Frame(self.ventana, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
        main_frame.pack(padx=15, pady=15, fill="both", expand=True)
        
        titulo = tk.Label(main_frame, text="PRODUCTION REGISTRATION", font=FUENTE_TITULO, 
                         bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
        titulo.pack(pady=10)
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=10, pady=5)
        
        frame = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
        frame.pack(padx=10, pady=10)
        
        if self.modo == "pedido":
            campos = [
                ("Fecha entrada", 0), ("Area", 1), ("Numero de parte", 2),
                ("Lote interno", 3), ("Fecha salida", 4), ("Linea", 5),
                ("No. piezas", 6), ("Piezas/bolsa", 7), ("No. bolsas", 8),
                ("Lote juliano", 9), ("Caja inicial", 10), ("FIFO", 11), ("Scrap", 12)
            ]
        else:
            campos = [
                ("Fecha entrada", 0), ("Area", 1), ("Numero de parte", 2),
                ("Lote interno", 3), ("Fecha salida", 4), ("Linea", 5),
                ("No. piezas", 6), ("Piezas/bolsa", 7), ("No. bolsas", 8),
                ("Lote juliano", 9), ("Caja inicial", 10), ("FIFO", 11), ("Sobrante", 12), ("Scrap", 13)
            ]
        
        for texto, col in campos:
            lbl = tk.Label(frame, text=texto, font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, 
                          fg=COLOR_LABEL)
            lbl.grid(row=0, column=col, sticky="w", padx=2, pady=2)
            e = tk.Entry(frame, width=14, font=FUENTE_CAMPO, bg=COLOR_ENTRY, 
                        relief=tk.SUNKEN, bd=1)
            e.grid(row=1, column=col, padx=4, pady=2)
            self.entries[texto] = e
        
        self.entries["No. bolsas"].config(state="readonly")
        
        if "Scrap" in self.entries:
            self.entries["Scrap"].insert(0, "0")
        
        self.entries["Fecha entrada"].bind("<KeyRelease>", formatear_fecha)
        
        def on_fecha_salida_keyrelease(event):
            formatear_fecha(event)
            self.generar_lote_base()
        
        self.entries["Fecha salida"].bind("<KeyRelease>", on_fecha_salida_keyrelease)
        
        self.entries["No. piezas"].bind("<KeyRelease>", self.calcular_bolsas)
        self.entries["Piezas/bolsa"].bind("<KeyRelease>", self.calcular_bolsas)
        
        if self.modo == "pedido" and self.parte_fija:
            self.combo_parte = ttk.Combobox(frame, values=[self.parte_fija], width=12, state="readonly")
            self.combo_parte.grid(row=1, column=2)
            self.combo_parte.set(self.parte_fija)
            self.entries["Numero de parte"] = self.combo_parte
            
            if self.destino_fijo:
                pass
        else:
            self.combo_parte = ttk.Combobox(frame, values=sorted(list(partes_data.keys())), width=12, state="readonly")
            self.combo_parte.grid(row=1, column=2)
            self.entries["Numero de parte"] = self.combo_parte
            self.combo_parte.bind("<<ComboboxSelected>>", self.al_seleccionar_parte)
        
        if self.modo == "pedido":
            self.label_progreso = tk.Label(frame, text="", font=FUENTE_LABEL, 
                                          bg=COLOR_FONDO_FRAME, fg="#27ae60")
            self.label_progreso.grid(row=2, column=0, columnspan=13, pady=8, sticky="w")
            self.actualizar_info_pedido()
            
            if self.destino_fijo:
                lbl_destino = tk.Label(frame, text="Destino: {0}".format(self.destino_fijo), 
                                      font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, fg="#2980b9")
                lbl_destino.grid(row=3, column=0, columnspan=13, pady=2, sticky="w")
        
        if self.modo == "pedido":
            entry_list = [self.entries["Fecha entrada"], self.entries["Area"], self.combo_parte,
                          self.entries["Lote interno"], self.entries["Fecha salida"], self.entries["Linea"],
                          self.entries["No. piezas"], self.entries["Piezas/bolsa"],
                          self.entries["Lote juliano"], self.entries["Caja inicial"],
                          self.entries["FIFO"], self.entries["Scrap"]]
        else:
            entry_list = [self.entries["Fecha entrada"], self.entries["Area"], self.combo_parte,
                          self.entries["Lote interno"], self.entries["Fecha salida"], self.entries["Linea"],
                          self.entries["No. piezas"], self.entries["Piezas/bolsa"],
                          self.entries["Lote juliano"], self.entries["Caja inicial"],
                          self.entries["FIFO"], self.entries["Sobrante"], self.entries["Scrap"]]
        
        for i in range(len(entry_list) - 1):
            if isinstance(entry_list[i], tk.Entry):
                entry_list[i].bind("<Return>", lambda e, prox=entry_list[i+1]: focus_siguiente(e, prox))
            elif isinstance(entry_list[i], ttk.Combobox):
                entry_list[i].bind("<Return>", lambda e, prox=entry_list[i+1]: focus_siguiente(e, prox))
        
        if self.modo == "pedido":
            self.entries["Fecha salida"].config(state="normal")
            self.entries["Fecha salida"].delete(0, tk.END)
            
            if self.fecha_salida_inicial:
                self.entries["Fecha salida"].insert(0, self.fecha_salida_inicial)
            
            self.entries["Caja inicial"].config(state="readonly")
            self.cargar_sobrante_si_existe()
        
        frame_botones = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
        frame_botones.pack(pady=15)
        
        btn_ver = tk.Button(frame_botones, text="VIEW LABELS", command=self.visor_secuencial,
                           bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                           relief=tk.RAISED, bd=1, padx=20, pady=5)
        btn_ver.pack(side="left", padx=10)
        
        btn_guardar = tk.Button(frame_botones, text="SAVE DATA", command=self.guardar_registro,
                               bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                               relief=tk.RAISED, bd=1, padx=20, pady=5)
        btn_guardar.pack(side="left", padx=10)
        
        btn_limpiar = tk.Button(frame_botones, text="CLEAR", command=self.limpiar_campos,
                               bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                               relief=tk.RAISED, bd=1, padx=20, pady=5)
        btn_limpiar.pack(side="left", padx=10)
        
        btn_cerrar = tk.Button(frame_botones, text="CLOSE", command=self.ventana.destroy,
                              bg="#c0392b", fg="white", font=FUENTE_BOTON,
                              relief=tk.RAISED, bd=1, padx=20, pady=5)
        btn_cerrar.pack(side="left", padx=10)
        
        self.entries["Fecha entrada"].focus_set()

    def calcular_bolsas(self, *args):
        try:
            p = int(self.entries["No. piezas"].get())
            por_b = int(self.entries["Piezas/bolsa"].get())
            resultado = (p + por_b - 1) // por_b
            self.entries["No. bolsas"].config(state="normal")
            self.entries["No. bolsas"].delete(0, tk.END)
            self.entries["No. bolsas"].insert(0, str(resultado))
            self.entries["No. bolsas"].config(state="readonly")
        except:
            pass
    
    def generar_lote_base(self, *args):
        try:
            parte = self.entries["Numero de parte"].get() if isinstance(self.entries["Numero de parte"], tk.Entry) else self.combo_parte.get()
            fecha_str = self.entries["Fecha salida"].get()
            comp, posicion = obtener_complemento(parte)
            
            if len(fecha_str) < 10:
                if comp and posicion == "inicio":
                    self.entries["Lote juliano"].delete(0, tk.END)
                    self.entries["Lote juliano"].insert(0, comp + " ")
                return
            
            try:
                fecha = datetime.strptime(fecha_str, "%d/%m/%Y")
            except:
                return
                
            dia, anio = fecha.strftime("%j"), fecha.strftime("%y")
            lote_num = "{0}{1}1".format(dia, anio)
            if comp and posicion == "final":
                lote_final = "{0}{1}".format(lote_num, comp)
            elif comp:
                lote_final = "{0} {1}".format(comp, lote_num)
            else:
                lote_final = lote_num
            self.entries["Lote juliano"].delete(0, tk.END)
            self.entries["Lote juliano"].insert(0, lote_final)
        except:
            pass
    
    def al_seleccionar_parte(self, event):
        self.generar_lote_base()
        if self.modo == "pedido":
            self.actualizar_info_pedido()
    
    def actualizar_info_pedido(self):
        parte = self.combo_parte.get()
        if parte and self.pedido_id:
            detalle = obtener_detalle_pedido(self.pedido_id, parte)
            if detalle:
                acumulado = obtener_acumulado_real(self.pedido_id, parte)
                faltante = max(0, detalle["meta"] - acumulado)
                ultimo_serial = obtener_ultimo_serial(self.pedido_id, parte)
                
                self.label_progreso.config(
                    text="Order: Target {0} | Produced: {1} | Remaining: {2} | Next serial: {3}".format(
                        detalle["meta"], acumulado, faltante, ultimo_serial + 1
                    ),
                    fg="green" if faltante > 0 else "blue"
                )
                
                if faltante <= 0:
                    self.pedido_completado = True
                    for campo in ["No. piezas", "Piezas/bolsa", "Lote interno", "Linea", "Area", "Fecha entrada", "FIFO", "Lote juliano", "Scrap"]:
                        if campo in self.entries:
                            self.entries[campo].config(state="disabled")
                    if self.modo == "manual":
                        self.entries["Sobrante"].config(state="disabled")
                    messagebox.showinfo("Order completed", "This order has reached its target.\nNo more parts can be produced.", parent=self.ventana)
                else:
                    self.pedido_completado = False
                    for campo in ["No. piezas", "Piezas/bolsa", "Lote interno", "Linea", "Area", "Fecha entrada", "FIFO", "Lote juliano", "Scrap"]:
                        if campo in self.entries:
                            self.entries[campo].config(state="normal")
                    if self.modo == "manual":
                        self.entries["Sobrante"].config(state="normal")
                
                self.entries["Caja inicial"].config(state="normal")
                self.entries["Caja inicial"].delete(0, tk.END)
                self.entries["Caja inicial"].insert(0, str(ultimo_serial + 1))
                self.entries["Caja inicial"].config(state="readonly")
                
                if detalle.get("fecha_salida_inicial"):
                    self.entries["Fecha salida"].config(state="readonly")
                    self.entries["Fecha salida"].delete(0, tk.END)
                    self.entries["Fecha salida"].insert(0, detalle["fecha_salida_inicial"])
                else:
                    self.entries["Fecha salida"].config(state="normal")
            else:
                self.label_progreso.config(text="Error: Part not found in this order", fg="red")
    
    def cargar_sobrante_si_existe(self):
        parte = self.combo_parte.get()
        if self.sobrante_ya_cargado:
            return
        
        if parte and self.pedido_id:
            sobrante_data = obtener_sobrante_pendiente(parte, self.pedido_id)
            if sobrante_data and sobrante_data["sobrante"] > 0:
                self.fecha_sobrante_original = sobrante_data["fecha_salida"]
                resp = messagebox.askyesno(
                    "Remainder detected",
                    "Remainder from previous order found: {0} pieces.\n\nLoad remainder data (internal lot, area, etc.)?".format(
                        sobrante_data["sobrante"]
                    ),
                    parent=self.ventana
                )
                if resp:
                    self.entries["Fecha entrada"].delete(0, tk.END)
                    self.entries["Fecha entrada"].insert(0, sobrante_data["fecha_entrada"])
                    self.entries["Area"].delete(0, tk.END)
                    self.entries["Area"].insert(0, sobrante_data["area"])
                    self.entries["Lote interno"].delete(0, tk.END)
                    self.entries["Lote interno"].insert(0, sobrante_data["lote_interno"])
                    self.entries["Linea"].delete(0, tk.END)
                    self.entries["Linea"].insert(0, sobrante_data["linea"])
                    self.entries["FIFO"].delete(0, tk.END)
                    self.entries["FIFO"].insert(0, sobrante_data["fifo"])
                    self.entries["Piezas/bolsa"].delete(0, tk.END)
                    self.entries["Piezas/bolsa"].insert(0, sobrante_data["piezas_por_bolsa"])
                    self.entries["No. piezas"].delete(0, tk.END)
                    self.entries["No. piezas"].insert(0, str(sobrante_data["sobrante"]))
                    
                    self.fecha_sobrante_usada = False
                    
                    if "Scrap" in self.entries:
                        self.entries["Scrap"].delete(0, tk.END)
                        self.entries["Scrap"].insert(0, "0")
                    
                    self.calcular_bolsas()
                    self.sobrante_ya_cargado = True
                    
                    self.ventana.after(100, lambda: self.visor_secuencial())
    
    def obtener_datos(self):
        i = self.entries["Caja inicial"].get()
        b = self.entries["No. bolsas"].get()
        try:
            serial_ini = int(i)
            serial_fin = serial_ini + int(b) - 1
        except (ValueError, TypeError):
            serial_ini = 0
            serial_fin = 0
        
        parte = self.entries["Numero de parte"].get() if isinstance(self.entries["Numero de parte"], tk.Entry) else self.combo_parte.get()
        
        scrap_valor = "0"
        if "Scrap" in self.entries:
            scrap_temp = self.entries["Scrap"].get().strip()
            scrap_valor = scrap_temp if scrap_temp else "0"
        
        datos = {
            "Fecha entrada": self.entries["Fecha entrada"].get(),
            "Fecha salida": self.entries["Fecha salida"].get(),
            "Area": self.entries["Area"].get(),
            "Numero de parte": parte,
            "Lote interno": self.entries["Lote interno"].get(),
            "Linea": self.entries["Linea"].get(),
            "No. piezas": self.entries["No. piezas"].get(),
            "Piezas por bolsa": self.entries["Piezas/bolsa"].get(),
            "No. bolsas": self.entries["No. bolsas"].get(),
            "Lote juliano": self.entries["Lote juliano"].get(),
            "FIFO": self.entries["FIFO"].get(),
            "Sobrante": "0",
            "Serial inicial": str(serial_ini),
            "Serial final": str(serial_fin),
            "Scrap": scrap_valor,
            "Destino": self.destino_fijo if self.destino_fijo else ""
        }
        
        if self.modo == "manual":
            datos["Sobrante"] = self.entries["Sobrante"].get()
            datos["Pedido ID"] = None
            datos["Fecha pedido"] = None
        else:
            datos["Pedido ID"] = self.pedido_id
            datos["Fecha pedido"] = self.pedido_fecha
        
        return datos
    
    def validar_campos(self, data):
        for campo_fecha in ["Fecha entrada", "Fecha salida"]:
            valor = str(data.get(campo_fecha, "")).strip()
            if not valor:
                messagebox.showwarning("Missing data", "Field '{0}' is required.".format(campo_fecha), parent=self.ventana)
                return False
            try:
                datetime.strptime(valor, "%d/%m/%Y")
            except ValueError:
                messagebox.showwarning("Invalid date", "Field '{0}' invalid format dd/mm/yyyy.".format(campo_fecha), parent=self.ventana)
                return False
        
        for k, v in data.items():
            if k.startswith("_"): 
                continue
            if not str(v).strip():
                if k == "Pedido ID" or k == "Fecha pedido" or k == "Sobrante" or k == "Scrap" or k == "Destino":
                    continue
                messagebox.showwarning("Missing data", "Field '{0}' is required.".format(k), parent=self.ventana)
                return False
        return True
    
    def guardar_registro(self):
        try:
            data = self.obtener_datos()
            if not self.validar_campos(data):
                return
            
            if self.modo == "pedido":
                if detectar_duplicado(data, self.pedido_id):
                    messagebox.showwarning("Duplicate", "A record with this Lot and Serial already exists in this order.", parent=self.ventana)
                    return
            else:
                if detectar_duplicado(data, None):
                    messagebox.showwarning("Duplicate", "A record with this Lot and Serial already exists.", parent=self.ventana)
                    return
            
            sobrante_calculado = 0
            nuevo_acumulado_guardar = None
            fecha_salida_inicial_guardar = None
            bolsas_ajustadas = None
            
            if self.modo == "pedido":
                parte = data["Numero de parte"]
                detalle = obtener_detalle_pedido(self.pedido_id, parte)
                if detalle:
                    acumulado_actual = obtener_acumulado_real(self.pedido_id, parte)
                    cantidad_producida = int(data["No. piezas"])
                    nuevo_acumulado = acumulado_actual + cantidad_producida
                    
                    if nuevo_acumulado > detalle["meta"] or cantidad_producida % int(data["Piezas por bolsa"]) != 0:
                        data_ajustada, sobrante_calculado, piezas_por_bolsa_original, piezas_ultima_bolsa, bolsas_ajustadas = ajustar_excedente(
                            data.copy(), detalle["meta"], acumulado_actual
                        )
                        data = data_ajustada
                        nuevo_acumulado_guardar = min(nuevo_acumulado, detalle["meta"])
                    else:
                        nuevo_acumulado_guardar = nuevo_acumulado
                    
                    if not detalle.get("fecha_salida_inicial"):
                        fecha_salida_inicial_guardar = data["Fecha salida"]
            
            if not messagebox.askyesno("Confirm", "Save data?", parent=self.ventana):
                if self.modo == "pedido" and self.sobrante_ya_cargado:
                    self.sobrante_ya_cargado = False
                return
            
            if bolsas_ajustadas is not None:
                self.entries["No. bolsas"].config(state="normal")
                self.entries["No. bolsas"].delete(0, tk.END)
                self.entries["No. bolsas"].insert(0, str(bolsas_ajustadas))
                self.entries["No. bolsas"].config(state="readonly")
                
                self.entries["No. piezas"].delete(0, tk.END)
                self.entries["No. piezas"].insert(0, data["No. piezas"])
            
            data["Sobrante"] = str(sobrante_calculado)
            
            guardar_db(data)
            
            if self.modo == "pedido" and nuevo_acumulado_guardar is not None:
                actualizar_acumulado_con_nuevo_registro(
                    self.pedido_id, parte, nuevo_acumulado_guardar,
                    int(data["Serial final"]), fecha_salida_inicial_guardar
                )
                if fecha_salida_inicial_guardar:
                    self.entries["Fecha salida"].config(state="readonly")
            
            if self.modo == "pedido" and sobrante_calculado > 0:
                messagebox.showinfo("Success", "Saved successfully.\n\nOrder target reached.\nRemainder: {0} pieces".format(sobrante_calculado), parent=self.ventana)
            else:
                messagebox.showinfo("Success", "Saved successfully.", parent=self.ventana)
            
            if self.modo == "pedido":
                self.entries["No. piezas"].delete(0, tk.END)
                self.entries["Piezas/bolsa"].delete(0, tk.END)
                self.entries["Lote interno"].delete(0, tk.END)
                self.entries["Linea"].delete(0, tk.END)
                self.entries["Area"].delete(0, tk.END)
                self.entries["Fecha entrada"].delete(0, tk.END)
                self.entries["FIFO"].delete(0, tk.END)
                self.entries["Lote juliano"].delete(0, tk.END)
                if "Scrap" in self.entries:
                    self.entries["Scrap"].delete(0, tk.END)
                    self.entries["Scrap"].insert(0, "0")
                self.entries["No. bolsas"].config(state="normal")
                self.entries["No. bolsas"].delete(0, tk.END)
                self.entries["No. bolsas"].config(state="readonly")
                
                self.actualizar_info_pedido()
            else:
                self.limpiar_campos()
            
            self.entries["Fecha entrada"].focus_set()
                
        except Exception as e:
            messagebox.showerror("Error", "Could not save: {0}".format(e), parent=self.ventana)
    
    def limpiar_campos(self):
        for key, entry in self.entries.items():
            if isinstance(entry, tk.Entry):
                entry.config(state="normal")
                entry.delete(0, tk.END)
        if "Scrap" in self.entries:
            self.entries["Scrap"].insert(0, "0")
        if self.modo == "manual":
            pass
        self.entries["No. bolsas"].config(state="readonly")
        if self.modo == "pedido":
            for campo in ["No. piezas", "Piezas/bolsa", "Lote interno", "Linea", "Area", "Fecha entrada", "FIFO", "Lote juliano", "Scrap"]:
                if campo in self.entries:
                    self.entries[campo].delete(0, tk.END)
            if "Scrap" in self.entries:
                self.entries["Scrap"].insert(0, "0")
            
            self.entries["Fecha salida"].config(state="normal")
            self.entries["Fecha salida"].delete(0, tk.END)
            if self.fecha_salida_inicial:
                self.entries["Fecha salida"].insert(0, self.fecha_salida_inicial)
            self.actualizar_info_pedido()
        self.entries["Fecha entrada"].focus_set()
    
    def visor_secuencial(self):
        data = self.obtener_datos()
        if not self.validar_campos(data):
            return
        
        piezas_ultima_bolsa = None
        
        if self.modo == "pedido" and self.pedido_id:
            parte = data["Numero de parte"]
            detalle = obtener_detalle_pedido(self.pedido_id, parte)
            if detalle:
                acumulado_actual = obtener_acumulado_real(self.pedido_id, parte)
                meta = detalle["meta"]
                cantidad_producida = int(data["No. piezas"])
                
                if acumulado_actual + cantidad_producida > meta or cantidad_producida % int(data["Piezas por bolsa"]) != 0:
                    data, sobrante, piezas_por_bolsa_original, piezas_ultima_bolsa, bolsas = ajustar_excedente(
                        data, meta, acumulado_actual
                    )
        
        try:
            inicio = int(data["Serial inicial"])
            total = int(data["No. bolsas"])
        except:
            messagebox.showerror("Error", "Initial box and Number of bags must be numbers.", parent=self.ventana)
            return
        
        if not messagebox.askyesno("Verification", "Data correct?", parent=self.ventana):
            return
        
        indice_actual = 0
        ventana_prev = tk.Toplevel(self.ventana)
        ventana_prev.title("Print Preview")
        ventana_prev.configure(bg=COLOR_FONDO)
        ventana_prev.attributes("-topmost", True)
        ventana_prev.resizable(False, False)

        frame_visor = tk.Frame(ventana_prev, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
        frame_visor.pack(padx=15, pady=15, fill="both", expand=True)
        
        lbl_imagen = tk.Label(frame_visor, bg="white", relief=tk.SUNKEN, bd=1)
        lbl_imagen.pack(padx=15, pady=15)
        
        lbl_info = tk.Label(frame_visor, text="", font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
        lbl_info.pack(pady=5)

        def refrescar_vista():
            nonlocal indice_actual
            serial_actual = inicio + indice_actual
            
            if piezas_ultima_bolsa is not None and indice_actual == total - 1:
                data_temp = data.copy()
                data_temp["Piezas por bolsa"] = str(piezas_ultima_bolsa)
                img_pil = generar_imagen_etiqueta(data_temp, serial_actual)
            else:
                img_pil = generar_imagen_etiqueta(data, serial_actual)

            nw, nh = int(img_pil.width * 0.6), int(img_pil.height * 0.6)
            try: 
                resample_filter = Image.ANTIALIAS
            except AttributeError: 
                resample_filter = Image.BILINEAR

            img_resizada = img_pil.resize((nw, nh), resample_filter)
            img_tk = ImageTk.PhotoImage(img_resizada)
            lbl_imagen.config(image=img_tk)
            lbl_imagen.image = img_tk

            lbl_info.config(text="Label {0} of {1} | Serial: {2}".format(indice_actual + 1, total, serial_actual))

            if indice_actual + 1 >= total:
                btn_imprimir.config(text="PRINT AND FINISH")
            else:
                btn_imprimir.config(text="PRINT LABEL {0}".format(serial_actual))

            ventana_prev.after(500, lambda: btn_imprimir.config(state="normal"))
            btn_imprimir.focus_set()

        def ejecutar_impresion():
            nonlocal indice_actual
            if btn_imprimir['state'] == 'disabled': 
                return
            btn_imprimir.config(state="disabled")

            serial_a_imprimir = inicio + indice_actual
            
            if piezas_ultima_bolsa is not None and indice_actual == total - 1:
                data_temp = data.copy()
                data_temp["Piezas por bolsa"] = str(piezas_ultima_bolsa)
                zpl_final = generar_codigo_zpl(data_temp, serial_a_imprimir)
            else:
                zpl_final = generar_codigo_zpl(data, serial_a_imprimir)

            if not enviar_a_zebra(zpl_final):
                btn_imprimir.config(state="normal")
                return

            if indice_actual + 1 >= total:
                ventana_prev.destroy()
                return

            indice_actual += 1
            refrescar_vista()

        def saltar_etiqueta():
            nonlocal indice_actual
            if indice_actual + 1 >= total:
                ventana_prev.destroy()
            else:
                indice_actual += 1
                refrescar_vista()

        frame_btns = tk.Frame(frame_visor, bg=COLOR_FONDO_FRAME)
        frame_btns.pack(pady=10)

        btn_imprimir = tk.Button(frame_btns, text="PRINT", command=ejecutar_impresion,
                               bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                               relief=tk.RAISED, bd=1, padx=15, pady=3)
        btn_imprimir.pack(side="left", padx=8)

        btn_saltar = tk.Button(frame_btns, text="SKIP SERIAL", command=saltar_etiqueta,
                              bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                              relief=tk.RAISED, bd=1, padx=15, pady=3)
        btn_saltar.pack(side="left", padx=8)

        btn_cancelar = tk.Button(frame_btns, text="Cancel", command=ventana_prev.destroy,
                                bg="#c0392b", fg="white", font=FUENTE_BOTON,
                                relief=tk.RAISED, bd=1, padx=15, pady=3)
        btn_cancelar.pack(side="left", padx=8)

        ventana_prev.bind("<Return>", lambda e: ejecutar_impresion())
        refrescar_vista()

# --- ORDERS VIEW WINDOW ---
def ventana_ver_pedidos(parent):
    ventana = tk.Toplevel(parent)
    ventana.title("View Orders")
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("1400x800")
    ventana.attributes("-topmost", True)
    
    frame_superior = tk.Frame(ventana, bg=COLOR_FONDO)
    frame_superior.pack(fill="x", padx=10, pady=5)
    
    mostrar_cancelados_var = tk.IntVar(value=0)
    chk_mostrar_cancelados = tk.Checkbutton(frame_superior, text="Show cancelled orders", 
                                            variable=mostrar_cancelados_var, bg=COLOR_FONDO,
                                            font=FUENTE_LABEL, command=lambda: cargar_lista_pedidos())
    chk_mostrar_cancelados.pack(side="left", padx=10)
    
    frame_filtros = tk.Frame(ventana, bg=COLOR_FONDO)
    frame_filtros.pack(fill="x", padx=10, pady=5)
    
    tk.Label(frame_filtros, text="Part:", font=FUENTE_LABEL, bg=COLOR_FONDO).pack(side="left", padx=5)
    entry_filtro_parte = tk.Entry(frame_filtros, width=20, font=FUENTE_CAMPO, bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_filtro_parte.pack(side="left", padx=5)
    
    tk.Label(frame_filtros, text="Order date:", font=FUENTE_LABEL, bg=COLOR_FONDO).pack(side="left", padx=5)
    entry_filtro_fecha_pedido = tk.Entry(frame_filtros, width=12, font=FUENTE_CAMPO, bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_filtro_fecha_pedido.pack(side="left", padx=5)
    entry_filtro_fecha_pedido.bind("<KeyRelease>", formatear_fecha)
    
    tk.Label(frame_filtros, text="Exit date:", font=FUENTE_LABEL, bg=COLOR_FONDO).pack(side="left", padx=5)
    entry_filtro_fecha_salida = tk.Entry(frame_filtros, width=12, font=FUENTE_CAMPO, bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_filtro_fecha_salida.pack(side="left", padx=5)
    entry_filtro_fecha_salida.bind("<KeyRelease>", formatear_fecha)
    
    notebook_main = ttk.Notebook(ventana)
    notebook_main.pack(fill="both", expand=True, padx=10, pady=10)

    tab_pedidos = ttk.Frame(notebook_main)
    notebook_main.add(tab_pedidos, text="Orders")

    frame_lista = tk.LabelFrame(tab_pedidos, text=" ORDER LIST ", 
                                font=FUENTE_SUBTITULO, bg=COLOR_FONDO_FRAME,
                                fg=COLOR_TITULO)
    frame_lista.pack(side="left", fill="both", expand=True, padx=10, pady=10)
    
    columnas_pedido = ("ID", "Order date", "Part", "Status", "Target", "Produced", "Remaining", "Exit date", "Destination")
    tree_pedidos = ttk.Treeview(frame_lista, columns=columnas_pedido, show="headings", height=20)
    tree_pedidos.heading("ID", text="ID")
    tree_pedidos.heading("Order date", text="Order date")
    tree_pedidos.heading("Part", text="Part number")
    tree_pedidos.heading("Status", text="Status")
    tree_pedidos.heading("Target", text="Target")
    tree_pedidos.heading("Produced", text="Produced")
    tree_pedidos.heading("Remaining", text="Remaining")
    tree_pedidos.heading("Exit date", text="Exit date")
    tree_pedidos.heading("Destination", text="Destination")
    tree_pedidos.column("ID", width=60)
    tree_pedidos.column("Order date", width=100)
    tree_pedidos.column("Part", width=150)
    tree_pedidos.column("Status", width=80)
    tree_pedidos.column("Target", width=70)
    tree_pedidos.column("Produced", width=70)
    tree_pedidos.column("Remaining", width=70)
    tree_pedidos.column("Exit date", width=100)
    tree_pedidos.column("Destination", width=100)
    tree_pedidos.pack(fill="both", expand=True, padx=5, pady=5)
    
    tree_pedidos.tag_configure("activo", background="#d4efdf", foreground="#1e8449")
    tree_pedidos.tag_configure("cerrado", background="#fadbd8", foreground="#922b21")
    tree_pedidos.tag_configure("cancelado", background="#ebedef", foreground="#7f8c8d")
    tree_pedidos.tag_configure("eliminado", background="#d5d8dc", foreground="#7f8c8d")
    
    frame_detalle = tk.LabelFrame(tab_pedidos, text=" ORDER DETAIL ", 
                                  font=FUENTE_SUBTITULO, bg=COLOR_FONDO_FRAME,
                                  fg=COLOR_TITULO)
    frame_detalle.pack(side="right", fill="both", expand=True, padx=10, pady=10)
    
    def cargar_lista_pedidos():
        for row in tree_pedidos.get_children():
            tree_pedidos.delete(row)
        
        parte_filtro = entry_filtro_parte.get().strip()
        fp_filtro = entry_filtro_fecha_pedido.get().strip()
        fs_filtro = entry_filtro_fecha_salida.get().strip()
        
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        
        query = '''
            SELECT DISTINCT p.id, p.fecha_pedido, d.numero_parte, p.estado, 
                   d.cantidad_meta, d.acumulado, (d.cantidad_meta - d.acumulado) as faltante,
                   (SELECT r2.fecha_salida FROM registros r2 
                    WHERE r2.pedido_id = p.id AND r2.numero_parte = d.numero_parte 
                    ORDER BY r2.id DESC LIMIT 1) as ultima_fecha_salida,
                   p.destino
            FROM pedidos p
            JOIN pedido_detalles d ON p.id = d.pedido_id
            LEFT JOIN registros r ON r.pedido_id = p.id AND r.numero_parte = d.numero_parte
            WHERE p.estado != 'eliminado'
        '''
        params = []
        
        if not mostrar_cancelados_var.get():
            query += " AND p.estado != 'cancelado'"
        if parte_filtro:
            query += " AND d.numero_parte LIKE ?"
            params.append("%{0}%".format(parte_filtro))
        if fp_filtro:
            query += " AND p.fecha_pedido = ?"
            params.append(fp_filtro)
        if fs_filtro:
            query += " AND r.fecha_salida = ?"
            params.append(fs_filtro)
        
        query += " ORDER BY p.id DESC"
        
        c.execute(query, params)
        pedidos = c.fetchall()
        conn.close()
        
        for p in pedidos:
            estado = p[3]
            if estado == "activo":
                tag = "activo"
            elif estado == "cerrado":
                tag = "cerrado"
            elif estado == "cancelado":
                tag = "cancelado"
            elif estado == "eliminado":
                tag = "eliminado"
            else:
                tag = ""
            destino_text = p[8] if len(p) > 8 and p[8] else ""
            tree_pedidos.insert("", "end", values=(p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7] if p[7] else "", destino_text), tags=(tag,))
    
    def on_select_pedido(event):
        seleccion = tree_pedidos.selection()
        if not seleccion:
            return
        valores = tree_pedidos.item(seleccion[0], "values")
        pedido_id = valores[0]
        pedido_fecha = valores[1]
        pedido_estado = valores[3]
        pedido_destino = valores[8] if len(valores) > 8 else ""
        
        for widget in frame_detalle.winfo_children():
            widget.destroy()
        
        lbl_titulo = tk.Label(frame_detalle, text="Order {0} - {1} - Destination: {2}".format(pedido_id, pedido_fecha, pedido_destino), 
                              font=FUENTE_SUBTITULO, bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
        lbl_titulo.pack(pady=10)
        
        frame_botones_accion = tk.Frame(frame_detalle, bg=COLOR_FONDO_FRAME)
        frame_botones_accion.pack(pady=5)
        
        if pedido_estado == "activo":
            def cancelar():
                if messagebox.askyesno("Cancel order", "Cancel this order? Records will be kept but the order will no longer be shown.", parent=ventana):
                    cancelar_pedido(pedido_id)
                    cargar_lista_pedidos()
                    for widget in frame_detalle.winfo_children():
                        widget.destroy()
                    tk.Label(frame_detalle, text="Order cancelled", font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, fg="red").pack(pady=20)
            btn_cancelar = tk.Button(frame_botones_accion, text="Cancel order", command=cancelar,
                                    bg="#c0392b", fg="white", font=FUENTE_BOTON,
                                    relief=tk.RAISED, bd=1, padx=15, pady=3)
            btn_cancelar.pack(side="left", padx=5)
        elif pedido_estado == "cancelado":
            def reactivar():
                if messagebox.askyesno("Reactivate order", "Reactivate this order? It will appear in the lists again.", parent=ventana):
                    reactivar_pedido(pedido_id)
                    cargar_lista_pedidos()
                    for widget in frame_detalle.winfo_children():
                        widget.destroy()
                    tk.Label(frame_detalle, text="Order reactivated", font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, fg="green").pack(pady=20)
            btn_reactivar = tk.Button(frame_botones_accion, text="Reactivate order", command=reactivar,
                                      bg="#27ae60", fg="white", font=FUENTE_BOTON,
                                      relief=tk.RAISED, bd=1, padx=15, pady=3)
            btn_reactivar.pack(side="left", padx=5)
        
        def editar_destino():
            ventana_editar_destino_pedido(ventana, pedido_id, pedido_destino, cargar_lista_pedidos)
        btn_editar_destino = tk.Button(frame_botones_accion, text="Edit destination", command=editar_destino,
                                      bg="#34495e", fg="white", font=FUENTE_BOTON,
                                      relief=tk.RAISED, bd=1, padx=15, pady=3)
        btn_editar_destino.pack(side="left", padx=5)
        
        def editar_fecha_salida():
            conn = sqlite3.connect(RUTA_DB)
            c = conn.cursor()
            c.execute("SELECT fecha_salida_inicial FROM pedido_detalles WHERE pedido_id=? LIMIT 1", (pedido_id,))
            fecha_actual = c.fetchone()
            conn.close()
            ventana_editar_fecha_salida_pedido(ventana, pedido_id, fecha_actual[0] if fecha_actual else None, cargar_lista_pedidos)
        btn_editar_fecha = tk.Button(frame_botones_accion, text="Edit exit date", command=editar_fecha_salida,
                                    bg="#34495e", fg="white", font=FUENTE_BOTON,
                                    relief=tk.RAISED, bd=1, padx=15, pady=3)
        btn_editar_fecha.pack(side="left", padx=5)
        
        if pedido_estado == "cerrado":
            def eliminar_cerrado():
                if messagebox.askyesno("Hide order", 
                    "Hide this order? Records will be kept but the order will no longer be shown.\n\nAssociated records will also be hidden.", 
                    parent=ventana):
                    eliminar_pedido(pedido_id)
                    cargar_lista_pedidos()
                    for widget in frame_detalle.winfo_children():
                        widget.destroy()
                    tk.Label(frame_detalle, text="Order hidden", font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, fg="red").pack(pady=20)
            btn_eliminar = tk.Button(frame_botones_accion, text="Hide order", command=eliminar_cerrado,
                                    bg="#c0392b", fg="white", font=FUENTE_BOTON,
                                    relief=tk.RAISED, bd=1, padx=15, pady=3)
            btn_eliminar.pack(side="left", padx=5)
        
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        c.execute('SELECT id, numero_parte, cantidad_meta, fecha_salida_inicial FROM pedido_detalles WHERE pedido_id=?', (pedido_id,))
        partes = c.fetchall()
        conn.close()
        
        if not partes:
            tk.Label(frame_detalle, text="This order has no parts assigned", font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, fg="red").pack(pady=20)
            return
        
        notebook_partes = ttk.Notebook(frame_detalle)
        notebook_partes.pack(fill="both", expand=True, padx=10, pady=10)
        
        for detalle_id, parte, meta, fecha_salida in partes:
            acumulado = obtener_acumulado_real(pedido_id, parte)
            
            tab = ttk.Frame(notebook_partes)
            notebook_partes.add(tab, text="{0} ({1}/{2})".format(parte, acumulado, meta))
            
            info_frame = tk.LabelFrame(tab, text=" INFORMATION ", font=FUENTE_LABEL, 
                                       bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
            info_frame.pack(fill="x", padx=10, pady=5)
            tk.Label(info_frame, text="Initial exit date: {0}".format(fecha_salida if fecha_salida else "Not defined"), 
                    font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME).pack(anchor="w", padx=10, pady=2)
            tk.Label(info_frame, text="Target: {0} | Produced: {1} | Remaining: {2}".format(meta, acumulado, meta - acumulado), 
                    font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME).pack(anchor="w", padx=10, pady=2)
            
            registros_frame = tk.LabelFrame(tab, text=" PRODUCTION RECORDS ", 
                                            font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
            registros_frame.pack(fill="both", expand=True, padx=10, pady=5)
            
            columnas_reg = ("Exit date", "Internal lot", "Pieces", "Pieces/bag", 
                           "Initial serial", "Final serial", "Remainder", "Scrap", "Order date", "Destination")
            tree_reg = ttk.Treeview(registros_frame, columns=columnas_reg, show="headings", height=8)
            for col in columnas_reg:
                tree_reg.heading(col, text=col)
                tree_reg.column(col, width=100)
            tree_reg.pack(fill="both", expand=True, padx=5, pady=5)
            
            conn = sqlite3.connect(RUTA_DB)
            c2 = conn.cursor()
            c2.execute('''
                SELECT r.fecha_salida, r.lote_interno, r.no_piezas, r.piezas_por_bolsa, 
                       r.serial_inicial, r.serial_final, r.sobrante, r.scrap, p2.fecha_pedido, r.destino
                FROM registros r
                LEFT JOIN pedidos p2 ON r.pedido_id = p2.id
                WHERE r.pedido_id=? AND r.numero_parte=?
                ORDER BY CAST(r.serial_inicial AS INTEGER) ASC
            ''', (pedido_id, parte))
            registros = c2.fetchall()
            conn.close()
            
            for reg in registros:
                tree_reg.insert("", "end", values=reg)
            
            def ver_seriales(p=parte, pid=pedido_id):
                ventana_seriales = tk.Toplevel(ventana)
                ventana_seriales.title("Serials - {0} (Order {1})".format(p, pid))
                ventana_seriales.configure(bg=COLOR_FONDO)
                ventana_seriales.geometry("700x500")
                
                frame_serial = tk.Frame(ventana_seriales, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
                frame_serial.pack(fill="both", expand=True, padx=10, pady=10)
                
                columnas_serial = ("Serial", "Internal lot", "Exit date", "Destination")
                tree_serial = ttk.Treeview(frame_serial, columns=columnas_serial, show="headings", height=20)
                tree_serial.heading("Serial", text="Serial")
                tree_serial.heading("Internal lot", text="Internal lot")
                tree_serial.heading("Exit date", text="Exit date")
                tree_serial.heading("Destination", text="Destination")
                tree_serial.column("Serial", width=100)
                tree_serial.column("Internal lot", width=150)
                tree_serial.column("Exit date", width=100)
                tree_serial.column("Destination", width=100)
                tree_serial.pack(fill="both", expand=True, padx=5, pady=5)
                
                scroll = ttk.Scrollbar(frame_serial, orient="vertical", command=tree_serial.yview)
                scroll.pack(side="right", fill="y")
                tree_serial.configure(yscrollcommand=scroll.set)
                
                conn = sqlite3.connect(RUTA_DB)
                c3 = conn.cursor()
                c3.execute('''
                    SELECT serial_inicial, serial_final, lote_interno, fecha_salida, destino
                    FROM registros
                    WHERE pedido_id=? AND numero_parte=?
                    ORDER BY CAST(serial_inicial AS INTEGER) ASC
                ''', (pid, p))
                registros_serial = c3.fetchall()
                conn.close()
                
                for ini, fin, lote, fecha, destino in registros_serial:
                    try:
                        for s in range(int(ini), int(fin) + 1):
                            tree_serial.insert("", "end", values=(s, lote, fecha, destino))
                    except:
                        tree_serial.insert("", "end", values=(ini, lote, fecha, destino))
                
                btn_cerrar = tk.Button(frame_serial, text="Close", command=ventana_seriales.destroy,
                                      bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                                      relief=tk.RAISED, bd=1, padx=15, pady=3)
                btn_cerrar.pack(pady=10)
            
            btn_seriales = tk.Button(tab, text="View all serials", command=ver_seriales,
                                    bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                                    relief=tk.RAISED, bd=1, padx=15, pady=3)
            btn_seriales.pack(pady=5)
    
    tree_pedidos.bind("<<TreeviewSelect>>", on_select_pedido)
    
    btn_filtrar = tk.Button(frame_filtros, text="Filter", command=cargar_lista_pedidos,
                           bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                           relief=tk.RAISED, bd=1, padx=10, pady=2)
    btn_filtrar.pack(side="left", padx=10)
    
    btn_limpiar_filtros = tk.Button(frame_filtros, text="Clear filters", command=lambda: [
        entry_filtro_parte.delete(0, tk.END),
        entry_filtro_fecha_pedido.delete(0, tk.END),
        entry_filtro_fecha_salida.delete(0, tk.END),
        cargar_lista_pedidos()
    ], bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON, relief=tk.RAISED, bd=1, padx=10, pady=2)
    btn_limpiar_filtros.pack(side="left", padx=5)
    
    tab_manuales = ttk.Frame(notebook_main)
    notebook_main.add(tab_manuales, text="Manual records without order")

    lbl_manual_titulo = tk.Label(tab_manuales, text="MANUAL RECORDS WITHOUT ORDER", 
                                 font=FUENTE_SUBTITULO, bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
    lbl_manual_titulo.pack(pady=10)

    frame_filtros_man = tk.Frame(tab_manuales, bg=COLOR_FONDO_FRAME)
    frame_filtros_man.pack(padx=10, pady=5, fill="x")

    tk.Label(frame_filtros_man, text="Filter by part:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(side="left", padx=5)
    entry_filtro_parte_man = tk.Entry(frame_filtros_man, width=20, font=FUENTE_CAMPO,
                                     bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_filtro_parte_man.pack(side="left", padx=5)

    tk.Label(frame_filtros_man, text="Date from:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(side="left", padx=5)
    entry_filtro_fecha_man = tk.Entry(frame_filtros_man, width=15, font=FUENTE_CAMPO,
                                     bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_filtro_fecha_man.pack(side="left", padx=5)
    entry_filtro_fecha_man.bind("<KeyRelease>", formatear_fecha)

    frame_tabla_man = tk.Frame(tab_manuales, bg=COLOR_FONDO)
    frame_tabla_man.pack(padx=10, pady=5, fill="both", expand=True)

    columnas_man = (
        "ID", "Entry date", "Exit date", "Area", "Part number",
        "Internal lot", "Line", "Pieces", "Pieces/bag",
        "Bags", "Julian lot", "FIFO", "Remainder",
        "Initial serial", "Final serial", "Scrap", "Destination"
    )

    scroll_y_man = ttk.Scrollbar(frame_tabla_man, orient="vertical")
    scroll_x_man = ttk.Scrollbar(frame_tabla_man, orient="horizontal")

    tabla_man = ttk.Treeview(frame_tabla_man, columns=columnas_man, show="headings",
                              yscrollcommand=scroll_y_man.set, xscrollcommand=scroll_x_man.set, height=18)

    scroll_y_man.config(command=tabla_man.yview)
    scroll_x_man.config(command=tabla_man.xview)
    scroll_y_man.pack(side="right", fill="y")
    scroll_x_man.pack(side="bottom", fill="x")
    tabla_man.pack(fill="both", expand=True)

    for col in columnas_man:
        tabla_man.heading(col, text=col)
        tabla_man.column(col, width=50 if col == "ID" else 100)

    def cargar_manuales_sin_pedido():
        for row in tabla_man.get_children():
            tabla_man.delete(row)
        try:
            conn = sqlite3.connect(RUTA_DB)
            c = conn.cursor()
            parte_f = entry_filtro_parte_man.get().strip()
            fecha_f = entry_filtro_fecha_man.get().strip()
            
            fecha_f_comp = None
            if fecha_f:
                try:
                    fecha_obj = datetime.strptime(fecha_f, "%d/%m/%Y")
                    fecha_f_comp = fecha_obj.strftime("%Y%m%d")
                except:
                    pass
            
            query = '''
                SELECT id, fecha_entrada, fecha_salida, area, numero_parte,
                       lote_interno, linea, no_piezas, piezas_por_bolsa,
                       no_bolsas, lote_juliano, fifo, sobrante,
                       serial_inicial, serial_final, scrap, destino
                FROM registros
                WHERE (pedido_id IS NULL OR pedido_id = '')
            '''
            params = []
            if parte_f:
                query += " AND numero_parte LIKE ?"
                params.append("%{0}%".format(parte_f))
            if fecha_f_comp:
                query += " AND (substr(fecha_salida, 7, 4) || substr(fecha_salida, 4, 2) || substr(fecha_salida, 1, 2)) >= ?"
                params.append(fecha_f_comp)
            query += " ORDER BY id DESC"
            c.execute(query, params)
            for fila in c.fetchall():
                tabla_man.insert("", "end", values=fila)
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=ventana)

    btn_filtrar_man = tk.Button(frame_filtros_man, text="Filter", command=cargar_manuales_sin_pedido,
                           bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                           relief=tk.RAISED, bd=1, padx=10, pady=2)
    btn_filtrar_man.pack(side="left", padx=10)
    
    btn_mostrar_man = tk.Button(frame_filtros_man, text="Show all", command=lambda: [
        entry_filtro_parte_man.delete(0, tk.END),
        entry_filtro_fecha_man.delete(0, tk.END),
        cargar_manuales_sin_pedido()
    ], bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON, relief=tk.RAISED, bd=1, padx=10, pady=2)
    btn_mostrar_man.pack(side="left", padx=5)

    cargar_manuales_sin_pedido()
    cargar_lista_pedidos()

    btn_cerrar_ventana = tk.Button(ventana, text="Close", command=ventana.destroy,
                                  bg="#c0392b", fg="white", font=FUENTE_BOTON,
                                  relief=tk.RAISED, bd=1, padx=20, pady=5)
    btn_cerrar_ventana.pack(side="bottom", pady=10)

# --- EDIT ORDER DESTINATION ---
def ventana_editar_destino_pedido(parent, pedido_id, destino_actual, callback_refrescar=None):
    ventana = tk.Toplevel(parent)
    ventana.title("Edit Order Destination")
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("400x250")
    ventana.attributes("-topmost", True)
    
    main_frame = tk.Frame(ventana, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame.pack(padx=15, pady=15, fill="both", expand=True)
    
    tk.Label(main_frame, text="EDIT ORDER DESTINATION", font=FUENTE_TITULO, 
            bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO).pack(pady=10)
    
    ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=10, pady=5)
    
    frame_datos = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_datos.pack(padx=20, pady=15, fill="x")
    
    tk.Label(frame_datos, text="Order ID: {0}".format(pedido_id), font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(anchor="w", pady=5)
    
    tk.Label(frame_datos, text="Current destination: {0}".format(destino_actual if destino_actual else "No destination"), 
            font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME).pack(anchor="w", pady=5)
    
    tk.Label(frame_datos, text="New destination:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(anchor="w", pady=5)
    
    cfg = cargar_config()
    todos_destinos = []
    for cliente, datos in cfg["clientes"].items():
        destinos = datos.get("destinos", [])
        todos_destinos.extend(destinos)
    todos_destinos = list(set(todos_destinos))
    
    combo_nuevo_destino = ttk.Combobox(frame_datos, values=sorted(todos_destinos), width=30, state="readonly")
    combo_nuevo_destino.pack(anchor="w", pady=5)
    if destino_actual in todos_destinos:
        combo_nuevo_destino.set(destino_actual)
    
    def guardar():
        nuevo_destino = combo_nuevo_destino.get().strip()
        if not nuevo_destino:
            messagebox.showwarning("Warning", "Select a destination", parent=ventana)
            return
        
        if nuevo_destino == destino_actual:
            messagebox.showinfo("Info", "The destination is already the same", parent=ventana)
            ventana.destroy()
            return
        
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        c.execute('UPDATE pedidos SET destino=? WHERE id=?', (nuevo_destino, pedido_id))
        c.execute('UPDATE registros SET destino=? WHERE pedido_id=?', (nuevo_destino, pedido_id))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("Success", "Destination updated successfully in order and all its records", parent=ventana)
        ventana.destroy()
        if callback_refrescar:
            callback_refrescar()
    
    frame_btns = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_btns.pack(pady=15)
    
    tk.Button(frame_btns, text="Save", command=guardar,
             bg="#27ae60", fg="white", font=FUENTE_BOTON, padx=20, pady=4).pack(side="left", padx=10)
    tk.Button(frame_btns, text="Cancel", command=ventana.destroy,
             bg="#c0392b", fg="white", font=FUENTE_BOTON, padx=20, pady=4).pack(side="left", padx=10)

# --- EDIT EXIT DATE ---
def ventana_editar_fecha_salida_pedido(parent, pedido_id, fecha_actual, callback_refrescar=None):
    ventana = tk.Toplevel(parent)
    ventana.title("Edit Exit Date")
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("400x230")
    ventana.attributes("-topmost", True)
    
    main_frame = tk.Frame(ventana, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame.pack(padx=15, pady=15, fill="both", expand=True)
    
    tk.Label(main_frame, text="EDIT INITIAL EXIT DATE", font=FUENTE_TITULO, 
            bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO).pack(pady=10)
    
    ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=10, pady=5)
    
    frame_datos = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_datos.pack(padx=20, pady=15, fill="x")
    
    tk.Label(frame_datos, text="Order ID: {0}".format(pedido_id), font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(anchor="w", pady=5)
    
    tk.Label(frame_datos, text="Current date: {0}".format(fecha_actual if fecha_actual else "Not defined"), 
            font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME).pack(anchor="w", pady=5)
    
    tk.Label(frame_datos, text="New date (dd/mm/yyyy):", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(anchor="w", pady=5)
    
    entry_fecha = tk.Entry(frame_datos, width=20, font=FUENTE_CAMPO,
                          bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_fecha.pack(anchor="w", pady=5)
    if fecha_actual:
        entry_fecha.insert(0, fecha_actual)
    
    entry_fecha.bind("<KeyRelease>", formatear_fecha)
    
    def guardar():
        nueva_fecha = entry_fecha.get().strip()
        if not nueva_fecha:
            messagebox.showwarning("Warning", "Enter a date", parent=ventana)
            return
        
        try:
            datetime.strptime(nueva_fecha, "%d/%m/%Y")
        except:
            messagebox.showwarning("Warning", "Invalid date (dd/mm/yyyy)", parent=ventana)
            return
        
        if nueva_fecha == fecha_actual:
            messagebox.showinfo("Info", "The date is already the same", parent=ventana)
            ventana.destroy()
            return
        
        actualizar_fecha_salida_pedido(pedido_id, nueva_fecha)
        
        messagebox.showinfo("Success", "Exit date updated successfully in order and all its records", parent=ventana)
        ventana.destroy()
        if callback_refrescar:
            callback_refrescar()
    
    frame_btns = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_btns.pack(pady=15)
    
    tk.Button(frame_btns, text="Save", command=guardar,
             bg="#27ae60", fg="white", font=FUENTE_BOTON, padx=20, pady=4).pack(side="left", padx=10)
    tk.Button(frame_btns, text="Cancel", command=ventana.destroy,
             bg="#c0392b", fg="white", font=FUENTE_BOTON, padx=20, pady=4).pack(side="left", padx=10)

# --- ORDER WINDOWS ---
def ventana_nuevo_pedido(parent):
    ventana = tk.Toplevel(parent)
    ventana.title("New Order")
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("500x550")
    ventana.attributes("-topmost", True)
    
    main_frame = tk.Frame(ventana, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame.pack(padx=15, pady=15, fill="both", expand=True)
    
    lbl_titulo = tk.Label(main_frame, text="NEW ORDER (ONE PART)", font=FUENTE_TITULO, 
                         bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
    lbl_titulo.pack(pady=10)
    
    ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=10, pady=5)
    
    frame_datos = tk.LabelFrame(main_frame, text=" ORDER DATA ", 
                                font=FUENTE_SUBTITULO, bg=COLOR_FONDO_FRAME,
                                fg=COLOR_TITULO)
    frame_datos.pack(padx=15, pady=10, fill="x")
    
    tk.Label(frame_datos, text="Order date (dd/mm/yyyy):", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=0, column=0, padx=8, pady=8, sticky="w")
    entry_fecha = tk.Entry(frame_datos, width=18, font=FUENTE_CAMPO,
                          bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_fecha.grid(row=0, column=1, padx=8, pady=8)
    entry_fecha.bind("<KeyRelease>", formatear_fecha)
    
    tk.Label(frame_datos, text="Client:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=1, column=0, padx=8, pady=8, sticky="w")
    cfg = cargar_config()
    clientes = list(cfg["clientes"].keys())
    combo_cliente = ttk.Combobox(frame_datos, values=clientes, width=20, state="readonly")
    combo_cliente.grid(row=1, column=1, padx=8, pady=8)
    if clientes:
        combo_cliente.set(clientes[0])
    
    tk.Label(frame_datos, text="Destination:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=2, column=0, padx=8, pady=8, sticky="w")
    combo_destino = ttk.Combobox(frame_datos, values=[], width=20, state="readonly")
    combo_destino.grid(row=2, column=1, padx=8, pady=8)
    
    tk.Label(frame_datos, text="Part number:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=3, column=0, padx=8, pady=8, sticky="w")
    combo_parte = ttk.Combobox(frame_datos, width=25, state="readonly")
    combo_parte.grid(row=3, column=1, padx=8, pady=8)
    
    tk.Label(frame_datos, text="Quantity to produce:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=4, column=0, padx=8, pady=8, sticky="w")
    entry_cantidad = tk.Entry(frame_datos, width=18, font=FUENTE_CAMPO,
                             bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_cantidad.grid(row=4, column=1, padx=8, pady=8)
    
    def actualizar_destinos(*args):
        cliente_sel = combo_cliente.get()
        if cliente_sel:
            destinos = obtener_destinos_de_cliente(cliente_sel)
            combo_destino.config(values=destinos)
            if destinos:
                combo_destino.set(destinos[0])
            else:
                combo_destino.set("")
        else:
            combo_destino.config(values=[])
            combo_destino.set("")
    
    def actualizar_partes(*args):
        cliente_sel = combo_cliente.get()
        if cliente_sel:
            partes_filtradas = obtener_partes_por_cliente(cliente_sel)
            combo_parte.config(values=sorted(partes_filtradas))
            combo_parte.set("")
        else:
            combo_parte.config(values=[])

    def actualizar_todo_cliente(*args):
        actualizar_destinos()
        actualizar_partes()

    combo_cliente.bind("<<ComboboxSelected>>", actualizar_todo_cliente)
    
    actualizar_destinos()
    actualizar_partes()
    
    entry_fecha.focus_set()
    
    def guardar_pedido():
        fecha = entry_fecha.get().strip()
        cliente = combo_cliente.get()
        destino = combo_destino.get()
        parte = combo_parte.get()
        cantidad = entry_cantidad.get().strip()
        
        if not fecha:
            messagebox.showwarning("Warning", "Enter order date", parent=ventana)
            return
        if not cliente:
            messagebox.showwarning("Warning", "Select a client", parent=ventana)
            return
        if not destino:
            messagebox.showwarning("Warning", "Select a destination", parent=ventana)
            return
        if not parte:
            messagebox.showwarning("Warning", "Select a part number", parent=ventana)
            return
        if not cantidad:
            messagebox.showwarning("Warning", "Enter quantity to produce", parent=ventana)
            return
        
        try:
            datetime.strptime(fecha, "%d/%m/%Y")
        except:
            messagebox.showwarning("Warning", "Invalid date (dd/mm/yyyy)", parent=ventana)
            return
        
        try:
            cantidad_int = int(cantidad)
        except:
            messagebox.showwarning("Warning", "Quantity must be a number", parent=ventana)
            return
        
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        c.execute('INSERT INTO pedidos (fecha_pedido, cliente, destino) VALUES (?, ?, ?)', (fecha, cliente, destino))
        pedido_id = c.lastrowid
        
        c.execute('''
            INSERT INTO pedido_detalles (pedido_id, numero_parte, cantidad_meta)
            VALUES (?, ?, ?)
        ''', (pedido_id, parte, cantidad_int))
        
        conn.commit()
        conn.close()
        
        messagebox.showinfo("Success", "Order {0} created successfully\nDestination: {1}".format(pedido_id, destino), parent=ventana)
        ventana.destroy()
    
    btn_guardar = tk.Button(main_frame, text="Save Order", command=guardar_pedido,
                           bg="#27ae60", fg="white", font=FUENTE_BOTON,
                           relief=tk.RAISED, bd=1, padx=25, pady=5)
    btn_guardar.pack(pady=15)
    
    btn_cancelar = tk.Button(main_frame, text="Cancel", command=ventana.destroy,
                            bg="#c0392b", fg="white", font=FUENTE_BOTON,
                            relief=tk.RAISED, bd=1, padx=25, pady=5)
    btn_cancelar.pack(pady=5)

def ventana_editar_pedido(parent, pedido_id, detalle_id, parte, meta_actual, acumulado, callback_refrescar):
    ventana = tk.Toplevel(parent)
    ventana.title("Edit Order - {0}".format(parte))
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("400x420")
    ventana.attributes("-topmost", True)
    
    try:
        acumulado = int(acumulado)
    except:
        acumulado = 0
    
    conn = sqlite3.connect(RUTA_DB)
    c = conn.cursor()
    c.execute("SELECT fecha_pedido, destino FROM pedidos WHERE id=?", (pedido_id,))
    resultado = c.fetchone()
    fecha_pedido_actual = resultado[0]
    destino_actual = resultado[1] if len(resultado) > 1 else ""
    conn.close()
    
    main_frame = tk.Frame(ventana, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame.pack(padx=15, pady=15, fill="both", expand=True)
    
    tk.Label(main_frame, text="EDIT ORDER", font=FUENTE_TITULO, 
            bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO).pack(pady=10)
    
    ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=10, pady=5)
    
    frame_datos = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_datos.pack(padx=20, pady=15, fill="x")
    
    tk.Label(frame_datos, text="Part:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=0, column=0, sticky="w", pady=8)
    tk.Label(frame_datos, text=parte, font=FUENTE_CAMPO, 
            bg=COLOR_FONDO_FRAME).grid(row=0, column=1, padx=10, sticky="w")
    
    tk.Label(frame_datos, text="Order date (dd/mm/yyyy):", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=1, column=0, sticky="w", pady=8)
    entry_fecha_pedido = tk.Entry(frame_datos, width=15, font=FUENTE_CAMPO,
                                 bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_fecha_pedido.grid(row=1, column=1, padx=10, pady=8)
    entry_fecha_pedido.insert(0, fecha_pedido_actual)
    entry_fecha_pedido.bind("<KeyRelease>", formatear_fecha)
    
    tk.Label(frame_datos, text="Destination:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=2, column=0, sticky="w", pady=8)
    
    cliente = obtener_cliente_de_parte(parte)
    destinos = obtener_destinos_de_cliente(cliente)
    combo_destino_edit = ttk.Combobox(frame_datos, values=destinos, width=15, state="readonly")
    combo_destino_edit.grid(row=2, column=1, padx=10, pady=8)
    if destino_actual in destinos:
        combo_destino_edit.set(destino_actual)
    elif destinos:
        combo_destino_edit.set(destinos[0])
    
    tk.Label(frame_datos, text="Target quantity:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=3, column=0, sticky="w", pady=8)
    entry_meta = tk.Entry(frame_datos, width=15, font=FUENTE_CAMPO,
                         bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_meta.grid(row=3, column=1, padx=10, pady=8)
    entry_meta.insert(0, str(meta_actual))
    
    if acumulado > 0:
        entry_meta.config(state="disabled")
        tk.Label(frame_datos, text="(Cannot change. Already {0} pieces produced)".format(acumulado), 
                font=("Segoe UI", 8), fg="red", bg=COLOR_FONDO_FRAME).grid(row=4, column=0, columnspan=2, pady=5)
    else:
        tk.Label(frame_datos, text="(Can change target because no production)", 
                font=("Segoe UI", 8), fg="green", bg=COLOR_FONDO_FRAME).grid(row=4, column=0, columnspan=2, pady=5)
    
    def guardar():
        nueva_fecha_pedido = entry_fecha_pedido.get().strip()
        nuevo_destino = combo_destino_edit.get().strip()
        nueva_meta = entry_meta.get().strip()
        cambios = False
        
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        
        if nueva_fecha_pedido:
            try:
                datetime.strptime(nueva_fecha_pedido, "%d/%m/%Y")
                c.execute("UPDATE pedidos SET fecha_pedido=? WHERE id=?", (nueva_fecha_pedido, pedido_id))
                cambios = True
            except:
                messagebox.showwarning("Invalid date", "Format must be dd/mm/yyyy", parent=ventana)
                conn.close()
                return
        
        if nuevo_destino and nuevo_destino != destino_actual:
            c.execute("UPDATE pedidos SET destino=? WHERE id=?", (nuevo_destino, pedido_id))
            c.execute("UPDATE registros SET destino=? WHERE pedido_id=?", (nuevo_destino, pedido_id))
            cambios = True
        
        if nueva_meta and acumulado == 0:
            try:
                nueva_meta_int = int(nueva_meta)
                c.execute("UPDATE pedido_detalles SET cantidad_meta=? WHERE id=?", (nueva_meta_int, detalle_id))
                cambios = True
            except:
                messagebox.showwarning("Invalid quantity", "Must be a number", parent=ventana)
                conn.close()
                return
        
        if cambios:
            conn.commit()
            if nueva_meta and acumulado == 0:
                actualizar_acumulado_pedido(pedido_id, parte)
            messagebox.showinfo("Success", "Order updated", parent=ventana)
            if callback_refrescar:
                callback_refrescar()
            ventana.destroy()
        else:
            conn.close()
            messagebox.showinfo("No changes", "No modifications were made", parent=ventana)
            ventana.destroy()
    
    frame_btns = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_btns.pack(pady=15)
    
    tk.Button(frame_btns, text="Save", command=guardar,
             bg="#27ae60", fg="white", font=FUENTE_BOTON, padx=20, pady=4).pack(side="left", padx=10)
    tk.Button(frame_btns, text="Cancel", command=ventana.destroy,
             bg="#c0392b", fg="white", font=FUENTE_BOTON, padx=20, pady=4).pack(side="left", padx=10)

def ventana_trabajar_pedido(parent):
    ventana = tk.Toplevel(parent)
    ventana.title("Work on Order")
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("950x600")
    ventana.attributes("-topmost", True)
    
    main_frame = tk.Frame(ventana, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame.pack(padx=15, pady=15, fill="both", expand=True)
    
    lbl_titulo = tk.Label(main_frame, text="PARTS WITH ACTIVE ORDERS", font=FUENTE_TITULO, 
                         bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
    lbl_titulo.pack(pady=10)
    
    ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=10, pady=5)
    
    frame_lista = tk.LabelFrame(main_frame, text=" PART LIST ", 
                                font=FUENTE_SUBTITULO, bg=COLOR_FONDO_FRAME,
                                fg=COLOR_TITULO)
    frame_lista.pack(padx=15, pady=10, fill="both", expand=True)
    
    columnas = ("Order ID", "Part", "Target", "Produced", "Remaining", "Order date", "Initial exit date", "Destination")
    tree = ttk.Treeview(frame_lista, columns=columnas, show="headings", height=15)
    tree.heading("Order ID", text="Order ID")
    tree.heading("Part", text="Part number")
    tree.heading("Target", text="Target")
    tree.heading("Produced", text="Produced")
    tree.heading("Remaining", text="Remaining")
    tree.heading("Order date", text="Order date")
    tree.heading("Initial exit date", text="Initial exit date")
    tree.heading("Destination", text="Destination")
    tree.column("Order ID", width=70)
    tree.column("Part", width=180)
    tree.column("Target", width=70)
    tree.column("Produced", width=80)
    tree.column("Remaining", width=80)
    tree.column("Order date", width=100)
    tree.column("Initial exit date", width=100)
    tree.column("Destination", width=100)
    tree.pack(padx=10, pady=10, fill="both", expand=True)
    
    def cargar_partes_activas():
        for row in tree.get_children():
            tree.delete(row)
        
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        c.execute('''
            SELECT p.id, d.numero_parte, d.cantidad_meta, 
                   d.acumulado, p.fecha_pedido, d.fecha_salida_inicial, p.destino
            FROM pedidos p
            JOIN pedido_detalles d ON p.id = d.pedido_id
            WHERE p.estado = 'activo'
            ORDER BY p.id DESC
        ''')
        resultados = c.fetchall()
        conn.close()
        
        for row in resultados:
            pedido_id, parte, meta, acumulado, fecha_pedido, fecha_salida_inicial, destino = row
            faltante = meta - acumulado
            fecha_pedido_text = fecha_pedido if fecha_pedido else "Not defined"
            fecha_salida_text = fecha_salida_inicial if fecha_salida_inicial else "Pending"
            destino_text = destino if destino else "No destination"
            tag = "rojo" if faltante <= 0 else "verde"
            tree.insert("", "end", values=(pedido_id, parte, meta, acumulado, faltante, fecha_pedido_text, fecha_salida_text, destino_text), tags=(tag,))
    
    tree.tag_configure("verde", background="#d4efdf", foreground="#1e8449")
    tree.tag_configure("rojo", background="#fadbd8", foreground="#922b21")
    
    cargar_partes_activas()
    
    frame_botones = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_botones.pack(pady=10)
    
    def producir_parte():
        seleccion = tree.selection()
        if not seleccion:
            messagebox.showwarning("Warning", "Select a part", parent=ventana)
            return
        valores = tree.item(seleccion[0], "values")
        pedido_id = valores[0]
        parte = valores[1]
        fecha_salida_inicial = valores[6] if valores[6] != "Pending" else None
        destino = valores[7] if len(valores) > 7 and valores[7] != "No destination" else None
        
        ventana.destroy()
        FormularioProduccion(parent, "pedido", pedido_id, parte, None, fecha_salida_inicial, destino).abrir()
    
    def editar_pedido():
        seleccion = tree.selection()
        if not seleccion:
            messagebox.showwarning("Warning", "Select a part", parent=ventana)
            return
        valores = tree.item(seleccion[0], "values")
        pedido_id = valores[0]
        parte = valores[1]
        meta = valores[2]
        acumulado = valores[3]
        
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        c.execute("SELECT id FROM pedido_detalles WHERE pedido_id=? AND numero_parte=?", (pedido_id, parte))
        detalle_id = c.fetchone()[0]
        conn.close()
        
        def refrescar():
            cargar_partes_activas()
        
        try:
            acumulado_int = int(acumulado)
        except:
            acumulado_int = 0
        
        ventana_editar_pedido(ventana, pedido_id, detalle_id, parte, meta, acumulado_int, refrescar)
    
    def refrescar():
        cargar_partes_activas()
    
    btn_producir = tk.Button(frame_botones, text="PRODUCE", command=producir_parte,
                            bg="#27ae60", fg="white", font=FUENTE_BOTON,
                            relief=tk.RAISED, bd=1, padx=30, pady=5)
    btn_producir.pack(side="left", padx=10)
    
    btn_editar = tk.Button(frame_botones, text="EDIT ORDER", command=editar_pedido,
                          bg="#34495e", fg="white", font=FUENTE_BOTON,
                          relief=tk.RAISED, bd=1, padx=30, pady=5)
    btn_editar.pack(side="left", padx=10)
    
    btn_refrescar = tk.Button(frame_botones, text="REFRESH", command=refrescar,
                             bg="#34495e", fg="white", font=FUENTE_BOTON,
                             relief=tk.RAISED, bd=1, padx=30, pady=5)
    btn_refrescar.pack(side="left", padx=10)
    
    btn_cerrar = tk.Button(frame_botones, text="CLOSE", command=ventana.destroy,
                          bg="#c0392b", fg="white", font=FUENTE_BOTON,
                          relief=tk.RAISED, bd=1, padx=30, pady=5)
    btn_cerrar.pack(side="left", padx=10)
    
    tree.bind("<Double-1>", lambda e: producir_parte())

# --- REPORTS ---
def abrir_reporte(parent):
    ventana = tk.Toplevel(parent)
    ventana.title("Generate Reports")
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("700x550")
    ventana.attributes("-topmost", True)
    
    notebook = ttk.Notebook(ventana)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)
    
    tab_general = ttk.Frame(notebook)
    notebook.add(tab_general, text="General Report")
    
    main_frame_general = tk.Frame(tab_general, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame_general.pack(padx=15, pady=15, fill="both", expand=True)
    
    lbl_titulo_general = tk.Label(main_frame_general, text="GENERAL FILTERS", font=FUENTE_TITULO, 
                                  bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
    lbl_titulo_general.pack(pady=10)
    
    ttk.Separator(main_frame_general, orient='horizontal').pack(fill='x', padx=10, pady=5)
    
    frame_filtros = tk.Frame(main_frame_general, bg=COLOR_FONDO_FRAME)
    frame_filtros.pack(padx=20, pady=15, fill="x")
    
    row = 0
    tk.Label(frame_filtros, text="Start date (dd/mm/yyyy):", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=row, column=0, sticky="w", pady=6)
    entry_fecha_ini = tk.Entry(frame_filtros, width=18, font=FUENTE_CAMPO,
                              bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_fecha_ini.grid(row=row, column=1, padx=10, pady=6)
    entry_fecha_ini.bind("<KeyRelease>", formatear_fecha)
    
    row += 1
    tk.Label(frame_filtros, text="End date (dd/mm/yyyy):", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=row, column=0, sticky="w", pady=6)
    entry_fecha_fin = tk.Entry(frame_filtros, width=18, font=FUENTE_CAMPO,
                              bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_fecha_fin.grid(row=row, column=1, padx=10, pady=6)
    entry_fecha_fin.bind("<KeyRelease>", formatear_fecha)
    
    row += 1
    tk.Label(frame_filtros, text="Client:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=row, column=0, sticky="w", pady=6)
    
    cfg = cargar_config()
    clientes_lista = ["All"] + list(cfg["clientes"].keys())
    combo_cliente = ttk.Combobox(frame_filtros, values=clientes_lista, width=25, state="readonly")
    combo_cliente.grid(row=row, column=1, padx=10, pady=6)
    combo_cliente.set("All")
    
    row += 1
    tk.Label(frame_filtros, text="Part number:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=row, column=0, sticky="w", pady=6)
    
    combo_parte = ttk.Combobox(frame_filtros, width=25, state="readonly")
    combo_parte.grid(row=row, column=1, padx=10, pady=6)
    
    def actualizar_partes_por_cliente_filtro(*args):
        cliente_sel = combo_cliente.get()
        if cliente_sel == "All":
            combo_parte.config(state="disabled")
            combo_parte.set("All")
        else:
            combo_parte.config(state="readonly")
            partes_filtradas = obtener_partes_por_cliente(cliente_sel)
            combo_parte.config(values=["All"] + sorted(partes_filtradas))
            combo_parte.set("All")
    
    combo_cliente.bind("<<ComboboxSelected>>", actualizar_partes_por_cliente_filtro)
    actualizar_partes_por_cliente_filtro()
    
    lbl_info_general = tk.Label(main_frame_general, text="", font=FUENTE_LABEL, 
                                bg=COLOR_FONDO_FRAME, fg="#2980b9")
    lbl_info_general.pack(pady=5)
    
    entry_fecha_ini.focus_set()

    def ejecutar_reporte_general():
        fecha_ini_str = entry_fecha_ini.get().strip()
        fecha_fin_str = entry_fecha_fin.get().strip()
        cliente_sel = combo_cliente.get().strip()
        parte_sel = combo_parte.get().strip()
        
        fecha_ini_comp = None
        fecha_fin_comp = None
        
        if fecha_ini_str:
            try:
                fecha_obj = datetime.strptime(fecha_ini_str, "%d/%m/%Y")
                fecha_ini_comp = fecha_obj.strftime("%Y%m%d")
            except:
                messagebox.showwarning("Invalid date", "Invalid start date", parent=ventana)
                return
        if fecha_fin_str:
            try:
                fecha_obj = datetime.strptime(fecha_fin_str, "%d/%m/%Y")
                fecha_fin_comp = fecha_obj.strftime("%Y%m%d")
            except:
                messagebox.showwarning("Invalid date", "Invalid end date", parent=ventana)
                return
        
        try:
            conn = sqlite3.connect(RUTA_DB)
            c = conn.cursor()
            
            query = '''
                SELECT r.fecha_entrada, r.fecha_salida, r.area, r.numero_parte,
                    r.lote_interno, r.linea, r.no_piezas, r.piezas_por_bolsa,
                    r.no_bolsas, r.lote_juliano, r.fifo, r.sobrante,
                    r.serial_inicial, r.serial_final, r.pedido_id, r.pedido_fecha, r.scrap, r.destino
                FROM registros r
                LEFT JOIN pedidos p ON r.pedido_id = p.id
                WHERE 1=1
                AND (p.estado IS NULL OR p.estado != 'eliminado')
            '''
            params = []
            
            if fecha_ini_comp:
                query += " AND (substr(r.fecha_salida, 7, 4) || substr(r.fecha_salida, 4, 2) || substr(r.fecha_salida, 1, 2)) >= ?"
                params.append(fecha_ini_comp)
            if fecha_fin_comp:
                query += " AND (substr(r.fecha_salida, 7, 4) || substr(r.fecha_salida, 4, 2) || substr(r.fecha_salida, 1, 2)) <= ?"
                params.append(fecha_fin_comp)
            
            if cliente_sel != "All":
                query += " AND (p.cliente = ? OR r.pedido_id IS NULL)"
                params.append(cliente_sel)
            
            if parte_sel != "All":
                query += " AND r.numero_parte = ?"
                params.append(parte_sel)
            
            query += " ORDER BY r.fecha_salida DESC, r.id DESC"
            
            info_text = "Generating report with filters: "
            if cliente_sel != "All":
                info_text = info_text + "Client: " + cliente_sel + " | "
            if parte_sel != "All":
                info_text = info_text + "Part: " + parte_sel + " | "
            if fecha_ini_str:
                info_text = info_text + "From: " + fecha_ini_str + " | "
            if fecha_fin_str:
                info_text = info_text + "To: " + fecha_fin_str + " | "
            if info_text.endswith(" | "):
                info_text = info_text[:-3]
            if info_text == "Generating report with filters: ":
                info_text = "Generating full report (no filters)"
            
            lbl_info_general.config(text=info_text)
            ventana.update()
            
            c.execute(query, params)
            resultados = c.fetchall()
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=ventana)
            return
        
        if not resultados:
            messagebox.showwarning("Report", "No records with these filters.", parent=ventana)
            return
        
        escritorio = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        carpeta = os.path.join(escritorio, "Control de Rollos")
        if not os.path.exists(carpeta): 
            os.makedirs(carpeta)
        
        nombre = "reporte_general_{0}.csv".format(datetime.now().strftime("%Y%m%d_%H%M%S"))
        ruta = os.path.join(carpeta, nombre)
        
        with open(ruta, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=COLUMNAS_CSV)
            writer.writeheader()
            for fila in resultados:
                writer.writerow({
                    "Fecha entrada": fila[0],
                    "Fecha salida": fila[1],
                    "Area": fila[2],
                    "Numero de parte": fila[3],
                    "Lote interno": fila[4],
                    "Linea": fila[5],
                    "No. piezas": fila[6],
                    "Piezas por bolsa": fila[7],
                    "No. bolsas": fila[8],
                    "Lote juliano": fila[9],
                    "FIFO": fila[10],
                    "Sobrante": fila[11],
                    "Serial inicial": fila[12],
                    "Serial final": fila[13],
                    "Pedido ID": fila[14] if len(fila) > 14 else "",
                    "Fecha pedido": fila[15] if len(fila) > 15 else "",
                    "Scrap": fila[16] if len(fila) > 16 else "",
                    "Destino": fila[17] if len(fila) > 17 else ""
                })
        
        messagebox.showinfo("Report", "Report generated:\n{0}\n\n{1} records exported.".format(nombre, len(resultados)), parent=ventana)    

    btn_generar_general = tk.Button(main_frame_general, text="Generate CSV report", command=ejecutar_reporte_general,
                                    bg="#27ae60", fg="white", font=FUENTE_BOTON,
                                    relief=tk.RAISED, bd=1, padx=30, pady=6)
    btn_generar_general.pack(pady=15)
    
    tab_pedido = ttk.Frame(notebook)
    notebook.add(tab_pedido, text="Order Report")
    
    main_frame_pedido = tk.Frame(tab_pedido, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame_pedido.pack(padx=15, pady=15, fill="both", expand=True)
    
    lbl_titulo_pedido = tk.Label(main_frame_pedido, text="SELECT ORDER", font=FUENTE_TITULO, 
                                 bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
    lbl_titulo_pedido.pack(pady=10)
    
    ttk.Separator(main_frame_pedido, orient='horizontal').pack(fill='x', padx=10, pady=5)
    
    frame_seleccion = tk.Frame(main_frame_pedido, bg=COLOR_FONDO_FRAME)
    frame_seleccion.pack(padx=20, pady=15, fill="x")
    
    tk.Label(frame_seleccion, text="Select order:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=0, column=0, sticky="w", pady=6)
    
    frame_opciones = tk.Frame(frame_seleccion, bg=COLOR_FONDO_FRAME)
    frame_opciones.grid(row=0, column=1, padx=10, pady=6, sticky="w")
    
    var_sel_pedido = tk.StringVar()
    var_sel_pedido.set("id")
    rb_id = tk.Radiobutton(frame_opciones, text="By ID", variable=var_sel_pedido, value="id",
                          bg=COLOR_FONDO_FRAME, font=FUENTE_LABEL)
    rb_id.pack(side="left", padx=5)
    rb_lista = tk.Radiobutton(frame_opciones, text="Select from list", variable=var_sel_pedido, value="lista",
                             bg=COLOR_FONDO_FRAME, font=FUENTE_LABEL)
    rb_lista.pack(side="left", padx=5)
    
    frame_id = tk.Frame(frame_seleccion, bg=COLOR_FONDO_FRAME)
    frame_id.grid(row=1, column=0, columnspan=2, pady=10, sticky="w")
    frame_id.grid_forget()
    
    tk.Label(frame_id, text="Order ID:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(side="left", padx=5)
    entry_pedido_id = tk.Entry(frame_id, width=12, font=FUENTE_CAMPO,
                              bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_pedido_id.pack(side="left", padx=5)
    
    def cargar_info_pedido(pedido_id_str):
        if not pedido_id_str:
            messagebox.showwarning("Warning", "Enter an order ID", parent=ventana)
            return
        try:
            pedido_id_int = int(pedido_id_str)
        except:
            messagebox.showwarning("Warning", "ID must be a number", parent=ventana)
            return
        
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        c.execute('SELECT id, fecha_pedido, cliente, estado, destino FROM pedidos WHERE id=? AND estado != "eliminado"', (pedido_id_int,))
        pedido = c.fetchone()
        
        if not pedido:
            messagebox.showwarning("Warning", "Order ID {0} does not exist".format(pedido_id_int), parent=ventana)
            conn.close()
            return
        
        c.execute('SELECT numero_parte, cantidad_meta FROM pedido_detalles WHERE pedido_id=?', (pedido_id_int,))
        partes = c.fetchall()
        conn.close()
        
        estado_texto = pedido[3]
        if estado_texto == "activo":
            estado_mostrar = "active"
        elif estado_texto == "cerrado":
            estado_mostrar = "completed"
        else:
            estado_mostrar = "cancelled"
        
        destino_texto = pedido[4] if len(pedido) > 4 and pedido[4] else "No destination"
        
        info_text = "Order {0} - Date: {1} - Client: {2} - Destination: {3} - Status: {4}\nParts: ".format(
            pedido[0], pedido[1], pedido[2], destino_texto, estado_mostrar)
        if partes:
            info_text = info_text + ", ".join(["{0} (target: {1})".format(p[0], p[1]) for p in partes])
        else:
            info_text = info_text + "No parts assigned"
        
        lbl_info_pedido.config(text=info_text)
        pedido_seleccionado[0] = pedido_id_int
    
    btn_buscar_id = tk.Button(frame_id, text="Search", 
                             command=lambda: cargar_info_pedido(entry_pedido_id.get().strip()),
                             bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                             relief=tk.RAISED, bd=1, padx=10, pady=2)
    btn_buscar_id.pack(side="left", padx=5)
    
    frame_lista_pedidos = tk.Frame(frame_seleccion, bg=COLOR_FONDO_FRAME)
    frame_lista_pedidos.grid(row=2, column=0, columnspan=2, pady=10, sticky="w")
    frame_lista_pedidos.grid_forget()
    
    frame_lista_scroll = tk.Frame(frame_lista_pedidos, bg=COLOR_FONDO_FRAME)
    frame_lista_scroll.pack(fill="x")
    
    scroll_lista = ttk.Scrollbar(frame_lista_scroll, orient="vertical")
    lista_pedidos = tk.Listbox(frame_lista_scroll, yscrollcommand=scroll_lista.set, height=8, width=70,
                              font=FUENTE_CAMPO, bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    scroll_lista.config(command=lista_pedidos.yview)
    lista_pedidos.pack(side="left", fill="both", expand=True)
    scroll_lista.pack(side="right", fill="y")
    
    frame_btn_lista = tk.Frame(frame_lista_pedidos, bg=COLOR_FONDO_FRAME)
    frame_btn_lista.pack(pady=5)
    
    def cargar_lista_pedidos_reporte():
        lista_pedidos.delete(0, tk.END)
        conn = sqlite3.connect(RUTA_DB)
        c = conn.cursor()
        c.execute('SELECT id, fecha_pedido, cliente, estado, destino FROM pedidos WHERE estado != "eliminado" AND estado != "cancelado" ORDER BY id DESC')
        pedidos = c.fetchall()
        conn.close()
        for p in pedidos:
            if p[3] == "cerrado":
                estado_texto = " (completed)"
            else:
                estado_texto = ""
            destino_texto = p[4] if len(p) > 4 and p[4] else "No destination"
            lista_pedidos.insert(tk.END, "ID: {0} | Date: {1} | Client: {2} | Destination: {3}{4}".format(p[0], p[1], p[2], destino_texto, estado_texto))
    
    btn_refrescar_lista = tk.Button(frame_btn_lista, text="Refresh order list", command=cargar_lista_pedidos_reporte,
                                   bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                                   relief=tk.RAISED, bd=1, padx=15, pady=2)
    btn_refrescar_lista.pack(side="left", padx=5)
    
    lbl_info_pedido = tk.Label(main_frame_pedido, text="", font=FUENTE_LABEL, 
                               bg=COLOR_FONDO_FRAME, fg="#2980b9", wraplength=500)
    lbl_info_pedido.pack(pady=5)
    
    pedido_seleccionado = [None]
    
    def on_select_lista(event):
        seleccion = lista_pedidos.curselection()
        if not seleccion:
            return
        texto = lista_pedidos.get(seleccion[0])
        import re
        match = re.search(r'ID: (\d+)', texto)
        if match:
            pedido_id = int(match.group(1))
            cargar_info_pedido(str(pedido_id))
    
    lista_pedidos.bind("<<ListboxSelect>>", on_select_lista)
    
    def actualizar_visibilidad_seleccion(*args):
        modo = var_sel_pedido.get()
        if modo == "id":
            frame_id.grid(row=1, column=0, columnspan=2, pady=10, sticky="w")
            frame_lista_pedidos.grid_forget()
        else:
            frame_id.grid_forget()
            frame_lista_pedidos.grid(row=2, column=0, columnspan=2, pady=10, sticky="w")
            cargar_lista_pedidos_reporte()
    
    var_sel_pedido.trace('w', actualizar_visibilidad_seleccion)
    actualizar_visibilidad_seleccion()
    
    def ejecutar_reporte_pedido():
        pedido_id = pedido_seleccionado[0]
        if not pedido_id:
            messagebox.showwarning("Warning", "Select an order first", parent=ventana)
            return
        
        try:
            conn = sqlite3.connect(RUTA_DB)
            c = conn.cursor()
            c.execute('''
                SELECT fecha_entrada, fecha_salida, area, numero_parte,
                       lote_interno, linea, no_piezas, piezas_por_bolsa,
                       no_bolsas, lote_juliano, fifo, sobrante,
                       serial_inicial, serial_final, pedido_id, pedido_fecha, scrap, destino
                FROM registros
                WHERE pedido_id = ?
                ORDER BY fecha_salida DESC, id DESC
            ''', (pedido_id,))
            resultados = c.fetchall()
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=ventana)
            return
        
        if not resultados:
            messagebox.showwarning("Report", "Order {0} has no associated records.".format(pedido_id), parent=ventana)
            return
        
        escritorio = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        carpeta = os.path.join(escritorio, "Control de Rollos")
        if not os.path.exists(carpeta): 
            os.makedirs(carpeta)
        
        nombre = "reporte_pedido_{0}_{1}.csv".format(pedido_id, datetime.now().strftime("%Y%m%d_%H%M%S"))
        ruta = os.path.join(carpeta, nombre)
        
        with open(ruta, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=COLUMNAS_CSV)
            writer.writeheader()
            for fila in resultados:
                writer.writerow({
                    "Fecha entrada": fila[0],
                    "Fecha salida": fila[1],
                    "Area": fila[2],
                    "Numero de parte": fila[3],
                    "Lote interno": fila[4],
                    "Linea": fila[5],
                    "No. piezas": fila[6],
                    "Piezas por bolsa": fila[7],
                    "No. bolsas": fila[8],
                    "Lote juliano": fila[9],
                    "FIFO": fila[10],
                    "Sobrante": fila[11],
                    "Serial inicial": fila[12],
                    "Serial final": fila[13],
                    "Pedido ID": fila[14] if len(fila) > 14 else "",
                    "Fecha pedido": fila[15] if len(fila) > 15 else "",
                    "Scrap": fila[16] if len(fila) > 16 else "",
                    "Destino": fila[17] if len(fila) > 17 else ""
                })
        
        messagebox.showinfo("Report", "Report generated:\n{0}\n\n{1} records exported.".format(nombre, len(resultados)), parent=ventana)
    
    btn_generar_pedido = tk.Button(main_frame_pedido, text="Generate CSV report", command=ejecutar_reporte_pedido,
                                   bg="#27ae60", fg="white", font=FUENTE_BOTON,
                                   relief=tk.RAISED, bd=1, padx=30, pady=6)
    btn_generar_pedido.pack(pady=15)
    
    btn_cerrar = tk.Button(ventana, text="Close", command=ventana.destroy,
                          bg="#c0392b", fg="white", font=FUENTE_BOTON,
                          relief=tk.RAISED, bd=1, padx=20, pady=5)
    btn_cerrar.pack(side="bottom", pady=10)
    
    cargar_lista_pedidos_reporte()

# --- REIMPRINT ---
def abrir_reimprimir(parent):
    ventana = tk.Toplevel(parent)
    ventana.title("Reimprimir etiquetas")
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("1400x550")
    ventana.attributes("-topmost", True)
    
    main_frame = tk.Frame(ventana, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame.pack(padx=10, pady=10, fill="both", expand=True)
    
    frame_busq = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_busq.pack(padx=10, pady=8, fill="x")
    
    tk.Label(frame_busq, text="Part number:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(side="left")
    entry_parte = tk.Entry(frame_busq, width=18, font=FUENTE_CAMPO,
                          bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_parte.pack(side="left", padx=5)
    entry_parte.focus_set()
    
    tk.Label(frame_busq, text="Internal lot:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(side="left", padx=10)
    entry_lote_interno = tk.Entry(frame_busq, width=15, font=FUENTE_CAMPO,
                         bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_lote_interno.pack(side="left", padx=5)
    
    tk.Label(frame_busq, text="FIFO:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(side="left", padx=10)
    entry_fifo = tk.Entry(frame_busq, width=15, font=FUENTE_CAMPO,
                         bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_fifo.pack(side="left", padx=5)
    
    columnas = (
        "id", "Entry date", "Exit date", "Area", "Part number",
        "Internal lot", "Line", "Pieces", "Pieces/bag",
        "Bags", "Julian lot", "FIFO", "Remainder",
        "Initial serial", "Final serial", "Order ID", "Scrap", "Destination"
    )
    
    frame_tabla = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_tabla.pack(padx=10, pady=5, fill="both", expand=True)
    
    scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical")
    scroll_x = ttk.Scrollbar(frame_tabla, orient="horizontal")
    
    tabla = ttk.Treeview(frame_tabla, columns=columnas, show="headings",
                         yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
    
    scroll_y.config(command=tabla.yview)
    scroll_x.config(command=tabla.xview)
    scroll_y.pack(side="right", fill="y")
    scroll_x.pack(side="bottom", fill="x")
    tabla.pack(fill="both", expand=True)
    
    for col in columnas:
        tabla.heading(col, text=col)
        tabla.column(col, width=40 if col == "id" else 100)
    
    mostrar_todos = [False]
    lbl_info = tk.Label(main_frame, text="", font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, fg="#2980b9")
    lbl_info.pack(pady=2)
    
    def buscar():
        for row in tabla.get_children():
            tabla.delete(row)
        try:
            conn = sqlite3.connect(RUTA_DB)
            c = conn.cursor()
            parte = entry_parte.get().strip()
            lote_interno = entry_lote_interno.get().strip()
            fifo = entry_fifo.get().strip()
            
            query = '''
                SELECT id, fecha_entrada, fecha_salida, area, numero_parte,
                       lote_interno, linea, no_piezas, piezas_por_bolsa,
                       no_bolsas, lote_juliano, fifo, sobrante,
                       serial_inicial, serial_final, pedido_id, scrap, destino
                FROM registros
                WHERE 1=1
                AND (pedido_id IS NULL OR pedido_id NOT IN (SELECT id FROM pedidos WHERE estado = 'eliminado'))
            '''
            params = []
            if parte:
                query += " AND numero_parte = ?"
                params.append(parte)
            if lote_interno:
                query += " AND lote_interno = ?"
                params.append(lote_interno)
            if fifo:
                query += " AND fifo = ?"
                params.append(fifo)
            query += " ORDER BY id DESC"
            
            if not mostrar_todos[0]:
                query += " LIMIT 20"
            
            c.execute(query, params)
            resultados = c.fetchall()
            
            hay_mas = False
            if not mostrar_todos[0]:
                query_check = query.replace(" LIMIT 20", " LIMIT 21")
                c.execute(query_check, params)
                hay_mas = len(c.fetchall()) > 20
            
            for fila in resultados:
                tabla.insert("", "end", values=fila)
            conn.close()
            
            if hay_mas and not mostrar_todos[0]:
                btn_ver_mas.pack(side="bottom", pady=5)
                lbl_info.config(text="Showing last 20 records. More records available.")
            elif mostrar_todos[0]:
                btn_ver_mas.pack_forget()
                lbl_info.config(text="Showing all records ({} results).".format(len(resultados)))
            else:
                btn_ver_mas.pack_forget()
                lbl_info.config(text="")
                
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=ventana)
    
    def ver_todo():
        mostrar_todos[0] = True
        buscar()
    
    btn_buscar = tk.Button(frame_busq, text="Search", command=lambda: buscar(),
                          bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                          relief=tk.RAISED, bd=1, padx=15, pady=2)
    btn_buscar.pack(side="left", padx=10)
    
    btn_limpiar = tk.Button(frame_busq, text="Clear", command=lambda: [
        entry_parte.delete(0, tk.END),
        entry_lote_interno.delete(0, tk.END),
        entry_fifo.delete(0, tk.END),
        buscar()
    ], bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON, relief=tk.RAISED, bd=1, padx=15, pady=2)
    btn_limpiar.pack(side="left", padx=5)
    
    btn_ver_mas = tk.Button(main_frame, text=" Show all records ", command=ver_todo,
                            bg="#2980b9", fg="white", font=FUENTE_BOTON,
                            relief=tk.RAISED, bd=1, padx=20, pady=3)
    
    registro_seleccionado = [None]
    
    def al_seleccionar(event):
        seleccion = tabla.selection()
        if not seleccion:
            return
        valores = tabla.item(seleccion[0], "values")
        registro_seleccionado[0] = {
            "Fecha entrada": valores[1],
            "Fecha salida": valores[2],
            "Area": valores[3],
            "Numero de parte": valores[4],
            "Lote interno": valores[5],
            "Linea": valores[6],
            "No. piezas": valores[7],
            "Piezas por bolsa": valores[8],
            "No. bolsas": valores[9],
            "Lote juliano": valores[10],
            "FIFO": valores[11],
            "Sobrante": valores[12],
            "Serial inicial": valores[13],
            "Serial final": valores[14],
            "Scrap": valores[16] if len(valores) > 16 else "",
            "Destino": valores[17] if len(valores) > 17 else ""
        }
    
    tabla.bind("<<TreeviewSelect>>", al_seleccionar)
    
    def ejecutar_reimprimir():
        if registro_seleccionado[0] is None:
            messagebox.showwarning("Warning", "Select a record first.", parent=ventana)
            return
        visor_secuencial_static(registro_seleccionado[0], ventana)
    
    frame_btns = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_btns.pack(pady=10)
    
    btn_reimprimir = tk.Button(frame_btns, text="Reimprimir lote seleccionado", command=ejecutar_reimprimir,
                              bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                              relief=tk.RAISED, bd=1, padx=20, pady=4)
    btn_reimprimir.pack(padx=10)
    
    buscar()

def visor_secuencial_static(data, parent_ventana):
    try:
        inicio = int(data["Serial inicial"])
        final = int(data["Serial final"])
        total = final - inicio + 1
    except:
        messagebox.showerror("Error", "Initial serial and Final serial must be numbers.", parent=parent_ventana)
        return
    
    piezas_por_bolsa_original = int(data["Piezas por bolsa"])
    no_piezas_total = int(data["No. piezas"])
    no_bolsas = int(data["No. bolsas"])
    
    piezas_ultima_bolsa = None
    if no_bolsas > 0:
        piezas_por_bolsa_normal = piezas_por_bolsa_original
        if no_piezas_total % piezas_por_bolsa_normal != 0:
            piezas_ultima_bolsa = no_piezas_total % piezas_por_bolsa_normal
    
    indice_actual = 0
    ventana_prev = tk.Toplevel(parent_ventana)
    ventana_prev.title("Reimpresion")
    ventana_prev.configure(bg=COLOR_FONDO)
    ventana_prev.attributes("-topmost", True)
    ventana_prev.resizable(False, False)
    
    frame_visor = tk.Frame(ventana_prev, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    frame_visor.pack(padx=15, pady=15, fill="both", expand=True)
    
    lbl_imagen = tk.Label(frame_visor, bg="white", relief=tk.SUNKEN, bd=1)
    lbl_imagen.pack(padx=15, pady=15)
    
    lbl_info = tk.Label(frame_visor, text="", font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME, fg=COLOR_TITULO)
    lbl_info.pack(pady=5)
    
    def refrescar_vista():
        nonlocal indice_actual
        serial_actual = inicio + indice_actual
        
        if piezas_ultima_bolsa is not None and indice_actual == total - 1:
            data_temp = data.copy()
            data_temp["Piezas por bolsa"] = str(piezas_ultima_bolsa)
            img_pil = generar_imagen_etiqueta(data_temp, serial_actual)
        else:
            img_pil = generar_imagen_etiqueta(data, serial_actual)
        
        nw, nh = int(img_pil.width * 0.6), int(img_pil.height * 0.6)
        try: 
            resample_filter = Image.ANTIALIAS
        except AttributeError: 
            resample_filter = Image.BILINEAR
        
        img_resizada = img_pil.resize((nw, nh), resample_filter)
        img_tk = ImageTk.PhotoImage(img_resizada)
        lbl_imagen.config(image=img_tk)
        lbl_imagen.image = img_tk
        
        lbl_info.config(text="Label {0} of {1} | Serial: {2}".format(indice_actual + 1, total, serial_actual))
        
        if indice_actual + 1 >= total:
            btn_imprimir.config(text="PRINT AND FINISH")
        else:
            btn_imprimir.config(text="PRINT LABEL {0}".format(serial_actual))
        
        ventana_prev.after(500, lambda: btn_imprimir.config(state="normal"))
        btn_imprimir.focus_set()
    
    def ejecutar_impresion():
        nonlocal indice_actual
        if btn_imprimir['state'] == 'disabled': 
            return
        btn_imprimir.config(state="disabled")
        
        serial_a_imprimir = inicio + indice_actual
        
        if piezas_ultima_bolsa is not None and indice_actual == total - 1:
            data_temp = data.copy()
            data_temp["Piezas por bolsa"] = str(piezas_ultima_bolsa)
            zpl_final = generar_codigo_zpl(data_temp, serial_a_imprimir)
        else:
            zpl_final = generar_codigo_zpl(data, serial_a_imprimir)
        
        if not enviar_a_zebra(zpl_final):
            btn_imprimir.config(state="normal")
            return
        
        if indice_actual + 1 >= total:
            ventana_prev.destroy()
            return
        
        indice_actual += 1
        refrescar_vista()
    
    def saltar_etiqueta():
        nonlocal indice_actual
        if indice_actual + 1 >= total:
            ventana_prev.destroy()
        else:
            indice_actual += 1
            refrescar_vista()
    
    frame_btns = tk.Frame(frame_visor, bg=COLOR_FONDO_FRAME)
    frame_btns.pack(pady=10)
    
    btn_imprimir = tk.Button(frame_btns, text="PRINT", command=ejecutar_impresion,
                            bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                            relief=tk.RAISED, bd=1, padx=15, pady=3)
    btn_imprimir.pack(side="left", padx=8)
    
    btn_saltar = tk.Button(frame_btns, text="SKIP SERIAL", command=saltar_etiqueta,
                          bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                          relief=tk.RAISED, bd=1, padx=15, pady=3)
    btn_saltar.pack(side="left", padx=8)
    
    btn_cancelar = tk.Button(frame_btns, text="Cancel", command=ventana_prev.destroy,
                            bg="#c0392b", fg="white", font=FUENTE_BOTON,
                            relief=tk.RAISED, bd=1, padx=15, pady=3)
    btn_cancelar.pack(side="left", padx=8)
    
    ventana_prev.bind("<Return>", lambda e: ejecutar_impresion())
    refrescar_vista()

# --- RECORD EDITOR ---
def pedir_password(parent):
    ventana_pass = tk.Toplevel(parent)
    ventana_pass.title("Restricted Access")
    ventana_pass.geometry("320x160")
    ventana_pass.attributes("-topmost", True)
    ventana_pass.resizable(False, False)
    ventana_pass.grab_set()
    ventana_pass.configure(bg="#2b2b2b")
    
    tk.Label(ventana_pass, text="RESTRICTED AREA", font=("Arial", 12, "bold"),
             bg="#2b2b2b", fg="#ff4444").pack(pady=(18, 4))
    tk.Label(ventana_pass, text="Enter password to continue",
             font=("Arial", 9), bg="#2b2b2b", fg="#aaaaaa").pack()
    
    entry_pass = tk.Entry(ventana_pass, show="*", width=22, font=("Arial", 11),
                          bg="#3c3c3c", fg="white", insertbackground="white",
                          relief="flat", bd=4)
    entry_pass.pack(pady=10)
    entry_pass.focus_set()
    
    resultado = [False]
    
    def verificar(event=None):
        if entry_pass.get() == PASSWORD_ADMIN:
            resultado[0] = True
            ventana_pass.destroy()
        else:
            entry_pass.config(bg="#5c2222")
            ventana_pass.after(500, lambda: entry_pass.config(bg="#3c3c3c"))
            entry_pass.delete(0, tk.END)
    
    entry_pass.bind("<Return>", verificar)
    tk.Button(ventana_pass, text="Enter", command=verificar,
              bg="#ff4444", fg="white", font=("Arial", 10, "bold"),
              relief="flat", padx=20, pady=4).pack()
    ventana_pass.wait_window()
    return resultado[0]

def abrir_editor(parent):
    if not pedir_password(parent):
        return
    
    ventana = tk.Toplevel(parent)
    ventana.title("Record Editor")
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("1400x600")
    ventana.attributes("-topmost", True)
    
    main_frame = tk.Frame(ventana, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame.pack(padx=10, pady=10, fill="both", expand=True)
    
    frame_busq = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_busq.pack(padx=10, pady=8, fill="x")
    
    tk.Label(frame_busq, text="Part number:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(side="left")
    entry_parte = tk.Entry(frame_busq, width=20, font=FUENTE_CAMPO,
                          bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_parte.pack(side="left", padx=5)
    entry_parte.focus_set()
    
    columnas = (
        "id", "Entry date", "Exit date", "Area", "Part number",
        "Internal lot", "Line", "Pieces", "Pieces/bag",
        "Bags", "Julian lot", "FIFO", "Remainder",
        "Initial serial", "Final serial", "Order ID", "Scrap", "Destination"
    )
    
    frame_tabla = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_tabla.pack(padx=10, pady=5, fill="both", expand=True)
    
    scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical")
    scroll_x = ttk.Scrollbar(frame_tabla, orient="horizontal")
    
    tabla = ttk.Treeview(frame_tabla, columns=columnas, show="headings",
                         yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
    
    scroll_y.config(command=tabla.yview)
    scroll_x.config(command=tabla.xview)
    scroll_y.pack(side="right", fill="y")
    scroll_x.pack(side="bottom", fill="x")
    tabla.pack(fill="both", expand=True)
    
    for col in columnas:
        tabla.heading(col, text=col)
        tabla.column(col, width=40 if col == "id" else 100)
    
    def buscar():
        for row in tabla.get_children():
            tabla.delete(row)
        try:
            conn = sqlite3.connect(RUTA_DB)
            c = conn.cursor()
            parte = entry_parte.get().strip()
            if parte:
                c.execute('''
                    SELECT id, fecha_entrada, fecha_salida, area, numero_parte,
                           lote_interno, linea, no_piezas, piezas_por_bolsa,
                           no_bolsas, lote_juliano, fifo, sobrante,
                           serial_inicial, serial_final, pedido_id, scrap, destino
                    FROM registros
                    WHERE numero_parte=?
                    AND (pedido_id IS NULL OR pedido_id NOT IN (SELECT id FROM pedidos WHERE estado = 'eliminado'))
                    ORDER BY id DESC
                ''', (parte,))
            else:
                c.execute('''
                    SELECT id, fecha_entrada, fecha_salida, area, numero_parte,
                           lote_interno, linea, no_piezas, piezas_por_bolsa,
                           no_bolsas, lote_juliano, fifo, sobrante,
                           serial_inicial, serial_final, pedido_id, scrap, destino
                    FROM registros
                    WHERE (pedido_id IS NULL OR pedido_id NOT IN (SELECT id FROM pedidos WHERE estado = 'eliminado'))
                    ORDER BY id DESC
                ''')
            for fila in c.fetchall():
                tabla.insert("", "end", values=fila)
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=ventana)
    
    btn_buscar = tk.Button(frame_busq, text="Search", command=buscar,
                          bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                          relief=tk.RAISED, bd=1, padx=15, pady=2)
    btn_buscar.pack(side="left", padx=10)
    
    frame_form = tk.LabelFrame(main_frame, text=" Edit selected record ", 
                               font=FUENTE_SUBTITULO, bg=COLOR_FONDO_FRAME,
                               fg=COLOR_TITULO)
    frame_form.pack(padx=10, pady=8, fill="x")
    
    campos_edit = [
        "Entry date", "Exit date", "Area", "Part number",
        "Internal lot", "Line", "Pieces", "Pieces/bag",
        "Bags", "Julian lot", "FIFO", "Remainder",
        "Initial serial", "Final serial", "Order ID", "Scrap", "Destination"
    ]
    
    entries_edit = {}
    for idx, campo in enumerate(campos_edit):
        tk.Label(frame_form, text=campo, font=FUENTE_LABEL, 
                bg=COLOR_FONDO_FRAME).grid(row=0, column=idx, padx=3, pady=5)
        e = tk.Entry(frame_form, width=12, font=FUENTE_CAMPO,
                    bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
        e.grid(row=1, column=idx, padx=3, pady=5)
        entries_edit[campo] = e
    
    for campo_f in ["Entry date", "Exit date"]:
        if campo_f in entries_edit:
            entries_edit[campo_f].bind("<KeyRelease>", formatear_fecha)
    
    id_seleccionado = [None]
    
    def al_seleccionar(event):
        seleccion = tabla.selection()
        if not seleccion: 
            return
        valores = tabla.item(seleccion[0], "values")
        id_seleccionado[0] = valores[0]
        for idx, campo in enumerate(campos_edit):
            entries_edit[campo].delete(0, tk.END)
            entries_edit[campo].insert(0, valores[idx + 1] if idx + 1 < len(valores) else "")
    
    tabla.bind("<<TreeviewSelect>>", al_seleccionar)
    
    frame_btns_ed = tk.Frame(main_frame, bg=COLOR_FONDO_FRAME)
    frame_btns_ed.pack(pady=10)
    
    def guardar_edicion():
        if id_seleccionado[0] is None:
            messagebox.showwarning("Warning", "Select a record first.", parent=ventana)
            return
        try:
            conn = sqlite3.connect(RUTA_DB)
            c = conn.cursor()
            c.execute('''
                UPDATE registros SET
                    fecha_entrada=?, fecha_salida=?, area=?, numero_parte=?,
                    lote_interno=?, linea=?, no_piezas=?, piezas_por_bolsa=?,
                    no_bolsas=?, lote_juliano=?, fifo=?, sobrante=?,
                    serial_inicial=?, serial_final=?, pedido_id=?, scrap=?,
                    destino=?
                WHERE id=?
            ''', (
                entries_edit["Entry date"].get(),
                entries_edit["Exit date"].get(),
                entries_edit["Area"].get(),
                entries_edit["Part number"].get(),
                entries_edit["Internal lot"].get(),
                entries_edit["Line"].get(),
                entries_edit["Pieces"].get(),
                entries_edit["Pieces/bag"].get(),
                entries_edit["Bags"].get(),
                entries_edit["Julian lot"].get(),
                entries_edit["FIFO"].get(),
                entries_edit["Remainder"].get(),
                entries_edit["Initial serial"].get(),
                entries_edit["Final serial"].get(),
                entries_edit["Order ID"].get() if entries_edit["Order ID"].get() else None,
                entries_edit["Scrap"].get(),
                entries_edit["Destination"].get(),
                id_seleccionado[0]
            ))
            conn.commit()
            conn.close()
            buscar()
            messagebox.showinfo("Success", "Record {0} updated successfully.".format(id_seleccionado[0]), parent=ventana)
            messagebox.showwarning(
                "WARNING",
                "If you modified the quantity of a record belonging to an active order,\n"
                "the order accumulated total will NOT be automatically updated.\n\n"
                "It is recommended NOT to modify quantities of orders in production.\n"
                "Use this function only for specific corrections and at your own risk.",
                parent=ventana
            )
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=ventana)
    
    def eliminar_registro():
        if id_seleccionado[0] is None:
            messagebox.showwarning("Warning", "Select a record first.", parent=ventana)
            return
        if not messagebox.askyesno("Confirm", "Delete record {0}? This action cannot be undone.".format(id_seleccionado[0]), parent=ventana):
            return
        try:
            conn = sqlite3.connect(RUTA_DB)
            c = conn.cursor()
            c.execute("DELETE FROM registros WHERE id=?", (id_seleccionado[0],))
            conn.commit()
            conn.close()
            id_eliminado = id_seleccionado[0]
            id_seleccionado[0] = None
            for campo in campos_edit:
                entries_edit[campo].delete(0, tk.END)
            buscar()
            messagebox.showinfo("Success", "Record {0} deleted successfully.".format(id_eliminado), parent=ventana)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=ventana)
    
    btn_guardar_ed = tk.Button(frame_btns_ed, text="Save changes", command=guardar_edicion,
                              bg="#27ae60", fg="white", font=FUENTE_BOTON,
                              relief=tk.RAISED, bd=1, padx=20, pady=4)
    btn_guardar_ed.pack(side="left", padx=10)
    
    btn_eliminar_ed = tk.Button(frame_btns_ed, text="Delete record", command=eliminar_registro,
                               bg="#c0392b", fg="white", font=FUENTE_BOTON,
                               relief=tk.RAISED, bd=1, padx=20, pady=4)
    btn_eliminar_ed.pack(side="left", padx=10)
    
    buscar()

# --- ADMINISTRATION ---
def abrir_admin(parent):
    if not pedir_password(parent):
        return
    
    ventana = tk.Toplevel(parent)
    ventana.title("Administration")
    ventana.configure(bg=COLOR_FONDO)
    ventana.geometry("700x650")
    ventana.attributes("-topmost", True)
    ventana.resizable(False, False)
    
    main_frame = tk.Frame(ventana, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
    main_frame.pack(padx=10, pady=10, fill="both", expand=True)
    
    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)
    
    tab_clientes = ttk.Frame(notebook)
    notebook.add(tab_clientes, text="Clients")
    
    frame_clientes = tk.Frame(tab_clientes, bg=COLOR_FONDO_FRAME)
    frame_clientes.pack(fill="both", expand=True, padx=10, pady=10)
    
    tk.Label(frame_clientes, text="Registered clients:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(pady=5)
    lista_clientes = tk.Listbox(frame_clientes, height=10, width=50, font=FUENTE_CAMPO,
                               bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    lista_clientes.pack(pady=5)
    
    def refrescar_clientes():
        lista_clientes.delete(0, tk.END)
        cfg = cargar_config()
        for clave, datos in cfg["clientes"].items():
            destinos = datos.get("destinos", [])
            destinos_str = ", ".join(destinos) if destinos else "No destinations"
            lista_clientes.insert(tk.END, "{0} - {1} - Destinations: {2}".format(clave, datos.get("nombre", ""), destinos_str))
    
    refrescar_clientes()
    
    frame_nuevo_cliente = tk.LabelFrame(frame_clientes, text=" New client ", 
                                        font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME,
                                        fg=COLOR_TITULO)
    frame_nuevo_cliente.pack(pady=10, fill="x")
    
    tk.Label(frame_nuevo_cliente, text="ID (e.g. CLIENT):", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    entry_clave = tk.Entry(frame_nuevo_cliente, width=15, font=FUENTE_CAMPO,
                          bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_clave.grid(row=0, column=1, padx=5, pady=5)
    
    tk.Label(frame_nuevo_cliente, text="Name:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=1, column=0, padx=5, pady=5, sticky="w")
    entry_nombre = tk.Entry(frame_nuevo_cliente, width=25, font=FUENTE_CAMPO,
                           bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_nombre.grid(row=1, column=1, padx=5, pady=5)
    
    tk.Label(frame_nuevo_cliente, text="Destinations (comma separated):", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=2, column=0, padx=5, pady=5, sticky="w")
    entry_destinos_cliente = tk.Entry(frame_nuevo_cliente, width=25, font=FUENTE_CAMPO,
                                      bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_destinos_cliente.grid(row=2, column=1, padx=5, pady=5)
    
    def agregar_cliente():
        clave = entry_clave.get().strip().upper()
        nombre = entry_nombre.get().strip()
        destinos_text = entry_destinos_cliente.get().strip()
        destinos_list = [d.strip() for d in destinos_text.split(",") if d.strip()]
        
        if not clave or not nombre:
            messagebox.showwarning("Warning", "Complete ID and name.", parent=ventana)
            return
        cfg = cargar_config()
        if clave in cfg["clientes"]:
            messagebox.showwarning("Warning", "That client already exists.", parent=ventana)
            return
        cfg["clientes"][clave] = {"nombre": nombre, "destinos": destinos_list}
        guardar_config(cfg)
        refrescar_clientes()
        entry_clave.delete(0, tk.END)
        entry_nombre.delete(0, tk.END)
        entry_destinos_cliente.delete(0, tk.END)
        messagebox.showinfo("Success", "Client '{0}' added with destinations: {1}".format(clave, ", ".join(destinos_list)), parent=ventana)
    
    btn_agregar_cliente = tk.Button(frame_nuevo_cliente, text="Add client", command=agregar_cliente,
                                   bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                                   relief=tk.RAISED, bd=1, padx=15, pady=3)
    btn_agregar_cliente.grid(row=3, column=0, columnspan=2, pady=10)
    
    tab_piezas = ttk.Frame(notebook)
    notebook.add(tab_piezas, text="Parts")
    
    frame_piezas = tk.Frame(tab_piezas, bg=COLOR_FONDO_FRAME)
    frame_piezas.pack(fill="both", expand=True, padx=10, pady=10)
    
    tk.Label(frame_piezas, text="Registered parts:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(pady=5)
    
    frame_lista_piezas = tk.Frame(frame_piezas, bg=COLOR_FONDO_FRAME)
    frame_lista_piezas.pack(fill="both", expand=True, padx=5, pady=5)
    
    scroll_piezas = ttk.Scrollbar(frame_lista_piezas, orient="vertical")
    lista_piezas = tk.Listbox(frame_lista_piezas, yscrollcommand=scroll_piezas.set, height=8, 
                             font=FUENTE_CAMPO, bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    scroll_piezas.config(command=lista_piezas.yview)
    lista_piezas.pack(side="left", fill="both", expand=True)
    scroll_piezas.pack(side="right", fill="y")
    
    def refrescar_lista_piezas():
        lista_piezas.delete(0, tk.END)
        cfg = cargar_config()
        for parte, datos in cfg["partes"].items():
            cliente = datos.get("cliente", "")
            complemento = datos.get("complemento", "")
            posicion = datos.get("posicion", "inicio")
            lista_piezas.insert(tk.END, "{0} - Client: {1} - Comp: {2} - Pos: {3}".format(
                parte, cliente, complemento if complemento else "(empty)", posicion))
    
    refrescar_lista_piezas()
    
    frame_acciones_piezas = tk.Frame(frame_piezas, bg=COLOR_FONDO_FRAME)
    frame_acciones_piezas.pack(pady=5, fill="x")
    
    def eliminar_pieza():
        seleccion = lista_piezas.curselection()
        if not seleccion:
            messagebox.showwarning("Warning", "Select a part to delete", parent=ventana)
            return
        pieza_texto = lista_piezas.get(seleccion[0])
        parte = pieza_texto.split(" - ")[0]
        
        if not messagebox.askyesno("Confirm", 
            "Delete part '{0}' from JSON?\n\nThis will NOT delete existing records in the database.".format(parte), 
            parent=ventana):
            return
        
        cfg = cargar_config()
        if parte in cfg["partes"]:
            del cfg["partes"][parte]
            guardar_config(cfg)
            global partes_data
            partes_data = cargar_partes()
            refrescar_lista_piezas()
            messagebox.showinfo("Success", "Part '{0}' removed from JSON.".format(parte), parent=ventana)
        else:
            messagebox.showwarning("Error", "The part does not exist in JSON.", parent=ventana)
    
    btn_eliminar_pieza = tk.Button(frame_acciones_piezas, text="Delete selected part", command=eliminar_pieza,
                                  bg="#c0392b", fg="white", font=FUENTE_BOTON,
                                  relief=tk.RAISED, bd=1, padx=15, pady=3)
    btn_eliminar_pieza.pack(side="left", padx=5)
    
    btn_refrescar_piezas = tk.Button(frame_acciones_piezas, text="Refresh list", command=refrescar_lista_piezas,
                                    bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                                    relief=tk.RAISED, bd=1, padx=15, pady=3)
    btn_refrescar_piezas.pack(side="left", padx=5)
    
    frame_nueva_pieza = tk.LabelFrame(frame_piezas, text=" New part ", 
                                      font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME,
                                      fg=COLOR_TITULO)
    frame_nueva_pieza.pack(pady=10, fill="x")
    
    tk.Label(frame_nueva_pieza, text="Part number:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    entry_nueva_parte = tk.Entry(frame_nueva_pieza, width=20, font=FUENTE_CAMPO,
                                 bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_nueva_parte.grid(row=0, column=1, padx=5, pady=5)
    
    tk.Label(frame_nueva_pieza, text="Client:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=1, column=0, padx=5, pady=5, sticky="w")
    combo_cliente = ttk.Combobox(frame_nueva_pieza, values=list(cargar_config()["clientes"].keys()), 
                                 width=25, state="readonly")
    combo_cliente.grid(row=1, column=1, padx=5, pady=5)
    
    tk.Label(frame_nueva_pieza, text="Complement:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=2, column=0, padx=5, pady=5, sticky="w")
    entry_complemento = tk.Entry(frame_nueva_pieza, width=20, font=FUENTE_CAMPO,
                                 bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
    entry_complemento.grid(row=2, column=1, padx=5, pady=5)
    
    tk.Label(frame_nueva_pieza, text="Complement position:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).grid(row=3, column=0, padx=5, pady=5, sticky="w")
    combo_posicion = ttk.Combobox(frame_nueva_pieza, values=["inicio", "final"], 
                                  width=10, state="readonly")
    combo_posicion.set("inicio")
    combo_posicion.grid(row=3, column=1, padx=5, pady=5, sticky="w")
    
    def agregar_pieza():
        parte = entry_nueva_parte.get().strip()
        cliente = combo_cliente.get().strip()
        complemento = entry_complemento.get().strip()
        posicion = combo_posicion.get().strip()
        if not parte or not cliente:
            messagebox.showwarning("Warning", "Part number and client are required.", parent=ventana)
            return
        cfg = cargar_config()
        if parte in cfg["partes"]:
            messagebox.showwarning("Warning", "That part already exists.", parent=ventana)
            return
        cfg["partes"][parte] = {"cliente": cliente, "complemento": complemento, "posicion": posicion}
        guardar_config(cfg)
        global partes_data
        partes_data = cargar_partes()
        refrescar_lista_piezas()
        messagebox.showinfo("Success", "Part '{0}' added.".format(parte), parent=ventana)
        entry_nueva_parte.delete(0, tk.END)
        entry_complemento.delete(0, tk.END)
        combo_cliente.set("")
        combo_posicion.set("inicio")
    
    btn_agregar_pieza = tk.Button(frame_nueva_pieza, text="Add part", command=agregar_pieza,
                                 bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                                 relief=tk.RAISED, bd=1, padx=15, pady=3)
    btn_agregar_pieza.grid(row=4, column=0, columnspan=2, pady=10)
    
    tab_config = ttk.Frame(notebook)
    notebook.add(tab_config, text="Configuration")
    
    frame_config = tk.Frame(tab_config, bg=COLOR_FONDO_FRAME)
    frame_config.pack(fill="both", expand=True, padx=10, pady=10)
    
    frame_impresora = tk.LabelFrame(frame_config, text=" Label Printer ", 
                                    font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME,
                                    fg=COLOR_TITULO)
    frame_impresora.pack(pady=10, fill="x")
    
    impresora_actual = obtener_impresora()
    tk.Label(frame_impresora, text="Current printer:", font=FUENTE_LABEL, 
            bg=COLOR_FONDO_FRAME).pack(anchor="w", padx=10, pady=5)
    lbl_actual = tk.Label(frame_impresora, text=impresora_actual, font=("Segoe UI", 9, "bold"), 
                         bg=COLOR_FONDO_FRAME, fg="#2980b9")
    lbl_actual.pack(anchor="w", padx=20, pady=5)
    
    def abrir_config_impresora():
        ventana_imp = tk.Toplevel(ventana)
        ventana_imp.title("Configure Printer")
        ventana_imp.configure(bg=COLOR_FONDO)
        ventana_imp.geometry("600x450")
        ventana_imp.attributes("-topmost", True)
        
        main_frame_imp = tk.Frame(ventana_imp, bg=COLOR_FONDO_FRAME, relief=tk.RIDGE, bd=1)
        main_frame_imp.pack(padx=15, pady=15, fill="both", expand=True)
        
        tk.Label(main_frame_imp, text="Select label printer:", 
                font=FUENTE_LABEL, bg=COLOR_FONDO_FRAME).pack(pady=10)
        
        frame = tk.Frame(main_frame_imp, bg=COLOR_FONDO_FRAME)
        frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        scroll = ttk.Scrollbar(frame)
        scroll.pack(side="right", fill="y")
        
        lista = tk.Listbox(frame, yscrollcommand=scroll.set, height=12, font=FUENTE_CAMPO,
                          bg=COLOR_ENTRY, relief=tk.SUNKEN, bd=1)
        lista.pack(side="left", fill="both", expand=True)
        scroll.config(command=lista.yview)
        
        def refrescar():
            lista.delete(0, tk.END)
            try:
                impresoras = [p[2] for p in win32print.EnumPrinters(2)]
                for imp in impresoras:
                    lista.insert(tk.END, imp)
                
                actual = obtener_impresora()
                for i, imp in enumerate(impresoras):
                    if imp == actual:
                        lista.selection_set(i)
                        lista.see(i)
            except Exception as e:
                messagebox.showerror("Error", "Could not get printers:\n{0}".format(str(e)), parent=ventana_imp)
        
        btn_refrescar = tk.Button(main_frame_imp, text="Refresh list", command=refrescar,
                                  bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                                  relief=tk.RAISED, bd=1, padx=15, pady=3)
        btn_refrescar.pack(pady=8)
        
        def guardar():
            seleccion = lista.curselection()
            if not seleccion:
                messagebox.showwarning("Warning", "Select a printer", parent=ventana_imp)
                return
            impresora = lista.get(seleccion[0])
            guardar_impresora(impresora)
            lbl_actual.config(text=impresora)
            messagebox.showinfo("Success", "Printer configured:\n{0}".format(impresora), parent=ventana_imp)
            ventana_imp.destroy()
        
        btn_guardar_imp = tk.Button(main_frame_imp, text="Save", command=guardar,
                                   bg="#27ae60", fg="white", font=FUENTE_BOTON,
                                   relief=tk.RAISED, bd=1, padx=20, pady=5)
        btn_guardar_imp.pack(pady=10, side="left", padx=10)
        
        btn_cancelar_imp = tk.Button(main_frame_imp, text="Cancel", command=ventana_imp.destroy,
                                     bg="#c0392b", fg="white", font=FUENTE_BOTON,
                                     relief=tk.RAISED, bd=1, padx=20, pady=5)
        btn_cancelar_imp.pack(pady=10, side="left", padx=10)
        
        refrescar()
    
    btn_cambiar_imp = tk.Button(frame_impresora, text="Change printer", command=abrir_config_impresora,
                               bg=COLOR_BOTON, fg="white", font=FUENTE_BOTON,
                               relief=tk.RAISED, bd=1, padx=20, pady=4)
    btn_cambiar_imp.pack(pady=10)

# --- MAIN MENU ---
class MenuPrincipal:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Control System - Production Management")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")
        
        ancho = 500
        alto = 650
        self.root.update_idletasks()
        pantalla_ancho = self.root.winfo_screenwidth()
        pantalla_alto = self.root.winfo_screenheight()
        x = (pantalla_ancho - ancho) // 2
        y = (pantalla_alto - alto) // 2 - 30
        self.root.geometry("{0}x{1}+{2}+{3}".format(ancho, alto, x, y))
        
        main_frame = tk.Frame(self.root, bg="#e8e8e8", relief=tk.RIDGE, bd=2)
        main_frame.pack(padx=15, pady=15, fill="both", expand=True)
        
        header_frame = tk.Frame(main_frame, bg="#2c3e50")
        header_frame.pack(fill="x", pady=0)
        
        tk.Label(header_frame, text="CONTROL SYSTEM", font=("Segoe UI", 24, "bold"), 
                bg="#2c3e50", fg="white").pack(pady=(15, 5))
        tk.Label(header_frame, text="Production Management", font=("Segoe UI", 10), 
                bg="#2c3e50", fg="#bdc3c7").pack(pady=(0, 15))
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=15, pady=10)
        
        botones_frame = tk.Frame(main_frame, bg="#e8e8e8")
        botones_frame.pack(pady=15, padx=20, fill="both", expand=True)
        
        def on_enter(event, btn, color):
            btn.config(bg=color)
        
        def on_leave(event, btn, color):
            btn.config(bg=color)
        
        botones = [
            ("NEW ORDER", lambda: ventana_nuevo_pedido(self.root), "#2c3e50", "#1abc9c"),
            ("WORK ON ORDER", lambda: ventana_trabajar_pedido(self.root), "#2c3e50", "#1abc9c"),
            ("REPRINT LABELS", lambda: abrir_reimprimir(self.root), "#2c3e50", "#1abc9c"),
            ("VIEW ORDERS", lambda: ventana_ver_pedidos(self.root), "#2c3e50", "#1abc9c"),
            ("EDIT RECORDS", lambda: abrir_editor(self.root), "#2c3e50", "#1abc9c"),
            ("ADMINISTRATION", lambda: abrir_admin(self.root), "#2c3e50", "#1abc9c"),
            ("MANUAL REGISTRATION", lambda: FormularioProduccion(self.root, "manual").abrir(), "#2c3e50", "#1abc9c"),
            ("REPORTS", lambda: abrir_reporte(self.root), "#2c3e50", "#1abc9c"),
            ("EXIT", self.root.quit, "#2c3e50", "#e74c3c")
        ]
        
        for texto, comando, color_base, color_hover in botones:
            btn = tk.Button(botones_frame, text=texto, command=comando,
                           bg=color_base, fg="white", font=("Segoe UI", 10, "bold"),
                           relief=tk.RAISED, bd=1, padx=15, pady=7, width=25,
                           activebackground=color_hover, activeforeground="white")
            btn.pack(pady=5)
            btn.bind("<Enter>", lambda e, b=btn, c=color_hover: on_enter(e, b, c))
            btn.bind("<Leave>", lambda e, b=btn, c=color_base: on_leave(e, b, c))
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=15, pady=10)
        
        footer_frame = tk.Frame(main_frame, bg="#e8e8e8")
        footer_frame.pack(pady=10)
        
        tk.Label(footer_frame, text="Company Name", 
                font=("Segoe UI", 8, "italic"), bg="#e8e8e8", fg="#7f8c8d").pack()
        tk.Label(footer_frame, text="Version 2.0", 
                font=("Segoe UI", 7), bg="#e8e8e8", fg="#7f8c8d").pack()
    
    def ejecutar(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = MenuPrincipal()
    app.ejecutar()