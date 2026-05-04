import os
import sys
import glob
import pandas as pd
import re
import numpy as np
import calendar
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import ttk, messagebox

# =====================================================================
#  SISTEMA DE TRAZABILIDAD (LOGS)
# =====================================================================
def escribir_log(mensaje, carpeta):
    """Registra la actividad del proceso para auditoría técnica."""
    ruta_log = os.path.join(carpeta, "ejecucion_proceso.log")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ruta_log, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {mensaje}\n")

# ======================================================================
# INTERFAZ DE USUARIO (GUI)
# =====================================================================

def seleccionar_mes_interfaz():
    """Ventana emergente para la selección del periodo contable."""
    ventana = tk.Tk()
    ventana.title("Gestor de Periodos - Automatización")
    ventana.geometry("350x180")
    ventana.attributes("-topmost", True)

    label = tk.Label(ventana, text="Seleccione el mes de cierre contable:", pady=10, font=("Arial", 10))
    label.pack()

    meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
             'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']

    combo = ttk.Combobox(ventana, values=meses, state="readonly")
    combo.set('abril') 
    combo.pack(pady=5)

    seleccion = {"mes": "abril"}

    def confirmar():
        seleccion["mes"] = combo.get()
        ventana.destroy()

    btn = tk.Button(ventana, text="Confirmar y Procesar", command=confirmar, bg="#2c3e50", fg="white", padx=10)
    btn.pack(pady=15)

    def al_cerrar():
        if messagebox.askokcancel("Salir", "¿Desea cancelar el proceso de automatización?"):
            ventana.destroy()
            sys.exit()

    ventana.protocol("WM_DELETE_WINDOW", al_cerrar)
    ventana.mainloop()
    return seleccion["mes"]

# =====================================================================
#  CONFIGURACIÓN DE ENTORNO
# =====================================================================
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
NOMBRE_PROYECTO = "Sistema_Procesamiento_Contable"
CARPETA_TRABAJO = os.path.join(BASE_DIR, NOMBRE_PROYECTO)

if not os.path.exists(CARPETA_TRABAJO):
    os.makedirs(CARPETA_TRABAJO)
    print(f"✅ Carpeta de trabajo creada: {NOMBRE_PROYECTO}")
    print("Coloque los archivos .xlsx de origen en esta carpeta.")
    input("\nPresione ENTER para salir...")
    sys.exit()

# Inicialización de Log
ruta_log_file = os.path.join(CARPETA_TRABAJO, "ejecucion_proceso.log")
if os.path.exists(ruta_log_file): os.remove(ruta_log_file)

escribir_log("🚀 --- INICIO DE MOTOR DE PROCESAMIENTO FINANCIERO ---", CARPETA_TRABAJO)

archivos_disponibles = [
    f for f in glob.glob(os.path.join(CARPETA_TRABAJO, "*.xlsx")) 
    if not os.path.basename(f).startswith("REPORTE_FINAL")
]

if not archivos_disponibles:
    escribir_log("❌ Error: No se detectaron archivos Excel para procesar.", CARPETA_TRABAJO)
    print(f"❌ Error: Cargue archivos .xlsx en la carpeta: {NOMBRE_PROYECTO}")
    input("\nPresione ENTER para salir...")
    sys.exit()

RUTA_ENTRADA = archivos_disponibles[0]
MES_A_PROCESAR = seleccionar_mes_interfaz()
ANIO_A_PROCESAR = datetime.now().year 

MESES_DIC = {'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
             'julio':7,'agosto':8,'septiembre':9,'setiembre':9,'octubre':10,
             'noviembre':11,'diciembre':12}

nombre_salida = f"REPORTE_AUTOMATIZADO_{MES_A_PROCESAR.upper()}.xlsx"
RUTA_SALIDA = os.path.join(CARPETA_TRABAJO, nombre_salida)

# =====================================================================
#  MOTOR LÓGICO DE TRANSFORMACIÓN
# =====================================================================

def procesar_logica_periodos(row):
    """Analiza y clasifica la data basándose en reglas de devengo y diferido."""
    id_registro = str(row[0]) if pd.notnull(row[0]) else "S/ID"

    if pd.isna(row[0]) or pd.isna(row[1]):
        return pd.Series([None] * 33)

    texto_completo = " ".join([str(x) for x in row.values if pd.notnull(x)]).lower()

    # 1. Gestión de Fechas de Emisión
    f_emision = None
    try:
        f_raw = row[1]
        if isinstance(f_raw, datetime): f_emision = f_raw
        else: f_emision = pd.to_datetime(str(f_raw).split()[0], dayfirst=True)
    except: 
        escribir_log(f"⚠️ Registro {id_registro}: Error en formato de fecha emisión.", CARPETA_TRABAJO)
        return pd.Series([None] * 33)

    # 2. Extracción de Periodos de Servicio (Regex)
    f_ini, f_fin = None, None
    patron_fechas = r'(\d{1,2})\s+de\s+([a-z]+)\s+(?:de\s+)?(\d{4})'
    coincidencias = re.findall(patron_fechas, texto_completo)

    if len(coincidencias) >= 2:
        try:
            d1, m1, a1 = coincidencias[0]; d2, m2, a2 = coincidencias[1]
            f_ini = datetime(int(a1), MESES_DIC[m1], int(d1))
            f_fin = datetime(int(a2), MESES_DIC[m2], int(d2))
        except: 
            f_ini = f_emision + timedelta(days=5)
            f_fin = f_ini + relativedelta(months=1, days=-1)
    else:
        f_ini = f_emision + timedelta(days=5)
        f_fin = f_ini + (timedelta(days=13) if "semana" in texto_completo else relativedelta(months=1, days=-1))

    escribir_log(f"✅ ID {id_registro} OK: Periodo {f_ini.date()} - {f_fin.date()}", CARPETA_TRABAJO)

    # 3. Limpieza de Montos y Moneda
    def limpiar_monto(r):
        numeros = []
        for v in r.values:
            try:
                s = re.sub(r'[^\d.]', '', str(v).replace(',', ''))
                if s and '.' in s: numeros.append(float(s))
            except: continue
        numeros = sorted(list(set(numeros)), reverse=True)
        # Lógica para extraer base imponible (asumiendo impuestos estándar)
        for t in numeros:
            for n in numeros:
                if t > n and abs((n * 1.18) - t) <= 0.05: return n
        return numeros[-1] if numeros else 0

    monto_neto = limpiar_monto(row)
    moneda = "USD" if "usd" in texto_completo or "$" in texto_completo else "PEN"

    # 4. Cálculo de Diferimiento
    intervalo = "mensual"
    divisor = 30

    if "anual" in texto_completo: intervalo, divisor = "anual", 365
    elif "semestral" in texto_completo: intervalo, divisor = "semestral", 180
    elif "trimestral" in texto_completo: intervalo, divisor = "trimestral", 90

    f_limite_corte = (datetime(ANIO_A_PROCESAR, MESES_DIC[MES_A_PROCESAR.lower()], 1) + relativedelta(months=1, days=-1))
    diferencia_dias = (f_fin - f_limite_corte).days
    dias_diferidos = diferencia_dias if diferencia_dias > 0 else 0

    dif_pen, dif_usd = "", ""
    if dias_diferidos > 0:
        calculo = (monto_neto / divisor) * dias_diferidos
        if moneda == "PEN": dif_pen = "{:.2f}".format(calculo)
        else: dif_usd = "{:.2f}".format(calculo)

    # Retorno de Estructura Normalizada (33 Columnas)
    return pd.Series([
        row[0], f_emision.strftime('%Y-%m-%d'), (f_emision + timedelta(days=5)).strftime('%Y-%m-%d'),
        f_ini.strftime('%Y-%m-%d'), f_fin.strftime('%Y-%m-%d'), "DATO_CLIENTE_ID", "RAZON_SOCIAL_GENERICA",
        re.sub(r'\[|\]', '', str(row[6])) if pd.notnull(row[6]) else "", "SISTEMA_ERP", "RUC_GENERICO",
        "NOMBRE_COMERCIAL", "{:.2f}".format(monto_neto), 1, moneda, "{:.2f}".format(monto_neto),
        "CATEGORIA_PRODUCTO", intervalo, "CONTABILIDAD", "LINEA_NEGOCIO", "ESTADO_VALIDO",
        "", "ACTIVO", "", 
        "{:.2f}".format(monto_neto) if moneda == "PEN" else "",     
        "{:.2f}".format(monto_neto) if moneda == "USD" else "",     
        dif_pen, dif_usd, dias_diferidos, f_limite_corte.strftime('%d-%b-%y'),
        "", "", "", ""
    ])

# --- PROCESO PRINCIPAL ---
try:
    print(f"🚀 Procesando periodo: {MES_A_PROCESAR.upper()}")

    excel_file = pd.ExcelFile(RUTA_ENTRADA)
    df_raw = pd.read_excel(RUTA_ENTRADA, sheet_name=excel_file.sheet_names[0], header=None)
    
    df_final = df_raw.apply(procesar_logica_periodos, axis=1)

    df_final.columns = [
        'ID_Registro','Fecha_Emision','Fecha_Vencimiento','Inicio_Servicio','Fin_Servicio',
        'ID_Fiscal_Receptor','Razon_Social','Descripcion_Servicio','Referencia_ERP','ID_Fiscal_Emisor',
        'Nombre_Emisor','Precio_Unitario','Cantidad','Moneda','Importe_Total',
        'Nombre_Producto','Frecuencia_Facturacion','Categoria','Segmento','Validacion_Interna',
        'Transacciones','Estado','Fecha_Suspension','Base_Imponible_PEN','Base_Imponible_USD',
        'Diferido_PEN','Diferido_USD','Dias_Pendientes','Fecha_Corte_Mensual',
        'Aux_1','Relacion_Doc','Base_Imponible_Ref','Doc_Referencia'
    ]

    # Limpieza final de registros vacíos
    df_final = df_final.dropna(how='all') 
    df_final = df_final[df_final['ID_Registro'].notna()]

    df_final.to_excel(RUTA_SALIDA, index=False)

    escribir_log(f"🎯 PROCESO EXITOSO. Archivo generado: {nombre_salida}", CARPETA_TRABAJO)
    print("\n" + "="*50)
    print(f"🎯 ¡FINALIZADO! Reporte exportado a: {NOMBRE_PROYECTO}")
    print("="*50)
except Exception as e:
    escribir_log(f"❌ ERROR CRÍTICO: {str(e)}", CARPETA_TRABAJO)
    print(f"\n❌ Se produjo un error: {e}")

input("\nProceso terminado. Presione ENTER para cerrar...")