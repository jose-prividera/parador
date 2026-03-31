import pandas as pd
import numpy as np
import traceback
import os
import sys
import warnings
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

# --- VERIFICACIÓN DE LIBRERÍAS ---
try:
    import xlwings as xw
except ImportError:
    print("ERROR CRÍTICO: Falta 'xlwings'. Ejecuta: pip install xlwings")
    sys.exit(1)

try:
    import xlsxwriter
except ImportError:
    print("ERROR CRÍTICO: Falta 'xlsxwriter'. Ejecuta: pip install xlsxwriter")
    sys.exit(1)

# Silenciar advertencias
warnings.filterwarnings("ignore")

# =============================================================================
# CONFIGURACIÓN
# =============================================================================
HOJA_COMANDAS_DESTINO = "Devoluciones"
HOJA_CAJA_ADICION_DESTINO = "Caja Adicion"
HOJA_PAGOS_MP_DESTINO = "Pagos MP" 
HOJA_NO_CONCILIADO_DESTINO = "No conciliado"

FILES_OUTPUT = {}

CONFIG_CONCILIACION = {
    'col_fecha_getnet': 'Fecha de operacion',
    'col_monto_getnet': 'Monto Bruto Transaccion',
    'col_id_getnet': 'Cod de Transaccion',
    'col_fecha_mp': 'FECHA DE ORIGEN',
    'col_monto_mp': 'VALOR DE LA COMPRA',
    'col_id_mp': 'ID DE OPERACIÓN EN MERCADO PAGO',
    'col_fecha_sis': 'Fecha',
    'col_id_sis': 'ID de venta',
    'col_monto_sis': 'Monto Bruto pago',
    'col_plataforma': "Medio de cobro",
    'col_turno': "TURNO",
    'val_getnet': "Getnet",
    'val_mp': "Mercado Pago",
    'val_efectivo': "Efectivo",
    'val_cta_cte': "Cta Cte"
}

REGLAS_CONCILIACION = [
    ("R0: 1min / $1", 1, 1),        
    ("R1: 5min / $5", 5, 5),        
    ("R2: 30min / $5", 30, 5),
    ("R3: Mismo Día / $5", 9999, 5) 
]

KEYWORDS_EXCLUIR = ['interno', 'ingreso extra']

# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================

def seleccionar_archivo(titulo_ventana):
    print(f"--> Por favor, selecciona el archivo: {titulo_ventana}")
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    ruta = filedialog.askopenfilename(title=f"SELECCIONA: {titulo_ventana}", filetypes=[("Archivos Excel", "*.xlsx;*.xls;*.xlsm;*.html")])
    root.destroy()
    if not ruta: sys.exit()
    print(f"    OK: {os.path.basename(ruta)}")
    return ruta

def seleccionar_carpeta(titulo_ventana):
    print(f"--> Por favor, selecciona la CARPETA DE DESTINO.")
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    ruta = filedialog.askdirectory(title=titulo_ventana)
    root.destroy()
    if not ruta: sys.exit()
    print(f"    OK: Carpeta {ruta}")
    return ruta

def parsear_fecha_mp_iso(series):
    s = series.astype(str).str.strip()
    s = s.str.slice(0, 19).str.replace("T", " ", regex=False)
    fechas = pd.to_datetime(s, format='%Y-%m-%d %H:%M:%S', errors='coerce')
    mask_nulos = fechas.isna()
    if mask_nulos.any():
        fechas.loc[mask_nulos] = pd.to_datetime(s[mask_nulos], dayfirst=True, errors='coerce')
    return fechas

def normalizar_fecha_argentina(series):
    if pd.api.types.is_datetime64_any_dtype(series): return series
    s = series.astype(str).str.lower().str.strip()
    basura = ['"', "'", 'p. m.', 'a. m.', 'p.m.', 'a.m.', 'p m', 'a m', '.']
    for b in basura:
        if b in ['p. m.', 'p.m.', 'p m']: s = s.str.replace(b, 'pm', regex=False)
        elif b in ['a. m.', 'a.m.', 'a m']: s = s.str.replace(b, 'am', regex=False)
        else: s = s.str.replace(b, '', regex=False)
    return pd.to_datetime(s, dayfirst=True, errors='coerce')

def convertir_a_string_visual(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return series.dt.strftime('%d/%m/%Y %H:%M:%S').fillna('')
    return series

def formato_visual_columna(df, col_name):
    if col_name in df.columns:
        if not pd.api.types.is_datetime64_any_dtype(df[col_name]):
             try: df[col_name] = pd.to_datetime(df[col_name], dayfirst=True, errors='coerce')
             except: pass
        if pd.api.types.is_datetime64_any_dtype(df[col_name]):
            return df[col_name].dt.strftime('%d/%m/%Y %H:%M:%S').fillna('')
    return df[col_name]

def limpiar_monto_general(valor):
    """
    Lógica de conversión corregida para formato Argentino (Getnet/Sistema).
    Prioriza la coma como separador decimal.
    """
    if pd.isna(valor) or str(valor).strip() == '': return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    
    s = str(valor).strip().replace('$', '').replace('ARS', '').replace(' ', '').replace('\xa0', '')
    
    if ',' in s:
        # Caso 79.200,08 -> Eliminamos el punto de miles, cambiamos coma por punto
        s = s.replace('.', '').replace(',', '.')
    else:
        # Caso 20.000 (sin coma) -> El punto suele ser separador de miles
        if '.' in s:
            partes = s.split('.')
            if len(partes) > 2 or len(partes[1]) == 3:
                s = s.replace('.', '')
            
    try: return float(s)
    except ValueError: return 0.0

def asignar_turno_desde_excel(df_reporte, col_fecha_dt, df_turnos_maestro, col_ap_maestro='Apertura_DT', col_ci_maestro='Cierre_DT'):
    """
    Asigna turno permitiendo elegir qué columnas de apertura/cierre del maestro usar.
    Por defecto usa 'Apertura_DT' y 'Cierre_DT' (Horario Principal).
    """
    if df_turnos_maestro.empty:
        df_reporte['TURNO'] = "Sin Turnos"
        return df_reporte
    
    if not pd.api.types.is_datetime64_any_dtype(df_reporte[col_fecha_dt]):
          df_reporte[col_fecha_dt] = pd.to_datetime(df_reporte[col_fecha_dt], dayfirst=True, errors='coerce')
          
    try:
        if df_reporte[col_fecha_dt].dt.tz is not None:
            df_reporte[col_fecha_dt] = df_reporte[col_fecha_dt].dt.tz_localize(None)
    except: pass

    def find_turno(transaction_time):
        if pd.isna(transaction_time): return "Fecha Nula"
        mask = (transaction_time >= df_turnos_maestro[col_ap_maestro]) & \
                (transaction_time <= df_turnos_maestro[col_ci_maestro])
        matching = df_turnos_maestro.loc[mask, 'TURNO']
        if not matching.empty: return matching.iloc[0]
        return "Fuera de turno"

    df_reporte['TURNO'] = df_reporte[col_fecha_dt].apply(find_turno)
    return df_reporte

def calcular_mascara_exclusion(series_clasificacion):
    s = series_clasificacion.fillna('').astype(str).str.lower().str.strip()
    for char_old, char_new in [('á','a'), ('é','e'), ('í','i'), ('ó','o'), ('ú','u')]:
        s = s.str.replace(char_old, char_new, regex=False)
    return s.isin(KEYWORDS_EXCLUIR)

# =============================================================================
# FASE 1: PROCESAMIENTO
# =============================================================================

def procesar_archivo_turnos(ruta_archivo):
    try: 
        # Leemos el archivo completo
        df = pd.read_excel(ruta_archivo, sheet_name=0)
    except: 
        return pd.DataFrame()
    
    # Normalizamos todas las columnas de fecha posibles (Standard y MP)
    cols_fecha = ['Fecha Apertura', 'Fecha Cierre', 'Fecha Apertura MP', 'Fecha Cierre MP']
    for c in cols_fecha:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce').dt.normalize()

    if 'TURNO' in df.columns:
        df['TURNO'] = df['TURNO'].astype(str).str.strip().str.upper()
    
    # Validamos que al menos existan las columnas principales
    columnas_req = ['Fecha Apertura', 'Hs Ap. Caja', 'Fecha Cierre', 'Hs Cierre Caja', 'TURNO']
    if not all(col in df.columns for col in columnas_req): 
        print("Faltan columnas requeridas en el archivo de Turnos.")
        return pd.DataFrame()
    
    try:
        # 1. Armado Turno STANDARD (Principal) - Se usa para Ventas, Getnet, Comandas, Caja
        ap_str = df['Fecha Apertura'].dt.strftime('%Y-%m-%d') + ' ' + df['Hs Ap. Caja'].astype(str)
        ci_str = df['Fecha Cierre'].dt.strftime('%Y-%m-%d') + ' ' + df['Hs Cierre Caja'].astype(str)
        df['Apertura_DT'] = pd.to_datetime(ap_str, errors='coerce')
        df['Cierre_DT'] = pd.to_datetime(ci_str, errors='coerce')

        # 2. Armado Turno MERCADO PAGO (Secundario) - Se usa SOLO para Pagos MP (Negativos)
        # Verificamos si existen las columnas específicas de MP en el Excel
        if 'Fecha Apertura MP' in df.columns and 'Hs Ap. Caja MP' in df.columns:
            ap_mp_str = df['Fecha Apertura MP'].dt.strftime('%Y-%m-%d') + ' ' + df['Hs Ap. Caja MP'].astype(str)
            ci_mp_str = df['Fecha Cierre MP'].dt.strftime('%Y-%m-%d') + ' ' + df['Hs Cierre Caja MP'].astype(str)
            df['Apertura_MP_DT'] = pd.to_datetime(ap_mp_str, errors='coerce')
            df['Cierre_MP_DT'] = pd.to_datetime(ci_mp_str, errors='coerce')
        else:
            # Fallback: Si no existen las columnas de MP, usamos el horario Standard como respaldo
            # para que el script no falle.
            df['Apertura_MP_DT'] = df['Apertura_DT']
            df['Cierre_MP_DT'] = df['Cierre_DT']

        return df
    except Exception as e: 
        print(f"Error interno procesando turnos: {e}")
        return pd.DataFrame()
def transformar_reporte_getnet(ruta_archivo):
    if not os.path.exists(ruta_archivo): return pd.DataFrame()
    try:
        df = pd.read_excel(ruta_archivo, sheet_name=0, dtype=str)
        
        # 1. NORMALIZACIÓN EXTREMA: Pasamos todo a minúsculas, quitamos espacios extra y borramos tildes
        df.columns = (df.columns.astype(str)
                      .str.strip()
                      .str.lower()
                      .str.replace('á', 'a').str.replace('é', 'e')
                      .str.replace('í', 'i').str.replace('ó', 'o')
                      .str.replace('ú', 'u'))
    except: return pd.DataFrame()
    
    # 2. RENOMBRADO OFICIAL: Forzamos las columnas normalizadas a los nombres exactos y bonitos
    mapeo_columnas = {
        'estado': 'Estado venta',           
        'estado venta': 'Estado venta',
        'fecha de operacion': 'Fecha de operacion',
        'monto bruto transaccion': 'Monto Bruto Transaccion',
        'arancel': 'Arancel',
        'iva arancel': 'IVA Arancel',
        'monto neto transaccion': 'Monto Neto Transaccion',
        'tipo de transaccion': 'Tipo de Transaccion',
        
        # --- VARIACIONES PARA EL ID DE TRANSACCIÓN ---
        'cod de transaccion': 'Cod de Transaccion',
        'codigo de transaccion': 'Cod de Transaccion',
        'id de transaccion': 'Cod de Transaccion',
        'nro de operacion': 'Cod de Transaccion',
        
        # --- NUEVAS COLUMNAS DE GETNET ---
        'nombre del producto': 'Nombre del Producto',
        'dni del pagador': 'DNI del Pagador',
        'correo electronico': 'Correo Electrónico',
        'nombre y apellido': 'Nombre y Apellido',
        'origen': 'Origen'
    }
    # Aplicamos el mapeo.
    df.rename(columns=mapeo_columnas, inplace=True)

    # --- SEGURO ANTI-FALLOS ---
    # Si después de todo el mapeo la columna sigue sin existir, la creamos vacía para evitar crasheos en la FASE 2
    if 'Cod de Transaccion' not in df.columns:
        print("   [!] ADVERTENCIA: No se encontró la columna de ID en Getnet. Se asignará 'Sin ID'.")
        df['Cod de Transaccion'] = "Sin ID"

    # 3. PROCESAMIENTO NORMAL DE FECHAS
    col_fecha = "Fecha de operacion"
    if col_fecha in df.columns:
        df['FECHA_DT'] = normalizar_fecha_argentina(df[col_fecha])
        df[col_fecha] = formato_visual_columna(df, 'FECHA_DT')
    else:
        print("   [!] ADVERTENCIA: No se encontró la columna de fechas en Getnet tras la normalización.")
        df['FECHA_DT'] = pd.NaT
    
    # 4. APLICAR LIMPIEZA DE MONTOS
    cols_money = ["Monto Bruto Transaccion", "Arancel", "IVA Arancel", "Monto Neto Transaccion"]
    for col in cols_money:
        if col in df.columns: df[col] = df[col].apply(limpiar_monto_general)

    # 5. LÓGICA DE ESTADOS Y ANULACIONES
    col_monto_bruto = "Monto Bruto Transaccion"
    col_tipo_tx = "Tipo de Transaccion"
    col_estado_venta = "Estado venta"

    if col_monto_bruto in df.columns:
        es_anulacion = pd.Series(False, index=df.index)
        if col_tipo_tx in df.columns:
            es_anulacion = df[col_tipo_tx].astype(str).str.lower().str.contains('anulaci', na=False)
        
        es_rechazado = pd.Series(False, index=df.index)
        if col_estado_venta in df.columns:
            es_rechazado = df[col_estado_venta].astype(str).str.lower().str.contains('rechazado', na=False)
        
        mask_negativo = es_anulacion | es_rechazado
        df.loc[mask_negativo, col_monto_bruto] = df.loc[mask_negativo, col_monto_bruto].abs() * -1

    return df

def procesar_caja_adicion(ruta_input, ruta_output_xlsm, nombre_hoja_destino, df_turnos_maestro):
    print(f"    -> Procesando Caja Adición hacia '{nombre_hoja_destino}'...")
    if not os.path.exists(ruta_input): return

    try:
        df = pd.read_excel(ruta_input, sheet_name="Hoja1")

        df['Fecha Modificación'] = df['Fecha Modificación'].astype(str).str.replace('"', '', regex=False).str.replace('p. m.', 'PM', regex=False).str.replace('a. m.', 'AM', regex=False).str.upper()
        df['Fecha Modificación'] = pd.to_datetime(df['Fecha Modificación'], dayfirst=True, errors='coerce')
        
        # Caja usa turno PRINCIPAL
        if not df_turnos_maestro.empty:
            df = asignar_turno_desde_excel(df, 'Fecha Modificación', df_turnos_maestro)
        else:
            df['TURNO'] = "Sin Data Turnos"

        df['Fecha'] = df['Fecha Modificación']
        df['Hora'] = df['Fecha Modificación']
        
        df['Fecha Contable'] = pd.to_datetime(df['Fecha Contable'], dayfirst=True, errors='coerce')
        df['Fecha Pago/Venc.'] = pd.to_datetime(df['Fecha Pago/Venc.'], dayfirst=True, errors='coerce')
        
        cols_numericas = ['Monto', 'Monto EDIT.', 'Q.REC', 'Q.FAC', 'PRECIO']
        for col in cols_numericas:
            if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce')

        movimientos_validos = ['Egreso de Dinero', 'Ingreso de Dinero']

        if 'Forma de Pago' in df.columns:
            df_filtrado = df[
                (df['Origen'] == 'Caja') & 
                (df['Proveedor / Para'].isin(movimientos_validos)) &
                (df['Forma de Pago'].astype(str).str.strip().str.lower() == 'efectivo')
            ].copy()
        else:
            df_filtrado = df[
                (df['Origen'] == 'Caja') & 
                (df['Proveedor / Para'].isin(movimientos_validos))
            ].copy()

        columnas_finales = ["Fecha Contable", "Fecha Modificación", "Fecha", "Hora", "TURNO", "Origen", "Clase", "Proveedor / Para", "Monto", "Comentario", "Usuario", "Tipo", "Forma de Pago", "Fecha Pago/Venc.", "N.I.F.", "# Doc", "Cuenta Contable", "Monto EDIT.", "LIN", "Q.REC", "Q.FAC", "PRECIO"]
        df_final = df_filtrado[[c for c in columnas_finales if c in df_filtrado.columns]]

        app = xw.App(visible=False)
        try:
            wb = app.books.open(ruta_output_xlsm)
            try: ws = wb.sheets[nombre_hoja_destino]; ws.clear()
            except: ws = wb.sheets.add(nombre_hoja_destino)
            
            if not df_final.empty:
                ws.range('A1').options(index=False).value = df_final
                ws.autofit()
                
                last_row = len(df_final) + 1
                ws.range(f'B2:B{last_row}').number_format = 'dd/mm/yyyy hh:mm:ss'
                ws.range(f'C2:C{last_row}').number_format = 'dd/mm/yyyy'
                ws.range(f'D2:D{last_row}').number_format = 'hh:mm:ss'
                
            wb.save()
            print("    -> Caja Adición exportada.")
        except Exception as ex: print(f"Error xlwings en Caja Adicion: {ex}")
        finally: 
            try: wb.close(); app.quit()
            except: pass
            
    except Exception as e: print(f"Error procesando Caja Adicion: {e}")

def obtener_df_pagos_mp_negativos(ruta_input_mp, df_turnos_maestro):
    """
    PROCESO ESPECIAL: Usa las columnas SECUNDARIAS de MP para asignar el turno.
    """
    print(f"    -> Procesando Pagos MP (Negativos) con HORARIO MP EXTENDIDO...")
    if not os.path.exists(ruta_input_mp): return pd.DataFrame()

    try:
        df = pd.read_excel(ruta_input_mp)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        col_fecha_mp = 'FECHA DE ORIGEN'
        col_id_mp = 'ID DE OPERACIÓN EN MERCADO PAGO'
        col_monto_mp = 'MONTO NETO DE LA OPERACIÓN QUE IMPACTÓ TU DINERO'

        if col_fecha_mp not in df.columns: return pd.DataFrame()

        df['FECHA_DT_TEMP'] = parsear_fecha_mp_iso(df[col_fecha_mp]) 
        
        # --- ASIGNACIÓN CON COLUMNAS SECUNDARIAS ---
        df = asignar_turno_desde_excel(df, "FECHA_DT_TEMP", df_turnos_maestro, 
                                       col_ap_maestro='Apertura_MP_DT', 
                                       col_ci_maestro='Cierre_MP_DT')
        
        df['Fecha y Hora'] = df['FECHA_DT_TEMP'] 
        df['Fecha'] = df['FECHA_DT_TEMP'].dt.normalize()
        df['Hora'] = df['FECHA_DT_TEMP'] 

        if col_monto_mp in df.columns: df[col_monto_mp] = pd.to_numeric(df[col_monto_mp], errors='coerce').fillna(0)
        
        df_negativos = df[df[col_monto_mp] < 0].copy()
        
        if df_negativos.empty: return pd.DataFrame()
        if col_id_mp in df_negativos.columns: df_negativos[col_id_mp] = df_negativos[col_id_mp].astype(str).str.replace(r'\.0$', '', regex=True)

        columnas_deseadas = ['Fecha y Hora', 'Fecha', 'Hora', col_id_mp, col_monto_mp, 'TURNO']
        return df_negativos[[c for c in columnas_deseadas if c in df_negativos.columns]]

    except Exception as e: 
        print(f"Error procesando Pagos MP: {e}")
        return pd.DataFrame()

def procesar_comandas(ruta_input, ruta_output_xlsm, nombre_hoja_destino, df_turnos_maestro):
    print(f"    -> Procesando Comandas hacia '{nombre_hoja_destino}'...")
    if not os.path.exists(ruta_input): return

    try:
        try: df = pd.read_excel(ruta_input, dtype=str)
        except:
            try: dfs = pd.read_html(ruta_input); df = max(dfs, key=len) if dfs else pd.DataFrame()
            except: return

        cols_actuales = [str(c).strip() for c in df.columns]
        if "ID Comanda" not in cols_actuales:
            for i in range(min(20, len(df))):
                fila = df.iloc[i].astype(str).tolist()
                if "ID Comanda" in fila:
                    df.columns = df.iloc[i]; df = df[i+1:].copy(); df.reset_index(drop=True, inplace=True); break
        df.columns = df.columns.astype(str).str.strip()
        
        if "Precios" in df.columns: df["Precios"] = df["Precios"].apply(limpiar_monto_general).fillna(0).astype('int64')
        if "ID Comanda" in df.columns: df["ID Comanda"] = df["ID Comanda"].astype(str).str.replace(r'\.0$', '', regex=True)

        col_fecha_pedido = "Fecha Hora pedido" 
        if "Hora pedido" in df.columns:
            df["Hora pedido DT"] = normalizar_fecha_argentina(df["Hora pedido"])
            df["Hora pedido"] = df["Hora pedido DT"].dt.strftime('%H:%M:%S').fillna('')
            df[col_fecha_pedido] = df["Hora pedido DT"]
        else: df[col_fecha_pedido] = pd.NaT

        # Comandas usan turno PRINCIPAL
        if not df_turnos_maestro.empty and col_fecha_pedido in df.columns:
            df = asignar_turno_desde_excel(df, col_fecha_pedido, df_turnos_maestro)
        else: df['TURNO'] = "Sin Data Fecha"
            
        cols_finales = ["ID Comanda", "TURNO", "Camarero Mesa", "Mesa", "Producto", "Precios", "Comentario", col_fecha_pedido, "Hora pedido", "Hora Anulación"]
        df_final = df[[c for c in cols_finales if c in df.columns]]

        app = xw.App(visible=False)
        try:
            wb = app.books.open(ruta_output_xlsm)
            try: ws = wb.sheets[nombre_hoja_destino]; ws.clear()
            except: ws = wb.sheets.add(nombre_hoja_destino)
            if not df_final.empty:
                ws.range('A1').options(index=False).value = df_final
                
                cols_validas = [c for c in cols_finales if c in df.columns]
                if col_fecha_pedido in cols_validas:
                    try:
                        idx = cols_validas.index(col_fecha_pedido)
                        letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                        letra = letras[idx] if idx < 26 else "H"
                        last_row = len(df_final) + 1
                        ws.range(f'{letra}2:{letra}{last_row}').number_format = 'dd/mm/yyyy hh:mm:ss'
                    except: pass
                ws.autofit()
            wb.save()
            print("    -> Comandas exportadas.")
        except Exception as ex: print(f"Error xlwings: {ex}")
        finally: 
            try: wb.close(); app.quit()
            except: pass
    except Exception as e: print(f"Error procesando comandas: {e}")

def auditar_duplicados_cruce(df_getnet, df_mp, df_sys):
    print("\n>>> Verificando integridad de cruces (Buscando Duplicados)...")
    alertas = []

    def chequear_tabla(df, nombre_tabla, col_cruce, col_id_principal):
        if df.empty or col_cruce not in df.columns: return
        
        # Filtramos los conciliados que tengan un ID de cruce válido
        mask_validos = (df['Estado'] == 'Conciliado') & df[col_cruce].notna() & (df[col_cruce].astype(str).str.strip() != '') & (df[col_cruce].astype(str).str.strip() != 'nan')
        df_conciliados = df[mask_validos].copy()
        if df_conciliados.empty: return

        # Buscamos si el mismo ID de cruce se repite más de una vez
        duplicados = df_conciliados[df_conciliados.duplicated(subset=[col_cruce], keep=False)].copy()
        
        if not duplicados.empty:
            cols_basicas = [col_id_principal, 'datetime_col', 'TURNO', 'monto_col_numeric', 'Estado', 'Tipo Match', col_cruce]
            cols_disponibles = [c for c in cols_basicas if c in duplicados.columns]

            df_rep = duplicados[cols_disponibles].copy()
            df_rep.insert(0, 'Tabla Origen', nombre_tabla)
            alertas.append(df_rep)

    # 1. Plataformas (Revisamos si apuntan a la misma venta de sistema)
    chequear_tabla(df_getnet, "Getnet", 'ID Venta Sistema (Conc.)', 'Cod de Transaccion')
    chequear_tabla(df_mp, "Mercado Pago", 'ID Venta Sistema (Conc.)', 'ID DE OPERACIÓN EN MERCADO PAGO')

    # 2. Sistema (Revisamos si apuntan a la misma operación de plataforma)
    chequear_tabla(df_sys, "Sistema (Match MP)", 'ID Operación MP (Conc.)', 'ID de venta')
    chequear_tabla(df_sys, "Sistema (Match Getnet)", 'ID Operación Getnet (Conc.)', 'ID de venta')

    if alertas:
        df_final = pd.concat(alertas, ignore_index=True)
        print(f"   [ALERTA ROJA] ⚠️ Se detectaron {len(df_final)} registros involucrados en cruces duplicados.")
        return df_final
    else:
        print("   [OK] ✅ Integridad perfecta: No hay cruces duplicados.")
        return pd.DataFrame()

# =============================================================================
# FASE 2: CONCILIACIÓN
# =============================================================================

def exportar_tablas_a_xlwings(tablas, ruta_macro, nombre_hoja):
    print(f"    -> Exportando tablas a MACRO ({nombre_hoja})...")
    if not os.path.exists(ruta_macro): return
    app = xw.App(visible=False)
    try:
        wb = app.books.open(ruta_macro)
        try: ws = wb.sheets[nombre_hoja]; ws.clear()
        except: ws = wb.sheets.add(nombre_hoja)
        current_col = 1
        for df_tabla, titulo in tablas:
            df_tabla = df_tabla.fillna('')
            ws.range((1, current_col)).value = titulo
            try: ws.range((1, current_col)).api.Font.Bold = True
            except: pass
            if not df_tabla.empty:
                ws.range((2, current_col)).options(index=False).value = df_tabla
                current_col += len(df_tabla.columns) + 1
            else:
                ws.range((3, current_col)).value = "Sin datos"; current_col += 2
        ws.autofit(); wb.save()
    except Exception as e: print(f"Error xlwings: {e}")
    finally: 
        try: wb.close(); app.quit()
        except: pass

def generar_reporte_plano(df_getnet, df_mp, df_sistema, writer, config):
    print("Generando Reporte Estadístico...")
    try:
        col_turno = config['col_turno']
        col_plat_sis = config['col_plataforma']
        def garantizar_cols(df, cols):
            for c in cols:
                if c not in df.columns: df[c] = "Sin Dato"
            return df

        df_g = garantizar_cols(df_getnet.copy(), [col_turno, 'Estado', 'datetime_col', 'monto_col_numeric'])
        df_m = garantizar_cols(df_mp.copy(), [col_turno, 'Estado', 'datetime_col', 'monto_col_numeric'])
        df_s = garantizar_cols(df_sistema.copy(), [col_turno, 'Estado', 'datetime_col', 'monto_col_numeric', col_plat_sis])
        
        df_g['Plataforma'] = 'Getnet'; df_g['Medio de cobro'] = 'Getnet'
        df_m['Plataforma'] = 'Mercado Pago'; df_m['Medio de cobro'] = 'Mercado Pago'
        df_s['Plataforma'] = 'Sistema'; df_s['Medio de cobro'] = df_s[col_plat_sis]

        df_g['Monto'] = df_g['monto_col_numeric']
        df_m['Monto'] = df_m['monto_col_numeric']
        df_s['Monto'] = df_s['monto_col_numeric']
        
        try: df_g['Fecha_Grp'] = df_g['datetime_col'].dt.strftime('%d/%m/%Y')
        except: df_g['Fecha_Grp'] = "Sin Fecha"
        try: df_m['Fecha_Grp'] = df_m['datetime_col'].dt.strftime('%d/%m/%Y')
        except: df_m['Fecha_Grp'] = "Sin Fecha"
        try: df_s['Fecha_Grp'] = df_s['datetime_col'].dt.strftime('%d/%m/%Y')
        except: df_s['Fecha_Grp'] = "Sin Fecha"
        
        df_g['Estado_Simple'] = df_g['Estado'].apply(lambda x: 'Conciliado' if x == 'Conciliado' else 'No Conciliado')
        df_m['Estado_Simple'] = df_m['Estado'].apply(lambda x: 'Conciliado' if x == 'Conciliado' else 'No Conciliado')
        df_s['Estado_Simple'] = df_s['Estado'].apply(lambda x: 'Conciliado' if x == 'Conciliado' else 'No Conciliado')

        cols_req = ['Fecha_Grp', 'TURNO', 'Plataforma', 'Medio de cobro', 'Monto', 'Estado_Simple']
        dfs_to_concat = [d[cols_req] for d in [df_g, df_m, df_s] if not d.empty]
        
        if dfs_to_concat:
            df_total = pd.concat(dfs_to_concat, ignore_index=True)
            df_total['Cantidad'] = 1

            df_agg = df_total.groupby(['Fecha_Grp', 'TURNO', 'Plataforma', 'Medio de cobro', 'Estado_Simple']).agg(
                Monto_Total=('Monto', 'sum'), Cantidad_Total=('Cantidad', 'count')
            ).reset_index()

            df_p_m = df_agg.pivot_table(index=['Fecha_Grp', 'TURNO', 'Plataforma', 'Medio de cobro'], columns='Estado_Simple', values='Monto_Total').fillna(0)
            df_p_c = df_agg.pivot_table(index=['Fecha_Grp', 'TURNO', 'Plataforma', 'Medio de cobro'], columns='Estado_Simple', values='Cantidad_Total').fillna(0)
            
            df_p_m.columns = [f"Total Monto {col}" for col in df_p_m.columns]
            df_p_c.columns = [f"Cantidad {col}" for col in df_p_c.columns]
            
            df_reporte = df_p_m.join(df_p_c).reset_index().rename(columns={'Fecha_Grp': 'Fecha'})
            
            for col in ['Total Monto Conciliado', 'Total Monto No Conciliado', 'Cantidad Conciliado', 'Cantidad No Conciliado']:
                if col not in df_reporte.columns: df_reporte[col] = 0
                
            df_reporte['Total Monto'] = df_reporte['Total Monto Conciliado'] + df_reporte['Total Monto No Conciliado']
            df_reporte['Total Cantidad'] = df_reporte['Cantidad Conciliado'] + df_reporte['Cantidad No Conciliado']
            
            df_reporte.fillna(0).to_excel(writer, sheet_name='Reporte_Estadistico_Plano', index=False)
    except Exception as e: print(f"Error Reporte: {e}"); traceback.print_exc()

def generar_hoja_auditoria(df_getnet, df_mp, df_sistema, writer, config):
    print("Generando Auditoría...")
    try:
        col_turno = config['col_turno']
        col_plat = config['col_plataforma']
        dfs = []

        if not df_getnet.empty:
            dg = df_getnet[df_getnet['Estado'] == 'No Conciliado'].copy()
            dg['Origen'] = 'Getnet'; dg['Fecha y Hora'] = convertir_a_string_visual(dg['datetime_col'])
            dg = dg.rename(columns={config['col_monto_getnet']: 'Monto', config['col_id_getnet']: 'ID_Operacion', 'Estado_Auditoria': 'Estado_Explicacion', col_turno: 'Turno'})
            dfs.append(dg)
        
        if not df_mp.empty:
            mask_excluir = calcular_mascara_exclusion(df_mp['Clasificacion'])
            dm = df_mp[(df_mp['Estado'] == 'No Conciliado') & (~mask_excluir)].copy()
            dm['Origen'] = 'Mercado Pago'; dm['Fecha y Hora'] = convertir_a_string_visual(dm['datetime_col'])
            dm = dm.rename(columns={config['col_monto_mp']: 'Monto', config['col_id_mp']: 'ID_Operacion', 'Estado_Auditoria': 'Estado_Explicacion', col_turno: 'Turno'})
            dfs.append(dm)

        if not df_sistema.empty:
            ds = df_sistema[df_sistema['Estado'] == 'No Conciliado'].copy()
            ds['Origen'] = 'Sistema (' + ds[col_plat].astype(str) + ')'; ds['Fecha y Hora'] = convertir_a_string_visual(ds['datetime_col'])
            ds = ds.rename(columns={config['col_monto_sis']: 'Monto', config['col_id_sis']: 'ID_Operacion', 'Estado_Auditoria': 'Estado_Explicacion', col_turno: 'Turno'})
            dfs.append(ds)
        
        if dfs:
            df_audit = pd.concat(dfs, ignore_index=True)
            df_audit.to_excel(writer, sheet_name='Auditoría_Pendientes', index=False)
    except Exception as e: print(f"Error Auditoria: {e}")

def generar_hoja_tablas_revision_y_exportar(df_getnet, df_mp, df_sistema, writer, config):
    print("Generando Tablas Revisión...")
    def preparar_tabla_export(df_source, col_id_primary, titulo, incluir_detalles_cruce=False):
        if df_source.empty:
            cols = ['ID', 'Fecha y Hora', 'Fecha', 'Hora', 'TURNO', 'Monto']
            if incluir_detalles_cruce: cols += ['Tipo Match', 'Plataforma Cruce']
            return pd.DataFrame(columns=cols), titulo
            
        df_temp = df_source.copy()
        if col_id_primary not in df_temp.columns: df_temp[col_id_primary] = "Sin ID"
        df_temp['ID'] = df_temp[col_id_primary]
        df_temp['Fecha y Hora'] = df_temp['datetime_col']
        try: df_temp['Fecha'] = df_temp['datetime_col'].dt.date
        except: df_temp['Fecha'] = pd.NaT
        try: df_temp['Hora'] = df_temp['datetime_col'].dt.strftime('%H:%M:%S')
        except: df_temp['Hora'] = ""
        df_temp['Monto'] = df_temp['monto_col_numeric']
        
        cols_finales = ['ID', 'Fecha y Hora', 'Fecha', 'Hora', 'TURNO', 'Monto']
        if incluir_detalles_cruce:
            def get_platform_name(row):
                if 'ID Operación MP (Conc.)' in row.index and pd.notna(row['ID Operación MP (Conc.)']): return "Mercado Pago"
                if 'ID Operación Getnet (Conc.)' in row.index and pd.notna(row['ID Operación Getnet (Conc.)']): return "Getnet"
                return ""
            df_temp['Plataforma Cruce'] = df_temp.apply(get_platform_name, axis=1)
            cols_finales.extend(['Tipo Match', 'Plataforma Cruce'])

        return df_temp[[c for c in cols_finales if c in df_temp.columns]], titulo

    mask_excluir = calcular_mascara_exclusion(df_mp['Clasificacion'])
    df_mp_no_conc = df_mp[(df_mp['Estado'] == 'No Conciliado') & (~mask_excluir)]
    
    t1, h1 = preparar_tabla_export(df_mp_no_conc, config['col_id_mp'], "MP NO CONCILIADO")
    t2, h2 = preparar_tabla_export(df_getnet[df_getnet['Estado'] == 'No Conciliado'], config['col_id_getnet'], "GETNET NO CONCILIADO")
    t3, h3 = preparar_tabla_export(df_sistema[(df_sistema[config['col_plataforma']] == config['val_mp']) & (df_sistema['Estado'] == 'No Conciliado')], config['col_id_sis'], "SISTEMA (MP) NO CONCILIADO")
    t4, h4 = preparar_tabla_export(df_sistema[(df_sistema[config['col_plataforma']] == config['val_getnet']) & (df_sistema['Estado'] == 'No Conciliado')], config['col_id_sis'], "SISTEMA (GETNET) NO CONCILIADO")
    
    df_cruzados = df_sistema[(df_sistema['Estado'] == 'Conciliado') & (df_sistema['Tipo Match'].astype(str).str.contains('Global', case=False, na=False))]
    t5, h5 = preparar_tabla_export(df_cruzados, config['col_id_sis'], "MATCHES CRUZADOS / GLOBALES", True)
    t6, h6 = preparar_tabla_export(df_sistema[(df_sistema[config['col_plataforma']] == config['val_efectivo']) & (df_sistema['Estado'] == 'Conciliado')], config['col_id_sis'], "EFECTIVO CONCILIADO", True)
    t7, h7 = preparar_tabla_export(df_sistema[(df_sistema[config['col_plataforma']] == config['val_cta_cte']) & (df_sistema['Estado'] == 'Conciliado')], config['col_id_sis'], "CTA CTE CONCILIADO", True)

    tablas = [(t1, h1), (t2, h2), (t3, h3), (t4, h4), (t5, h5), (t6, h6), (t7, h7)]
    exportar_tablas_a_xlwings(tablas, FILES_OUTPUT['archivo_macro'], HOJA_NO_CONCILIADO_DESTINO)

def marcar_match(df_p, idx_p, df_s, idx_s, tipo, col_id_p, col_id_s, col_conc_p, col_conc_s):
    df_p.at[idx_p, 'Estado'] = 'Conciliado'
    df_s.at[idx_s, 'Estado'] = 'Conciliado'
    df_p.at[idx_p, 'Tipo Match'] = tipo
    df_s.at[idx_s, 'Tipo Match'] = tipo
    df_p.at[idx_p, col_conc_p] = df_s.at[idx_s, col_id_s]
    df_s.at[idx_s, col_conc_s] = df_p.at[idx_p, col_id_p]

def correr_conciliacion(df_plat, df_sis, nombre_fase, val_sis, config, indices_usados, ignorar_medio=False):
    col_turno = config['col_turno']; col_plat_sis = config['col_plataforma']
    print(f"  Fase {nombre_fase}...")
    for regla, min_tol, money_tol in REGLAS_CONCILIACION:
        matches = 0; is_date_only_match = regla.startswith("R3:")
        for i_p in df_plat.index:
            if df_plat.at[i_p, 'Estado'] != 'No Conciliado': continue
            t_p = df_plat.at[i_p, 'datetime_col']; m_p = df_plat.at[i_p, 'monto_col_numeric']; turno_p = df_plat.at[i_p, col_turno]
            if pd.isna(t_p) or pd.isna(m_p): continue

            mask = ~df_sis.index.isin(indices_usados) & (df_sis['Estado'] == 'No Conciliado') & (df_sis[col_turno] == turno_p)
            if not ignorar_medio: mask = mask & (df_sis[col_plat_sis] == val_sis)
            candidatos = df_sis[mask]
            
            if is_date_only_match:
                if t_p is pd.NaT: continue
                match = candidatos[(candidatos['datetime_col'].dt.date == t_p.date()) & (candidatos['monto_col_numeric'].between(m_p - money_tol, m_p + money_tol))]
            else:
                match = candidatos[(candidatos['datetime_col'].between(t_p - pd.Timedelta(minutes=min_tol), t_p + pd.Timedelta(minutes=min_tol))) & (candidatos['monto_col_numeric'].between(m_p - money_tol, m_p + money_tol))]
            
            if not match.empty:
                i_s = match.index[0]
                es_mp = 'MP' in nombre_fase
                col_id_p = config['col_id_mp'] if es_mp else config['col_id_getnet']
                col_conc_p = 'ID Venta Sistema (Conc.)'; col_conc_s = 'ID Operación MP (Conc.)' if es_mp else 'ID Operación Getnet (Conc.)'
                marcar_match(df_plat, i_p, df_sis, i_s, f"{nombre_fase} ({regla})", col_id_p, config['col_id_sis'], col_conc_p, col_conc_s)
                indices_usados.add(i_s); matches += 1
        if matches > 0: print(f"    -> {matches} matches con regla: {regla}")

if __name__ == "__main__":
    print("=================================================================")
    print("   ASISTENTE DE CONCILIACIÓN (V49 - FINAL FULL DEFINITIVA)")
    print("=================================================================\n")

    ruta_turnos = seleccionar_archivo("Archivo de TURNOS")
    ruta_ventas = seleccionar_archivo("Reporte de VENTAS (Sistema)")
    ruta_getnet = seleccionar_archivo("Reporte de GETNET")
    ruta_mp = seleccionar_archivo("Reporte de MERCADO PAGO")
    ruta_comandas = seleccionar_archivo("Archivo de COMANDAS / DEVOLUCIONES")
    ruta_caja_adicion = seleccionar_archivo("Archivo CAJA ADICION")
    ruta_macro = seleccionar_archivo("Archivo MACRO EXCEL")
    output_folder = seleccionar_carpeta("Carpeta donde guardar el REPORTE FINAL")
    
    FILES_OUTPUT['turnos_proc'] = os.path.join(output_folder, "Turnos_Procesados.xlsx")
    FILES_OUTPUT['informe_gral'] = os.path.join(output_folder, "informes_parador.xlsx")
    FILES_OUTPUT['reporte_final'] = os.path.join(output_folder, "Resultado_Conciliacion.xlsx")
    FILES_OUTPUT['archivo_macro'] = ruta_macro

    dict_clasificaciones_previas = {}
    if os.path.exists(FILES_OUTPUT['reporte_final']):
        print(">>> Detectado archivo previo de conciliación. Guardando clasificaciones manuales...")
        try:
            df_old = pd.read_excel(FILES_OUTPUT['reporte_final'], sheet_name="MP Conciliado")
            col_clasif = 'Clasificacion'
            col_id_mp = CONFIG_CONCILIACION['col_id_mp']
            if col_clasif in df_old.columns and col_id_mp in df_old.columns:
                df_con_datos = df_old[df_old[col_clasif].notna() & (df_old[col_clasif].astype(str).str.strip() != '')].copy()
                if not df_con_datos.empty:
                    df_con_datos[col_id_mp] = df_con_datos[col_id_mp].astype(str).str.replace(r'\.0$', '', regex=True)
                    dict_clasificaciones_previas = df_con_datos.set_index(col_id_mp)[col_clasif].to_dict()
                    print(f"    -> Se recuperaron {len(dict_clasificaciones_previas)} clasificaciones manuales.")
        except Exception as e:
            print(f"    [AVISO] No se pudo leer el archivo previo: {e}")

    print("\n--- INICIANDO PROCESO ---")
    try:
        print(">>> FASE 1: PROCESAMIENTO")
        df_turnos = procesar_archivo_turnos(ruta_turnos)
        if df_turnos.empty: raise Exception("Error procesando Turnos.")
        
        # Exportar turnos incluyendo columnas nuevas para control
        with pd.ExcelWriter(FILES_OUTPUT['turnos_proc']) as w: 
            df_turnos.to_excel(w, index=False)
            print("    -> Archivo 'Turnos_Procesados.xlsx' actualizado (Todas las columnas incluidas).")

        # 1. GETNET (PRINCIPAL - Con Fix Decimales)
        df_gn = transformar_reporte_getnet(ruta_getnet)
        if not df_gn.empty: df_gn = asignar_turno_desde_excel(df_gn, "FECHA_DT", df_turnos)

        # 2. SISTEMA (PRINCIPAL)
        try:
            print("Procesando Ventas Sistema...")
            df_v = pd.read_excel(ruta_ventas, dtype={'FechaCierre': str, 'Comanda': str})
            
            if 'FechaCierre' in df_v.columns:
                df_v['FechaCierre'] = df_v['FechaCierre'].astype(str).str.replace('"', '', regex=False).str.strip()
                df_v['Fecha_DT'] = normalizar_fecha_argentina(df_v['FechaCierre']) 
                df_v['Fecha_Visual'] = convertir_a_string_visual(df_v['Fecha_DT']) 
            
            # --- CORRECCIÓN: LÓGICA ORIGINAL DE DINAMIZACIÓN ---
            # 1. Definimos las columnas que deben mantenerse FIJAS (no convertirse en filas)
            cols_id = ["FechaCierre", "Fecha_DT", "Fecha_Visual", "Comanda", "Pago", "Total", "Descuentos", "A Pagar", "Propina", "Pagos", "Boleta"]
            # Seleccionamos solo las que realmente vienen en el Excel
            vars_id_existentes = [c for c in cols_id if c in df_v.columns]
            
            # 2. Hacemos el melt respetando esas columnas
            df_v = pd.melt(df_v, id_vars=vars_id_existentes, var_name="Atributo", value_name="Valor")
            
            # 3. Limpieza de valores
            df_v['Valor'] = pd.to_numeric(df_v['Valor'], errors='coerce')
            df_v = df_v[df_v['Valor'] != 0].dropna(subset=['Valor'])
            
            # 4. Eliminamos filas basura que se generan si el Excel trae columnas extras
            df_v = df_v[~df_v['Atributo'].isin(['Caja', '#Boleta', 'Total', 'A Pagar'])]
            
            # 5. Renombrado y Asignación de Turno
            df_v = df_v.rename(columns={"Fecha_Visual": "Fecha", "Comanda": "ID de venta", "Valor": "Monto Bruto pago", "Atributo": "Medio de cobro"})
            df_v = asignar_turno_desde_excel(df_v, "Fecha_DT", df_turnos) 
            
        except Exception as e: 
            print(f"Error Ventas: {e}")
            traceback.print_exc()
            df_v = pd.DataFrame()
        # 3. MERCADO PAGO (PRINCIPAL - Para Conciliación)
        try:
            print("Procesando Mercado Pago (Conciliación)...")
            df_mp = pd.read_excel(ruta_mp)
            df_mp.columns = df_mp.columns.astype(str).str.strip().str.upper()

            # Exclusión de POS específico
            serial_excluir = 'SMARTPOS1493608722'
            if 'NÚMERO DE SERIE DEL LECTOR (S/N)' in df_mp.columns:
                df_mp = df_mp[df_mp['NÚMERO DE SERIE DEL LECTOR (S/N)'].astype(str).str.strip() != serial_excluir]

            col_fecha_mp = 'FECHA DE ORIGEN'
            if col_fecha_mp in df_mp.columns:
                df_mp['FECHA_DT'] = parsear_fecha_mp_iso(df_mp[col_fecha_mp])
                df_mp[col_fecha_mp] = convertir_a_string_visual(df_mp['FECHA_DT'])
            else: df_mp['FECHA_DT'] = pd.NaT
            
            if 'VALOR DE LA COMPRA' in df_mp.columns: 
                df_mp['VALOR DE LA COMPRA'] = pd.to_numeric(df_mp['VALOR DE LA COMPRA'], errors='coerce').fillna(0)
            
            mask_positivo = df_mp['VALOR DE LA COMPRA'] > 0
            df_mp = df_mp[mask_positivo & (df_mp['MEDIO DE PAGO'].astype(str).str.lower() != "nan")]
            if 'ID DE OPERACIÓN EN MERCADO PAGO' in df_mp.columns: 
                df_mp['ID DE OPERACIÓN EN MERCADO PAGO'] = df_mp['ID DE OPERACIÓN EN MERCADO PAGO'].astype(str).str.replace(r'\.0$', '', regex=True)
            df_mp = asignar_turno_desde_excel(df_mp, "FECHA_DT", df_turnos)
        except Exception as e: print(f"Error MP: {e}"); df_mp = pd.DataFrame()

        # 4. MERCADO PAGO (SECUNDARIO - Pagos Negativos)
        df_mp_negativos = obtener_df_pagos_mp_negativos(ruta_mp, df_turnos)

        # GUARDAR INFO GRAL CON FORMATO
        with pd.ExcelWriter(FILES_OUTPUT['informe_gral'], engine='xlsxwriter', datetime_format='dd/mm/yyyy hh:mm:ss') as writer:
            if not df_gn.empty: df_gn.to_excel(writer, sheet_name="Ventas_Getnet", index=False)
            if not df_v.empty: df_v.to_excel(writer, sheet_name="Ventas_Sistema", index=False)
            if not df_mp.empty: df_mp.to_excel(writer, sheet_name="Ventas_MP", index=False)
            if not df_mp_negativos.empty: 
                df_mp_negativos.to_excel(writer, sheet_name="Pagos_MP_Para_Copiar", index=False)
                # Formato visual
                workbook = writer.book
                worksheet = writer.sheets["Pagos_MP_Para_Copiar"]
                fmt_fh = workbook.add_format({'num_format': 'dd/mm/yyyy hh:mm:ss'})
                worksheet.set_column('A:A', 22, fmt_fh)
        
        # PROCESAR MACROS (PRINCIPAL)
        procesar_comandas(ruta_comandas, FILES_OUTPUT['archivo_macro'], HOJA_COMANDAS_DESTINO, df_turnos)
        procesar_caja_adicion(ruta_caja_adicion, FILES_OUTPUT['archivo_macro'], HOJA_CAJA_ADICION_DESTINO, df_turnos)
        
        print(">>> FASE 1 OK.")
    except Exception as e: print(f"ERROR FASE 1: {e}"); traceback.print_exc(); sys.exit(1)

    try:
        print("\n>>> FASE 2: CONCILIACIÓN")
        cfg = CONFIG_CONCILIACION
        xls_obj = pd.ExcelFile(FILES_OUTPUT['informe_gral'])
        def leer_hoja_segura(xls, nombre): return pd.read_excel(xls, sheet_name=nombre) if nombre in xls.sheet_names else pd.DataFrame()

        df_getnet = leer_hoja_segura(xls_obj, "Ventas_Getnet")
        df_mp = leer_hoja_segura(xls_obj, "Ventas_MP")
        df_sis = leer_hoja_segura(xls_obj, "Ventas_Sistema")

        for df, c_fecha, c_monto in [(df_getnet, cfg['col_fecha_getnet'], cfg['col_monto_getnet']), (df_mp, cfg['col_fecha_mp'], cfg['col_monto_mp']), (df_sis, 'Fecha', cfg['col_monto_sis'])]:
            if cfg['col_turno'] not in df.columns: df[cfg['col_turno']] = "Sin Turno"
            df['datetime_col'] = pd.to_datetime(df.get(c_fecha, pd.Series()), dayfirst=True, errors='coerce') 
            df['monto_col_numeric'] = pd.to_numeric(df.get(c_monto, pd.Series()), errors='coerce')
            df['Estado'] = 'No Conciliado'; df['Tipo Match'] = pd.NA
            if df is df_sis: df[cfg['col_plataforma']] = df.get(cfg['col_plataforma'], pd.Series()).fillna("-").astype(str)

        indices_usados = set()
        if not df_mp.empty and not df_sis.empty: correr_conciliacion(df_mp, df_sis, "F1: MP", cfg['val_mp'], cfg, indices_usados)
        if not df_getnet.empty and not df_sis.empty: correr_conciliacion(df_getnet, df_sis, "F2: GN", cfg['val_getnet'], cfg, indices_usados)
        if not df_mp.empty and not df_sis.empty: correr_conciliacion(df_mp, df_sis, "F3: MP vs Efec", "Efectivo", cfg, indices_usados)
        if not df_getnet.empty and not df_sis.empty: correr_conciliacion(df_getnet, df_sis, "F4: GN vs Efec", "Efectivo", cfg, indices_usados)
        if not df_mp.empty and not df_sis.empty: correr_conciliacion(df_mp, df_sis, "F5: MP vs Cta Cte", cfg['val_cta_cte'], cfg, indices_usados)
        if not df_getnet.empty and not df_sis.empty: correr_conciliacion(df_getnet, df_sis, "F6: GN vs Cta Cte", cfg['val_cta_cte'], cfg, indices_usados)
        if not df_mp.empty and not df_sis.empty: correr_conciliacion(df_mp, df_sis, "F7: MP Global", None, cfg, indices_usados, ignorar_medio=True)
        if not df_getnet.empty and not df_sis.empty: correr_conciliacion(df_getnet, df_sis, "F8: GN Global", None, cfg, indices_usados, ignorar_medio=True)
        
        if dict_clasificaciones_previas:
            df_mp['Clasificacion'] = df_mp[cfg['col_id_mp']].astype(str).str.replace(r'\.0$', '', regex=True).map(dict_clasificaciones_previas)
        else: df_mp['Clasificacion'] = ""
        
        # --- NUEVO: AUDITORÍA DE DUPLICADOS ---
        df_alertas = auditar_duplicados_cruce(df_getnet, df_mp, df_sis)

        with pd.ExcelWriter(FILES_OUTPUT['reporte_final']) as writer:
            
            # --- NUEVO: GUARDAR HOJA DE ALERTAS ---
            if not df_alertas.empty:
                df_alertas.to_excel(writer, sheet_name='🚨 Alertas Duplicados', index=False)
            else:
                pd.DataFrame({'Estado de Auditoría': ['✅ No se encontraron cruces duplicados. Integridad perfecta.']}).to_excel(writer, sheet_name='🚨 Alertas Duplicados', index=False)
            # --------------------------------------
            
            generar_reporte_plano(df_getnet, df_mp, df_sis, writer, cfg)
            generar_hoja_auditoria(df_getnet, df_mp, df_sis, writer, cfg)
            
            # Limpieza visual final
            df_getnet['Fecha (Solo)'] = df_getnet['datetime_col'].dt.strftime('%d/%m/%Y')
            df_mp['Fecha (Solo)'] = df_mp['datetime_col'].dt.strftime('%d/%m/%Y')
            df_sis['Fecha (Solo)'] = df_sis['datetime_col'].dt.strftime('%d/%m/%Y')
            col_id_mp = cfg['col_id_mp']
            if col_id_mp in df_mp.columns: df_mp[col_id_mp] = df_mp[col_id_mp].astype(str).str.replace(r'\.0$', '', regex=True)
            if 'ID Operación MP (Conc.)' in df_sis.columns: df_sis['ID Operación MP (Conc.)'] = df_sis['ID Operación MP (Conc.)'].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')
            
            cols_drop = ['datetime_col', 'monto_col_numeric', 'Estado_Auditoria', 'FECHA_DT', 'Fecha_DT'] 
            df_getnet.drop(columns=cols_drop, errors='ignore').to_excel(writer, sheet_name="Getnet Conciliado", index=False)
            df_mp.drop(columns=cols_drop, errors='ignore').to_excel(writer, sheet_name="MP Conciliado", index=False)
            df_sis.drop(columns=cols_drop, errors='ignore').to_excel(writer, sheet_name="Sistema Conciliado", index=False)

            # Formato de ancho para que la hoja de alertas se vea bien
            for sheet in writer.sheets.values(): sheet.set_column('A:A', 20)
            
            generar_hoja_tablas_revision_y_exportar(df_getnet, df_mp, df_sis, writer, cfg)