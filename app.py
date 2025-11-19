import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import io
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import re
import numpy as np

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="BLUSH - Sistema de Comisiones",
    page_icon="💇‍♀️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS ---
st.markdown("""
<style>
    .main-header {
        font-family: 'Helvetica Neue', sans-serif;
        font-size: 2.5rem;
        font-weight: 700;
        color: #FFFFFF;
        text-align: center;
        padding: 1.5rem;
        background: linear-gradient(90deg, #E91E63 0%, #FF4081 100%);
        border-radius: 15px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(233, 30, 99, 0.3);
    }
    .stButton>button {
        background-color: #E91E63;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        border: none;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #C2185B;
        transform: translateY(-2px);
    }
</style>
""", unsafe_allow_html=True)

# --- CONSTANTES ---
PALABRAS_PRODUCTO = [
    'MASCARILLA', 'SHAMPOO', 'SHAMPO', 'ACONDICIONADOR',
    'CREMA', 'SERUM', 'AMPOLLA', 'SPRAY', 'GEL',
    'LOTION', 'REDKEN', 'LOREAL', 'TIGI', 'KERASTASE',
    'X250ML', 'X300ML', 'X500ML', 'ML', 'GR',
    'BED HEAD', 'ALL SOFT', 'FRIZZ DISMISS', 'TRATAMIENTO'
]
REGEX_PRODUCTOS = '|'.join([re.escape(p) for p in PALABRAS_PRODUCTO])

# --- LÓGICA DE NEGOCIO INTELIGENTE ---

@st.cache_data(show_spinner=False)
def procesar_datos(uploaded_file):
    try:
        xl = pd.ExcelFile(uploaded_file)
        # Buscar hoja que contenga "ventas" o "hoja1"
        sheet_name = next((s for s in xl.sheet_names if 'ventas' in s.lower() or 'hoja1' in s.lower()), xl.sheet_names[0])
        
        # 1. DETECCIÓN INTELIGENTE DE CABECERA
        # Leemos las primeras 20 filas sin cabecera para encontrar dónde están los títulos
        df_preview = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, nrows=20)
        
        header_row_idx = None
        for idx, row in df_preview.iterrows():
            row_text = row.astype(str).str.upper().str.strip().tolist()
            # Buscamos fila que tenga 'EMPLEADO' y alguna variante de 'TOTAL'
            if 'EMPLEADO' in row_text and any(x in row_text for x in ['TOTAL', 'TOTAL COMP', 'IMPORTE']):
                header_row_idx = idx
                break
        
        if header_row_idx is None:
            # Si falla la detección automática, probamos el default (fila 9)
            header_row_idx = 9

        # 2. CARGA REAL DE DATOS
        uploaded_file.seek(0) # Resetear archivo
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=header_row_idx)
        
        # 3. NORMALIZACIÓN DE COLUMNAS (Aquí estaba el error)
        # Convertimos todo a mayúsculas y quitamos espacios extra
        df.columns = df.columns.str.strip().str.upper()
        
        # Diccionario de sinónimos para arreglar nombres incorrectos
        renames = {
            'TOTAL COMP': 'TOTAL',   # <--- ESTO ARREGLA TU ERROR
            'IMPORTE': 'TOTAL',
            'FECHA REGISTRO': 'FECHA',
            'PRODUCTO/SERVICIO': 'PRODUCTO / SERVICIO'
        }
        df = df.rename(columns=renames)
        
        # 4. VALIDACIÓN
        if 'TOTAL' not in df.columns or 'EMPLEADO' not in df.columns:
            st.error(f"⚠️ No encontré las columnas 'EMPLEADO' o 'TOTAL/TOTAL COMP'. Columnas detectadas: {list(df.columns)}")
            return pd.DataFrame()

        # 5. LIMPIEZA Y CÁLCULOS (Igual que antes)
        df = df[df['EMPLEADO'].notna()].copy()
        df['EMPLEADO'] = df['EMPLEADO'].astype(str).str.strip().str.title()
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce').ffill()
        df['MONTO'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
        
        if 'CLASE' not in df.columns: df['CLASE'] = ''
        if 'PRODUCTO / SERVICIO' not in df.columns: df['PRODUCTO / SERVICIO'] = ''
        
        # Detección Productos
        df['ITEM_UPPER'] = df['PRODUCTO / SERVICIO'].astype(str).str.upper()
        df['CLASE_UPPER'] = df['CLASE'].astype(str).str.upper().str.strip()
        
        cond_clase = df['CLASE_UPPER'] == 'PRODUCTO'
        cond_key = df['ITEM_UPPER'].str.contains(REGEX_PRODUCTOS, regex=True, na=False)
        df['ES_PRODUCTO'] = cond_clase | cond_key
        
        # Comisiones
        c_producto = df['ES_PRODUCTO']
        c_julio_serv = (df['EMPLEADO'] == 'Julio') & (~df['ES_PRODUCTO'])
        c_jy_corte = (df['EMPLEADO'].isin(['Jhon', 'Yuri'])) & (~df['ES_PRODUCTO']) & (
            df['ITEM_UPPER'].str.contains('CORTE|BARBERIA', regex=True)
        )
        
        condiciones = [c_producto, c_julio_serv, c_jy_corte]
        pcts = [0.10, 0.40, 0.35]
        labels = ["Producto 10%", "Servicio 40%", "Corte 35%"]
        
        df['PORCENTAJE'] = np.select(condiciones, pcts, default=0.25)
        df['TIPO_COMISION'] = np.select(condiciones, labels, default="Servicio 25%")
        df['COMISION'] = df['MONTO'] * df['PORCENTAJE']
        
        return df.drop(columns=['ITEM_UPPER', 'CLASE_UPPER'])

    except Exception as e:
        st.error(f"Error leyendo el archivo: {str(e)}")
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
def generar_resumen(df):
    if df.empty: return pd.DataFrame()
    
    resumen = df.groupby('EMPLEADO').agg(
        TOTAL_PRODUCCION=('MONTO', 'sum'),
        TOTAL_COMISION=('COMISION', 'sum'),
        NUM_TRANSACCIONES=('MONTO', 'count'),
        PROD_SERVICIOS=('MONTO', lambda x: x[~df.loc[x.index, 'ES_PRODUCTO']].sum()),
        COM_SERVICIOS=('COMISION', lambda x: x[~df.loc[x.index, 'ES_PRODUCTO']].sum()),
        PROD_PRODUCTOS=('MONTO', lambda x: x[df.loc[x.index, 'ES_PRODUCTO']].sum()),
        COM_PRODUCTOS=('COMISION', lambda x: x[df.loc[x.index, 'ES_PRODUCTO']].sum())
    ).reset_index()
    
    resumen['PARTICIPACION'] = (resumen['TOTAL_PRODUCCION'] / resumen['TOTAL_PRODUCCION'].sum())
    resumen = resumen.sort_values('TOTAL_COMISION', ascending=False)
    return resumen

@st.cache_data(show_spinner=False)
def crear_excel_lujo(df, resumen_df):
    wb = Workbook()
    
    # --- HOJA DASHBOARD ---
    ws = wb.active
    ws.title = "RESUMEN EJECUTIVO"
    ws.sheet_view.showGridLines = False
    
    # Estilos
    color_brand = "E91E63"
    fill_header = PatternFill(start_color=color_brand, end_color=color_brand, fill_type="solid")
    font_header = Font(bold=True, color="FFFFFF", size=11)
    
    # Header
    ws['B2'] = "BLUSH HAIR & MAKE-UP SALON"
    ws['B2'].font = Font(size=20, bold=True, color=color_brand)
    ws['B3'] = f"Reporte Generado: {datetime.now().strftime('%d/%m/%Y')}"
    
    # Tabla Resumen
    headers = ['PUESTO', 'EMPLEADO', 'VENTA TOTAL', 'COMISIÓN A PAGAR', '% PART.']
    for col, h in enumerate(headers, 2):
        c = ws.cell(6, col, h)
        c.fill = fill_header
        c.font = font_header
        c.alignment = Alignment(horizontal='center')

    for idx, row in enumerate(resumen_df.itertuples(), 1):
        r = 6 + idx
        ws.cell(r, 2, f"#{idx}").alignment = Alignment(horizontal='center')
        ws.cell(r, 3, row.EMPLEADO)
        ws.cell(r, 4, row.TOTAL_PRODUCCION).number_format = '"S/" #,##0.00'
        ws.cell(r, 5, row.TOTAL_COMISION).number_format = '"S/" #,##0.00'
        ws.cell(r, 5).font = Font(bold=True)
        ws.cell(r, 6, row.PARTICIPACION).number_format = '0.0%'
        
    # Ajustar anchos
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 20

    # --- HOJA DETALLE ---
    ws2 = wb.create_sheet("DETALLE")
    headers_det = ['FECHA', 'EMPLEADO', 'ITEM', 'TIPO', 'MONTO', 'REGLA', 'COMISION']
    for col, h in enumerate(headers_det, 1):
        c = ws2.cell(1, col, h)
        c.fill = fill_header
        c.font = font_header
    
    for r, row in enumerate(df.itertuples(), 2):
        ws2.cell(r, 1, row.FECHA).number_format = 'dd/mm/yyyy'
        ws2.cell(r, 2, row.EMPLEADO)
        ws2.cell(r, 3, getattr(row, '_4')) # Item
        ws2.cell(r, 4, "Producto" if row.ES_PRODUCTO else "Servicio")
        ws2.cell(r, 5, row.MONTO).number_format = '#,##0.00'
        ws2.cell(r, 6, row.TIPO_COMISION)
        ws2.cell(r, 7, row.COMISION).number_format = '#,##0.00'

    ws2.column_dimensions['C'].width = 40
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- UI ---
st.markdown('<div class="main-header">💇‍♀️ BLUSH SYSTEM v3.1</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube tu Excel aquí", type=['xlsx', 'xls'])

if uploaded_file:
    with st.spinner("Procesando..."):
        df = procesar_datos(uploaded_file)
        
    if not df.empty:
        resumen = generar_resumen(df)
        
        # Métricas
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Ventas", f"S/ {resumen['TOTAL_PRODUCCION'].sum():,.2f}")
        col2.metric("Total Comisiones", f"S/ {resumen['TOTAL_COMISION'].sum():,.2f}")
        col3.metric("Transacciones", len(df))
        
        st.divider()
        
        # Tablas
        tab1, tab2 = st.tabs(["📋 Nómina", "🔍 Detalle"])
        with tab1:
            st.dataframe(resumen[['EMPLEADO', 'TOTAL_PRODUCCION', 'TOTAL_COMISION', 'PARTICIPACION']].style.format({'TOTAL_PRODUCCION': 'S/ {:,.2f}', 'TOTAL_COMISION': 'S/ {:,.2f}', 'PARTICIPACION': '{:.1%}'}), use_container_width=True)
        with tab2:
            st.dataframe(df, use_container_width=True)
            
        # Descarga
        excel_data = crear_excel_lujo(df, resumen)
        st.download_button("📥 Descargar Reporte Excel", excel_data, f"Reporte_Blush_{datetime.now().strftime('%Y%m%d')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        st.warning("No se pudieron procesar los datos. Revisa que el Excel tenga las columnas EMPLEADO y TOTAL COMP.")