import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import io
from datetime import datetime
import plotly.express as px
import numpy as np  # Necesario para limpiar celdas vacías

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
        font-family: 'Arial', sans-serif;
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
    .metric-card {
        background-color: #F8F9FA;
        border-left: 5px solid #E91E63;
        padding: 15px;
        border-radius: 5px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

# --- FUNCIONES DE PROCESAMIENTO ---
@st.cache_data(ttl=3600)
def load_data(file):
    try:
        # 1. Leer sin encabezado para buscar la fila real
        df_raw = pd.read_excel(file, header=None)
        
        header_idx = -1
        for i, row in df_raw.head(30).iterrows(): # Aumenté el rango de búsqueda
            row_str = " ".join(row.astype(str).values).upper()
            if "PRODUCTO" in row_str and "TOTAL" in row_str:
                header_idx = i
                break
        
        if header_idx == -1:
            st.error("❌ No se encontró la cabecera. Verifica que existan columnas 'PRODUCTO / SERVICIO' y 'TOTAL'.")
            return None

        # 2. Cargar datos desde la fila correcta
        df = pd.read_excel(file, header=header_idx)
        
        # 3. Normalizar columnas
        df.columns = df.columns.str.strip().str.upper()
        
        col_map = {}
        for col in df.columns:
            if "PRODUCTO" in col: col_map[col] = 'PRODUCTO'
            elif "EMPLEADO" in col: col_map[col] = 'EMPLEADO'
            # Ojo: Evitamos "TOTAL COMPROBANTE" o "SUB TOTAL", buscamos la columna de importe final
            elif "TOTAL" in col and "COMP" not in col and "SUB" not in col: col_map[col] = 'TOTAL'
            elif "TV" in col: col_map[col] = 'TV'
            elif "FECHA" in col and "REGISTRO" not in col: col_map[col] = 'FECHA'
            elif "CLIENTE" in col and "T-" not in col: col_map[col] = 'CLIENTE'
            elif "PEDIDO" in col: col_map[col] = 'PEDIDO'

        df = df.rename(columns=col_map)

        # --- CORRECCIÓN CRÍTICA: LIMPIEZA DE VACÍOS ---
        # Convertir espacios vacíos, cadenas vacías o solo espacios en NaN real
        # Esto es vital para que el ffill funcione
        df = df.replace(r'^\s*$', np.nan, regex=True)

        # 4. Rellenar datos hacia abajo (Propagar Fecha, TV, Cliente a los items secundarios)
        cols_to_fill = [c for c in ['FECHA', 'TV', 'PEDIDO', 'CLIENTE'] if c in df.columns]
        df[cols_to_fill] = df[cols_to_fill].ffill()

        # 5. Filtros de Lógica de Negocio
        
        # A) Eliminar filas basura (Totales del reporte, encabezados repetidos)
        palabras_basura = ['TOTAL', 'SUBTOTAL', 'RESUMEN', 'REGISTRO', 'PRODUCTO', 'SERVICIO', 'SEDE', 'DESDE']
        df = df[~df['PRODUCTO'].astype(str).str.upper().isin(palabras_basura)]
        
        # B) Asegurar que hay un producto escrito
        df = df.dropna(subset=['PRODUCTO'])
        
        # C) Filtrar por columna TV (Solo Ventas, ignorar Anulados)
        # Importante: Hacemos esto DESPUÉS del ffill para no borrar items válidos
        if 'TV' in df.columns:
            df = df[df['TV'].astype(str).str.upper() == 'V']

        # 6. Limpieza numérica y de empleados
        # Convertir Total a número
        df['TOTAL'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
        
        if 'EMPLEADO' in df.columns:
            df['EMPLEADO'] = df['EMPLEADO'].fillna('Sin Asignar')
            
        return df

    except Exception as e:
        st.error(f"Error procesando el archivo: {str(e)}")
        return None

def procesar_nomina(df, comision_base):
    # Agrupar por empleado sumando sus ventas
    resumen = df.groupby('EMPLEADO').agg(
        TOTAL_PRODUCCION=('TOTAL', 'sum'),
        CONTEO_SERVICIOS=('PRODUCTO', 'count')
    ).reset_index()
    
    resumen['PORCENTAJE'] = comision_base / 100
    resumen['TOTAL_COMISION'] = resumen['TOTAL_PRODUCCION'] * resumen['PORCENTAJE']
    
    total_global = resumen['TOTAL_PRODUCCION'].sum()
    resumen['PARTICIPACION'] = resumen['TOTAL_PRODUCCION'] / total_global if total_global > 0 else 0
        
    return resumen.sort_values('TOTAL_PRODUCCION', ascending=False)

def crear_excel(df_detalle, df_resumen):
    output = io.BytesIO()
    wb = Workbook()
    
    header_style = PatternFill(start_color="E91E63", end_color="E91E63", fill_type="solid")
    font_white = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Hoja 1: Resumen
    ws1 = wb.active
    ws1.title = "Resumen Nómina"
    headers = ["EMPLEADO", "SERVICIOS", "VENTA TOTAL", "% COMISIÓN", "A PAGAR", "PARTICIPACIÓN"]
    ws1.append(headers)
    
    for col in range(1, len(headers) + 1):
        cell = ws1.cell(row=1, column=col)
        cell.fill = header_style
        cell.font = font_white
        cell.alignment = Alignment(horizontal='center')

    for r, row in enumerate(df_resumen.itertuples(), 2):
        ws1.cell(row=r, column=1, value=row.EMPLEADO).border = border
        ws1.cell(row=r, column=2, value=row.CONTEO_SERVICIOS).border = border
        ws1.cell(row=r, column=3, value=row.TOTAL_PRODUCCION).number_format = '"S/" #,##0.00'
        ws1.cell(row=r, column=4, value=row.PORCENTAJE).number_format = '0%'
        ws1.cell(row=r, column=5, value=f"=C{r}*D{r}").number_format = '"S/" #,##0.00'
        ws1.cell(row=r, column=6, value=row.PARTICIPACION).number_format = '0.0%'

    # Hoja 2: Detalle
    ws2 = wb.create_sheet("Detalle Procesado")
    # Columnas que queremos exportar
    cols_export = ['FECHA', 'EMPLEADO', 'PRODUCTO', 'TOTAL', 'CLIENTE', 'TV', 'PEDIDO']
    cols_existentes = [c for c in cols_export if c in df_detalle.columns]
    
    ws2.append(cols_existentes)
    for col in range(1, len(cols_existentes) + 1):
        ws2.cell(row=1, column=col).fill = header_style
        ws2.cell(row=1, column=col).font = font_white

    for r_idx, row in enumerate(df_detalle[cols_existentes].itertuples(index=False), 2):
        for c_idx, value in enumerate(row, 1):
            ws2.cell(row=r_idx, column=c_idx, value=value)

    wb.save(output)
    return output.getvalue()

# --- INTERFAZ PRINCIPAL ---
st.markdown('<div class="main-header">💇‍♀️ BLUSH - Nómina (Corregido Final)</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("Configuración")
    comision_input = st.slider("Comisión General (%)", 0, 100, 50)

uploaded_file = st.file_uploader("Sube el Excel aquí", type=['xlsx', 'xls'])

if uploaded_file:
    df = load_data(uploaded_file)
    
    if df is not None:
        resumen = procesar_nomina(df, comision_input)
        
        # Métricas
        total_ventas = resumen['TOTAL_PRODUCCION'].sum()
        total_servicios = resumen['CONTEO_SERVICIOS'].sum()
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Ventas Totales", f"S/ {total_ventas:,.2f}")
        col2.metric("Comisión Total", f"S/ {total_ventas * (comision_input/100):,.2f}")
        col3.metric("Items Procesados", int(total_servicios))
        
        st.markdown("---")
        
        tab1, tab2 = st.tabs(["📊 Resultados", "📝 Lista Completa de Servicios"])
        
        with tab1:
            st.dataframe(resumen.style.format({'TOTAL_PRODUCCION': 'S/ {:,.2f}', 'TOTAL_COMISION': 'S/ {:,.2f}'}), use_container_width=True)
            
        with tab2:
            st.success(f"✅ Se encontraron {len(df)} productos/servicios válidos.")
            st.dataframe(df[['FECHA', 'EMPLEADO', 'PRODUCTO', 'TOTAL', 'TV']], use_container_width=True)
            
        excel_data = crear_excel(df, resumen)
        st.download_button(
            label="📥 DESCARGAR REPORTE VALIDADO",
            data=excel_data,
            file_name="Nomina_Blush_Ok.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )