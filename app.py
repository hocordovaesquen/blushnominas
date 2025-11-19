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

# --- LÓGICA DE NEGOCIO ---

@st.cache_data(show_spinner=False)
def procesar_datos(uploaded_file):
    try:
        # 1. DETECCIÓN INTELIGENTE DE CABECERA
        # Leemos las primeras 25 filas sin cabecera para buscar dónde están los títulos reales
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = next((s for s in xl.sheet_names if 'ventas' in s.lower() or 'hoja1' in s.lower()), xl.sheet_names[0])
        
        df_preview = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, nrows=25)
        
        header_row_idx = None
        for idx, row in df_preview.iterrows():
            # Convertimos la fila a texto y mayúsculas para buscar palabras clave
            row_text = row.astype(str).str.upper().str.strip().tolist()
            
            # Buscamos una fila que tenga 'EMPLEADO' y ('TOTAL' o 'TOTAL COMP')
            tiene_empleado = 'EMPLEADO' in row_text
            tiene_total = any(x in row_text for x in ['TOTAL', 'TOTAL COMP', 'IMPORTE', 'TOTAL VENTA'])
            
            if tiene_empleado and tiene_total:
                header_row_idx = idx
                break
        
        if header_row_idx is None:
            st.error("No pude encontrar la fila de encabezados (que tenga 'EMPLEADO' y 'TOTAL COMP'). Intentando fila 9 por defecto.")
            header_row_idx = 9

        # 2. CARGA REAL
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=header_row_idx)
        
        # 3. LIMPIEZA DE COLUMNAS
        df.columns = df.columns.str.strip().str.upper()
        
        # Mapeo de nombres extraños a nombres estándar
        renames = {
            'TOTAL COMP': 'TOTAL',       # Esto es lo que fallaba antes
            'IMPORTE': 'TOTAL',
            'FECHA REGISTRO': 'FECHA',
            'PRODUCTO/SERVICIO': 'PRODUCTO / SERVICIO'
        }
        df = df.rename(columns=renames)
        
        # Validación
        if 'TOTAL' not in df.columns or 'EMPLEADO' not in df.columns:
            st.error(f"⚠️ Columnas no encontradas. Detectadas: {list(df.columns)}")
            return pd.DataFrame()

        # 4. LIMPIEZA DE DATOS
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
        st.error(f"Error procesando el archivo: {str(e)}")
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

# --- GENERADOR DE EXCEL CON FÓRMULAS ---
@st.cache_data(show_spinner=False)
def crear_excel_con_formulas(df, resumen_df):
    wb = Workbook()
    
    # Estilos Globales
    color_brand = "E91E63" # Rosa Blush
    color_header = "FCE4EC" # Rosa claro
    color_total = "F5F5F5"  # Gris claro
    
    header_style = PatternFill(start_color=color_brand, end_color=color_brand, fill_type="solid")
    font_white = Font(bold=True, color="FFFFFF", size=11, name='Calibri')
    font_bold = Font(bold=True, name='Calibri')
    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # ==========================================
    # HOJA 1: NOMINA (CON FÓRMULAS)
    # ==========================================
    ws = wb.active
    ws.title = "NOMINA"
    
    # Título
    ws.merge_cells('A1:K1')
    ws['A1'] = "NOMINA QUINCENAL - BLUSH HAIR & MAKE-UP"
    ws['A1'].font = Font(size=14, bold=True, color=color_brand)
    ws['A1'].alignment = Alignment(horizontal='center')
    
    # Cabeceras
    headers = [
        '#', 'EMPLEADO', 
        'PROD. SERVICIOS', 'COM. SERVICIOS', 
        'PROD. PRODUCTOS', 'COM. PRODUCTOS', 
        'TOTAL PROD.', 'TOTAL COM.', 
        'DESCUENTOS', 'EXTRAS', 'A PAGAR'
    ]
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.value = header
        cell.fill = header_style
        cell.font = font_white
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_thin

    # Llenado de datos
    start_row = 4
    current_row = start_row
    
    for idx, row in enumerate(resumen_df.itertuples(), 1):
        # Datos Estáticos (Valores calculados en Python)
        ws.cell(current_row, 1, idx)
        ws.cell(current_row, 2, row.EMPLEADO)
        ws.cell(current_row, 3, row.PROD_SERVICIOS).number_format = '#,##0.00'
        ws.cell(current_row, 4, row.COM_SERVICIOS).number_format = '#,##0.00'
        ws.cell(current_row, 5, row.PROD_PRODUCTOS).number_format = '#,##0.00'
        ws.cell(current_row, 6, row.COM_PRODUCTOS).number_format = '#,##0.00'
        
        # --- AQUÍ ESTÁN LAS FÓRMULAS FILA POR FILA ---
        
        # Col G: TOTAL PROD = Prod.Servicios (C) + Prod.Productos (E)
        ws.cell(current_row, 7, f"=C{current_row}+E{current_row}")
        ws.cell(current_row, 7).number_format = '"S/" #,##0.00'
        
        # Col H: TOTAL COM = Com.Servicios (D) + Com.Productos (F)
        ws.cell(current_row, 8, f"=D{current_row}+F{current_row}")
        ws.cell(current_row, 8).number_format = '"S/" #,##0.00'
        ws.cell(current_row, 8).font = font_bold
        
        # Col I y J: Descuentos y Extras (Vacíos o 0 para que el usuario llene)
        ws.cell(current_row, 9, 0).number_format = '#,##0.00'
        ws.cell(current_row, 10, 0).number_format = '#,##0.00'
        
        # Col K: A PAGAR = Total Com (H) - Descuentos (I) + Extras (J)
        ws.cell(current_row, 11, f"=H{current_row}-I{current_row}+J{current_row}")
        ws.cell(current_row, 11).number_format = '"S/" #,##0.00'
        ws.cell(current_row, 11).fill = PatternFill(start_color="E0F7FA", fill_type="solid") # Azulito para resaltar pago
        ws.cell(current_row, 11).font = font_bold
        
        # Bordes
        for c in range(1, 12):
            ws.cell(current_row, c).border = border_thin
            
        current_row += 1
        
    # FILA DE TOTALES (CON FÓRMULAS DE SUMA)
    ws.cell(current_row, 2, "TOTAL GENERAL").font = font_bold
    
    # Columnas C hasta K
    cols_letras = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
    for i, letra in enumerate(cols_letras):
        idx_col = 3 + i
        # Fórmula: =SUM(L4:L10)
        ws.cell(current_row, idx_col, f"=SUM({letra}{start_row}:{letra}{current_row-1})")
        ws.cell(current_row, idx_col).font = font_bold
        ws.cell(current_row, idx_col).fill = PatternFill(start_color=color_total, fill_type="solid")
        ws.cell(current_row, idx_col).number_format = '"S/" #,##0.00'
        ws.cell(current_row, idx_col).border = border_thin

    # Ajustar anchos
    ws.column_dimensions['B'].width = 20
    for c in cols_letras:
        ws.column_dimensions[c].width = 15

    # ==========================================
    # HOJA 2: RESUMEN EJECUTIVO (Dashboard)
    # ==========================================
    ws2 = wb.create_sheet("RESUMEN EJECUTIVO")
    ws2.sheet_view.showGridLines = False
    
    ws2['B2'] = "RESUMEN EJECUTIVO DEL PERIODO"
    ws2['B2'].font = Font(size=16, bold=True, color=color_brand)
    
    # Métricas Grandes
    ws2['B4'] = "VENTA TOTAL"
    ws2['C4'] = f"=NOMINA!G{current_row}" # Jala el total de la otra hoja
    ws2['C4'].number_format = '"S/" #,##0.00'
    ws2['C4'].font = Font(size=14, bold=True)
    
    ws2['B5'] = "COMISIÓN TOTAL"
    ws2['C5'] = f"=NOMINA!H{current_row}" # Jala el total de la otra hoja
    ws2['C5'].number_format = '"S/" #,##0.00'
    ws2['C5'].font = Font(size=14, bold=True)

    # ==========================================
    # HOJA 3: DETALLE (Datos crudos)
    # ==========================================
    ws3 = wb.create_sheet("DETALLE")
    headers_det = ['FECHA', 'EMPLEADO', 'ITEM', 'TIPO', 'MONTO', 'REGLA', 'COMISION']
    
    for col, h in enumerate(headers_det, 1):
        c = ws3.cell(1, col, h)
        c.fill = header_style
        c.font = font_white
    
    for r, row in enumerate(df.itertuples(), 2):
        ws3.cell(r, 1, row.FECHA).number_format = 'dd/mm/yyyy'
        ws3.cell(r, 2, row.EMPLEADO)
        ws3.cell(r, 3, getattr(row, '_4')) # Producto/Servicio
        ws3.cell(r, 4, "Producto" if row.ES_PRODUCTO else "Servicio")
        ws3.cell(r, 5, row.MONTO).number_format = '#,##0.00'
        ws3.cell(r, 6, row.TIPO_COMISION)
        ws3.cell(r, 7, row.COMISION).number_format = '#,##0.00'
        
    ws3.column_dimensions['C'].width = 40

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- INTERFAZ DE USUARIO ---
st.markdown('<div class="main-header">💇‍♀️ BLUSH SYSTEM v4.0 (Final)</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube tu Excel de Ventas aquí", type=['xlsx', 'xls'])

if uploaded_file:
    with st.spinner("Analizando archivo y detectando estructura..."):
        df = procesar_datos(uploaded_file)
        
    if not df.empty:
        resumen = generar_resumen(df)
        
        # Métricas en Pantalla
        col1, col2, col3 = st.columns(3)
        total_v = resumen['TOTAL_PRODUCCION'].sum()
        total_c = resumen['TOTAL_COMISION'].sum()
        col1.metric("Venta Total", f"S/ {total_v:,.2f}")
        col2.metric("Comisión Total", f"S/ {total_c:,.2f}")
        col3.metric("Transacciones", len(df))
        
        st.markdown("---")
        
        tab1, tab2 = st.tabs(["📋 Vista Previa Nómina", "🔍 Detalle Transacciones"])
        
        with tab1:
            st.info("💡 Esta es una vista previa. El Excel descargado incluirá las fórmulas editables.")
            # Formato visual para la web
            st.dataframe(
                resumen[['EMPLEADO', 'TOTAL_PRODUCCION', 'TOTAL_COMISION', 'PARTICIPACION']].style
                .format({'TOTAL_PRODUCCION': 'S/ {:,.2f}', 'TOTAL_COMISION': 'S/ {:,.2f}', 'PARTICIPACION': '{:.1%}'})
                .background_gradient(subset=['TOTAL_COMISION'], cmap='RdPu'),
                use_container_width=True
            )
            
        with tab2:
            st.dataframe(df, use_container_width=True)
            
        st.markdown("### 👇 DESCARGA OFICIAL")
        
        # Generar el Excel con fórmulas
        excel_data = crear_excel_con_formulas(df, resumen)
        
        st.download_button(
            label="📥 DESCARGAR NOMINA (CON FÓRMULAS)",
            data=excel_data,
            file_name=f"Nomina_Blush_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    else:
        st.warning("No se pudieron procesar los datos.")