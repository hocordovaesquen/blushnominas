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
    .metric-card {
        background-color: #FFF;
        border-left: 5px solid #E91E63;
        padding: 1rem;
        border-radius: 5px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    .stButton>button {
        background-color: #E91E63;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        border: none;
        width: 100%;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #C2185B;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
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

# --- LÓGICA DE NEGOCIO OPTIMIZADA ---

@st.cache_data(show_spinner=False)
def procesar_datos(uploaded_file):
    try:
        # Intentar leer Excel (tolerancia a nombres de hoja)
        xl = pd.ExcelFile(uploaded_file)
        # Buscar hoja que contenga "ventas" o "hoja1" o usar la primera
        sheet_name = next((s for s in xl.sheet_names if 'ventas' in s.lower()), xl.sheet_names[0])
        
        # Leer con cabecera en fila 9 (ajustable según tu formato original)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=9)
        
        # Validación básica
        cols_necesarias = ['EMPLEADO', 'TOTAL', 'FECHA', 'PRODUCTO / SERVICIO']
        if not any(col in df.columns for col in cols_necesarias):
            # Intento de rescate: leer sin skiprows si el formato cambió
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            if not any(col in df.columns for col in cols_necesarias):
                return pd.DataFrame() # Falló

        # Limpieza Vectorizada
        df = df[df['EMPLEADO'].notna()].copy()
        df['EMPLEADO'] = df['EMPLEADO'].astype(str).str.strip().str.title()
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce').ffill()
        df['MONTO'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
        
        # Relleno de columnas faltantes
        if 'CLASE' not in df.columns: df['CLASE'] = ''
        if 'PRODUCTO / SERVICIO' not in df.columns: df['PRODUCTO / SERVICIO'] = ''
        
        # Detección de productos (Vectorizada = Rápida)
        df['ITEM_UPPER'] = df['PRODUCTO / SERVICIO'].astype(str).str.upper()
        df['CLASE_UPPER'] = df['CLASE'].astype(str).str.upper().str.strip()
        
        # Lógica: Es producto si CLASE dice PRODUCTO o si el nombre contiene palabra clave
        cond_clase = df['CLASE_UPPER'] == 'PRODUCTO'
        cond_key = df['ITEM_UPPER'].str.contains(REGEX_PRODUCTOS, regex=True, na=False)
        df['ES_PRODUCTO'] = cond_clase | cond_key
        
        # Cálculo de Comisiones (Vectorizado con Numpy Select)
        # Condiciones
        c_producto = df['ES_PRODUCTO']
        c_julio_serv = (df['EMPLEADO'] == 'Julio') & (~df['ES_PRODUCTO'])
        # Jhon y Yuri Corte (palabra 'CORTE' o 'BARBERIA')
        c_jy_corte = (df['EMPLEADO'].isin(['Jhon', 'Yuri'])) & (~df['ES_PRODUCTO']) & (
            df['ITEM_UPPER'].str.contains('CORTE|BARBERIA', regex=True)
        )
        
        # Asignación
        condiciones = [c_producto, c_julio_serv, c_jy_corte]
        pcts = [0.10, 0.40, 0.35]
        labels = ["Producto 10%", "Servicio 40%", "Corte 35%"]
        
        df['PORCENTAJE'] = np.select(condiciones, pcts, default=0.25)
        df['TIPO_COMISION'] = np.select(condiciones, labels, default="Servicio 25%")
        
        df['COMISION'] = df['MONTO'] * df['PORCENTAJE']
        
        # Limpiar temporales
        return df.drop(columns=['ITEM_UPPER', 'CLASE_UPPER'])
        
    except Exception as e:
        st.error(f"Error técnico: {e}")
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
def generar_resumen(df):
    if df.empty: return pd.DataFrame()
    
    # Agrupación vectorizada
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

# --- GENERADOR DE EXCEL PREMIUN ---
def estilo_header(cell):
    cell.fill = PatternFill(start_color="E91E63", end_color="E91E63", fill_type="solid")
    cell.font = Font(bold=True, color="FFFFFF", size=11, name='Calibri')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = Border(bottom=Side(style='medium', color='880E4F'))

def estilo_moneda(cell, bold=False):
    cell.number_format = '"S/" #,##0.00'
    cell.font = Font(name='Calibri', size=11, bold=bold)

@st.cache_data(show_spinner=False)
def crear_excel_lujo(df, resumen_df):
    wb = Workbook()
    
    # --- HOJA 1: RESUMEN EJECUTIVO (DASHBOARD) ---
    ws_dash = wb.active
    ws_dash.title = "RESUMEN EJECUTIVO"
    ws_dash.sheet_view.showGridLines = False # Look limpio
    
    # Título Principal
    ws_dash.merge_cells('B2:H2')
    title = ws_dash['B2']
    title.value = "BLUSH HAIR & MAKE-UP SALON"
    title.font = Font(size=24, bold=True, color="E91E63", name="Arial")
    title.alignment = Alignment(horizontal='center')
    
    ws_dash.merge_cells('B3:H3')
    sub = ws_dash['B3']
    sub.value = f"Reporte de Comisiones - {df['FECHA'].min():%d/%m/%Y} al {df['FECHA'].max():%d/%m/%Y}"
    sub.font = Font(size=14, color="555555", name="Arial")
    sub.alignment = Alignment(horizontal='center')
    
    # Tarjetas de Métricas
    metrics = [
        ("PRODUCCION TOTAL", resumen_df['TOTAL_PRODUCCION'].sum(), "B5"),
        ("COMISIONES TOTALES", resumen_df['TOTAL_COMISION'].sum(), "E5"),
        ("TICKET PROMEDIO", df['MONTO'].mean(), "H5")
    ]
    
    for titulo, valor, celda in metrics:
        ws_dash[celda] = titulo
        ws_dash[celda].font = Font(bold=True, color="888888")
        ws_dash[celda].alignment = Alignment(horizontal='center')
        
        # Valor debajo
        row_val = int(celda[1]) + 1
        col_val = celda[0]
        ws_dash[f"{col_val}{row_val}"] = valor
        ws_dash[f"{col_val}{row_val}"].number_format = '"S/" #,##0.00'
        ws_dash[f"{col_val}{row_val}"].font = Font(size=18, bold=True, color="333333")
        ws_dash[f"{col_val}{row_val}"].alignment = Alignment(horizontal='center')
        
        # Borde decorativo
        ws_dash[f"{col_val}{row_val}"].border = Border(bottom=Side(style='thick', color='E91E63'))

    # Tabla Ranking
    ws_dash['B9'] = "TOP EMPLEADOS DEL PERIODO"
    ws_dash['B9'].font = Font(size=14, bold=True, color="E91E63")
    
    headers_rank = ['PUESTO', 'EMPLEADO', 'PRODUCCION', 'COMISION', '% PART.']
    for i, h in enumerate(headers_rank, 2):
        cell = ws_dash.cell(row=11, column=i)
        cell.value = h
        estilo_header(cell)
    
    for idx, row in enumerate(resumen_df.head(5).itertuples(), 1):
        r = 11 + idx
        ws_dash.cell(r, 2).value = f"#{idx}"
        ws_dash.cell(r, 2).alignment = Alignment(horizontal='center')
        ws_dash.cell(r, 3).value = row.EMPLEADO
        ws_dash.cell(r, 4).value = row.TOTAL_PRODUCCION
        estilo_moneda(ws_dash.cell(r, 4))
        ws_dash.cell(r, 5).value = row.TOTAL_COMISION
        estilo_moneda(ws_dash.cell(r, 5), bold=True)
        ws_dash.cell(r, 6).value = row.PARTICIPACION
        ws_dash.cell(r, 6).number_format = '0.0%'

    # Anchos
    ws_dash.column_dimensions['B'].width = 10
    ws_dash.column_dimensions['C'].width = 25
    ws_dash.column_dimensions['D'].width = 18
    ws_dash.column_dimensions['E'].width = 18
    ws_dash.column_dimensions['F'].width = 15
    
    # --- HOJA 2: NÓMINA (DETALLADA) ---
    ws_nom = wb.create_sheet("NÓMINA")
    
    headers_nom = ['EMPLEADO', 'PROD. SERVICIOS', 'COM. SERVICIOS', 'PROD. PRODUCTOS', 
                   'COM. PRODUCTOS', 'TOTAL PROD.', 'TOTAL COM.', 'A PAGAR']
    
    for i, h in enumerate(headers_nom, 1):
        cell = ws_nom.cell(row=1, column=i)
        cell.value = h
        estilo_header(cell)
        
    for r, row in enumerate(resumen_df.itertuples(), 2):
        ws_nom.cell(r, 1, row.EMPLEADO)
        ws_nom.cell(r, 2, row.PROD_SERVICIOS).number_format = '#,##0.00'
        ws_nom.cell(r, 3, row.COM_SERVICIOS).number_format = '#,##0.00'
        ws_nom.cell(r, 4, row.PROD_PRODUCTOS).number_format = '#,##0.00'
        ws_nom.cell(r, 5, row.COM_PRODUCTOS).number_format = '#,##0.00'
        ws_nom.cell(r, 6, row.TOTAL_PRODUCCION).number_format = '"S/" #,##0.00'
        ws_nom.cell(r, 6).font = Font(bold=True)
        ws_nom.cell(r, 7, row.TOTAL_COMISION).number_format = '"S/" #,##0.00'
        ws_nom.cell(r, 7).fill = PatternFill(start_color="E1F5FE", fill_type="solid")
        ws_nom.cell(r, 7).font = Font(bold=True, color="01579B")
        ws_nom.cell(r, 8, f"=G{r}") # Formula simple para A Pagar
        
    # Totales footer
    last_r = len(resumen_df) + 2
    ws_nom.cell(last_r, 1, "TOTALES").font = Font(bold=True)
    for c in range(2, 9):
        col_letter = get_column_letter(c)
        ws_nom.cell(last_r, c, f"=SUM({col_letter}2:{col_letter}{last_r-1})")
        ws_nom.cell(last_r, c).font = Font(bold=True)
        ws_nom.cell(last_r, c).number_format = '"S/" #,##0.00'
        ws_nom.cell(last_r, c).fill = PatternFill(start_color="F5F5F5", fill_type="solid")

    # Auto-ajuste anchos Nómina
    for col in ws_nom.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        ws_nom.column_dimensions[column].width = max_length + 2

    # --- HOJA 3: DETALLE TRANSACCIONES ---
    ws_det = wb.create_sheet("DETALLE")
    headers_det = ['FECHA', 'EMPLEADO', 'ITEM', 'TIPO', 'MONTO', 'REGLA', 'COMISION']
    
    for i, h in enumerate(headers_det, 1):
        cell = ws_det.cell(1, i)
        cell.value = h
        estilo_header(cell)
        # Color diferente para detalle
        cell.fill = PatternFill(start_color="7B1FA2", end_color="7B1FA2", fill_type="solid")
    
    for r, row in enumerate(df.itertuples(), 2):
        ws_det.cell(r, 1, row.FECHA).number_format = 'dd/mm/yyyy'
        ws_det.cell(r, 2, row.EMPLEADO)
        ws_det.cell(r, 3, getattr(row, '_4')) # Producto/Servicio column name is tricky in tuple
        ws_det.cell(r, 4, "Producto" if row.ES_PRODUCTO else "Servicio")
        ws_det.cell(r, 5, row.MONTO).number_format = '#,##0.00'
        ws_det.cell(r, 6, row.TIPO_COMISION)
        ws_det.cell(r, 7, row.COMISION).number_format = '#,##0.00'

    ws_det.column_dimensions['C'].width = 40
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- UI PRINCIPAL ---

st.markdown('<div class="main-header">💇‍♀️ BLUSH SYSTEM v3.0</div>', unsafe_allow_html=True)

with st.sidebar:
    st.title("Panel de Control")
    st.info("Sube el Excel de ventas (Hoja 1 o Ventas) para calcular automáticamente.")
    uploaded_file = st.file_uploader("Cargar Excel", type=['xlsx', 'xls'])

if uploaded_file:
    with st.spinner("🚀 Optimizando datos..."):
        df = procesar_datos(uploaded_file)
    
    if not df.empty:
        resumen = generar_resumen(df)
        
        # MÉTRICAS SUPERIOR
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ventas Totales", f"S/ {resumen['TOTAL_PRODUCCION'].sum():,.2f}")
        c2.metric("Comisiones a Pagar", f"S/ {resumen['TOTAL_COMISION'].sum():,.2f}")
        c3.metric("Transacciones", len(df))
        c4.metric("Top Vendedor", resumen.iloc[0]['EMPLEADO'])
        
        st.markdown("---")
        
        # TABS
        tab1, tab2 = st.tabs(["📊 Dashboard & Nómina", "📝 Detalle Completo"])
        
        with tab1:
            col_izq, col_der = st.columns([2, 1])
            
            with col_izq:
                st.subheader("Nómina Calculada")
                st.dataframe(
                    resumen[['EMPLEADO', 'TOTAL_PRODUCCION', 'TOTAL_COMISION', 'PARTICIPACION']].style
                    .format({
                        'TOTAL_PRODUCCION': 'S/ {:,.2f}', 
                        'TOTAL_COMISION': 'S/ {:,.2f}',
                        'PARTICIPACION': '{:.1%}'
                    })
                    .background_gradient(subset=['TOTAL_COMISION'], cmap='RdPu'),
                    use_container_width=True
                )
            
            with col_der:
                st.subheader("Distribución")
                fig = px.pie(resumen, values='TOTAL_COMISION', names='EMPLEADO', hole=0.4)
                fig.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.dataframe(df[['FECHA', 'EMPLEADO', 'PRODUCTO / SERVICIO', 'MONTO', 'COMISION', 'TIPO_COMISION']], use_container_width=True)
        
        # BOTÓN DE DESCARGA DELUXE
        st.markdown("### 👇 Descarga Final")
        excel_data = crear_excel_lujo(df, resumen)
        
        st.download_button(
            label="📥 DESCARGAR REPORTE EXCEL PROFESIONAL",
            data=excel_data,
            file_name=f"Reporte_Blush_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    else:
        st.error("No pudimos leer el archivo. Asegúrate de que sea el Excel correcto.")
else:
    st.markdown("""
    <div style='text-align:center; margin-top: 50px; opacity: 0.6'>
        <h3>Esperando archivo...</h3>
        <p>Arrastra el reporte de ventas aquí</p>
    </div>
    """, unsafe_allow_html=True)
```

### 2. Archivo: `requirements.txt`
Asegúrate de que este archivo tenga lo siguiente (si ya lo tienes, solo verifica):

```text
streamlit
pandas
openpyxl
plotly
numpy