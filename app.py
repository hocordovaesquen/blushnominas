import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import io
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import re
import numpy as np

# CONFIGURACION DE PAGINA
st.set_page_config(
    page_title="BLUSH - Calculador de Comisiones",
    page_icon="💇‍♀️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS PERSONALIZADO
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #E91E63;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #FCE4EC 0%, #F8BBD0 100%);
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .stButton>button {
        background-color: #E91E63;
        color: white;
        font-weight: bold;
        border-radius: 10px;
        padding: 0.5rem 2rem;
        border: none;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #C2185B;
    }
</style>
""", unsafe_allow_html=True)

# --- CONSTANTES Y CONFIGURACIÓN ---
PALABRAS_PRODUCTO = [
    'MASCARILLA', 'SHAMPOO', 'SHAMPO', 'ACONDICIONADOR',
    'CREMA', 'SERUM', 'AMPOLLA', 'SPRAY', 'GEL',
    'LOTION', 'REDKEN', 'LOREAL', 'TIGI', 'KERASTASE',
    'X250ML', 'X300ML', 'X500ML', 'ML', 'GR',
    'BED HEAD', 'ALL SOFT', 'FRIZZ DISMISS'
]
# Pre-compilamos el regex para que sea super rápido
REGEX_PRODUCTOS = '|'.join([re.escape(p) for p in PALABRAS_PRODUCTO])

# --- FUNCIONES CACHEADAS (Optimizan la velocidad) ---

@st.cache_data(show_spinner=False)
def procesar_datos(uploaded_file):
    """Procesa el archivo subido con optimización vectorizada"""
    
    # Leer Excel
    try:
        df = pd.read_excel(uploaded_file, sheet_name='Hoja1', skiprows=9)
    except:
        try:
            # Intento alternativo si la hoja se llama diferente
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, skiprows=9)
        except Exception as e:
            st.error(f"Error leyendo el Excel: {e}")
            return pd.DataFrame()
    
    # Validar columnas mínimas necesarias
    required_cols = ['EMPLEADO', 'FECHA', 'TOTAL']
    if not all(col in df.columns for col in required_cols):
        # Intentar buscar la fila de cabecera correcta si falló skiprows
        return pd.DataFrame() # Retorna vacío para manejar error arriba

    # Limpiar datos básicos
    df = df[df['EMPLEADO'].notna()].copy()
    df['EMPLEADO'] = df['EMPLEADO'].astype(str).str.strip()
    
    # Procesar FECHA
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce').ffill()
    
    # Limpiar MONTO
    df['MONTO'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
    
    # Asegurar columnas de texto
    if 'PRODUCTO / SERVICIO' not in df.columns:
        df['PRODUCTO / SERVICIO'] = ''
    if 'CLASE' not in df.columns:
        df['CLASE'] = ''

    # --- OPTIMIZACIÓN: DETECCION VECTORIZADA DE PRODUCTO ---
    # 1. Normalizar texto a mayúsculas para búsquedas
    df['PROD_SERV_UPPER'] = df['PRODUCTO / SERVICIO'].astype(str).str.upper()
    df['CLASE_UPPER'] = df['CLASE'].astype(str).str.upper().str.strip()
    
    # 2. Crear máscaras booleanas (True/False)
    es_clase_producto = df['CLASE_UPPER'] == 'PRODUCTO'
    contiene_palabra_clave = df['PROD_SERV_UPPER'].str.contains(REGEX_PRODUCTOS, regex=True, na=False)
    
    # 3. Asignar resultado
    df['ES_PRODUCTO'] = es_clase_producto | contiene_palabra_clave
    
    # --- CÁLCULO DE COMISIONES ---
    cond_producto = df['ES_PRODUCTO']
    cond_julio_serv = (df['EMPLEADO'] == 'Julio') & (~df['ES_PRODUCTO'])
    cond_jhon_yuri_corte = (df['EMPLEADO'].isin(['Jhon', 'Yuri'])) & (~df['ES_PRODUCTO']) & (df['PROD_SERV_UPPER'].str.contains('CORTE'))
    
    conditions = [
        cond_producto,          # 10%
        cond_julio_serv,        # 40%
        cond_jhon_yuri_corte    # 35%
    ]
    choices_pct = [0.10, 0.40, 0.35]
    choices_tipo = ["Producto 10%", "Servicio 40%", "Corte 35%"]
    
    df['PORCENTAJE'] = np.select(conditions, choices_pct, default=0.25)
    df['TIPO_COMISION'] = np.select(conditions, choices_tipo, default="Servicio 25%")
    
    df['COMISION'] = df['MONTO'] * df['PORCENTAJE']
    
    # Limpieza
    df = df.drop(columns=['PROD_SERV_UPPER', 'CLASE_UPPER'])
    
    return df

@st.cache_data(show_spinner=False)
def crear_resumen(df):
    if df.empty:
        return pd.DataFrame()

    resumen_lista = []
    for emp in sorted(df['EMPLEADO'].unique()):
        d_emp = df[df['EMPLEADO'] == emp]
        servicios = d_emp[~d_emp['ES_PRODUCTO']]
        productos = d_emp[d_emp['ES_PRODUCTO']]
        
        resumen_lista.append({
            'EMPLEADO': emp,
            'PRODUCCION_SERVICIOS': servicios['MONTO'].sum(),
            'COMISION_SERVICIOS': servicios['COMISION'].sum(),
            'PRODUCCION_PRODUCTOS': productos['MONTO'].sum(),
            'COMISION_PRODUCTOS': productos['COMISION'].sum(),
            'TOTAL_PRODUCCION': d_emp['MONTO'].sum(),
            'TOTAL_COMISION': d_emp['COMISION'].sum(),
            'CANTIDAD_SERVICIOS': len(servicios),
            'CANTIDAD_PRODUCTOS': len(productos),
            'TOTAL_TRANSACCIONES': len(d_emp),
            'TICKET_PROMEDIO': d_emp['MONTO'].mean()
        })
    
    resumen_df = pd.DataFrame(resumen_lista)
    if not resumen_df.empty:
        resumen_df['PARTICIPACION'] = (resumen_df['TOTAL_PRODUCCION'] / resumen_df['TOTAL_PRODUCCION'].sum() * 100)
        return resumen_df.sort_values('TOTAL_COMISION', ascending=False)
    return resumen_df

@st.cache_data(show_spinner=False)
def crear_excel_profesional(df, resumen_df):
    wb = Workbook()
    
    rosa_blush = "E91E63"
    rosa_claro = "FCE4EC"
    verde_exito = "4CAF50"
    
    header_fill = PatternFill(start_color=rosa_blush, end_color=rosa_blush, fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11, name='Arial')
    total_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    total_font = Font(bold=True, size=11, name='Arial')
    highlight_fill = PatternFill(start_color=verde_exito, end_color=verde_exito, fill_type="solid")
    highlight_font = Font(bold=True, color="FFFFFF", size=11, name='Arial')
    
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # HOJA NOMINA
    ws = wb.active
    ws.title = "NOMINA"
    
    if not df.empty:
        fecha_min = df['FECHA'].min().strftime('%d/%m/%Y')
        fecha_max = df['FECHA'].max().strftime('%d/%m/%Y')
    else:
        fecha_min = "-"
        fecha_max = "-"
    
    ws.merge_cells('A1:K1')
    ws['A1'] = f'NOMINA QUINCENAL - {fecha_min} AL {fecha_max}'
    ws['A1'].font = Font(bold=True, size=14, name='Arial')
    ws['A1'].alignment = Alignment(horizontal='center')
    ws['A1'].fill = PatternFill(start_color=rosa_claro, end_color=rosa_claro, fill_type='solid')
    
    headers = ['#', 'EMPLEADO', 'PROD. SERVICIOS', 'COM. SERVICIOS', 'PROD. PRODUCTOS', 
               'COM. PRODUCTOS', 'TOTAL PROD.', 'TOTAL COM.', 'DESCUENTOS', 'EXTRAS', 'A PAGAR']
    
    for col, h in enumerate(headers, 1):
        c = ws.cell(3, col)
        c.value = h
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = border
    
    fila = 4
    if not resumen_df.empty:
        for idx, row in resumen_df.iterrows():
            ws.cell(fila, 1, idx+1)
            ws.cell(fila, 2, row['EMPLEADO'])
            ws.cell(fila, 3, row['PRODUCCION_SERVICIOS'])
            ws.cell(fila, 4, row['COMISION_SERVICIOS'])
            ws.cell(fila, 5, row['PRODUCCION_PRODUCTOS'])
            ws.cell(fila, 6, row['COMISION_PRODUCTOS'])
            ws.cell(fila, 7, row['TOTAL_PRODUCCION'])
            ws.cell(fila, 8, row['TOTAL_COMISION'])
            ws.cell(fila, 9, 0)
            ws.cell(fila, 10, 0)
            ws.cell(fila, 11, f'=H{fila}-I{fila}+J{fila}')
            
            for col in range(1, 12):
                c = ws.cell(fila, col)
                c.border = border
                if col > 2:
                    c.number_format = '#,##0.00'
                    c.alignment = Alignment(horizontal='right')
                if col in [7, 8]:
                    c.fill = highlight_fill
                    c.font = highlight_font
            fila += 1
        
        ws.cell(fila, 2, 'TOTAL').font = total_font
        for col in range(3, 12):
            if col in [9, 10]:
                ws.cell(fila, col, 0)
            else:
                ws.cell(fila, col, f'=SUM({chr(64+col)}4:{chr(64+col)}{fila-1})')
            c = ws.cell(fila, col)
            c.fill = total_fill
            c.font = total_font
            c.number_format = '#,##0.00'
            c.border = border
    
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 20
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        ws.column_dimensions[col].width = 14
    
    # HOJA DETALLE
    ws2 = wb.create_sheet("DETALLE")
    ws2.merge_cells('A1:H1')
    ws2['A1'] = 'DETALLE DE CALCULOS'
    ws2['A1'].font = Font(bold=True, size=14, name='Arial')
    ws2['A1'].alignment = Alignment(horizontal='center')
    
    headers2 = ['FECHA', 'EMPLEADO', 'SERVICIO/PRODUCTO', 'TIPO', 'MONTO', '%', 'COMISION', 'REGLA']
    
    for col, h in enumerate(headers2, 1):
        c = ws2.cell(3, col)
        c.value = h
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center', wrap_text=True)
        c.border = border
    
    fila = 4
    if not df.empty:
        for _, row in df.iterrows():
            ws2.cell(fila, 1, row['FECHA'].strftime('%d/%m/%Y'))
            ws2.cell(fila, 2, row['EMPLEADO'])
            ws2.cell(fila, 3, row['PRODUCTO / SERVICIO'])
            ws2.cell(fila, 4, 'Producto' if row['ES_PRODUCTO'] else 'Servicio')
            ws2.cell(fila, 5, row['MONTO'])
            ws2.cell(fila, 6, f"{row['PORCENTAJE']*100:.0f}%")
            ws2.cell(fila, 7, row['COMISION'])
            ws2.cell(fila, 8, row['TIPO_COMISION'])
            
            for col in range(1, 9):
                c = ws2.cell(fila, col)
                c.border = border
                if col in [5, 7]:
                    c.number_format = '#,##0.00'
                    c.alignment = Alignment(horizontal='right')
            fila += 1
        
        ws2.cell(fila, 4, 'TOTAL').font = total_font
        ws2.cell(fila, 5, f'=SUM(E4:E{fila-1})')
        ws2.cell(fila, 7, f'=SUM(G4:G{fila-1})')
        
        for col in [5, 7]:
            c = ws2.cell(fila, col)
            c.fill = total_fill
            c.font = total_font
            c.number_format = '#,##0.00'
            c.border = border
    
    ws2.column_dimensions['A'].width = 12
    ws2.column_dimensions['B'].width = 15
    ws2.column_dimensions['C'].width = 40
    ws2.column_dimensions['D'].width = 12
    ws2.column_dimensions['E'].width = 12
    ws2.column_dimensions['F'].width = 10
    ws2.column_dimensions['G'].width = 12
    ws2.column_dimensions['H'].width = 18
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- INTERFAZ PRINCIPAL ---

# HEADER
st.markdown('<div class="main-header">💇‍♀️ BLUSH HAIR & MAKE-UP<br>Calculador de Comisiones</div>', unsafe_allow_html=True)
st.markdown("### 📍 Los Olivos, Lima - Perú")

# SIDEBAR
with st.sidebar:
    st.markdown("""
    <div style='text-align: center; padding: 20px; background: linear-gradient(135deg, #E91E63 0%, #9C27B0 100%); border-radius: 10px;'>
        <h1 style='color: white; margin: 0;'>💇‍♀️</h1>
        <h2 style='color: white; margin: 10px 0;'>BLUSH</h2>
        <p style='color: white; margin: 0;'>Hair & Make-Up</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### 📋 Instrucciones")
    st.markdown("""
    1. Sube tu archivo Excel de **Registro de Ventas**
    2. Revisa el **Dashboard** con los resultados
    3. Descarga el **Reporte Completo** en Excel
    
    ✨ Detección inteligente de productos vs servicios
    """)
    
    st.markdown("---")
    st.markdown("### 💰 Reglas de Comisión")
    st.markdown("""
    **Jhon y Yuri:**
    - Corte: 35%
    - Otros servicios: 25%
    - Productos: 10%
    
    **Julio Luna:**
    - Servicios: 40%
    - Productos: 10%
    
    **Otros:**
    - Servicios: 25%
    - Productos: 10%
    """)

# UPLOAD
uploaded_file = st.file_uploader(
    "📤 Arrastra aquí tu archivo Excel",
    type=['xlsx', 'xls'],
    help="Archivo de Registro de Ventas"
)

if uploaded_file:
    try:
        with st.spinner('⏳ Procesando datos...'):
            df = procesar_datos(uploaded_file)
            
            if df.empty:
                 st.error("⚠️ No se encontraron datos válidos o la estructura del Excel no es correcta.")
            else:
                resumen_df = crear_resumen(df)
                st.success('✅ Datos procesados correctamente!')
                
                # Info del archivo
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.info(f"📊 Se procesaron **{len(df)}** transacciones de **{len(resumen_df)}** empleados")
                with col2:
                    fecha_range = f"{df['FECHA'].min().strftime('%d/%m')} - {df['FECHA'].max().strftime('%d/%m/%Y')}"
                    st.info(f"📅 {fecha_range}")
                
                # METRICAS
                col1, col2, col3, col4 = st.columns(4)
                
                total_prod = resumen_df['TOTAL_PRODUCCION'].sum()
                total_com = resumen_df['TOTAL_COMISION'].sum()
                num_empleados = len(resumen_df)
                promedio_com = total_com / num_empleados if num_empleados > 0 else 0
                
                with col1:
                    st.metric("💵 Producción Total", f"S/ {total_prod:,.2f}")
                with col2:
                    st.metric("💰 Comisiones Totales", f"S/ {total_com:,.2f}")
                with col3:
                    st.metric("👥 Empleados", num_empleados)
                with col4:
                    st.metric("📊 Promedio", f"S/ {promedio_com:,.2f}")
                
                st.markdown("---")
                
                # TABS
                tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "📋 Nómina", "🔍 Detalle"])
                
                with tab1:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("#### 🏆 Top Empleados")
                        fig = px.bar(
                            resumen_df.head(5),
                            x='EMPLEADO',
                            y='TOTAL_COMISION',
                            color='TOTAL_COMISION',
                            color_continuous_scale=['#FCE4EC', '#E91E63'],
                            text='TOTAL_COMISION'
                        )
                        fig.update_traces(texttemplate='S/ %{text:,.2f}', textposition='outside')
                        fig.update_layout(showlegend=False, height=400)
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with col2:
                        st.markdown("#### 📈 Participación")
                        fig = px.pie(
                            resumen_df,
                            values='TOTAL_PRODUCCION',
                            names='EMPLEADO',
                            color_discrete_sequence=px.colors.sequential.RdPu
                        )
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)
                    
                    st.markdown("#### 📊 Servicios vs Productos")
                    
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        name='Servicios',
                        x=resumen_df['EMPLEADO'],
                        y=resumen_df['PRODUCCION_SERVICIOS'],
                        marker_color='#E91E63'
                    ))
                    fig.add_trace(go.Bar(
                        name='Productos',
                        x=resumen_df['EMPLEADO'],
                        y=resumen_df['PRODUCCION_PRODUCTOS'],
                        marker_color='#9C27B0'
                    ))
                    fig.update_layout(barmode='stack', height=400)
                    st.plotly_chart(fig, use_container_width=True)
                
                with tab2:
                    st.markdown("#### 💰 Nómina Completa")
                    
                    display_df = resumen_df[[
                        'EMPLEADO', 'PRODUCCION_SERVICIOS', 'COMISION_SERVICIOS',
                        'PRODUCCION_PRODUCTOS', 'COMISION_PRODUCTOS', 
                        'TOTAL_PRODUCCION', 'TOTAL_COMISION'
                    ]].copy()
                    
                    display_df.columns = ['Empleado', 'Prod. Servicios', 'Com. Servicios',
                                           'Prod. Productos', 'Com. Productos', 'Total Prod.', 'Total Comisión']
                    
                    st.dataframe(
                        display_df.style.format({
                            'Prod. Servicios': 'S/ {:,.2f}',
                            'Com. Servicios': 'S/ {:,.2f}',
                            'Prod. Productos': 'S/ {:,.2f}',
                            'Com. Productos': 'S/ {:,.2f}',
                            'Total Prod.': 'S/ {:,.2f}',
                            'Total Comisión': 'S/ {:,.2f}'
                        }).background_gradient(cmap='RdPu', subset=['Total Comisión']),
                        use_container_width=True,
                        height=400
                    )
                
                with tab3:
                    st.markdown("#### 🔍 Detalle de Transacciones")
                    
                    empleado_filter = st.multiselect(
                        'Filtrar por empleado:',
                        options=df['EMPLEADO'].unique(),
                        default=list(df['EMPLEADO'].unique())[:3]
                    )
                    
                    if not empleado_filter:
                        df_filtrado = df
                    else:
                        df_filtrado = df[df['EMPLEADO'].isin(empleado_filter)]
                    
                    display_detalle = df_filtrado[[
                        'FECHA', 'EMPLEADO', 'PRODUCTO / SERVICIO', 
                        'MONTO', 'TIPO_COMISION', 'COMISION', 'ES_PRODUCTO'
                    ]].copy()
                    
                    display_detalle['TIPO_REAL'] = display_detalle['ES_PRODUCTO'].apply(lambda x: 'Producto' if x else 'Servicio')
                    
                    display_detalle = display_detalle.rename(columns={
                        'FECHA': 'Fecha',
                        'EMPLEADO': 'Empleado',
                        'PRODUCTO / SERVICIO': 'Servicio/Producto',
                        'MONTO': 'Monto',
                        'TIPO_COMISION': 'Regla',
                        'COMISION': 'Comisión',
                        'TIPO_REAL': 'Tipo'
                    })
                    
                    display_detalle = display_detalle[['Fecha', 'Empleado', 'Servicio/Producto', 
                                                       'Tipo', 'Monto', 'Regla', 'Comisión']]
                    
                    st.dataframe(
                        display_detalle.style.format({
                            'Monto': 'S/ {:,.2f}',
                            'Comisión': 'S/ {:,.2f}'
                        }),
                        use_container_width=True,
                        height=500
                    )
                
                # BOTON DESCARGA
                st.markdown("---")
                excel_data = crear_excel_profesional(df, resumen_df)
                
                fecha_min = df['FECHA'].min().strftime('%d-%m-%Y')
                fecha_max = df['FECHA'].max().strftime('%d-%m-%Y')
                nombre_archivo = f"NOMINA_BLUSH_{fecha_min}_al_{fecha_max}.xlsx"
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.download_button(
                        label="📥 DESCARGAR REPORTE COMPLETO",
                        data=excel_data,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
    except Exception as e:
        st.error(f"❌ Error crítico: {str(e)}")

else:
    st.info("👆 Sube tu archivo Excel para comenzar")

# FOOTER
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>💇‍♀️ <b>BLUSH Hair & Make-Up Salon</b> | Los Olivos, Lima</p>
    <p style='font-size: 0.8rem;'>Sistema de Cálculo de Comisiones v2.0</p>
</div>
""", unsafe_allow_html=True)