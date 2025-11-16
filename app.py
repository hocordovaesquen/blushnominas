import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, Reference
import io
from datetime import datetime

# --- LÓGICA DE CÁLCULO (Tu código, sin cambios) ---
# He movido tu lógica a una función para mantener el script ordenado

def process_excel(excel_file):
    """
    Toma un archivo Excel cargado, procesa las comisiones 
    y devuelve un objeto Workbook de openpyxl.
    """
    
    # LEER Y PROCESAR DATOS
    df = pd.read_excel(excel_file, sheet_name='Hoja1', skiprows=9)
    df = df[df['EMPLEADO'].notna()].copy()
    df['EMPLEADO'] = df['EMPLEADO'].str.strip()
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce').ffill()
    df['ES_PRODUCTO'] = df['CLASE'].fillna('Servicio') == 'Producto'
    df['MONTO'] = df['TOTAL'].fillna(0)
    
    def calcular_comision_y_porcentaje(row):
        empleado = row['EMPLEADO']
        servicio = row['PRODUCTO / SERVICIO']
        monto = row['MONTO']
        es_producto = row['ES_PRODUCTO']
        
        if es_producto:
            porcentaje = 0.10
            tipo_comision = "Producto 10%"
        elif empleado in ['Jhon', 'Yuri']:
            if pd.notna(servicio) and 'corte' in str(servicio).lower():
                porcentaje = 0.35
                tipo_comision = "Corte 35%"
            else:
                porcentaje = 0.25
                tipo_comision = "Servicio 25%"
        elif empleado == 'Julio':
            porcentaje = 0.40
            tipo_comision = "Servicio 40%"
        else:
            porcentaje = 0.25
            tipo_comision = "Servicio 25%"
        
        comision = monto * porcentaje
        
        return pd.Series({
            'COMISION': comision,
            'PORCENTAJE': porcentaje,
            'TIPO_COMISION': tipo_comision
        })
    
    df[['COMISION', 'PORCENTAJE', 'TIPO_COMISION']] = df.apply(calcular_comision_y_porcentaje, axis=1)
    
    # RESUMEN POR EMPLEADO
    resumen_lista = []
    for emp in sorted(df['EMPLEADO'].unique()):
        d_emp = df[df['EMPLEADO'] == emp]
        servicios = d_emp[~d_emp['ES_PRODUCTO']]
        productos = d_emp[d_emp['ES_PRODUCTO']]
        
        total_transacciones = len(d_emp)
        ticket_promedio = d_emp['MONTO'].mean() if total_transacciones > 0 else 0
        
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
            'TOTAL_TRANSACCIONES': total_transacciones,
            'TICKET_PROMEDIO': ticket_promedio
        })
    
    resumen_df = pd.DataFrame(resumen_lista)
    resumen_df['PARTICIPACION'] = (resumen_df['TOTAL_PRODUCCION'] / resumen_df['TOTAL_PRODUCCION'].sum() * 100)
    resumen_df = resumen_df.sort_values('TOTAL_COMISION', ascending=False).reset_index(drop=True) # Reset index para Top 3
    
    # ESTADISTICAS GENERALES
    fecha_min = df['FECHA'].min()
    fecha_max = df['FECHA'].max()
    dias_periodo = (fecha_max - fecha_min).days + 1
    
    ventas_por_dia = df.groupby(df['FECHA'].dt.date)['MONTO'].sum()
    mejor_dia = ventas_por_dia.idxmax()
    venta_mejor_dia = ventas_por_dia.max()
    
    total_produccion = resumen_df['TOTAL_PRODUCCION'].sum()
    total_comisiones = resumen_df['TOTAL_COMISION'].sum()
    promedio_diario = total_produccion / dias_periodo
    
    # CREAR EXCEL PROFESIONAL (El mismo código tuyo)
    wb = Workbook()
    
    # COLORES CORPORATIVOS BLUSH
    rosa_blush = "E91E63"
    rosa_claro = "FCE4EC"
    purpura = "9C27B0"
    gris_oscuro = "424242"
    verde_exito = "4CAF50"
    azul_info = "2196F3"
    
    # ESTILOS
    header_fill = PatternFill(start_color=rosa_blush, end_color=rosa_blush, fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11, name='Arial')
    
    subheader_fill = PatternFill(start_color=rosa_claro, end_color=rosa_claro, fill_type="solid")
    subheader_font = Font(bold=True, color=gris_oscuro, size=10, name='Arial')
    
    total_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    total_font = Font(bold=True, size=11, name='Arial')
    
    highlight_fill = PatternFill(start_color=verde_exito, end_color=verde_exito, fill_type="solid")
    highlight_font = Font(bold=True, color="FFFFFF", size=11, name='Arial')
    
    border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    # ================================================================
    # HOJA 1: RESUMEN EJECUTIVO
    # ================================================================
    ws1 = wb.active
    ws1.title = "RESUMEN EJECUTIVO"
    
    # Encabezado corporativo
    ws1.merge_cells('A1:F1')
    ws1['A1'] = 'BLUSH HAIR & MAKE-UP SALON'
    ws1['A1'].font = Font(bold=True, size=18, color=rosa_blush, name='Arial')
    ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
    
    ws1.merge_cells('A2:F2')
    ws1['A2'] = 'Los Olivos - Lima, Peru'
    ws1['A2'].font = Font(size=11, color=gris_oscuro, name='Arial')
    ws1['A2'].alignment = Alignment(horizontal='center')
    
    ws1.merge_cells('A3:F3')
    ws1['A3'] = f'REPORTE DE COMISIONES - {fecha_min.strftime("%d/%m/%Y")} al {fecha_max.strftime("%d/%m/%Y")}'
    ws1['A3'].font = Font(bold=True, size=14, name='Arial')
    ws1['A3'].alignment = Alignment(horizontal='center')
    
    ws1['A4'] = f'Generado: {datetime.now().strftime("%d/%m/%Y %H:%M")}'
    ws1['A4'].font = Font(size=9, italic=True, color='666666')
    
    # KPIs Principales
    fila = 6
    kpis = [
        ('PRODUCCION TOTAL', f'S/ {total_produccion:,.2f}', verde_exito),
        ('COMISIONES TOTALES', f'S/ {total_comisiones:,.2f}', rosa_blush),
        ('PROMEDIO DIARIO', f'S/ {promedio_diario:,.2f}', azul_info),
        ('DIAS DEL PERIODO', str(dias_periodo), purpura),
    ]
    
    col = 1
    for titulo, valor, color in kpis:
        ws1.cell(fila, col, titulo).font = Font(bold=True, size=9, name='Arial')
        ws1.cell(fila, col).alignment = Alignment(horizontal='center')
        
        ws1.cell(fila+1, col, valor).font = Font(bold=True, size=16, color=color, name='Arial')
        ws1.cell(fila+1, col).alignment = Alignment(horizontal='center')
        
        ws1.cell(fila, col).fill = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')
        ws1.cell(fila+1, col).fill = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')
        
        for r in [fila, fila+1]:
            ws1.cell(r, col).border = border
        
        col += 2
    
    # Mejor dia
    fila = 9
    ws1.merge_cells(f'A{fila}:F{fila}')
    ws1[f'A{fila}'] = f'MEJOR DIA: {mejor_dia.strftime("%d/%m/%Y")} - S/ {venta_mejor_dia:,.2f}'
    ws1[f'A{fila}'].font = Font(bold=True, size=12, color=verde_exito, name='Arial')
    ws1[f'A{fila}'].alignment = Alignment(horizontal='center')
    
    # Top 3
    fila = 11
    ws1.merge_cells(f'A{fila}:F{fila}')
    ws1[f'A{fila}'] = 'TOP 3 EMPLEADOS DEL PERIODO'
    ws1[f'A{fila}'].font = Font(bold=True, size=13, name='Arial')
    ws1[f'A{fila}'].alignment = Alignment(horizontal='center')
    ws1[f'A{fila}'].fill = subheader_fill
    
    fila = 13
    headers_top = ['PUESTO', 'EMPLEADO', 'PRODUCCION', 'COMISION', 'PARTICIPACION', 'SERVICIOS']
    for col, h in enumerate(headers_top, 1):
        c = ws1.cell(fila, col)
        c.value = h
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center')
        c.border = border
    
    medallas = ['🥇', '🥈', '🥉']
    fila = 14
    for idx, (i, row) in enumerate(resumen_df.head(3).iterrows()):
        ws1.cell(fila, 1, f'{medallas[idx]} {idx+1}')
        ws1.cell(fila, 2, row['EMPLEADO'])
        ws1.cell(fila, 3, row['TOTAL_PRODUCCION'])
        ws1.cell(fila, 4, row['TOTAL_COMISION'])
        ws1.cell(fila, 5, f"{row['PARTICIPACION']:.1f}%")
        ws1.cell(fila, 6, row['CANTIDAD_SERVICIOS'])
        
        for col in range(1, 7):
            c = ws1.cell(fila, col)
            c.border = border
            c.alignment = Alignment(horizontal='center' if col in [1, 2, 5, 6] else 'right')
            if col in [3, 4]:
                c.number_format = '#,##0.00'
        
        fila += 1
    
    # Anchos
    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 16
    ws1.column_dimensions['D'].width = 16
    ws1.column_dimensions['E'].width = 14
    ws1.column_dimensions['F'].width = 14
    
    # ================================================================
    # HOJA 2: NOMINA COMPLETA
    # ================================================================
    ws2 = wb.create_sheet("NOMINA")
    
    # Titulo
    ws2.merge_cells('A1:K1')
    ws2['A1'] = f'NOMINA QUINCENAL - {fecha_min.strftime("%d/%m/%Y")} AL {fecha_max.strftime("%d/%m/%Y")}'
    ws2['A1'].font = Font(bold=True, size=14, name='Arial')
    ws2['A1'].alignment = Alignment(horizontal='center')
    ws2['A1'].fill = PatternFill(start_color=rosa_claro, end_color=rosa_claro, fill_type='solid')
    
    # Encabezados
    headers = ['', 'EMPLEADO', 'PROD. SERVICIOS', 'COM. SERVICIOS', 'PROD. PRODUCTOS', 
             'COM. PRODUCTOS', 'TOTAL PROD.', 'TOTAL COM.', 'DESCUENTOS', 'EXTRAS', 'A PAGAR']
    
    for col, h in enumerate(headers, 1):
        c = ws2.cell(3, col)
        c.value = h
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = border
    
    # Datos
    fila = 4
    for idx, row in resumen_df.iterrows():
        ws2.cell(fila, 1, idx+1)
        ws2.cell(fila, 2, row['EMPLEADO'])
        ws2.cell(fila, 3, row['PRODUCCION_SERVICIOS'])
        ws2.cell(fila, 4, row['COMISION_SERVICIOS'])
        ws2.cell(fila, 5, row['PRODUCCION_PRODUCTOS'])
        ws2.cell(fila, 6, row['COMISION_PRODUCTOS'])
        ws2.cell(fila, 7, row['TOTAL_PRODUCCION'])
        ws2.cell(fila, 8, row['TOTAL_COMISION'])
        ws2.cell(fila, 9, 0)
        ws2.cell(fila, 10, 0)
        ws2.cell(fila, 11, f'=H{fila}-I{fila}+J{fila}')
        
        for col in range(1, 12):
            c = ws2.cell(fila, col)
            c.border = border
            if col > 2:
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal='right')
            if col in [7, 8, 11]:
                c.fill = highlight_fill
                c.font = highlight_font
        
        fila += 1
    
    # Totales
    ws2.cell(fila, 2, 'TOTAL').font = total_font
    for col in range(3, 12):
        if col in [9, 10]:
             ws2.cell(fila, col, 0)
        else:
             ws2.cell(fila, col, f'=SUM({chr(64+col)}4:{chr(64+col)}{fila-1})')
        c = ws2.cell(fila, col)
        c.fill = total_fill
        c.font = total_font
        c.number_format = '#,##0.00'
        c.border = border
    
    # Anchos
    ws2.column_dimensions['A'].width = 5
    ws2.column_dimensions['B'].width = 20
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        ws2.column_dimensions[col].width = 14
    
    # ================================================================
    # HOJA 3: RANKING Y ESTADISTICAS
    # ================================================================
    ws3 = wb.create_sheet("RANKING")
    
    ws3.merge_cells('A1:G1')
    ws3['A1'] = 'RANKING DE DESEMPENO'
    ws3['A1'].font = Font(bold=True, size=14, name='Arial')
    ws3['A1'].alignment = Alignment(horizontal='center')
    ws3['A1'].fill = subheader_fill
    
    headers_rank = ['POS', 'EMPLEADO', 'TOTAL PROD.', 'TOTAL COM.', 'SERVICIOS', 'TICKET PROM.', 'PARTICIPACION']
    
    for col, h in enumerate(headers_rank, 1):
        c = ws3.cell(3, col)
        c.value = h
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center')
        c.border = border
    
    fila = 4
    for idx, row in resumen_df.iterrows():
        ws3.cell(fila, 1, idx+1)
        ws3.cell(fila, 2, row['EMPLEADO'])
        ws3.cell(fila, 3, row['TOTAL_PRODUCCION'])
        ws3.cell(fila, 4, row['TOTAL_COMISION'])
        ws3.cell(fila, 5, row['TOTAL_TRANSACCIONES'])
        ws3.cell(fila, 6, row['TICKET_PROMEDIO'])
        ws3.cell(fila, 7, f"{row['PARTICIPACION']:.1f}%")
        
        for col in range(1, 8):
            c = ws3.cell(fila, col)
            c.border = border
            if col in [3, 4, 6]:
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal='right')
            elif col in [1, 5, 7]:
                c.alignment = Alignment(horizontal='center')
        
        if idx < 3:
            for col in range(1, 8):
                ws3.cell(fila, col).fill = PatternFill(start_color='FFF9C4', end_color='FFF9C4', fill_type='solid')
        
        fila += 1
    
    ws3.column_dimensions['A'].width = 8
    ws3.column_dimensions['B'].width = 20
    for col in ['C', 'D', 'E', 'F', 'G']:
        ws3.column_dimensions[col].width = 16
    
    # ================================================================
    # HOJA 4: DETALLE DE CALCULOS
    # ================================================================
    ws4 = wb.create_sheet("DETALLE_CALCULOS")
    
    ws4.merge_cells('A1:H1')
    ws4['A1'] = 'DETALLE COMPLETO DE CALCULOS'
    ws4['A1'].font = Font(bold=True, size=14, name='Arial')
    ws4['A1'].alignment = Alignment(horizontal='center')
    ws4['A1'].fill = subheader_fill
    
    headers_det = ['FECHA', 'EMPLEADO', 'SERVICIO/PRODUCTO', 'TIPO', 'MONTO', 
                   '%', 'COMISION', 'REGLA']
    
    for col, h in enumerate(headers_det, 1):
        c = ws4.cell(3, col)
        c.value = h
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center', wrap_text=True)
        c.border = border
    
    fila = 4
    for _, row in df.iterrows():
        ws4.cell(fila, 1, row['FECHA'].strftime('%d/%m/%Y'))
        ws4.cell(fila, 2, row['EMPLEADO'])
        ws4.cell(fila, 3, row['PRODUCTO / SERVICIO'])
        ws4.cell(fila, 4, 'Producto' if row['ES_PRODUCTO'] else 'Servicio')
        ws4.cell(fila, 5, row['MONTO'])
        ws4.cell(fila, 6, f"{row['PORCENTAJE']*100:.0f}%")
        ws4.cell(fila, 7, row['COMISION'])
        ws4.cell(fila, 8, row['TIPO_COMISION'])
        
        for col in range(1, 9):
            c = ws4.cell(fila, col)
            c.border = border
            if col in [5, 7]:
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal='right')
            elif col == 6:
                c.alignment = Alignment(horizontal='center')
        
        fila += 1
    
    # Totales
    ws4.cell(fila, 4, 'TOTAL').font = total_font
    ws4.cell(fila, 5, f'=SUM(E4:E{fila-1})')
    ws4.cell(fila, 7, f'=SUM(G4:G{fila-1})')
    
    for col in [5, 7]:
        c = ws4.cell(fila, col)
        c.fill = total_fill
        c.font = total_font
        c.number_format = '#,##0.00'
        c.border = border
    
    ws4.column_dimensions['A'].width = 12
    ws4.column_dimensions['B'].width = 15
    ws4.column_dimensions['C'].width = 35
    ws4.column_dimensions['D'].width = 12
    ws4.column_dimensions['E'].width = 12
    ws4.column_dimensions['F'].width = 10
    ws4.column_dimensions['G'].width = 12
    ws4.column_dimensions['H'].width = 16

    # Información de resumen para Streamlit
    resumen_kpis = {
        "periodo": f"{fecha_min.strftime('%d/%m/%Y')} al {fecha_max.strftime('%d/%m/%Y')} ({dias_periodo} dias)",
        "total_produccion": total_produccion,
        "total_comisiones": total_comisiones,
        "promedio_diario": promedio_diario,
        "mejor_dia": f"{mejor_dia.strftime('%d/%m/%Y')} (S/ {venta_mejor_dia:,.2f})"
    }
    
    nombre_salida = f"NOMINA_BLUSH_{fecha_min.strftime('%d-%m-%Y')}_al_{fecha_max.strftime('%d-%m-%Y')}.xlsx"
    
    return wb, resumen_df.head(3), resumen_kpis, nombre_salida

# --- INTERFAZ DE LA APP STREAMLIT ---

st.set_page_config(page_title="Comisiones Blush", page_icon="💅")

st.title("💅 CALCULADOR PROFESIONAL DE COMISIONES BLUSH")
st.subheader("Los Olivos, Lima")
st.markdown("---")

st.header("1. Sube tu archivo de ventas")
st.write("Asegúrate de subir el archivo `Listado_De_Registro_VentasV2_...xlsx`")

uploaded_file = st.file_uploader(
    "Selecciona el archivo Excel:", 
    type=["xlsx"],
    help="Solo se aceptan archivos .xlsx"
)

if uploaded_file is not None:
    st.success(f"Archivo recibido: **{uploaded_file.name}**")
    st.markdown("---")
    st.header("2. Procesar y Descargar Reporte")
    
    if st.button("Generar Reporte Profesional"):
        try:
            with st.spinner("Procesando datos y creando reporte... Esto puede tardar un momento..."):
                
                # Procesar el archivo
                workbook, top_3, kpis, filename = process_excel(uploaded_file)
                
                # Mostrar Resumen Ejecutivo en la App
                st.subheader("Resumen Ejecutivo del Periodo")
                st.write(f"**Periodo:** {kpis['periodo']}")
                
                col1, col2 = st.columns(2)
                col1.metric("Producción Total", f"S/ {kpis['total_produccion']:,.2f}")
                col2.metric("Comisiones Totales", f"S/ {kpis['total_comisiones']:,.2f}")
                
                col1.metric("Promedio Diario", f"S/ {kpis['promedio_diario']:,.2f}")
                col2.metric("Mejor Día", kpis['mejor_dia'])

                st.subheader("Top 3 Empleados")
                # Preparar top_3 para mostrarlo
                top_3_display = top_3[['EMPLEADO', 'TOTAL_PRODUCCION', 'TOTAL_COMISION', 'PARTICIPACION']].copy()
                top_3_display['PARTICIPACION'] = top_3_display['PARTICIPACION'].apply(lambda x: f"{x:.1f}%")
                st.dataframe(top_3_display)
                
                # Preparar el archivo para descarga
                output_buffer = io.BytesIO()
                workbook.save(output_buffer)
                
                st.markdown("---")
                st.header("3. Descargar")
                
                st.download_button(
                    label="Clic aquí para descargar el Reporte Excel",
                    data=output_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success(f"¡Reporte '{filename}' generado con éxito!")

        except Exception as e:
            st.error(f"Ha ocurrido un error al procesar el archivo: {e}")
            st.warning("Verifica que el archivo subido sea el correcto y tenga el formato esperado (ej. `Listado_De_Registro_VentasV2...`).")

else:
    st.info("Esperando que subas el archivo Excel...")