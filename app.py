import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Cash Flow Diario", page_icon="💰", layout="wide")

# --- FUNCIONES DE PROCESAMIENTO ---

def limpiar_y_estandarizar(df, fuente):
    """
    Toma un dataframe, selecciona las columnas clave y las renombra
    para que todos tengan el mismo formato (Fecha, Importe, Categoria).
    """
    df_limpio = pd.DataFrame()
    
    if fuente == 'cheques':
        # --- LÓGICA PARA CHEQUES ---
        # Fecha: F.VTO
        # Categoria: Fijo como "Cheques" (pedido por usuario)
        if 'F.VTO' in df.columns and 'IMPORTE' in df.columns:
            df_limpio['Fecha'] = pd.to_datetime(df['F.VTO'], errors='coerce')
            df_limpio['Importe'] = pd.to_numeric(df['IMPORTE'], errors='coerce')
            
            # CAMBIO SOLICITADO: Categoría fija
            df_limpio['Categoria'] = 'Cheques' 
            
            # Guardamos el proveedor como dato extra en 'Detalle' por si acaso se necesita auditar
            df_limpio['Detalle'] = df['PROVEEDOR'] if 'PROVEEDOR' in df.columns else ''
        else:
            return None, "Faltan columnas 'F.VTO' o 'IMPORTE' en el archivo de Cheques."
            
    elif fuente == 'proyeccion':
        # --- LÓGICA PARA PROYECCIÓN ---
        # Fecha: Fecha Cobro
        # Categoria: Columna "Tipo" (pedido por usuario)
        if 'Fecha Cobro' in df.columns and 'Importe' in df.columns:
            df_limpio['Fecha'] = pd.to_datetime(df['Fecha Cobro'], errors='coerce')
            df_limpio['Importe'] = pd.to_numeric(df['Importe'], errors='coerce')
            
            # CAMBIO SOLICITADO: Usar columna 'Tipo'
            df_limpio['Categoria'] = df['Tipo'] if 'Tipo' in df.columns else 'Obligación Gral'
            
            df_limpio['Detalle'] = df['Detalle'] if 'Detalle' in df.columns else ''
        else:
            return None, "Faltan columnas 'Fecha Cobro' o 'Importe' en el archivo de Proyección."
    
    # Eliminar filas donde la fecha o el importe sean inválidos (NaT o NaN)
    df_limpio = df_limpio.dropna(subset=['Fecha', 'Importe'])
    return df_limpio, "OK"

def generar_cash_flow(df_cheques, df_proy):
    # 1. Estandarizar ambos archivos
    clean_cheques, msg_c = limpiar_y_estandarizar(df_cheques, 'cheques')
    clean_proy, msg_p = limpiar_y_estandarizar(df_proy, 'proyeccion')
    
    if clean_cheques is None: return None, msg_c
    if clean_proy is None: return None, msg_p
    
    # 2. Unificar en una sola lista
    df_total = pd.concat([clean_cheques, clean_proy])
    
    # 3. Lógica de Negocio (Vencido vs Futuro)
    hoy = pd.Timestamp(datetime.date.today())
    df_total['Estado'] = df_total['Fecha'].apply(lambda x: 'Vencido (Pasado)' if x < hoy else 'Futuro')
    
    # 4. Agregar nombres de días
    dias_es = {
        'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
        'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
    }
    df_total['Dia_Semana'] = df_total['Fecha'].dt.day_name().map(dias_es)
    
    # 5. Generar Cuadro Resumen (Agrupado por día)
    # Agrupamos solo por Fecha para ver el total diario a pagar
    resumen = df_total.groupby(['Estado', 'Fecha', 'Dia_Semana'])['Importe'].sum().reset_index()
    resumen = resumen.sort_values(by='Fecha')
    
    return df_total, resumen

# --- INTERFAZ GRÁFICA (FRONTEND) ---

st.title("💸 Automatización de Cash Flow")
st.markdown("""
Sube tus archivos para unificar **Cheques** y **Proyecciones**.
""")

col1, col2 = st.columns(2)
file_cheques = col1.file_uploader("📂 1. Subir Cheques.xlsx", type=['xlsx', 'xls'])
file_proy = col2.file_uploader("📂 2. Subir Proyeccion Pagos.xlsx", type=['xlsx', 'xls'])

if file_cheques and file_proy:
    if st.button("Generar Informe Unificado", type="primary"):
        with st.spinner("Procesando datos..."):
            try:
                # Cargar Excel
                df_c = pd.read_excel(file_cheques)
                df_p = pd.read_excel(file_proy)
                
                # Procesar
                df_detallado, df_resumen = generar_cash_flow(df_c, df_p)
                
                if df_detallado is not None:
                    st.success("✅ ¡Informe generado correctamente!")
                    
                    # Métricas
                    total_vencido = df_resumen[df_resumen['Estado'].str.contains('Vencido')]['Importe'].sum()
                    total_futuro = df_resumen[df_resumen['Estado'] == 'Futuro']['Importe'].sum()
                    
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Deuda Vencida", f"${total_vencido:,.0f}")
                    c2.metric("Deuda Futura", f"${total_futuro:,.0f}")
                    c3.metric("Total General", f"${(total_vencido + total_futuro):,.0f}")

                    # Tabla Resumen
                    st.subheader("📅 Resumen Diario de Caja")
                    st.dataframe(df_resumen.style.format({'Importe': '${:,.2f}'}), use_container_width=True)
                    
                    # Excel para descargar
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # Hoja 1: Resumen Diario
                        df_resumen.to_excel(writer, sheet_name='Resumen_Diario', index=False)
                        
                        # Hoja 2: Detalle (Aquí verás la columna "Categoria")
                        # Ordenamos las columnas para que Categoria quede al principio
                        cols_order = ['Fecha', 'Dia_Semana', 'Importe', 'Categoria', 'Estado', 'Detalle']
                        df_detallado[cols_order].sort_values(by='Fecha').to_excel(writer, sheet_name='Detalle_Completo', index=False)
                        
                        # Formato Moneda
                        workbook = writer.book
                        money_fmt = workbook.add_format({'num_format': '$ #,##0.00'})
                        
                        ws1 = writer.sheets['Resumen_Diario']
                        ws1.set_column('D:D', 18, money_fmt) # Columna Importe en Resumen
                        
                        ws2 = writer.sheets['Detalle_Completo']
                        ws2.set_column('C:C', 18, money_fmt) # Columna Importe en Detalle

                    st.download_button(
                        label="📥 Descargar Excel Final",
                        data=output.getvalue(),
                        file_name=f"CashFlow_{datetime.date.today()}.xlsx",
                        mime="application/vnd.ms-excel"
                    )
                else:
                    st.error(f"Error: {df_resumen}") 
            
            except Exception as e:
                st.error(f"Error inesperado: {e}")