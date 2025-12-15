
import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io

# ==========================================
# PARTE 1: PROCESAMIENTO DE DATOS
# ==========================================

fecha_hoy = pd.to_datetime(datetime.now().date())
# fecha_hoy = pd.to_datetime('2025-12-02') # Descomentar para probar con fecha fija

# Data de correlación incrustada directamente
data_nombres = {
    'Cheques': [
        'BBVA FRANCES BYC', 'BBVA FRANCES MPZ', 'BBVA FRANCES MBZ', 'BBVA FRANCES MGX',
        'BBVA FRANCES RG2', 'CREDICOOP BYC', 'CREDICOOP MGX', 'CREDICOOP MBZ',
        'CREDICOOP TMX', 'DE LA NACION ARG. BYC', 'DE LA NACION ARG MGX',
        'PATAGONIA MBZ', 'SANTANDER RIO BYC', 'SANTANDER RIO MBZ',
        'SANTANDER MGXD', 'MERCADO PAGO BYC', 'MERCADO PAGO MGX', 'MERCADO PAGO MBZ'
    ],
    'Proyeccion Pagos': [
        'Bco BBVA BYC SA', 'Bco BBVA MPZ BYC SA', 'Bco BBVA MBZ SRL', 'Bco BBVA MGXD SRL',
        'Bco BBVA RG2', 'Bco Credicoop BYC SA', 'Bco Credicoop MGXD SRL', 'Bco Credicoop MBZ SRL',
        'Bco Credicoop TMX SRL', 'Bco Nacion BYC SA', 'Bco Nacion MGXD SRL',
        'Bco Patagonia MBZ SRL', 'Bco Santander BYC SA', 'Bco Santander MBZ SRL',
        'Bco Santander MGXD SRL', 'MercadoPago BYC', 'MercadoPago MGX', 'MercadoPago MBZ'
    ],
    'EMPRESA': [
        'BYC', 'BYC', 'MBZ', 'MGX',
        'BYC', 'BYC', 'MGX', 'MBZ',
        'TMX', 'BYC', 'MGX',
        'MBZ', 'BYC', 'MBZ',
        'MGX', 'BYC', 'MGX', 'MBZ'
    ]
}
nombres_df = pd.DataFrame(data_nombres)

# Create a robust mapping dictionary from nombres_df
bank_mapping_dict = {}
for idx, row in nombres_df.iterrows():
    canonical_banco = row['Proyeccion Pagos'].strip() # Assume this is the canonical name
    empresa = row['EMPRESA'].strip()
    
    # Map from Cheques name to (canonical_banco, empresa)
    raw_cheque_name = row['Cheques'].strip()
    bank_mapping_dict[raw_cheheque_name] = (canonical_banco, empresa)
    
    # Map from Proyeccion Pagos name to (canonical_banco, empresa)
    bank_mapping_dict[canonical_banco] = (canonical_banco, empresa)
    
# Function to apply the mapping consistently
def apply_bank_mapping(raw_bank_name):
    mapped_info = bank_mapping_dict.get(raw_bank_name.strip())
    if mapped_info:
        return mapped_info[0], mapped_info[1] # Banco_Limpio, Empresa
    return raw_bank_name, 'UNKNOWN' # Fallback if no match is found

def procesar_archivo(file_object_or_path, col_banco, col_fecha, col_importe, tipo_origen, nombres_map_df):
    df = pd.read_excel(file_object_or_path)
    df_clean = pd.DataFrame({
        'Banco_Raw': df.iloc[:, col_banco].astype(str).str.strip(),
        'Fecha': pd.to_datetime(df.iloc[:, col_fecha], errors='coerce'),
        'Importe': pd.to_numeric(df.iloc[:, col_importe], errors='coerce'),
        'Origen': tipo_origen
    })
    df_clean = df_clean.dropna(subset=['Importe', 'Banco_Raw'])

    # Apply the centralized mapping
    df_clean[['Banco_Limpio', 'Empresa']] = df_clean['Banco_Raw'].apply(lambda x: pd.Series(apply_bank_mapping(x)))

    return df_clean

# ========================================== Streamlit UI ==========================================
st.title("Generador de Reporte de Cashflow")
st.write("Sube tus archivos de Excel para generar un reporte detallado.")

# Cargadores de archivos en la página principal
st.header("Cargar Archivos")
uploaded_file_proyeccion = st.file_uploader(
    "Sube el archivo 'Proyeccion Pagos.xlsx'",
    type=["xlsx"],
    key="proyeccion_pagos"
)
uploaded_file_cheques = st.file_uploader(
    "Sube el archivo 'Cheques.xlsx'",
    type=["xlsx"],
    key="cheques"
)
uploaded_file_saldos = st.file_uploader(
    "Sube el archivo 'Saldos.xlsx' (Col A: Banco, Col B: Saldo FCI, Col C: Saldo Banco)",
    type=["xlsx"],
    key="saldos"
)

if uploaded_file_proyeccion is not None and uploaded_file_cheques is not None and uploaded_file_saldos is not None:
    with st.spinner('Procesando datos y generando reporte...'):
        archivo_proyeccion_io = io.BytesIO(uploaded_file_proyeccion.getvalue())
        archivo_cheques_io = io.BytesIO(uploaded_file_cheques.getvalue())
        archivo_saldos_io = io.BytesIO(uploaded_file_saldos.getvalue())

        df_proy = procesar_archivo(archivo_proyeccion_io, 0, 2, 9, 'Proyeccion', nombres_df)
        df_cheq = procesar_archivo(archivo_cheques_io, 3, 1, 14, 'Cheques', nombres_df)
        df_total = pd.concat([df_proy, df_cheq])

        # Cargar saldos iniciales del archivo Saldos.xlsx
        df_saldos = pd.read_excel(archivo_saldos_io)

        # Map original column indices to new names as per instruction
        df_saldos_clean = pd.DataFrame({
            'Banco_Raw_Saldos': df_saldos.iloc[:, 0].astype(str).str.strip(), # Column A for Banco
            'Saldo FCI': pd.to_numeric(df_saldos.iloc[:, 1], errors='coerce'),  # Column B for Saldo FCI
            'Saldo Banco': pd.to_numeric(df_saldos.iloc[:, 2], errors='coerce')  # Column C for Saldo Banco
        })
        df_saldos_clean = df_saldos_clean.dropna(subset=['Saldo FCI', 'Saldo Banco'])

        # Apply the centralized mapping to saldos data
        df_saldos_clean[['Banco_Limpio', 'Empresa']] = df_saldos_clean['Banco_Raw_Saldos'].apply(lambda x: pd.Series(apply_bank_mapping(x)))

        df_saldos_clean = df_saldos_clean[['Empresa', 'Banco_Limpio', 'Saldo FCI', 'Saldo Banco']].drop_duplicates()
        df_saldos_clean = df_saldos_clean.set_index(['Empresa', 'Banco_Limpio'])

        # Periodos
        fecha_limite_semana = fecha_hoy + timedelta(days=5)

        # 1. Vencido
        filtro_vencido = df_total['Fecha'] < fecha_hoy
        df_vencido = df_total[filtro_vencido].groupby(['Empresa', 'Banco_Limpio'])[['Importe']].sum() 
        df_vencido.columns = ['Vencido']

        # 2. Semana (Días)
        filtro_semana = (df_total['Fecha'] >= fecha_hoy) & (df_total['Fecha'] <= fecha_limite_semana)
        df_semana_data = df_total[filtro_semana].copy()
        dias_es_full = {0:'Lunes', 1:'Martes', 2:'Miércoles', 3:'Jueves', 4:'Viernes', 5:'Sábado', 6:'Domingo'}

        expected_day_columns = []
        for i in range(6):
            current_date = fecha_hoy + timedelta(days=i)
            expected_day_columns.append(f"{current_date.strftime('%d-%b')}\n{dias_es_full[current_date.weekday()]}")

        df_semana_data['Nombre_Dia'] = df_semana_data['Fecha'].apply(lambda x: f"{x.strftime('%d-%b')}\n{dias_es_full[x.weekday()]}")

        df_semana_pivot = pd.pivot_table(
            df_semana_data, index=['Empresa', 'Banco_Limpio'], columns='Nombre_Dia', values='Importe', aggfunc='sum', fill_value=0
        )

        df_semana_pivot = df_semana_pivot.reindex(columns=expected_day_columns, fill_value=0)

        df_semana_pivot['Total Semana'] = df_semana_pivot.sum(axis=1)

        # 3. Emitidos (Futuro solo cheques)
        filtro_emitidos = (df_total['Fecha'] > fecha_limite_semana) & (df_total['Origen'] == 'Cheques')
        df_emitidos = df_total[filtro_emitidos].groupby(['Empresa', 'Banco_Limpio'])[['Importe']].sum()
        df_emitidos.columns = ['Emitidos']

        # Unir todo usando left merges, con df_saldos_clean como base
        reporte_final = df_saldos_clean.copy() # Start with all banks from Saldos.xlsx

        reporte_final = pd.merge(
            reporte_final,
            df_vencido, # Merge Vencido
            left_index=True,
            right_index=True,
            how='left'
        )

        reporte_final = pd.merge(
            reporte_final,
            df_semana_pivot, # Merge Semana
            left_index=True,
            right_index=True,
            how='left'
        )

        reporte_final = pd.merge(
            reporte_final,
            df_emitidos, # Merge Emitidos
            left_index=True,
            right_index=True,
            how='left'
        )

        # After all merges, fill NaN values with 0
        reporte_final = reporte_final.fillna(0)

        # Calcular 'A Cubrir Vencido' como (Saldo Banco - Vencido)
        reporte_final['A Cubrir Vencido'] = reporte_final['Saldo Banco'] - reporte_final['Vencido']

        # Calculate 'A Cubrir Semana'
        reporte_final['A Cubrir Semana'] = reporte_final['Saldo Banco'] - reporte_final['Vencido'] - reporte_final['Total Semana']

        # Reordenar columnas para colocar 'Saldo Banco' y 'Saldo FCI' antes de 'Vencido'
        # y luego 'A Cubrir Vencido' e 'A Cubrir Semana' al final.
        cols = reporte_final.columns.tolist()

        # Define lists for new column order
        new_order_cols = []

        # 1. Add static leading columns
        if 'Saldo Banco' in cols: new_order_cols.append('Saldo Banco')
        if 'Saldo FCI' in cols: new_order_cols.append('Saldo FCI')
        if 'Vencido' in cols: new_order_cols.append('Vencido')

        # 2. Add daily columns in their specific order
        for col in expected_day_columns:
            if col in cols:
                new_order_cols.append(col)

        # 3. Add 'Total Semana' and 'Emitidos'
        if 'Total Semana' in cols: new_order_cols.append('Total Semana')
        if 'Emitidos' in cols: new_order_cols.append('Emitidos')

        # 4. Collect other existing columns that are not 'A Cubrir Vencido' or 'A Cubrir Semana'
        #    and have not been added yet. This handles any unforeseen columns and ensures
        #    ACV and ACS are indeed last.
        for col in cols:
            if col not in new_order_cols and col != 'A Cubrir Vencido' and col != 'A Cubrir Semana':
                new_order_cols.append(col)

        # 5. Append 'A Cubrir Vencido' and 'A Cubrir Semana' at the very end in the specified order
        if 'A Cubrir Vencido' in cols:
            new_order_cols.append('A Cubrir Vencido')
        if 'A Cubrir Semana' in cols:
            new_order_cols.append('A Cubrir Semana')

        reporte_final = reporte_final[new_order_cols]

        # ========================================== Streamlit Output ==========================================
        st.subheader("Reporte de Cashflow Generado")
        st.dataframe(reporte_final)

        # Para la descarga del Excel
        output_excel_data = io.BytesIO()
        writer = pd.ExcelWriter(output_excel_data, engine='xlsxwriter')
        workbook = writer.book
        worksheet = workbook.add_worksheet('Resumen')

        # --- DEFINICIÓN DE FORMATOS ---
        # Define default font for all formats
        default_font_properties = {'font_name': 'Bahnshift SemiLight'}

        fmt_header = workbook.add_format({
            **default_font_properties,
            'bold': True, 'font_color': 'white', 'bg_color': '#ED7D31',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'text_wrap': True
        })
        fmt_subtotal = workbook.add_format({
            **default_font_properties,
            'bold': True, 'bg_color': '#FCE4D6', 'num_format': '$ #,##0',
            'border': 1
        })
        fmt_currency = workbook.add_format({
            **default_font_properties,
            'num_format': '$ #,##0', 'border': 1
        })
        fmt_text = workbook.add_format({
            **default_font_properties,
            'border': 1
        })

        # New formats for conditional formatting on 'A Cubrir Vencido' and 'A Cubrir Semana'
        fmt_positive_acv = workbook.add_format({
            **default_font_properties,
            'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '$ #,##0', 'border': 1
        })
        fmt_negative_acv = workbook.add_format({
            **default_font_properties,
            'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '$ #,##0', 'border': 1
        })
        
        # New format for the grand total row
        fmt_grand_total = workbook.add_format({
            **default_font_properties,
            'bold': True, 'bg_color': '#BFBFBF', 'num_format': '$ #,##0',
            'border': 1, 'align': 'left', 'valign': 'vcenter'
        })

        # --- ESCRIBIR ENCABEZADOS ---
        worksheet.write('A1', 'Resumen Cashflow', workbook.add_format({**default_font_properties, 'bold': True, 'font_size': 14}))
        worksheet.write('A2', f"Fecha Actual: {fecha_hoy.strftime('%d/%m/%Y')}")

        fila_actual = 3
        col_bancos = 0
        worksheet.write(fila_actual, col_bancos, "Etiquetas de fila", fmt_header)

        columnas_datos = reporte_final.columns.tolist()

        # Find the index of 'A Cubrir Vencido' for conditional formatting
        acv_col_idx = -1
        if 'A Cubrir Vencido' in columnas_datos:
            acv_col_idx = columnas_datos.index('A Cubrir Vencido') + 1 # +1 because of the bank column at index 0
        
        # Find the index of 'A Cubrir Semana' for conditional formatting
        acs_col_idx = -1
        if 'A Cubrir Semana' in columnas_datos:
            acs_col_idx = columnas_datos.index('A Cubrir Semana') + 1 # +1 because of the bank column at index 0

        for i, col_name in enumerate(columnas_datos):
            worksheet.write(fila_actual, i + 1, col_name, fmt_header)

        fila_actual += 1

        # --- ESCRIBIR DATOS POR GRUPO (EMPRESA) ---
        empresas_unicas = reporte_final.index.get_level_values(0).unique()

        for empresa in empresas_unicas:
            datos_empresa = reporte_final.loc[empresa]

            if isinstance(datos_empresa, pd.Series):
                banco_limpio_idx = datos_empresa.name[1]
                datos_empresa = pd.DataFrame(datos_empresa).T
                datos_empresa.index = [banco_limpio_idx]
                datos_empresa.index.name = 'Banco_Limpio'

            for banco, row in datos_empresa.iterrows():
                worksheet.write(fila_actual, 0, banco, fmt_text)

                for i, val in enumerate(row):
                    current_col_excel_idx = i + 1
                    if current_col_excel_idx == acv_col_idx or current_col_excel_idx == acs_col_idx:
                        if val > 0:
                            worksheet.write(fila_actual, current_col_excel_idx, val, fmt_positive_acv)
                        elif val < 0:
                            worksheet.write(fila_actual, current_col_excel_idx, val, fmt_negative_acv)
                        else:
                            worksheet.write(fila_actual, current_col_excel_idx, val, fmt_currency) # Default for 0
                    else:
                        worksheet.write(fila_actual, current_col_excel_idx, val, fmt_currency)

                fila_actual += 1

            # --- CREAR FILA DE SUBTOTAL ---
            worksheet.write(fila_actual, 0, f"Total {empresa}", fmt_subtotal)

            sumas = datos_empresa.sum()
            for i, val in enumerate(sumas):
                # Apply subtotal format for 'A Cubrir Vencido' as well, without specific conditional coloring
                worksheet.write(fila_actual, i + 1, val, fmt_subtotal)

            fila_actual += 1
        
        # --- CREAR FILA DE TOTAL BANCOS ---
        # Sum all numeric columns for the grand total row
        grand_totals_series = reporte_final.select_dtypes(include=['number']).sum()
        grand_totals = {col: grand_totals_series.get(col, '') for col in columnas_datos}

        worksheet.write(fila_actual, 0, "TOTAL BANCOS", fmt_grand_total)

        for i, col_name in enumerate(columnas_datos):
            val = grand_totals.get(col_name, "") # Get calculated total or empty string
            worksheet.write(fila_actual, i + 1, val, fmt_grand_total)

        fila_actual += 1

        # Ajustar ancho de columnas
        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, len(columnas_datos), 15)

        writer.close()
        output_excel_data.seek(0)

        st.download_button(
            label="Descargar Reporte de Cashflow Formateado",
            data=output_excel_data,
            file_name="Resumen_Cashflow_Formateado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("¡Listo! Archivo generado y disponible para descarga.")

else:
    st.info("Por favor, sube los archivos para generar el reporte de cashflow.")