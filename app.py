
import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io
from fpdf import FPDF

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
    bank_mapping_dict[raw_cheque_name] = (canonical_banco, empresa)

    # Map from Proyeccion Pagos name to (canonical_banco, empresa)
    bank_mapping_dict[canonical_banco] = (canonical_banco, empresa)

# Function to apply the mapping consistently
def apply_bank_mapping(raw_bank_name):
    mapped_info = bank_mapping_dict.get(raw_bank_name.strip())
    if mapped_info:
        return mapped_info[0], mapped_info[1] # Banco_Limpio, Empresa
    return raw_bank_name, 'UNKNOWN' # Fallback if no match is found

def procesar_archivo(file_object_or_path, col_banco, col_fecha, col_importe, tipo_origen, nombres_map_df, col_detalle=None, col_numero_cheque=None):
    df = pd.read_excel(file_object_or_path)

    # NEW CONDITIONAL FILTER: Only for 'Proyeccion' files, filter where column H (index 7) is empty
    if tipo_origen == 'Proyeccion':
        df = df[df.iloc[:, 7].isnull()].copy() # Filter rows where column H is NaN

    df_clean = pd.DataFrame({
        'Banco_Raw': df.iloc[:, col_banco].astype(str).str.strip(),
        'Fecha': pd.to_datetime(df.iloc[:, col_fecha], errors='coerce'),
        'Importe': pd.to_numeric(df.iloc[:, col_importe], errors='coerce'),
        'Origen': tipo_origen
    })
    df_clean = df_clean.dropna(subset=['Importe', 'Banco_Raw'])

    if col_detalle is not None:
        df_clean['Detalle'] = df.iloc[:, col_detalle].astype(str).str.strip()
    else:
        df_clean['Detalle'] = '' # Default empty string if no detail column

    if col_numero_cheque is not None:
        df_clean['Numero_Cheque'] = df.iloc[:, col_numero_cheque].astype(str).str.strip()
    else:
        df_clean['Numero_Cheque'] = '' # Default empty string if no cheque number column

    # Apply the centralized mapping
    df_clean[['Banco_Limpio', 'Empresa']] = df_clean['Banco_Raw'].apply(lambda x: pd.Series(apply_bank_mapping(x)))

    return df_clean

def procesar_archivo_impuestos(file_object_or_path):
    df = pd.read_excel(file_object_or_path)

    # Extract data from specified columns
    df_impuestos_clean = pd.DataFrame({
        'Empresa_Raw': df.iloc[:, 2].astype(str).str.strip(), # Column C
        'Fecha': pd.to_datetime(df.iloc[:, 5], errors='coerce'), # Column F
        'Importe': pd.to_numeric(df.iloc[:, 6], errors='coerce'), # Column G
        'Estado': df.iloc[:, 11].astype(str).str.strip(), # Column L
        'Detalle': df.iloc[:, 1].astype(str).str.strip() # Column B for Detalle
    })

    # Filter based on 'Estado'
    df_impuestos_clean = df_impuestos_clean[df_impuestos_clean['Estado'].isin(['VENCIDO', 'A PAGAR'])].copy()

    # Convert 'Importe' to numeric
    # df_impuestos_clean['Importe'] = df_impuestos_clean['Importe'] * -1 # REMOVED: User wants positive sign

    # Add 'Origen' column
    df_impuestos_clean['Origen'] = 'Impuestos'

    # Add empty 'Numero_Cheque' column for consistency
    df_impuestos_clean['Numero_Cheque'] = ''

    # Create mapping from nombres_df for 'Empresa' to 'Banco_Limpio'
    # Group by 'EMPRESA' and take the first 'Proyeccion Pagos' as the default 'Banco_Limpio'
    empresa_to_default_bank = nombres_df.groupby('EMPRESA')['Proyeccion Pagos'].first().to_dict()

    # Apply mapping to create 'Banco_Limpio' and handle 'UNKNOWN'
    df_impuestos_clean['Banco_Limpio'] = df_impuestos_clean['Empresa_Raw'].map(empresa_to_default_bank)
    df_impuestos_clean['Banco_Limpio'] = df_impuestos_clean['Banco_Limpio'].fillna('UNKNOWN')

    # Rename Empresa_Raw to Empresa for consistency and select final columns
    df_impuestos_clean = df_impuestos_clean.rename(columns={'Empresa_Raw': 'Empresa'})
    df_impuestos_clean = df_impuestos_clean[['Empresa', 'Banco_Limpio', 'Fecha', 'Importe', 'Origen', 'Detalle', 'Numero_Cheque']]
    df_impuestos_clean = df_impuestos_clean.dropna(subset=['Importe', 'Empresa', 'Banco_Limpio', 'Fecha'])

    return df_impuestos_clean

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
uploaded_file_impuestos = st.file_uploader(
    "Sube el archivo 'Calendario de Vencimientos Impositivos.xlsx'",
    type=["xlsx"],
    key="calendario_impositivos"
)

if uploaded_file_proyeccion is not None and uploaded_file_cheques is not None and uploaded_file_saldos is not None and uploaded_file_impuestos is not None:
    with st.spinner('Procesando datos y generando reporte...'):
        archivo_proyeccion_io = io.BytesIO(uploaded_file_proyeccion.getvalue())
        archivo_cheques_io = io.BytesIO(uploaded_file_cheques.getvalue())
        archivo_saldos_io = io.BytesIO(uploaded_file_saldos.getvalue())
        archivo_impuestos_io = io.BytesIO(uploaded_file_impuestos.getvalue())

        df_proy = procesar_archivo(archivo_proyeccion_io, 0, 2, 9, 'Proyeccion', nombres_df, col_detalle=6)
        df_cheq = procesar_archivo(archivo_cheques_io, 3, 1, 14, 'Cheques', nombres_df, col_detalle=10, col_numero_cheque=2)
        df_impuestos = procesar_archivo_impuestos(archivo_impuestos_io)
        
        # Create df_total from the three processed dataframes
        df_total = pd.concat([df_proy, df_cheq, df_impuestos])

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

        # Calculate 'Disponible Futuro'
        reporte_final['Disponible Futuro'] = reporte_final['Saldo Banco'] - reporte_final['Vencido'] - reporte_final['Total Semana']

        # Reordenar columnas para colocar 'Saldo Banco' y 'Saldo FCI' antes de 'Vencido'
        # y luego 'A Cubrir Vencido' e 'Disponible Futuro' al final.
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

        # 4. Collect other existing columns that are not 'A Cubrir Vencido' or 'Disponible Futuro'
        #    and have not been added yet. This handles any unforeseen columns and ensures
        #    ACV and Disponible Futuro are indeed last.
        for col in cols:
            if col not in new_order_cols and col != 'A Cubrir Vencido' and col != 'Disponible Futuro':
                new_order_cols.append(col)

        # 5. Append 'A Cubrir Vencido' and 'Disponible Futuro' at the very end in the specified order
        if 'A Cubrir Vencido' in cols:
            new_order_cols.append('A Cubrir Vencido')
        if 'Disponible Futuro' in cols:
            new_order_cols.append('Disponible Futuro')

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
        # Subtotal LABEL format (e.g., "Total BYC")
        fmt_subtotal_label = workbook.add_format({
            **default_font_properties,
            'bold': True, 'bg_color': '#FCE4D6',
            'border': 1, 'align': 'left', 'valign': 'vcenter'
        })
        # Subtotal VALUE format
        fmt_subtotal_value = workbook.add_format({
            **default_font_properties,
            'bold': True, 'bg_color': '#FCE4D6', 'num_format': '$ #,##0',
            'border': 1, 'align': 'right', 'valign': 'vcenter'
        })
        fmt_currency = workbook.add_format({
            **default_font_properties,
            'num_format': '$ #,##0', 'border': 1, 'align': 'right'
        })
        fmt_text = workbook.add_format({
            **default_font_properties,
            'border': 1
        })

        # New formats for conditional formatting on 'A Cubrir Vencido' and 'Disponible Futuro'
        fmt_positive_acv = workbook.add_format({
            **default_font_properties,
            'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '$ #,##0', 'border': 1, 'align': 'right'
        })
        fmt_negative_acv = workbook.add_format({
            **default_font_properties,
            'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '$ #,##0', 'border': 1, 'align': 'right'
        })

        # New format for the grand total row *label* "TOTAL BANCOS"
        fmt_grand_total_label = workbook.add_format({
            **default_font_properties,
            'bold': True, 'bg_color': '#BFBFBF',
            'border': 1, 'align': 'left', 'valign': 'vcenter'
        })

        # New format for the grand total row *values*
        fmt_grand_total_value = workbook.add_format({
            **default_font_properties,
            'bold': True, 'bg_color': '#BFBFBF', 'num_format': '$ #,##0',
            'border': 1, 'align': 'right', 'valign': 'vcenter'
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

        # Find the index of 'Disponible Futuro' for conditional formatting
        df_col_idx = -1
        if 'Disponible Futuro' in columnas_datos:
            df_col_idx = columnas_datos.index('Disponible Futuro') + 1 # +1 because of the bank column at index 0

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
                    if current_col_excel_idx == acv_col_idx or current_col_excel_idx == df_col_idx:
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
            worksheet.write(fila_actual, 0, f"Total {empresa}", fmt_subtotal_label) # Apply specific label format

            sumas = datos_empresa.sum()
            for i, val in enumerate(sumas): # Loop through subtotal values
                current_col_excel_idx = i + 1
                # Apply conditional formatting to subtotal rows as well
                if current_col_excel_idx == acv_col_idx or current_col_excel_idx == df_col_idx:
                    if val > 0:
                        worksheet.write(fila_actual, current_col_excel_idx, val, fmt_positive_acv) # Already bold and right-aligned
                    elif val < 0:
                        worksheet.write(fila_actual, current_col_excel_idx, val, fmt_negative_acv) # Already bold and right-aligned
                    else:
                        worksheet.write(fila_actual, current_col_excel_idx, val, fmt_subtotal_value) # Default for 0, now bold and right-aligned
                else:
                    worksheet.write(fila_actual, i + 1, val, fmt_subtotal_value) # Use subtotal value format

            fila_actual += 1

        # --- CREAR FILA DE TOTAL BANCOS ---
        # Sum all numeric columns for the grand total row
        grand_totals_series = reporte_final.select_dtypes(include=['number']).sum()

        worksheet.write(fila_actual, 0, "TOTAL BANCOS", fmt_grand_total_label) # Use specific label format

        for i, col_name in enumerate(columnas_datos):
            val = grand_totals_series.get(col_name, "") # Get calculated total or empty string
            current_col_excel_idx = i + 1
            # Apply conditional formatting to grand total row as well
            if current_col_excel_idx == acv_col_idx or current_col_excel_idx == df_col_idx:
                if isinstance(val, (int, float)):
                    if val > 0:
                        worksheet.write(fila_actual, current_col_excel_idx, val, fmt_positive_acv) # Already bold and right-aligned
                    elif val < 0:
                        worksheet.write(fila_actual, current_col_excel_idx, val, fmt_negative_acv) # Already bold and right-aligned
                    else:
                        worksheet.write(fila_actual, current_col_excel_idx, val, fmt_grand_total_value) # Default for 0
                else:
                     worksheet.write(fila_actual, current_col_excel_idx, val, fmt_grand_total_value) # For non-numeric or empty string
            else:
                worksheet.write(fila_actual, i + 1, val, fmt_grand_total_value) # Use grand total value format

        fila_actual += 1

        # Ajustar ancho de columnas
        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, len(columnas_datos), 15)

        # --- Hoja 'Base' ---
        # Prepare the data for the 'Base' sheet
        # Reset index to make 'Empresa' and 'Banco_Limpio' regular columns
        # Create df_base_raw by concatenating the raw (processed) dataframes
        df_base_raw = pd.concat([df_proy, df_cheq, df_impuestos], ignore_index=True)

        # Define the desired columns for the 'Base' sheet
        base_columns = ['Empresa', 'Banco_Limpio', 'Fecha', 'Importe', 'Origen', 'Detalle', 'Numero_Cheque']

        # Ensure all base_columns exist in df_base_raw, filling missing ones with empty string or 0
        for col in base_columns:
            if col not in df_base_raw.columns:
                df_base_raw[col] = '' # Or 0 for numeric columns if preferred

        df_base = df_base_raw[base_columns].copy()

        # Write the DataFrame to the 'Base' sheet
        df_base.to_excel(writer, sheet_name='Base', index=False)

        writer.close()
        output_excel_data.seek(0)

        st.download_button(
            label="Descargar Reporte de Cashflow Formateado",
            data=output_excel_data,
            file_name="Resumen_Cashflow_Formateado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Para la descarga del PDF
        output_pdf_data = io.BytesIO()

        class PDF(FPDF):
            def header(self):
                self.set_font('Arial', 'B', 12)
                self.cell(0, 10, 'Resumen Cashflow', 0, 1, 'C')
                self.set_font('Arial', '', 10)
                self.cell(0, 10, f"Fecha Actual: {fecha_hoy.strftime('%d/%m/%Y')}", 0, 1, 'L')
                self.ln(5)

            def footer(self):
                self.set_y(-15)
                self.set_font('Arial', 'I', 8)
                self.cell(0, 10, 'Page %s' % self.page_no(), 0, 0, 'C')

        pdf = PDF(orientation='L') # Landscape orientation
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font('Arial', '', 8)

        # Prepare data for PDF table
        reporte_final_for_pdf = reporte_final.reset_index()
        reporte_final_for_pdf['Banco'] = reporte_final_for_pdf['Empresa'] + ' - ' + reporte_final_for_pdf['Banco_Limpio']
        reporte_final_for_pdf = reporte_final_for_pdf.drop(columns=['Empresa', 'Banco_Limpio'])

        # Reorder columns for PDF display (Banco first, then original order from reporte_final)
        col_names_pdf_ordered = ['Banco'] + reporte_final.columns.tolist()
        reporte_final_for_pdf = reporte_final_for_pdf[col_names_pdf_ordered]

        # Column headers for PDF (keep \n characters for multi_cell)
        processed_col_names = col_names_pdf_ordered # Use original names including \n

        # Determine max height for the header row
        max_header_height = 0
        line_height_base = 5 # Use the same height as in the multi_cell call for consistency

        # Allocate fixed width for 'Banco' column and distribute remaining width for others
        page_width = pdf.w - 2 * pdf.l_margin
        fixed_banco_width = 45
        num_data_cols = len(processed_col_names) - 1
        col_widths = [fixed_banco_width] + [(page_width - fixed_banco_width) / num_data_cols] * num_data_cols

        # Capture initial X and Y for height calculation loop
        initial_x_calc = pdf.get_x()
        initial_y_calc = pdf.get_y()

        current_x_pos_calc = initial_x_calc

        for i, header_text in enumerate(processed_col_names):
            # Temporarily set position for dry_run
            pdf.set_xy(current_x_pos_calc, initial_y_calc)
            # Use dry_run to get the actual number of lines multi_cell will generate
            # Setting a generous height to ensure it calculates lines correctly even if text wraps a lot
            lines_count = pdf.multi_cell(col_widths[i], line_height_base, header_text, 0, 'C', 0, 1, dry_run=True, output='S') # output='S' to return string
            # Calculate the height needed for this specific cell
            height_for_this_cell = lines_count.count('\n') + 1 * line_height_base if lines_count else line_height_base # Count newlines for height
            max_header_height = max(max_header_height, height_for_this_cell)
            current_x_pos_calc += col_widths[i] # Advance X for next header calculation

        # Ensure a minimum height if no text causes wrapping (e.g., all single line)
        if max_header_height == 0:
            max_header_height = line_height_base # Default to single line height

        # Restore original Y position and X position after height calculation
        pdf.set_xy(initial_x_calc, initial_y_calc)

        # Write header row
        pdf.set_fill_color(237, 125, 49) # Orange header color
        pdf.set_text_color(255, 255, 255) # White text
        pdf.set_font('Arial', 'B', 8)

        # Store starting X and Y for the actual header drawing
        current_x_draw = pdf.get_x()
        current_y_draw = pdf.get_y()

        for i, header in enumerate(processed_col_names):
            # Set position explicitly for each cell to ensure alignment
            pdf.set_xy(current_x_draw, current_y_draw)
            pdf.multi_cell(col_widths[i], max_header_height / (lines_count.count('\n') + 1), header, 1, 'C', 1, 0) # Adjusted height per line for multi_cell
            current_x_draw += col_widths[i] # Advance X for the next cell in the row

        pdf.ln(max_header_height) # Move to the next line after the entire header row

        # Write data rows and subtotals
        pdf.set_font('Arial', '', 8)

        # Get indices for conditional formatting columns in the `reporte_final` (original) DataFrame
        acv_col_idx = -1
        if 'A Cubrir Vencido' in reporte_final.columns:
            acv_col_idx = reporte_final.columns.get_loc('A Cubrir Vencido')

        df_col_idx = -1
        if 'Disponible Futuro' in reporte_final.columns:
            df_col_idx = reporte_final.columns.get_loc('Disponible Futuro')

        empresas_unicas = reporte_final.index.get_level_values(0).unique()

        for empresa in empresas_unicas:
            datos_empresa = reporte_final.loc[empresa]

            if isinstance(datos_empresa, pd.Series):
                banco_limpio_idx = datos_empresa.name[1]
                datos_empresa = pd.DataFrame(datos_empresa).T
                datos_empresa.index = [banco_limpio_idx]

            for banco, row in datos_empresa.iterrows():
                # Write bank name (first column in PDF)
                pdf.set_font('Arial', '', 8) # Regular font for data
                pdf.cell(col_widths[0], 6, str(banco), 1, 0, 'L')

                # Write numeric data
                for i, col_name_orig in enumerate(reporte_final.columns):
                    val = row[col_name_orig]

                    # Determine text to display: blank if 0, otherwise formatted value
                    display_text = '' if val == 0 else f"${val:,.0f}"

                    fill_cell = 0 # No fill by default
                    text_color = (0,0,0) # Black by default
                    fill_color = (255,255,255) # White by default

                    if col_name_orig == 'A Cubrir Vencido' or col_name_orig == 'Disponible Futuro':
                        if val > 0:
                            fill_color = (198, 239, 206) # Light Green
                            text_color = (0, 97, 0)     # Dark Green
                            fill_cell = 1
                        elif val < 0:
                            fill_color = (255, 199, 206) # Light Red
                            text_color = (156, 0, 6)    # Dark Red
                            fill_cell = 1

                    pdf.set_text_color(*text_color)
                    pdf.set_fill_color(*fill_color)
                    pdf.cell(col_widths[i+1], 6, display_text, 1, 0, 'R', fill_cell)

                    pdf.set_text_color(0,0,0) # Reset colors for next cell
                    pdf.set_fill_color(255,255,255)
                pdf.ln()

            # Subtotal row
            pdf.set_font('Arial', 'B', 8) # Bold for subtotal
            pdf.set_fill_color(252, 228, 214) # Light orange background
            pdf.cell(col_widths[0], 6, f"Total {empresa}", 1, 0, 'L', 1) # Label cell

            sumas = datos_empresa.sum()
            for i, col_name_orig in enumerate(reporte_final.columns):
                val = sumas[col_name_orig]

                # Determine text to display: blank if 0, otherwise formatted value
                display_text = '' if val == 0 else f"${val:,.0f}"

                fill_cell = 1 # Always fill subtotal cells
                text_color = (0,0,0) # Black by default
                fill_color = (252, 228, 214) # Light orange by default

                if col_name_orig == 'A Cubrir Vencido' or col_name_orig == 'Disponible Futuro':
                    if val > 0:
                        fill_color = (198, 239, 206) # Light Green
                        text_color = (0, 97, 0)     # Dark Green
                    elif val < 0:
                        fill_color = (255, 199, 206) # Light Red
                        text_color = (156, 0, 6)    # Dark Red

                pdf.set_text_color(*text_color)
                pdf.set_fill_color(*fill_color)
                pdf.cell(col_widths[i+1], 6, display_text, 1, 0, 'R', fill_cell)

                pdf.set_text_color(0,0,0) # Reset colors
                pdf.set_fill_color(252, 228, 214)
            pdf.ln()
            pdf.ln(2) # Small break between companies

        # Grand Total row
        pdf.set_font('Arial', 'B', 8) # Bold for grand total
        pdf.set_fill_color(191, 191, 191) # Grey background
        pdf.cell(col_widths[0], 6, "TOTAL BANCOS", 1, 0, 'L', 1) # Label cell

        # Sum all numeric columns for the grand total row
        grand_totals_series = reporte_final.select_dtypes(include=['number']).sum()

        for i, col_name_orig in enumerate(reporte_final.columns):
            val = grand_totals_series.get(col_name_orig, "") # Get calculated total or empty string

            # Determine text to display: blank if 0, otherwise formatted value
            if isinstance(val, (int, float)):
                display_text = '' if val == 0 else f"${val:,.0f}"
            else:
                display_text = str(val)

            fill_cell = 1 # Always fill grand total cells
            text_color = (0,0,0) # Black by default
            fill_color = (191, 191, 191) # Grey by default

            if col_name_orig == 'A Cubrir Vencido' or col_name_orig == 'Disponible Futuro':
                if isinstance(val, (int, float)):
                    if val > 0:
                        fill_color = (198, 239, 206) # Light Green
                        text_color = (0, 97, 0)     # Dark Green
                    elif val < 0:
                        fill_color = (255, 199, 206) # Light Red
                        text_color = (156, 0, 6)    # Dark Red

            pdf.set_text_color(*text_color)
            pdf.set_fill_color(*fill_color)

            if isinstance(val, (int, float)):
                pdf.cell(col_widths[i+1], 6, display_text, 1, 0, 'R', fill_cell)
            else:
                pdf.cell(col_widths[i+1], 6, str(val), 1, 0, 'R', fill_cell)

            pdf.set_text_color(0,0,0) # Reset colors
            pdf.set_fill_color(191, 191, 191)
        pdf.ln()

        pdf.output(output_pdf_data)
        output_pdf_data.seek(0)

        st.download_button(
            label="Descargar Reporte de Cashflow Formateado (PDF)",
            data=output_pdf_data,
            file_name="Resumen_Cashflow_Formateado.pdf",
            mime="application/pdf"
        )
        st.success("¡Listo! Archivo generado y disponible para descarga.")

else:
    st.info("Por favor, sube los archivos para generar el reporte de cashflow.")
