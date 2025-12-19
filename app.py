
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

def procesar_archivo(file_object_or_path, col_banco, col_fecha, col_importe, tipo_origen, nombres_map_df):
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

        # Column headers for PDF (replace \n with space for FPDF)
        processed_col_names = [col.replace('\n', ' ') for col in col_names_pdf_ordered]

        # Determine column widths dynamically
        # Max width available: pdf.w - 2 * pdf.l_margin (landscape A4 is 297mm)
        page_width = pdf.w - 2 * pdf.l_margin
        # Allocate fixed width for 'Banco' column and distribute remaining width for others
        fixed_banco_width = 45
        num_data_cols = len(processed_col_names) - 1
        data_col_width = (page_width - fixed_banco_width) / num_data_cols
        col_widths = [fixed_banco_width] + [data_col_width] * num_data_cols

        # Write header row
        pdf.set_fill_color(237, 125, 49) # Orange header color
        pdf.set_text_color(255, 255, 255) # White text
        pdf.set_font('Arial', 'B', 8)
        for i, header in enumerate(processed_col_names):
            pdf.multi_cell(col_widths[i], 5, header, 1, 'C', 1, 0)
        pdf.ln()
        pdf.set_fill_color(255, 255, 255) # Reset fill color for data rows
        pdf.set_text_color(0, 0, 0) # Reset text color

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
                pdf.set_font('Arial', '', 8)
                pdf.cell(col_widths[0], 6, str(banco), 1, 0, 'L')

                # Write numeric data
                for i, col_name_orig in enumerate(reporte_final.columns):
                    val = row[col_name_orig]

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
                    pdf.cell(col_widths[i+1], 6, f"${val:,.0f}", 1, 0, 'R', fill_cell)

                    pdf.set_text_color(0,0,0) # Reset colors for next cell
                    pdf.set_fill_color(255,255,255)
                pdf.ln()

            # Subtotal row
            pdf.set_font('Arial', 'B', 8)
            pdf.set_fill_color(252, 228, 214) # Light orange background
            pdf.cell(col_widths[0], 6, f"Total {empresa}", 1, 0, 'L', 1)

            sumas = datos_empresa.sum()
            for i, col_name_orig in enumerate(reporte_final.columns):
                val = sumas[col_name_orig]

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
                pdf.cell(col_widths[i+1], 6, f"${val:,.0f}", 1, 0, 'R', fill_cell)

                pdf.set_text_color(0,0,0) # Reset colors
                pdf.set_fill_color(252, 228, 214)
            pdf.ln()
            pdf.ln(2)

        # Grand Total row
        pdf.set_font('Arial', 'B', 8)
        pdf.set_fill_color(191, 191, 191) # Grey background
        pdf.cell(col_widths[0], 6, "TOTAL BANCOS", 1, 0, 'L', 1)

        # Sum all numeric columns for the grand total row
        grand_totals_series = reporte_final.select_dtypes(include=['number']).sum()

        for i, col_name_orig in enumerate(reporte_final.columns):
            val = grand_totals_series.get(col_name_orig, "") # Get calculated total or empty string

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
                pdf.cell(col_widths[i+1], 6, f"${val:,.0f}", 1, 0, 'R', fill_cell)
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
