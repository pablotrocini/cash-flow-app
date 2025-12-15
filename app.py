
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

def procesar_archivo(file_object_or_path, col_banco, col_fecha, col_importe, tipo_origen, nombres_map_df):
    df = pd.read_excel(file_object_or_path)
    df_clean = pd.DataFrame({
        'Banco_Raw': df.iloc[:, col_banco].astype(str).str.strip(),
        'Fecha': pd.to_datetime(df.iloc[:, col_fecha], errors='coerce'),
        'Importe': pd.to_numeric(df.iloc[:, col_importe], errors='coerce'),
        'Origen': tipo_origen
    })
    df_clean = df_clean.dropna(subset=['Importe', 'Banco_Raw'])

    if tipo_origen == 'Proyeccion':
        merge_on_col = 'Proyeccion Pagos'
        nombres_map_df_cleaned = nombres_map_df[['Proyeccion Pagos', 'EMPRESA']].copy()
        nombres_map_df_cleaned['Proyeccion Pagos'] = nombres_map_df_cleaned['Proyeccion Pagos'].astype(str).str.strip()
    elif tipo_origen == 'Cheques':
        merge_on_col = 'Cheques'
        nombres_map_df_cleaned = nombres_map_df[['Cheques', 'EMPRESA']].copy()
        nombres_map_df_cleaned['Cheques'] = nombres_map_df_cleaned['Cheques'].astype(str).str.strip()
    else:
        df_clean['Banco_Limpio'] = df_clean['Banco_Raw']
        df_clean['Empresa'] = 'UNKNOWN'
        return df_clean

    df_clean = pd.merge(
        df_clean,
        nombres_map_df_cleaned,
        left_on='Banco_Raw',
        right_on=merge_on_col,
        how='left'
    )

    df_clean['Banco_Limpio'] = df_clean[merge_on_col].fillna(df_clean['Banco_Raw'])
    df_clean['Empresa'] = df_clean['EMPRESA'].fillna('UNKNOWN')

    df_clean = df_clean.drop(columns=[merge_on_col, 'EMPRESA'])

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

        # Merge with nombres_df to get Empresa for saldos
        # Need to determine which column in nombres_df corresponds to 'Banco_Raw_Saldos'
        # For now, let's assume it should match either 'Cheques' or 'Proyeccion Pagos'
        # Since df_total's Banco_Limpio already uses the mapped name, we should map Saldos banks too
        # This mapping is complex as 'Banco_Raw_Saldos' doesn't directly map to 'Empresa' here.
        # A simpler approach for now is to try matching based on the Banco_Limpio from df_total, assuming they are consistent.

        # First, ensure df_total has unique (Empresa, Banco_Limpio) combinations to get the mapping
        unique_bancos_empresas_from_total = df_total[['Empresa', 'Banco_Limpio']].drop_duplicates().set_index('Banco_Limpio')

        # Map 'Banco_Raw_Saldos' to 'Banco_Limpio' using the names_df
        df_saldos_clean = pd.merge(
            df_saldos_clean,
            nombres_df[['Cheques', 'EMPRESA']].rename(columns={'Cheques': 'Banco_Limpio', 'EMPRESA': 'Empresa_Mapped'}),
            left_on='Banco_Raw_Saldos',
            right_on='Banco_Limpio',
            how='left'
        )
        df_saldos_clean = pd.merge(
            df_saldos_clean,
            nombres_df[['Proyeccion Pagos', 'EMPRESA']].rename(columns={'Proyeccion Pagos': 'Banco_Limpio', 'EMPRESA': 'Empresa_Mapped'}),
            left_on='Banco_Raw_Saldos',
            right_on='Banco_Limpio',
            how='left',
            suffixes=('_cheques', '_proyeccion')
        )

        # Consolidate Empresa from either merge
        df_saldos_clean['Empresa'] = df_saldos_clean['Empresa_Mapped_cheques'].fillna(df_saldos_clean['Empresa_Mapped_proyeccion']).fillna('UNKNOWN')
        # Consolidate Banco_Limpio from either merge
        df_saldos_clean['Banco_Limpio'] = df_saldos_clean['Banco_Limpio_cheques'].fillna(df_saldos_clean['Banco_Limpio_proyeccion']).fillna(df_saldos_clean['Banco_Raw_Saldos'])

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

        # Unir todo
        reporte_final = pd.concat([df_vencido, df_semana_pivot, df_emitidos], axis=1).fillna(0)

        # Merge 'Saldo FCI' and 'Saldo Banco' with reporte_final
        reporte_final = pd.merge(
            reporte_final,
            df_saldos_clean, # Usar el DataFrame de saldos limpios con las nuevas columnas
            left_index=True,
            right_index=True,
            how='left'
        )
        # Rellenar NaN en las columnas de saldos con 0
        reporte_final['Saldo FCI'] = reporte_final['Saldo FCI'].fillna(0)
        reporte_final['Saldo Banco'] = reporte_final['Saldo Banco'].fillna(0)

        # Calcular 'A Cubrir Vencido' como (Saldo Banco - Vencido)
        reporte_final['Vencido_temp'] = reporte_final['Vencido'].fillna(0)
        reporte_final['A Cubrir Vencido'] = reporte_final['Saldo Banco'] - reporte_final['Vencido_temp']
        reporte_final = reporte_final.drop(columns=['Vencido_temp'])

        # Reordenar columnas para colocar 'Saldo Banco' y 'Saldo FCI' antes de 'Vencido' y 'A Cubrir Vencido' después de 'Vencido'
        cols = reporte_final.columns.tolist()

        # Define lists for new column order
        new_order_cols = []
        if 'Saldo Banco' in cols: new_order_cols.append('Saldo Banco')
        if 'Saldo FCI' in cols: new_order_cols.append('Saldo FCI')
        if 'Vencido' in cols: new_order_cols.append('Vencido')
        if 'A Cubrir Vencido' in cols: new_order_cols.append('A Cubrir Vencido')

        # Add other columns in their original relative order, skipping those already added
        for col in cols:
            if col not in new_order_cols and col not in ['Saldo Banco', 'Saldo FCI', 'Vencido', 'A Cubrir Vencido']:
                new_order_cols.append(col)

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
        fmt_header = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#ED7D31',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'text_wrap': True
        })
        fmt_subtotal = workbook.add_format({
            'bold': True, 'bg_color': '#FCE4D6', 'num_format': '$ #,##0',
            'border': 1
        })
        fmt_currency = workbook.add_format({
            'num_format': '$ #,##0', 'border': 1
        })
        fmt_text = workbook.add_format({'border': 1})

        # --- ESCRIBIR ENCABEZADOS ---
        worksheet.write('A1', 'Resumen Cashflow', workbook.add_format({'bold': True, 'font_size': 14}))
        worksheet.write('A2', f"Fecha Actual: {fecha_hoy.strftime('%d/%m/%Y')}")

        fila_actual = 3
        col_bancos = 0
        worksheet.write(fila_actual, col_bancos, "Etiquetas de fila", fmt_header)

        columnas_datos = reporte_final.columns.tolist()
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
                    worksheet.write(fila_actual, i + 1, val, fmt_currency)

                fila_actual += 1

            # --- CREAR FILA DE SUBTOTAL ---
            worksheet.write(fila_actual, 0, f"Total {empresa}", fmt_subtotal)

            sumas = datos_empresa.sum()
            for i, val in enumerate(sumas):
                worksheet.write(fila_actual, i + 1, val, fmt_subtotal)

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
