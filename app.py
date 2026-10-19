import streamlit as st
import pandas as pd
import io
import calendar
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from motor_pulp import resolver_turnos

st.set_page_config(page_title="Generador de Turnos Hospitalarios", layout="wide", page_icon="🏥")

st.title("🏥 Sistema de Optimización de Turnos Médicos")
st.write("Carga el archivo de sugerencias para generar el cuadro de turnos optimizado.")

def crear_excel(df_res, mes=11, ano=2026):
    salida = io.BytesIO()
    
    dias_espanol = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']
    num_dias = calendar.monthrange(ano, mes)[1]
    primer_dia_semana, _ = calendar.monthrange(ano, mes)
    
    renombrados = {}
    es_fin_de_semana = {}
    
    for d in range(1, num_dias + 1):
        dia_idx = (primer_dia_semana + d - 1) % 7
        nombre_dia = dias_espanol[dia_idx]
        col_name = f"{d} {nombre_dia}"
        renombrados[d] = col_name
        es_fin_de_semana[col_name] = (dia_idx in [5, 6])

    df_export = df_res.copy()
    df_export = df_export.rename(columns=renombrados)
    
    colores_turnos = {
        "C": "F4D35E",  # Amarillo (Corrido)
        "N": "5DADE2",  # Celeste (Noche)
        "P": "F5B7B1",  # Rosado (Post-turno)
        "D": "D5D8DC",  # Gris (Descanso)
        "V": "BB8FCE",  # Morado (Vacaciones)
        "L": "82E0AA",  # Verde (Libre)
    }

    fill_header_norm = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    fill_header_weekend = PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True, name="Calibri", size=10)
    font_dark = Font(color="202020", name="Calibri", size=11)
    align_center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='D0D3D4'),
        right=Side(style='thin', color='D0D3D4'),
        top=Side(style='thin', color='D0D3D4'),
        bottom=Side(style='thin', color='D0D3D4')
    )

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df_export.to_excel(writer, sheet_name="Turnos", index=False)
        diagnostico = df_res.attrs.get("diagnostico", {"estado": df_res.attrs.get("status", "OK")})
        pd.DataFrame([diagnostico]).to_excel(writer, sheet_name="Diagnostico", index=False)

        wb = writer.book
        ws = wb["Turnos"]
        
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.alignment = align_center
            cell.font = font_header
            col_val = str(cell.value)
            
            if col_val in es_fin_de_semana and es_fin_de_semana[col_val]:
                cell.fill = fill_header_weekend
            else:
                cell.fill = fill_header_norm

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = align_center
                cell.font = font_dark
                cell.border = thin_border
                val = str(cell.value).strip().upper()
                
                if val in colores_turnos:
                    cell.fill = PatternFill(start_color=colores_turnos[val], end_color=colores_turnos[val], fill_type="solid")

    salida.seek(0)
    return salida.getvalue(), f"Turnos_{mes:02d}_{ano}.xlsx"

uploaded_file = st.file_uploader("Subir archivo Excel de Sugerencias", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_sugerencias = pd.read_excel(uploaded_file)
        st.success("Plantilla cargada exitosamente.")
        st.dataframe(df_sugerencias, use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            mes = st.number_input("Mes de programación", min_value=1, max_value=12, value=11)
        with col2:
            ano = st.number_input("Año", min_value=2024, max_value=2030, value=2026)
            
        if st.button("🚀 Generar Cuadro de Turnos", type="primary"):
            with st.spinner("Optimizando asignación de turnos..."):
                df_res = resolver_turnos(df_sugerencias, mes=mes, ano=ano)
                status = df_res.attrs.get('status', 'Desconocido')
                
                if status in ['Optimal', 'Feasible']:
                    st.success(f"¡Cuadro generado exitosamente! Estado: {status}")
                    st.dataframe(df_res, use_container_width=True)
                    
                    bytes_data, nombre_archivo = crear_excel(df_res, mes=mes, ano=ano)
                    st.download_button(
                        label="📥 Descargar Cuadro de Turnos en Excel",
                        data=bytes_data,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error(f"No se pudo encontrar una solución factible. Estado del optimizador: {status}")
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
