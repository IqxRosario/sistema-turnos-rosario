import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import io

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Gestor Automático de Turnos", page_icon="🏥", layout="wide")

PERSONAL_OFICIAL = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES OSPINO", 
    "KELLY JOHANA JURADO", "JOAN SEBASTIAN AGUDELO"
]

def obtener_festivos_colombia(ano, mes):
    festivos = {
        1: [1, 6], 3: [23], 4: [2, 3, 16, 17], 5: [1, 18],
        6: [8, 15, 29], 7: [20], 8: [7, 17], 10: [12], 11: [2, 16], 12: [8, 25]
    }
    return festivos.get(mes, [])

def procesar_historial(file):
    if file is None: return {}, {}
    try:
        df = pd.read_csv(file, skiprows=9) if file.name.endswith('.csv') else pd.read_excel(file, skiprows=9)
        df.columns = df.columns.str.strip().str.upper()
        res_nom, ult_est = {}, {}
        for _, row in df.iterrows():
            nombre = str(row['NOMBRE']).strip().upper()
            if "GINELAP" in nombre: nombre = "GINELAP"
            if nombre not in PERSONAL_OFICIAL: continue
            d_cols = [c for c in df.columns if c.isdigit() and int(c) >= 21]
            res_nom[nombre] = sum(1 for d in d_cols if 'C1' in str(row[d]).upper())
            d_f = [c for c in df.columns if c.isdigit()][-3:]
            act = [str(row[d]).upper() for d in d_f]
            ult_est[nombre] = {'termino_noche': 'N' in act[-1], 'seguidos': sum(1 for x in act if any(t in x for t in ['C', 'N']))}
        return res_nom, ult_est
    except: return {}, {}

def procesar_sugerencias_link(link):
    sugerencias = {p: {} for p in PERSONAL_OFICIAL}
    if not link: return sugerencias
    try:
        if "/edit" in link:
            csv_link = link.split('/edit')[0] + '/export?format=csv'
        else:
            csv_link = link
        df_sug = pd.read_csv(csv_link)
        for _, row in df_sug.iterrows():
            nombre = str(row.get('NOMBRE', '')).strip().upper()
            if "GINELAP" in nombre: nombre = "GINELAP"
            fecha = str(row.get('FECHA', '')).strip()
            fecha = ''.join(filter(str.isdigit, fecha)) 
            solicitud = str(row.get('SOLICITUD', '')).strip().upper()
            if nombre in PERSONAL_OFICIAL and fecha and solicitud != 'NAN':
                sugerencias[nombre][fecha] = solicitud
    except Exception as e:
        st.sidebar.warning("⚠️ No se pudo leer el link de sugerencias.")
    return sugerencias

def aplicar_colores(val):
    if val in ['L', 'D']: return 'background-color: #b6d7a8'
    if val == 'P': return 'background-color: #f4cccc'
    if 'N' in str(val): return 'background-color: #cfe2f3'
    if 'C1' in str(val): return 'background-color: #fff2cc'
    return ''

def generar_cuadro_maestro(mes, ano, h_nomina, h_estado, sugerencias):
    dias_mes = calendar.monthrange(ano, mes)[1]
    festivos = obtener_festivos_colombia(ano, mes)
    df = pd.DataFrame(index=PERSONAL_OFICIAL, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    c1_quincena = {p: 0 for p in PERSONAL_OFICIAL}
    ultima_noche = {p: -5 for p in PERSONAL_OFICIAL}

    for d in range(1, dias_mes + 1):
        dia_str, fecha = str(d), datetime(ano, mes, d)
        wd, es_f = fecha.weekday(), (d in festivos or wd == 6)

        if d == 1:
            for p in PERSONAL_OFICIAL:
                if h_estado.get(p, {}).get('termino_noche'): df.at[p, '1'], ultima_noche[p] = 'P', 0
                elif h_estado.get(p, {}).get('seguidos', 0) >= 3: df.at[p, '1'] = 'D'

        if wd == 2: df.at["ANGIE BERNAL", dia_str] = "C6"
        if wd == 1: 
            df.at["JHON RIOS", dia_str] = "N1"; ultima_noche["JHON RIOS"] = d
            if d < dias_mes: df.at["JHON RIOS", str(d+1)] = "P"
        if wd in [0, 1, 5]: df.at["IVETTE VALENCIA", dia_str] = "L"
        if wd in [3, 4]: df.at["GINELAP", dia_str] = "L"
        if wd == 0: df.at["GERLIS DOMINGUEZ", dia_str] = "L"; df.at["ZARIANA REYES", dia_str] = "L"
        if wd == 3: df.at["MARCELA CASTRO", dia_str] = "L"; df.at["JUAN CAMILO PEREZ", dia_str] = "L"
        if wd == 1 and df.at["JUAN CAMILO PEREZ", dia_str] == "": df.at["JUAN CAMILO PEREZ", dia_str] = "C6"

        for p in PERSONAL_OFICIAL:
            req = sugerencias.get(p, {}).get(dia_str)
            if req and df.at[p, dia_str] != 'P': 
                df.at[p, dia_str] = req
                if 'N' in req:
                    ultima_noche[p] = d
                    if d < dias_mes: df.at[p, str(d+1)] = 'P'

        turnos_dia = ['N1', 'N2', 'C1', 'C2']
        if not es_f and wd < 5: turnos_dia.extend(['C3', 'C4', 'C5', 'C6'])
        elif wd == 5: turnos_dia.extend(['C3', 'C4', 'C5'])

        for t in turnos_dia:
            if (df[dia_str] == t).any() and t not in ['C1', 'C2']: continue 
            disp = [p for p in PERSONAL_OFICIAL if df.at[p, dia_str] == ""]
            if not disp: continue
            if 'N' in t:
                cand_n = [p for p in disp if (d - ultima_noche[p]) > 1]
                if cand_n:
                    cand_n.sort(key=lambda x: ultima_noche[x])
                    el = cand_n[0]
                    df.at[el, dia_str], ultima_noche[el] = t, d
                    if d < dias_mes: df.at[el, str(d+1)] = 'P'
                continue
            if t == 'C1':
                disp.sort(key=lambda x: (c1_quincena[x], h_nomina.get(x, 0)))
                el = disp[0]
                df.at[el, dia_str] = t
                c1_quincena[el] += 1
            elif t == 'C5' and wd == 5:
                el = "ERNESTO MUSKUS" if "ERNESTO MUSKUS" in disp else disp[0]
                df.at[el, dia_str] = t
            else:
                el = disp[np.random.randint(len(disp))]
                df.at[el, dia_str] = t
        if d == 15: c1_quincena = {p: 0 for p in PERSONAL_OFICIAL}

    df['TOTAL C1'] = df.apply(lambda row: sum(1 for x in row if 'C1' in str(x)), axis=1)
    df['TOTAL NOCHES'] = df.apply(lambda row: sum(1 for x in row if 'N' in str(x)), axis=1)
    df['TOTAL TURNOS'] = df.apply(lambda row: sum(1 for x in row if any(t in str(x) for t in ['C', 'N'])), axis=1)
    return df.replace("", "D")

# --- INTERFAZ ---
st.title("🏥 Gestor Automático de Turnos - Instrumentación")
with st.sidebar:
    st.header("1. Historial Base")
    archivo = st.file_uploader("Subir Cuadro Mes Anterior", type=['csv', 'xlsx'])
    
    st.header("2. Peticiones del Equipo")
    link_oficial = "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?gid=0#gid=0"
    st.link_button("📝 Abrir Excel de Sugerencias", link_oficial)
    link_sheet = st.text_input("Link interno del sistema:", link_oficial)
    
    st.header("3. Mes a Proyectar")
    mes_n = st.selectbox("Mes", range(1, 13), index=3)

if st.button("🚀 GENERAR CUADRO INTELIGENTE", type="primary", use_container_width=True):
    with st.spinner('Conectando con Google Sheets y analizando historial...'):
        h_nom, h_est = procesar_historial(archivo)
        sug_dict = procesar_sugerencias_link(link_sheet)
        resultado = generar_cuadro_maestro(mes_n, 2026, h_nom, h_est, sug_dict)
        
        cols_dias = [c for c in resultado.columns if c.isdigit()]
        st.dataframe(resultado.style.applymap(aplicar_colores, subset=cols_dias), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resultado.to_excel(writer, sheet_name='Turnos_Oficial')
        st.download_button(label="📥 Descargar para WhatsApp", data=output.getvalue(), file_name=f"Cuadro_Final_{mes_n}.xlsx")
