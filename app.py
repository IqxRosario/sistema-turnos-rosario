import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import io

# --- 1. CONFIGURACIÓN Y PERSONAL ---
st.set_page_config(page_title="Gestor Automático de Turnos", page_icon="🏥", layout="wide")

PERSONAL_OFICIAL = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES OSPINO", 
    "KELLY JOHANA JURADO", "JOAN SEBASTIAN AGUDELO"
]

TURNOS_ALTA_EFECTIVIDAD = ['C1', 'C2', 'C3', 'C4', 'N1']
TURNOS_BAJA_EFECTIVIDAD = ['C5', 'C6', 'N2']

# --- 2. FUNCIONES DE APOYO ---
def obtener_festivos_colombia(ano, mes):
    festivos = {1: [1, 6], 3: [23], 4: [2, 3, 16, 17], 5: [1, 18], 6: [8, 15, 29], 7: [20], 8: [7, 17], 10: [12], 11: [2, 16], 12: [8, 25]}
    return festivos.get(mes, [])

def procesar_historial(file):
    if file is None: return {}, {}
    try:
        df = pd.read_csv(file, skiprows=9) if file.name.endswith('.csv') else pd.read_excel(file, skiprows=9)
        df.columns = df.columns.str.strip().str.upper()
        h_cont, h_est = {}, {}
        for _, row in df.iterrows():
            nombre = str(row['NOMBRE']).strip().upper()
            if "GINELAP" in nombre: nombre = "GINELAP"
            if nombre not in PERSONAL_OFICIAL: continue
            
            d_cols = [c for c in df.columns if str(c).isdigit() and int(c) >= 21]
            h_cont[nombre] = {
                'alta': sum(1 for d in d_cols if any(t in str(row[d]).upper() for t in TURNOS_ALTA_EFECTIVIDAD) and 'P' not in str(row[d]).upper()),
                'total': sum(1 for d in d_cols if any(t in str(row[d]).upper() for t in ['C', 'N']) and 'P' not in str(row[d]).upper())
            }
            d_f = [c for c in df.columns if str(c).isdigit()][-3:]
            act = [str(row[d]).upper() for d in d_f]
            h_est[nombre] = {'termino_noche': 'N' in act[-1], 'seguidos': sum(1 for x in act if any(t in x for t in ['C', 'N']) and 'P' not in x)}
        return h_cont, h_est
    except: return {}, {}

def procesar_sugerencias_link(link):
    sugerencias = {p: {} for p in PERSONAL_OFICIAL}
    if not link: return sugerencias
    try:
        csv_link = link.split('/edit')[0] + '/export?format=csv' if "/edit" in link else link
        df_sug = pd.read_csv(csv_link)
        for _, row in df_sug.iterrows():
            nombre = str(row.get('NOMBRE', '')).strip().upper()
            if "GINELAP" in nombre: nombre = "GINELAP"
            fecha = ''.join(filter(str.isdigit, str(row.get('FECHA', ''))))
            sol = str(row.get('SOLICITUD', '')).strip().upper()
            if nombre in PERSONAL_OFICIAL and fecha and sol != 'NAN':
                sugerencias[nombre][fecha] = sol
    except: pass
    return sugerencias

def aplicar_colores(val):
    if val in ['L', 'D']: return 'background-color: #b6d7a8'
    if val == 'P': return 'background-color: #f4cccc'
    if 'N' in str(val): return 'background-color: #cfe2f3'
    if any(t in str(val) for t in ['C1', 'C2', 'C3', 'C4']): return 'background-color: #fff2cc'
    return ''

# --- 3. MOTOR DE GENERACIÓN ROBUSTO ---
def generar_cuadro_maestro(mes, ano, h_cont, h_est, sugerencias):
    dias_mes = calendar.monthrange(ano, mes)[1]
    festivos = obtener_festivos_colombia(ano, mes)
    df = pd.DataFrame(index=PERSONAL_OFICIAL, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    c1_acu = {p: h_cont.get(p, {}).get('alta', 0) for p in PERSONAL_OFICIAL}
    total_acu = {p: h_cont.get(p, {}).get('total', 0) for p in PERSONAL_OFICIAL}
    ult_n = {p: -5 for p in PERSONAL_OFICIAL}
    consec = {p: h_est.get(p, {}).get('seguidos', 0) for p in PERSONAL_OFICIAL}

    for d in range(1, dias_mes + 1):
        dia_s, fecha = str(d), datetime(ano, mes, d)
        wd, es_f = fecha.weekday(), (d in festivos or fecha.weekday() == 6)
        
        if d == 21:
            for p in PERSONAL_OFICIAL: c1_acu[p], total_acu[p] = 0, 0

        for p in PERSONAL_OFICIAL:
            if d == 1 and h_est.get(p, {}).get('termino_noche'):
                df.at[p, '1'], ult_n[p], consec[p] = 'P', 0, 0
            if d > 1 and 'N' in str(df.at[p, str(d-1)]):
                df.at[p, dia_s], consec[p] = 'P', 0

        # Reglas Fijas
        if wd == 2: df.at["ANGIE BERNAL", dia_s] = "C6"
        if wd == 1: 
            df.at["JHON RIOS", dia_s] = "N1"; total_acu["JHON RIOS"] += 1; ult_n["JHON RIOS"] = d
        if wd in [0, 1, 5]: df.at["IVETTE VALENCIA", dia_s] = "L"
        if wd in [3, 4]: df.at["GINELAP", dia_s] = "L"
        if wd == 0: df.at["GERLIS DOMINGUEZ", dia_s] = "L"; df.at["ZARIANA REYES", dia_s] = "L"
        if wd == 3: df.at["MARCELA CASTRO", dia_s] = "L"; df.at["JUAN CAMILO PEREZ", dia_s] = "L"

        for p in PERSONAL_OFICIAL:
            req = sugerencias.get(p, {}).get(dia_s)
            if req and df.at[p, dia_s] == "":
                df.at[p, dia_s] = req
                if any(t in req for t in ['C', 'N']) and 'P' not in req: 
                    total_acu[p] += 1
                    if any(t in req for t in TURNOS_ALTA_EFECTIVIDAD): c1_acu[p] += 1
                if 'N' in req: ult_n[p] = d

        t_hoy = ['N1', 'N2', 'C1', 'C2', 'C3', 'C4']
        if not es_f: t_hoy += (['C5', 'C6'] if wd < 5 else ['C5'])

        for t in t_hoy:
            if (df[dia_s] == t).any() and t not in ['C1', 'C2']: continue
            disp = [p for p in PERSONAL_OFICIAL if df.at[p, dia_s] == "" and consec[p] < 3]
            
            if 'N' in t:
                disp = [p for p in disp if (d - ult_n[p]) > 2]
                if disp:
                    disp.sort(key=lambda x: total_acu[x])
                    el = disp[0]
                    df.at[el, dia_s], total_acu[el], ult_n[el], consec[el] = t, total_acu[el]+1, d, consec[el]+1
            elif t == 'C1':
                if d > 1: disp = [p for p in disp if str(df.at[p, str(d-1)]) != 'C1']
                if disp:
                    disp.sort(key=lambda x: (c1_acu[x], total_acu[x]))
                    el = disp[0]
                    df.at[el, dia_s], c1_acu[el], total_acu[el], consec[el] = t, c1_acu[el]+1, total_acu[el]+1, consec[el]+1
            else:
                if disp:
                    disp.sort(key=lambda x: total_acu[x])
                    el = disp[0]
                    df.at[el, dia_s], total_acu[el], consec[el] = t, total_acu[el]+1, consec[el]+1

    df['TOTAL MES'] = df.apply(lambda r: sum(1 for x in r if any(t in str(x) for t in ['C', 'N']) and 'P' not in str(x)), axis=1)
    df['NÓMINA (21-20)'] = df.apply(lambda r: h_cont.get(r.name, {}).get('total', 0) + sum(1 for d in range(1, 21) if any(t in str(r[str(d)]) for t in ['C', 'N']) and 'P' not in str(r[str(d)])), axis=1)
    return df.replace("", "D")

# --- 4. INTERFAZ COMPLETA (EL CUERPO DE LA APP) ---
st.title("🏥 Gestor de Turnos Inteligente - Versión Emprendedor")

with st.sidebar:
    st.header("1. Cargar Historial")
    archivo = st.file_uploader("Subir Cuadro Mes Anterior (Excel/CSV)", type=['xlsx', 'csv'])
    
    st.header("2. Sugerencias")
    link_sheet = st.text_input("Link de Google Sheets:", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?gid=0#gid=0")
    
    st.header("3. Mes a Generar")
    mes_n = st.selectbox("Seleccione Mes", range(1, 13), index=datetime.now().month - 1)

if st.button("🚀 GENERAR CUADRO Y CALCULAR NÓMINA", type="primary", use_container_width=True):
    with st.spinner('Aplicando matemática de cuotas y reglas de salud...'):
        h_cont, h_est = procesar_historial(archivo)
        sug_dict = procesar_sugerencias_link(link_sheet)
        resultado = generar_cuadro_maestro(mes_n, 2026, h_cont, h_est, sug_dict)
        
        # Mostrar en pantalla
        cols_dias = [c for c in resultado.columns if c.isdigit()]
        st.dataframe(resultado.style.applymap(aplicar_colores, subset=cols_dias), use_container_width=True)
        
        # Botón de Descarga
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resultado.to_excel(writer, sheet_name='Turnos_Equitativos')
        st.download_button(label="📥 Descargar Excel para Nómina", data=output.getvalue(), file_name=f"Cuadro_Final_Mes_{mes_n}.xlsx", use_container_width=True)

st.info("💡 Este sistema asegura que los turnos C1-C4 y Noches se repartan equitativamente entre el 21 del mes pasado y el 20 de este mes.")
