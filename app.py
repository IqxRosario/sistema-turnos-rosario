import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import io

# --- 1. CONFIGURACIÓN Y PERSONAL ---
st.set_page_config(page_title="Gestor de Turnos Inteligente", page_icon="🏥", layout="wide")

PERSONAL_OFICIAL = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES OSPINO", 
    "KELLY JOHANA JURADO", "JOAN SEBASTIAN AGUDELO"
]

TURNOS_ALTA = ['C1', 'C2', 'C3', 'C4', 'N1']
TURNOS_BAJA = ['C5', 'C6', 'N2']

# --- 2. FUNCIONES DE APOYO Y PROCESAMIENTO ---
def obtener_festivos_colombia(ano, mes):
    festivos = {1: [1, 6], 3: [23], 4: [2, 3, 16, 17], 5: [1, 18], 6: [8, 15, 29], 7: [20], 8: [7, 17], 10: [12], 11: [2, 16], 12: [8, 25]}
    return festivos.get(mes, [])

def procesar_historial_seguro(file):
    if file is None: return {}, {}
    try:
        df = pd.read_excel(file, skiprows=9) if file.name.endswith('.xlsx') else pd.read_csv(file, skiprows=9)
        df.columns = df.columns.str.strip().str.upper()
        h_cont, h_est = {}, {}
        for _, row in df.iterrows():
            nom = str(row.get('NOMBRE', '')).strip().upper()
            if "GINELAP" in nom: nom = "GINELAP"
            if nom not in PERSONAL_OFICIAL: continue
            
            d_cols = [c for c in df.columns if str(c).isdigit() and int(c) >= 21]
            h_cont[nom] = {
                'alta': sum(1 for d in d_cols if any(t in str(row[d]).upper() for t in TURNOS_ALTA) and 'P' not in str(row[d]).upper()),
                'total': sum(1 for d in d_cols if any(t in str(row[d]).upper() for t in ['C', 'N']) and 'P' not in str(row[d]).upper())
            }
            d_f = [c for c in df.columns if str(c).isdigit()][-3:]
            act = [str(row[d]).upper() for d in d_f]
            h_est[nom] = {'noche': 'N' in act[-1], 'consec': sum(1 for x in act if any(t in x for t in ['C', 'N']) and 'P' not in x)}
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

# --- 3. MOTOR DE GENERACIÓN CON TODO ---
def generar_cuadro_maestro(mes, ano, h_cont, h_est, sugerencias):
    dias_mes = calendar.monthrange(ano, mes)[1]
    festivos = obtener_festivos_colombia(ano, mes)
    df = pd.DataFrame(index=PERSONAL_OFICIAL, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    # Cuotas de Nómina
    c1_acu = {p: h_cont.get(p, {}).get('alta', 0) for p in PERSONAL_OFICIAL}
    total_acu = {p: h_cont.get(p, {}).get('total', 0) for p in PERSONAL_OFICIAL}
    ult_n = {p: -5 for p in PERSONAL_OFICIAL}
    consec = {p: h_est.get(p, {}).get('consec', 0) for p in PERSONAL_OFICIAL}

    for d in range(1, dias_mes + 1):
        dia_s, fecha = str(d), datetime(ano, mes, d)
        wd, es_f = fecha.weekday(), (d in festivos or fecha.weekday() == 6)
        
        if d == 21: # Reset nómina
            for p in PERSONAL_OFICIAL: c1_acu[p], total_acu[p] = 0, 0

        # A. SEGURIDAD Y EMPALME (DÍA 1 Y POSTURNOS)
        for p in PERSONAL_OFICIAL:
            if d == 1 and h_est.get(p, {}).get('noche'):
                df.at[p, '1'], ult_n[p], consec[p] = 'P', 0, 0
            elif d == 1 and h_est.get(p, {}).get('consec', 0) >= 3:
                df.at[p, '1'], consec[p] = 'D', 0
            if d > 1 and 'N' in str(df.at[p, str(d-1)]):
                df.at[p, dia_s], consec[p] = 'P', 0

        # B. REGLAS FIJAS
        if wd == 2: df.at["ANGIE BERNAL", dia_s] = "C6"
        if wd == 1: 
            df.at["JHON RIOS", dia_s] = "N1"; total_acu["JHON RIOS"] += 1; ult_n["JHON RIOS"] = d
        if wd in [0, 1, 5]: df.at["IVETTE VALENCIA", dia_s] = "L"
        if wd in [3, 4]: df.at["GINELAP", dia_s] = "L"
        if wd == 0: df.at["GERLIS DOMINGUEZ", dia_s] = "L"; df.at["ZARIANA REYES", dia_s] = "L"
        if wd == 3: df.at["MARCELA CASTRO", dia_s] = "L"; df.at["JUAN CAMILO PEREZ", dia_s] = "L"

        # C. SUGERENCIAS
        for p in PERSONAL_OFICIAL:
            req = sugerencias.get(p, {}).get(dia_s)
            if req and df.at[p, dia_s] == "":
                df.at[p, dia_s] = req
                if any(t in req for t in ['C', 'N']) and 'P' not in req: 
                    total_acu[p] += 1
                    if any(t in req for t in TURNOS_ALTA): c1_acu[p] += 1
                if 'N' in req: ult_n[p] = d

        # D. REPARTO MATEMÁTICO (Cuotas + Válvula de Escape)
        t_hoy = ['N1', 'N2', 'C1', 'C2', 'C3', 'C4']
        if not es_f: t_hoy += (['C5', 'C6'] if wd < 5 else ['C5'])

        for t in t_hoy:
            if (df[dia_s] == t).any() and t not in ['C1', 'C2']: continue
            
            # Intento 1: Candidato IDEAL (Cumple todas las reglas)
            disp = [p for p in PERSONAL_OFICIAL if df.at[p, dia_s] == "" and consec[p] < 3]
            if 'N' in t: disp = [p for p in disp if (d - ult_n[p]) > 2] # Anti NP-N
            if t == 'C1' and d > 1: disp = [p for p in disp if str(df.at[p, str(d-1)]) != 'C1'] # Espaciado

            # VÁLVULA DE ESCAPE: Si no hay nadie ideal, relajamos espaciado pero mantenemos Salud (consec < 3 y P)
            if not disp:
                disp = [p for p in PERSONAL_OFICIAL if df.at[p, dia_s] == "" and consec[p] < 3]

            if disp:
                # Prioridad por alcancía vacía
                disp.sort(key=lambda x: (c1_acu[x] if t in TURNOS_ALTA else total_acu[x], total_acu[x]))
                el = disp[0]
                df.at[el, dia_s] = t
                total_acu[el] += 1
                consec[el] += 1
                if t in TURNOS_ALTA: c1_acu[el] += 1
                if 'N' in t: ult_n[el] = d

    # E. TOTALES
    df['TOTAL MES'] = df.apply(lambda r: sum(1 for x in r if any(t in str(x) for t in ['C', 'N']) and 'P' not in str(x)), axis=1)
    df['NÓMINA (21-20)'] = df.apply(lambda r: h_cont.get(r.name, {}).get('total', 0) + sum(1 for d in range(1, 21) if any(t in str(r[str(d)]) for t in ['C', 'N']) and 'P' not in str(r[str(d)])), axis=1)
    df['EFECTIVOS (C1-C4)'] = df.apply(lambda r: sum(1 for x in r if any(t in str(x) for t in ['C1', 'C2', 'C3', 'C4'])), axis=1)
    
    return df.replace("", "D")

# --- 4. INTERFAZ STREAMLIT ---
st.title("🏥 Optimizador de Turnos - Emprendimiento")

with st.sidebar:
    st.header("1. Datos de Entrada")
    archivo = st.file_uploader("Subir Historial Mes Pasado", type=['xlsx', 'csv'])
    link_sheet = st.text_input("Link de Sugerencias (Google Sheets):", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?gid=0#gid=0")
    
    st.header("2. Configuración")
    mes_n = st.selectbox("Mes a Proyectar", range(1,
