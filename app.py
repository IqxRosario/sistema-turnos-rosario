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

# --- 2. FUNCIONES DE APOYO ---
def obtener_festivos_colombia(ano, mes):
    festivos = {1: [1, 6], 3: [23], 4: [2, 3, 16, 17], 5: [1, 18], 6: [8, 15, 29], 7: [20], 8: [7, 17], 10: [12], 11: [2, 16], 12: [8, 25]}
    return festivos.get(mes, [])

def procesar_historial_seguro(file):
    if file is None: return {}, {}
    try:
        # Intentamos leer saltando las 9 filas de encabezado oficial
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

# --- 3. MOTOR DE GENERACIÓN ---
def generar_cuadro_maestro(mes, ano, h_cont, h_est, sugerencias):
    dias_mes = calendar.monthrange(ano, mes)[1]
    festivos = obtener_festivos_colombia(ano, mes)
    df = pd.DataFrame(index=PERSONAL_OFICIAL, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    c1_acu = {p: h_cont.get(p, {}).get('alta', 0) for p in PERSONAL_OFICIAL}
    total_acu = {p: h_cont.get(p, {}).get('total', 0) for p in PERSONAL_OFICIAL}
    ult_n = {p: -5 for p in PERSONAL_OFICIAL}
    consec = {p: h_est.get(p, {}).get('consec', 0) for p in PERSONAL_OFICIAL}

    for d in range(1, dias_mes + 1):
        dia_s, fecha = str(d), datetime(ano, mes, d)
        wd, es_f = fecha.weekday(), (d in festivos or wd == 6)
        
        if d == 21: # Reset nómina
            for p in PERSONAL_OFICIAL: c1_acu[p], total_acu[p] = 0, 0

        # Seguridad y Empalme
        for p in PERSONAL_OFICIAL:
            if d == 1 and h_est.get(p, {}).get('noche'):
                df.at[p, '1'], ult_n[p], consec[p] = 'P', 0, 0
            elif d == 1 and h_est.get(p, {}).get('consec', 0) >= 3:
                df.at[p, '1'], consec[p] = 'D', 0
            if d > 1 and 'N' in str(df.at[p, str(d-1)]):
                df.at[p, dia_s], consec[p] = 'P', 0

        # Reglas Fijas
        if wd == 2: df.at["ANGIE BERNAL", dia_s] = "C6"
        if wd == 1: 
            df.at["JHON RIOS", dia_s] = "N1"; total_acu["JHON RIOS"] += 1; c1_acu["JHON RIOS"] += 1; ult_n["JHON RIOS"] = d
        if wd in [0, 1, 5]: df.at["IVETTE VALENCIA", dia_s] = "L"
        if wd in [3, 4]: df.at["GINELAP", dia_s] = "L"
        if wd == 0: df.at["GERLIS DOMINGUEZ", dia_s] = "L"; df.at["ZARIANA REYES", dia_s] = "L"
        if wd == 3: df.at["MARCELA CASTRO", dia_s] = "L"; df.at["JUAN CAMILO PEREZ", dia_s] = "L"

        # Sugerencias
        for p in PERSONAL_OFICIAL:
            req = sugerencias.get(p, {}).get(dia_s)
            if req and df.at[p, dia_s] == "":
                df.at[p, dia_s] = req
                if any(t in req for t in ['C', 'N']) and 'P' not in req: 
                    total_acu[p] += 1
                    if any(t
