import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import io

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Pro", page_icon="🏥", layout="wide")

PERSONAL_OFICIAL = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES OSPINO", 
    "KELLY JOHANA JURADO", "JOAN SEBASTIAN AGUDELO"
]

TURNOS_ALTA = ['C1', 'C2', 'C3', 'C4', 'N1']

# --- 2. FUNCIONES MEJORADAS ---
def procesar_historial_seguro(file):
    if file is None: return {}, {}
    try:
        # Probamos leer el archivo (ajustar skiprows si es necesario)
        df = pd.read_excel(file, skiprows=9) if file.name.endswith('.xlsx') else pd.read_csv(file, skiprows=9)
        df.columns = df.columns.str.strip().str.upper()
        
        h_cont, h_est = {}, {}
        for _, row in df.iterrows():
            nom = str(row.get('NOMBRE', '')).strip().upper()
            if "GINELAP" in nom: nom = "GINELAP"
            if nom not in PERSONAL_OFICIAL: continue
            
            # Contar días del 21 al 31
            d_cols = [c for c in df.columns if str(c).isdigit() and int(c) >= 21]
            turnos_viejos = sum(1 for d in d_cols if any(t in str(row[d]).upper() for t in ['C', 'N']) and 'P' not in str(row[d]).upper())
            c1_viejos = sum(1 for d in d_cols if 'C1' in str(row[d]).upper())
            
            h_cont[nom] = {'total': turnos_viejos, 'c1': c1_viejos}
            
            # Estado de salida
            ultimos = [str(row[c]).upper() for c in [cols for cols in df.columns if str(cols).isdigit()][-3:]]
            h_est[nom] = {'noche': 'N' in ultimos[-1], 'consec': sum(1 for x in ultimos if any(t in x for t in ['C', 'N']) and 'P' not in x)}
        
        return h_cont, h_est
    except Exception as e:
        st.error(f"Error leyendo el historial: {e}")
        return {}, {}

# --- 3. MOTOR CON VÁLVULA DE ESCAPE ---
def generar_cuadro_maestro(mes, ano, h_cont, h_est, sugerencias):
    dias_mes = calendar.monthrange(ano, mes)[1]
    festivos = {1: [1, 6], 3: [23], 4: [2, 3, 16, 17], 5: [1, 18], 6: [8, 15, 29], 7: [20], 8: [7, 17], 10: [12], 11: [2, 16], 12: [8, 25]}.get(mes, [])
    df = pd.DataFrame(index=PERSONAL_OFICIAL, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    # Contadores
    c1_acu = {p: h_cont.get(p, {}).get('c1', 0) for p in PERSONAL_OFICIAL}
    total_acu = {p: h_cont.get(p, {}).get('total', 0) for p in PERSONAL_OFICIAL}
    ult_n = {p: -5 for p in PERSONAL_OFICIAL}
    consec = {p: h_est.get(p, {}).get('consec', 0) for p in PERSONAL_OFICIAL}

    for d in range(1, dias_mes + 1):
        dia_s = str(d)
        fecha = datetime(ano, mes, d)
        wd, es_f = fecha.weekday(), (d in festivos or fecha.weekday() == 6)
        if d == 21: # Reinicio de ciclo
            for p in PERSONAL_OFICIAL: c1_acu[p], total_acu[p] = 0, 0

        # Bloqueos de Seguridad
        for p in PERSONAL_OFICIAL:
            if d == 1 and h_est.get(p, {}).get('noche'): df.at[p, '1'], ult_n[p], consec[p] = 'P', 0, 0
            if d > 1 and 'N' in str(df.at[p, str(d-1)]): df.at[p, dia_s], consec[p] = 'P', 0

        # Turnos del día
        t_hoy = ['N1', 'N2', 'C1', 'C2', 'C3', 'C4']
        if not es_f: t_hoy += (['C5', 'C6'] if wd < 5 else ['C5'])

        # Reparto (Lógica de Válvula)
        for t in t_hoy:
            if (df[dia_s] == t).any() and t not in ['C1', 'C2']: continue
            
            # 1. Candidatos ideales (cumplen TODAS las reglas)
            candidatos = [p for p in PERSONAL_OFICIAL if df.at[p, dia_s] == "" and consec[p] < 3]
            
            if 'N' in t: candidatos = [p for p in candidatos if (d - ult_n[p]) > 2]
            if t == 'C1' and d > 1: candidatos = [p for p in candidatos if str(df.at[p, str(d-1)]) != 'C1']
            
            # VÁLVULA DE ESCAPE: Si no hay ideales, relajamos reglas de espaciado
            if not candidatos:
                candidatos = [p for p in PERSONAL_OFICIAL if df.at[p, dia_s] == "" and consec[p] < 3]

            if candidatos:
                # Prioridad al que menos lleva en nómina
                candidatos.sort(key=lambda x: (c1_acu[x] if t == 'C1' else total_acu[x]))
                el = candidatos[0]
                df.at[el, dia_s] = t
                total_acu[el] += 1
                consec[el] += 1
                if t == 'C1': c1_acu[el] += 1
                if 'N' in t: ult_n[el] = d
            
    # Totales para auditoría
    df['TOTAL NÓMINA (21-20)'] = df.apply(lambda r: h_cont.get(r.name, {}).get('total', 0) + sum(1 for d in range(1, 21) if any(t in str(r[str(d)]) for t in ['C', 'N']) and 'P' not in str(r[str(d)])), axis=1)
    return df.replace("", "D")

# --- 4. INTERFAZ ---
st.title("🏥 Optimizador de Turnos - Modo Rescate")
with st.sidebar:
    archivo = st.file_uploader("Subir Historial", type=['xlsx', 'csv'])
    mes_n = st.selectbox("Mes", range(1, 13), index=2)

if st.button("🚀 GENERAR CUADRO", type="primary", use_container_width=True):
    h_cont, h_est = procesar_historial_seguro(archivo)
    
    # Diagnóstico para Camilo
    if h_cont:
        st.success(f"✅ Historial cargado. Ejemplo: {list(h_cont.keys())[0]} trae {h_cont[list(h_cont.keys())[0]]['total']} turnos.")
    else:
        st.warning("⚠️ El historial no se leyó correctamente. Revisa el formato del Excel.")

    resultado = generar_cuadro_maestro(mes_n, 2026, h_cont, h_est, {})
    st.dataframe(resultado, use_container_width=True)
