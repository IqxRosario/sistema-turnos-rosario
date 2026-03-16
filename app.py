import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import io

# --- 1. IDENTIDAD DEL GRUPO ---
st.set_page_config(page_title="Gestor de Instrumentación", layout="wide")

INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES OSPINO", 
    "KELLY JOHANA JURADO", "JOAN SEBASTIAN AGUDELO"
]

# Definición de valor de turnos para rotación justa
TURNOS_VALIOSOS = ['C1', 'C2', 'C3', 'C4', 'N1'] 

# --- 2. FUNCIONES DE LÓGICA ---
def procesar_historial(file):
    """Extrae el estado del 21 al 31 para empalmar y contar nómina"""
    if not file: return {}, {}
    try:
        df = pd.read_excel(file, skiprows=9) if file.name.endswith('.xlsx') else pd.read_csv(file, skiprows=9)
        df.columns = df.columns.str.strip().str.upper()
        h_nomina, h_salud = {}, {}
        for _, r in df.iterrows():
            nom = str(r.get('NOMBRE','')).strip().upper()
            if "GINELAP" in nom: nom = "GINELAP"
            if nom not in INTEGRANTES: continue
            
            # Conteo para nómina (Periodo 21 al 31 del mes pasado)
            cols_21_31 = [c for c in df.columns if str(c).isdigit() and int(c) >= 21]
            h_nomina[nom] = sum(1 for c in cols_21_31 if any(t in str(r[c]).upper() for t in TURNOS_VALIOSOS) and 'P' not in str(r[c]).upper())
            
            # Estado físico para el día 1
            ult_dias = [str(r[c]).upper() for c in [x for x in df.columns if str(x).isdigit()][-3:]]
            h_salud[nom] = {'noche_ayer': 'N' in ult_dias[-1], 'consecutivos': sum(1 for x in ult_dias if any(t in x for t in ['C','N']) and 'P' not in x)}
        return h_nomina, h_salud
    except: return {}, {}

def aplicar_colores(v):
    if v in ['L','D']: return 'background-color: #d9ead3' # Verde suave
    if v == 'P': return 'background-color: #f4cccc'      # Rojo suave
    if 'N' in str(v): return 'background-color: #cfe2f3' # Azul suave
    if any(t in str(v) for t in ['C1','C2','C3','C4']): return 'background-color: #fff2cc' # Amarillo
    return ''

# --- 3. EL MOTOR DE EQUIDAD Y ROTACIÓN ---
def generar_cuadro(mes, ano, h_nom, h_sal, sug_link):
    dias_mes = calendar.monthrange(ano, mes)[1]
    festivos = {1:[1,6], 3:[23], 4:[2,3,16,17], 5:[1,18], 6:[8,15,29], 7:[20], 8:[7,17], 10:[12], 11:[2,16], 12:[8,25]}.get(mes, [])
    
    df = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    # Contadores vivos
    conteo_21_20 = {p: h_nom.get(p, 0) for p in INTEGRANTES}
    u_noche, consec = {p: -5 for p in INTEGRANTES}, {p: h_sal.get(p,{}).get('consecutivos', 0) for p in INTEGRANTES}

    for d in range(1, dias_mes + 1):
        ds, fecha = str(d), datetime(ano, mes, d)
        wd = fecha.weekday()
        es_festivo = (d in festivos or wd == 6)
        
        if d == 21: # Reinicio de periodo de facturación
            conteo_21_20 = {p: 0 for p in INTEGRANTES}

        # A. REGLAS DE ORO (SALUD Y FIJOS)
        for p in INTEGRANTES:
            # Empalme Día 1
            if d == 1 and h_sal.get(p,{}).get('noche_ayer'): df.at[p, '1'], u_noche[p], consec[p] = 'P', 0, 0
            # Posturnos automáticos
            if d > 1 and 'N' in str(df.at[p, str(d-1)]): df.at[p, ds], consec[p] = 'P', 0
            
            # Restricciones Fijas de Disponibilidad
            if wd == 0 and p in ["GERLIS DOMINGUEZ", "ZARIANA REYES"]: df.at[p, ds] = "L"
            if wd == 1 and p == "IVETTE VALENCIA": df.at[p, ds] = "L"
            if wd == 2 and p == "ANGIE BERNAL": df.at[p, ds] = "C6"
            if wd == 3 and p in ["MARCELA CASTRO", "JUAN CAMILO PEREZ"]: df.at[p, ds] = "L"
            if wd in [3, 4] and p == "GINELAP": df.at[p, ds] = "L"
            if wd in [0, 5] and p == "IVETTE VALENCIA": df.at[p, ds] = "L"

        # B. REPARTO DE TURNOS VALIOSOS (Equidad y Rotación)
        turnos_dia = ['N1', 'C1', 'C2', 'C3', 'C4']
        for t in turnos_dia:
            if (df[ds] == t).any() and t not in ['C1', 'C2']: continue
            
            # Buscar instrumentadores disponibles
            disp = []
            for p in INTEGRANTES:
                if df.at[p, ds] == "" and consec[p] < 3:
                    if 'N' in t and (d - u_noche[p]) <= 2: continue # Anti N-P-N
                    if t == 'C1' and d > 1 and str(df.at[p, str(d-1)]) == 'C1': continue # Anti C1-C1
                    disp.append(p)
            
            if disp:
                # Transparencia: Se le da al que menos turnos valiosos lleva en el periodo 21-20
                disp.sort(key=lambda x: conteo_21_20[x])
                elegido = disp[0]
                df.at[elegido, ds] = t
                conteo_21_20[elegido] += 1
                consec[elegido] += 1
                if 'N' in t: u_noche[elegido] = d

        # C. RELLENO DE TURNOS RESTANTES (C5, C6, N2)
        t_relleno = (['N2', 'C5', 'C6'] if not es_festivo and wd < 5 else (['N2', 'C5'] if not es_festivo else ['N2']))
        for t in t_relleno:
            if (df[ds] == t).any(): continue
            disp = [p for p in INTEGRANTES if df.at[p, ds] == "" and consec[p] < 3]
            if disp:
                disp.sort(key=lambda x: conteo_21_20[x])
                el = disp[0]
                df.at[el, ds], consec[el] = t, consec[el] + 1
                if 'N' in t: u_noche[el] = d

    # Totales de auditoría
    df['TOTAL 21-20'] = df.apply(lambda r: h_nom.get(r.name, 0) + sum(1 for d_i in range(1, 21) if any(t in str(r[str(d_i)]) for t in TURNOS_VALIOSOS)), axis=1)
    return df.replace("", "D")

# --- 4. INTERFAZ PARA EL EQUIPO ---
st.title("🏥 Cuadro Automático de Instrumentación")
st.markdown("Generación de turnos basada en **Equidad (21-20)** y **Salud Ocupacional**.")

with st.sidebar:
    st.header("Entradas del Sistema")
    archivo = st.file_uploader("Historial Mes Pasado", type=['xlsx', 'csv'])
    mes_sel = st.selectbox("Mes a Generar", range(1, 13), index=datetime.now().month-1)
    st.info("El sistema respeta automáticamente las L de Ivette, Ginelap, Gerlis y Zariana.")

if st.button("🚀 GENERAR CUADRO DEL MES", type="primary", use_container_width=True):
    h_nom, h_sal = procesar_historial(archivo)
    resultado = generar_cuadro(mes_sel, 2026, h_nom, h_sal, "")
    
    # Mostrar Cuadro
    cols_dias = [c for c in resultado.columns if c.isdigit()]
    st.dataframe(resultado.style.applymap(aplicar_colores, subset=cols_dias), use_container_width=True)
    
    # Descargar
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        resultado.to_excel(writer, index=True)
    st.download_button("📥 Descargar Excel para el Grupo", output.getvalue(), f"Turnos_Instrumentacion_Mes_{mes_sel}.xlsx", use_container_width=True)
