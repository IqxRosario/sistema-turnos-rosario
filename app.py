import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import random
import io

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Instrumentación", layout="wide")

INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES OSPINO", 
    "KELLY JOHANA JURADO", "JOAN SEBASTIAN AGUDELO"
]

# Festivos Colombia 2026
def es_festivo(dia, mes):
    festivos_2026 = {1: [1, 12], 3: [23], 4: [2, 3], 5: [1, 18], 6: [8, 15, 29], 7: [20], 8: [7, 17], 10: [12], 11: [2, 16], 12: [8, 25]}
    return dia in festivos_2026.get(mes, [])

def aplicar_colores(v):
    if v in ['L', 'D']: return 'background-color: #d9ead3' # Verde
    if v == 'P': return 'background-color: #f4cccc'      # Rojo
    if 'N' in str(v): return 'background-color: #cfe2f3' # Azul
    if 'C' in str(v): return 'background-color: #fff2cc' # Amarillo
    return ''

# --- MOTOR DE GENERACIÓN ---
def generar_cuadro_equitativo(mes, ano):
    dias_mes = calendar.monthrange(ano, mes)[1]
    df = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    # Contadores para equidad
    turnos_totales = {p: 0 for p in INTEGRANTES}
    noches_totales = {p: 0 for p in INTEGRANTES}
    finde_totales = {p: 0 for p in INTEGRANTES}

    for d in range(1, dias_mes + 1):
        ds = str(d)
        fecha = datetime(ano, mes, d)
        wd = fecha.weekday()
        es_finde_o_festivo = wd >= 5 or es_festivo(d, mes)

        # 1. APLICAR REGLAS FIJAS DE DESCANSO Y POSTURNOS
        for p in INTEGRANTES:
            # Posturno automático si ayer hizo Noche
            if d > 1 and 'N' in str(df.at[p, str(d-1)]): 
                df.at[p, ds] = 'P'
                continue
            
            # Reglas Fijas del Grupo
            if wd == 0 and p in ["GERLIS DOMINGUEZ", "ZARIANA REYES"]: df.at[p, ds] = "L"
            if wd == 1 and p == "IVETTE VALENCIA": df.at[p, ds] = "L"
            if wd == 2 and p == "IVETTE VALENCIA": df.at[p, ds] = "L"
            if wd == 5 and p == "IVETTE VALENCIA": df.at[p, ds] = "L"
            if wd in [3, 4] and p == "GINELAP": df.at[p, ds] = "L"
            if wd == 3 and p in ["MARCELA CASTRO", "JUAN CAMILO PEREZ"]: df.at[p, ds] = "L"
            if wd == 1 and p == "JUAN CAMILO PEREZ": df.at[p, ds] = "L"

        # 2. DEFINIR NECESIDADES DEL DÍA
        turnos_noche = ['N1', 'N2']
        if es_festivo(d, mes) or wd == 6: # Domingo o Festivo
            turnos_dia = ['C1', 'C2']
        elif wd == 5: # Sábado
            turnos_dia = ['C1', 'C2', 'C3', 'C4', 'C5']
        else: # Lunes a Viernes
            turnos_dia = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6']

        # 3. ASIGNACIONES FIJAS DE TURNO
        if wd == 1 and df.at["JHON RIOS", ds] == "": # Martes
            df.at["JHON RIOS", ds] = "N1"
            turnos_noche.remove("N1")
            turnos_totales["JHON RIOS"] += 1
            noches_totales["JHON RIOS"] += 1

        if wd == 2 and df.at["ANGIE BERNAL", ds] == "": # Miércoles
            if "C6" in turnos_dia:
                df.at["ANGIE BERNAL", ds] = "C6"
                turnos_dia.remove("C6")
                turnos_totales["ANGIE BERNAL"] += 1

        # 4. REPARTIR NOCHES RESTANTES (Priorizando a quien tiene menos noches)
        for t in turnos_noche:
            disp_noches = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            
            # Evitar asignar Noche si mañana tienen una "L" fija que impide Posturno
            if d < dias_mes:
                manana_wd = (wd + 1) % 7
                disp_noches = [p for p in disp_noches if not (
                    (manana_wd == 0 and p in ["GERLIS DOMINGUEZ", "ZARIANA REYES"]) or
                    (manana_wd == 1 and p == "IVETTE VALENCIA") or
                    (manana_wd == 2 and p == "IVETTE VALENCIA") or
                    (manana_wd == 5 and p == "IVETTE VALENCIA") or
                    (manana_wd in [3, 4] and p == "GINELAP") or
                    (manana_wd == 3 and p in ["MARCELA CASTRO", "JUAN CAMILO PEREZ"]) or
                    (manana_wd == 1 and p == "JUAN CAMILO PEREZ")
                )]

            if disp_noches:
                # Mezclar para romper patrones, luego ordenar por quien tiene menos noches
                random.shuffle(disp_noches)
                disp_noches.sort(key=lambda x: (noches_totales[x], turnos_totales[x]))
                elegido = disp_noches[0]
                df.at[elegido, ds] = t
                turnos_totales[elegido] += 1
                noches_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 5. REPARTIR CORRIDOS (Priorizando a quien tiene menos turnos totales)
        for t in turnos_dia:
            disp_dia = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            
            # Si es fin de semana, intentar no pasar de 4 días de fin de semana al mes (2 findes)
            if es_finde_o_festivo:
                disp_dia.sort(key=lambda x: (finde_totales[x], turnos_totales[x]))
            else:
                disp_dia.sort(key=lambda x: turnos_totales[x])

            if disp_dia:
                elegido = disp_dia[0]
                df.at[elegido, ds] = t
                turnos_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 6. RELLENAR CON DESCANSO (D)
        for p in INTEGRANTES:
            if df.at[p, ds] == "":
                df.at[p, ds] = "D"

    # AGREGAR COLUMNAS DE AUDITORÍA
    df['TOTAL TURNOS'] = df.apply(lambda r: sum(1 for c in df.columns if any(t in str(r[c]) for t in ['C', 'N'])), axis=1)
    df['TOTAL NOCHES'] = df.apply(lambda r: sum(1 for c in df.columns if 'N' in str(r[c])), axis=1)
    df['FINES DE SEMANA'] = df.apply(lambda r: finde_totales[r.name], axis=1)
    
    return df

# --- INTERFAZ ---
st.title("🏥 Cuadro de Instrumentación (Equidad Simple)")
st.markdown("Generación automática basándose en **conteo total de turnos, noches y fines de semana.**")

mes_sel = st.sidebar.selectbox("Mes a Generar (2026)", range(1, 13), index=datetime.now().month-1)

if st.button("🚀 GENERAR CUADRO DEL MES", type="primary"):
    with st.spinner("Balanceando cargas..."):
        resultado = generar_cuadro_equitativo(mes_sel, 2026)
        
        cols_dias = [str(d) for d in range(1, calendar.monthrange(2026, mes_sel)[1] + 1)]
        st.dataframe(resultado.style.applymap(aplicar_colores, subset=cols_dias), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resultado.to_excel(writer, index=True)
        st.download_button("📥 Descargar Excel", output.getvalue(), f"Turnos_{mes_sel}_2026.xlsx")
