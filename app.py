import streamlit as st
import pandas as pd
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

def es_festivo(dia, mes):
    festivos_2026 = {1: [1, 12], 3: [23], 4: [2, 3], 5: [1, 18], 6: [8, 15, 29], 7: [20], 8: [7, 17], 10: [12], 11: [2, 16], 12: [8, 25]}
    return dia in festivos_2026.get(mes, [])

def aplicar_colores(v):
    if v in ['L', 'D']: return 'background-color: #d9ead3' 
    if v == 'P': return 'background-color: #f4cccc'      
    if 'N' in str(v): return 'background-color: #cfe2f3' 
    if 'C' in str(v): return 'background-color: #fff2cc' 
    return ''

# --- LECTORES DE DATOS EXTERNOS ---
def procesar_historial_empalme(file):
    historial = {p: ["", "", ""] for p in INTEGRANTES} 
    if not file: return historial
    try:
        df_hist = pd.read_excel(file, skiprows=9) if file.name.endswith('.xlsx') else pd.read_csv(file, skiprows=9)
        df_hist.columns = df_hist.columns.str.strip().str.upper()
        cols_dias = [c for c in df_hist.columns if str(c).isdigit()]
        if len(cols_dias) >= 3:
            ultimas_3 = cols_dias[-3:]
            for _, r in df_hist.iterrows():
                nom = str(r.get('NOMBRE','')).strip().upper()
                if "GINELAP" in nom: nom = "GINELAP"
                if nom in INTEGRANTES:
                    historial[nom] = [str(r[c]).upper() for c in ultimas_3]
    except Exception as e:
        st.sidebar.error(f"Error leyendo el historial: {e}")
    return historial

def procesar_sugerencias(link):
    sugerencias = {p: {} for p in INTEGRANTES}
    if not link: return sugerencias
    try:
        csv_link = link.split('/edit')[0] + '/export?format=csv' if "/edit" in link else link
        df_sug = pd.read_csv(csv_link)
        for _, r in df_sug.iterrows():
            nom = str(r.get('NOMBRE','')).strip().upper()
            if "GINELAP" in nom: nom = "GINELAP"
            fecha = ''.join(filter(str.isdigit, str(r.get('FECHA',''))))
            sol = str(r.get('SOLICITUD','')).strip().upper()
            if nom in INTEGRANTES and fecha and sol != 'NAN':
                sugerencias[nom][fecha] = sol
    except Exception as e:
        st.sidebar.warning("No se pudo leer el link de sugerencias.")
    return sugerencias

# --- MOTOR DE GENERACIÓN ---
def generar_cuadro_equitativo(mes, ano, historial_previo, sugerencias_dict):
    dias_mes = calendar.monthrange(ano, mes)[1]
    df = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    turnos_totales = {p: 0 for p in INTEGRANTES}
    noches_totales = {p: 0 for p in INTEGRANTES}
    finde_totales = {p: 0 for p in INTEGRANTES}

    def turno_en_dia(persona, dia_req):
        if dia_req > 0: return str(df.at[persona, str(dia_req)])
        else:
            idx = 2 + dia_req
            hist = historial_previo.get(persona, ["", "", ""])
            if 0 <= idx < 3: return str(hist[idx])
            return ""

    for d in range(1, dias_mes + 1):
        ds = str(d)
        wd = datetime(ano, mes, d).weekday()
        es_finde_o_festivo = wd >= 5 or es_festivo(d, mes)

        def necesita_descanso(persona):
            return all(any(t in turno_en_dia(persona, past_d) for t in ['C', 'N']) for past_d in range(d-3, d))

        # 1. DEFINIR NECESIDADES DEL DÍA
        turnos_noche = ['N1', 'N2']
        if es_festivo(d, mes) or wd == 6: turnos_dia = ['C1', 'C2']
        elif wd == 5: turnos_dia = ['C1', 'C2', 'C3', 'C4', 'C5']
        else: turnos_dia = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6']

        # 2. APLICAR REGLAS FIJAS, POSTURNOS Y SUGERENCIAS
        for p in INTEGRANTES:
            if 'N' in turno_en_dia(p, d-1): 
                df.at[p, ds] = 'P'
                continue
            
            if wd == 0 and p in ["GERLIS DOMINGUEZ", "ZARIANA REYES"]: df.at[p, ds] = "L"
            if wd in [1, 2, 5] and p == "IVETTE VALENCIA": df.at[p, ds] = "L"
            if wd in [3, 4] and p == "GINELAP": df.at[p, ds] = "L"
            if wd == 3 and p in ["MARCELA CASTRO", "JUAN CAMILO PEREZ"]: df.at[p, ds] = "L"
            if wd == 1 and p == "JUAN CAMILO PEREZ": df.at[p, ds] = "L"

            req = sugerencias_dict.get(p, {}).get(ds)
            if req and df.at[p, ds] == "":
                df.at[p, ds] = req
                if req in turnos_noche: turnos_noche.remove(req)
                if req in turnos_dia: turnos_dia.remove(req)
                if any(t in req for t in ['C', 'N']) and 'P' not in req:
                    turnos_totales[p] += 1
                    if es_finde_o_festivo: finde_totales[p] += 1
                    if 'N' in req: noches_totales[p] += 1

        # 3. ASIGNACIONES FIJAS DE TURNO
        if wd == 1 and df.at["JHON RIOS", ds] == "":
            df.at["JHON RIOS", ds] = "N1"
            if "N1" in turnos_noche: turnos_noche.remove("N1")
            turnos_totales["JHON RIOS"] += 1
            noches_totales["JHON RIOS"] += 1

        if wd == 2 and df.at["ANGIE BERNAL", ds] == "" and "C6" in turnos_dia:
            df.at["ANGIE BERNAL", ds] = "C6"
            turnos_dia.remove("C6")
            turnos_totales["ANGIE BERNAL"] += 1

        # 4. REPARTIR NOCHES RESTANTES
        for t in turnos_noche:
            disp_noches = [p for p in INTEGRANTES if df.at[p, ds] == "" and not necesita_descanso(p)]
            disp_noches = [p for p in disp_noches if 'N' not in turno_en_dia(p, d-2)] # Anti N-P-N

            if d < dias_mes:
                m_wd = (wd + 1) % 7
                disp_noches = [p for p in disp_noches if not (
                    (m_wd == 0 and p in ["GERLIS DOMINGUEZ", "ZARIANA REYES"]) or
                    (m_wd in [1, 2, 5] and p == "IVETTE VALENCIA") or
                    (m_wd in [3, 4] and p == "GINELAP") or
                    (m_wd == 3 and p in ["MARCELA CASTRO", "JUAN CAMILO PEREZ"]) or
                    (m_wd == 1 and p == "JUAN CAMILO PEREZ")
                )]

            if not disp_noches: disp_noches = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if disp_noches:
                random.shuffle(disp_noches) # Desempate aleatorio
                disp_noches.sort(key=lambda x: (noches_totales[x], turnos_totales[x]))
                elegido = disp_noches[0]
                df.at[elegido, ds] = t
                turnos_totales[elegido] += 1
                noches_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 5. REPARTIR CORRIDOS (CON ROTACIÓN FORZADA)
        for t in turnos_dia:
            disp_dia = [p for p in INTEGRANTES if df.at[p, ds] == "" and not necesita_descanso(p)]
            
            # ---> REGLA NUEVA: Evitar que repita exactamente el mismo turno de ayer
            disp_variada = [p for p in disp_dia if turno_en_dia(p, d-1) != t]
            if disp_variada:
                disp_dia = disp_variada # Usar la lista variada si hay opciones

            if not disp_dia: disp_dia = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            
            if disp_dia:
                random.shuffle(disp_dia) # Desempate aleatorio para evitar bloques
                if es_finde_o_festivo: disp_dia.sort(key=lambda x: (finde_totales[x], turnos_totales[x]))
                else: disp_dia.sort(key=lambda x: turnos_totales[x])
                
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
    df['TOTAL NOCHES'] = df.apply(lambda r: noches_totales[r.name], axis=1)
    df['FINES DE SEMANA'] = df.apply(lambda r: finde_totales[r.name], axis=1)
    return df

# --- INTERFAZ ---
st.title("🏥 Cuadro de Instrumentación (Rotación Dinámica)")

with st.sidebar:
    st.header("1. Cargar Historial")
    archivo_previo = st.file_uploader("Sube el Excel del mes anterior para el empalme", type=['xlsx', 'csv'])
    
    st.header("2. Peticiones de Turno")
    link_sheet = st.text_input("Link de Sugerencias (Google Sheets):", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?gid=0#gid=0")
    
    st.header("3. Configurar Mes")
    mes_sel = st.selectbox("Mes a Generar (2026)", range(1, 13), index=datetime.now().month-1)

if st.button("🚀 GENERAR CUADRO DEL MES", type="primary"):
    with st.spinner("Creando rotación sin repeticiones..."):
        historial_leido = procesar_historial_empalme(archivo_previo)
        sugerencias_leidas = procesar_sugerencias(link_sheet)
        
        resultado = generar_cuadro_equitativo(mes_sel, 2026, historial_leido, sugerencias_leidas)
        
        cols_dias = [str(d) for d in range(1, calendar.monthrange(2026, mes_sel)[1] + 1)]
        st.dataframe(resultado.style.applymap(aplicar_colores, subset=cols_dias), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resultado.to_excel(writer, index=True)
        st.download_button("📥 Descargar Excel", output.getvalue(), f"Turnos_Rotados_{mes_sel}_2026.xlsx")
