import streamlit as st
import pandas as pd
from datetime import datetime
import calendar
import random
import io
import holidays

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Instrumentación", layout="wide")

INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES OSPINO", 
    "KELLY JOHANA JURADO", "JOAN SEBASTIAN AGUDELO"
]

def es_festivo(dia, mes, ano):
    co_holidays = holidays.CO(years=ano)
    return datetime(ano, mes, dia).date() in co_holidays

def aplicar_colores(v):
    if v in ['L', 'D']: return 'background-color: #d9ead3; color: #000000;' 
    if v == 'P': return 'background-color: #f4cccc; color: #000000;'      
    if 'N' in str(v): return 'background-color: #cfe2f3; color: #000000;' 
    if 'C' in str(v): return 'background-color: #fff2cc; color: #000000;' 
    return ''

# --- LECTORES DE DATOS ---
def procesar_historial_empalme(file):
    historial = {p: ["", "", ""] for p in INTEGRANTES} 
    if not file: return historial
    try:
        df = pd.read_excel(file) if file.name.endswith('.xlsx') else pd.read_csv(file)
        if not any(str(c).isdigit() for c in df.columns):
            df = pd.read_excel(file, skiprows=9) if file.name.endswith('.xlsx') else pd.read_csv(file, skiprows=9)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        col_nombre = next((c for c in df.columns if 'NOMBRE' in c or 'UNNAMED: 0' in c), None)
        cols_dias = [c for c in df.columns if c.isdigit()]
        if len(cols_dias) >= 3:
            ultimas_3 = sorted(cols_dias, key=int)[-3:]
            for _, r in df.iterrows():
                nom = str(r[col_nombre]).strip().upper() if col_nombre else str(r.name).strip().upper()
                if "GINELAP" in nom: nom = "GINELAP"
                if nom in INTEGRANTES: historial[nom] = [str(r[c]).upper() for c in ultimas_3]
    except Exception as e: st.sidebar.error("Error leyendo el historial de empalme.")
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
            if nom in INTEGRANTES and fecha and sol != 'NAN': sugerencias[nom][fecha] = sol
    except: st.sidebar.warning("No se pudo leer el link de sugerencias.")
    return sugerencias

def procesar_configuracion(link):
    """Lee la hoja de Google Sheets con los Libres Fijos de cada persona"""
    libres_fijos = {p: [] for p in INTEGRANTES}
    if not link: return libres_fijos
    try:
        csv_link = link.split('/edit')[0] + '/export?format=csv'
        if "gid=" in link: csv_link += "&gid=" + link.split("gid=")[1].split("&")[0]
        
        df_conf = pd.read_csv(csv_link)
        for _, r in df_conf.iterrows():
            nom = str(r.get('NOMBRE','')).strip().upper()
            if "GINELAP" in nom: nom = "GINELAP"
            dias_str = str(r.get('DIAS_LIBRES',''))
            
            if nom in INTEGRANTES and dias_str and dias_str.lower() != 'nan':
                # Convierte "0, 1, 5" en una lista de números [0, 1, 5]
                libres_fijos[nom] = [int(x.strip()) for x in dias_str.split(',') if x.strip().isdigit()]
    except Exception as e: st.sidebar.warning("No se pudo leer la Configuración.")
    return libres_fijos

# --- MOTOR DE GENERACIÓN ---
def generar_cuadro_equitativo(mes, ano, historial_previo, sugerencias_dict, config_dict):
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
        es_finde_o_festivo = wd >= 5 or es_festivo(d, mes, ano)

        def racha_actual(persona):
            streak = 0
            for past_d in range(d - 1, d - 10, -1):
                t_pasado = turno_en_dia(persona, past_d)
                if any(t in t_pasado for t in ['C', 'N']) and 'P' not in t_pasado: streak += 1
                else: break
            return streak

        def necesita_descanso(persona):
            return racha_actual(persona) >= 3

        cuota_n = 2
        if es_festivo(d, mes, ano) or wd == 6: cuota_c = 2
        elif wd == 5: cuota_c = 5
        else: cuota_c = 6

        # 1. POSTURNOS Y SUGERENCIAS
        for p in INTEGRANTES:
            if 'N' in turno_en_dia(p, d-1): 
                df.at[p, ds] = 'P'
                continue
            
            req = sugerencias_dict.get(p, {}).get(ds)
            if req:
                turno_limpio = req
                if 'C' in req and 'P' not in req: turno_limpio = 'C'
                elif 'N' in req and 'P' not in req: turno_limpio = 'N'
                
                df.at[p, ds] = turno_limpio
                if turno_limpio == 'N': cuota_n -= 1
                if turno_limpio == 'C': cuota_c -= 1
                if turno_limpio in ['C', 'N']:
                    turnos_totales[p] += 1
                    if es_finde_o_festivo: finde_totales[p] += 1
                    if turno_limpio == 'N': noches_totales[p] += 1

        # 2. LIBRES FIJOS (LEÍDOS DESDE GOOGLE SHEETS)
        for p in INTEGRANTES:
            if df.at[p, ds] == "" and wd in config_dict.get(p, []):
                # Excepción: Juan Camilo los jueves (wd=3) puede hacer turno si se necesita, 
                # pero por defecto se asume Libre si no pidió turno en sugerencias.
                if p == "JUAN CAMILO PEREZ" and wd == 3 and cuota_c > 0:
                    pass # Se deja disponible por si hace falta para salvar el día
                else:
                    df.at[p, ds] = "L"

        # 3. ASIGNACIONES FIJAS BLINDADAS (Jhon y Angie)
        if wd == 1 and df.at["JHON RIOS", ds] == "" and cuota_n > 0: 
            df.at["JHON RIOS", ds] = "N"
            cuota_n -= 1
            turnos_totales["JHON RIOS"] += 1
            noches_totales["JHON RIOS"] += 1

        if wd == 2 and df.at["ANGIE BERNAL", ds] == "" and cuota_c > 0: 
            df.at["ANGIE BERNAL", ds] = "C"
            cuota_c -= 1
            turnos_totales["ANGIE BERNAL"] += 1

        # 4. REPARTIR NOCHES
        for _ in range(max(0, cuota_n)):
            disp_noches = [p for p in INTEGRANTES if df.at[p, ds] == "" and not necesita_descanso(p)]
            disp_noches = [p for p in disp_noches if 'N' not in turno_en_dia(p, d-2)] # Anti N-P-N

            if d < dias_mes:
                m_wd = (wd + 1) % 7
                # EL ESCUDO ANTI-POSTURNO: Si mañana es su día libre fijo, NO puede hacer noche hoy
                disp_noches = [p for p in disp_noches if m_wd not in config_dict.get(p, [])]

            if not disp_noches: disp_noches = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if disp_noches:
                random.shuffle(disp_noches)
                disp_noches.sort(key=lambda x: (noches_totales[x], racha_actual(x) >= 2, turnos_totales[x]))
                elegido = disp_noches[0]
                df.at[elegido, ds] = "N"
                turnos_totales[elegido] += 1
                noches_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 5. REPARTIR CORRIDOS
        for _ in range(max(0, cuota_c)):
            disp_dia = [p for p in INTEGRANTES if df.at[p, ds] == "" and not necesita_descanso(p)]
            if not disp_dia: disp_dia = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if disp_dia:
                random.shuffle(disp_dia)
                if es_finde_o_festivo: disp_dia.sort(key=lambda x: (racha_actual(x) >= 2, finde_totales[x], turnos_totales[x]))
                else: disp_dia.sort(key=lambda x: (racha_actual(x) >= 2, turnos_totales[x], racha_actual(x)))
                
                elegido = disp_dia[0]
                df.at[elegido, ds] = "C"
                turnos_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 6. RELLENAR CON DESCANSO
        for p in INTEGRANTES:
            if df.at[p, ds] == "": df.at[p, ds] = "D"

    df['TOTAL TURNOS'] = df.apply(lambda r: sum(1 for c in df.columns if any(t in str(r[c]) for t in ['C', 'N'])), axis=1)
    df['TOTAL NOCHES'] = df.apply(lambda r: noches_totales[r.name], axis=1)
    df['FINES DE SEMANA'] = df.apply(lambda r: finde_totales[r.name], axis=1)
    return df

# --- INTERFAZ ---
st.title("🏥 Cuadro de Instrumentación Automático")

with st.sidebar:
    st.header("1. Empalme (Mes Anterior)")
    archivo_previo = st.file_uploader("Sube el Excel:", type=['xlsx', 'csv'])
    
    st.header("2. Google Sheets")
    link_sheet = st.text_input("Link Pestaña SUGERENCIAS:", "https://docs.google.com/spreadsheets/d/...")
    link_config = st.text_input("Link Pestaña CONFIGURACIÓN (Libres Fijos):", "https://docs.google.com/spreadsheets/d/...")
    
    st.header("3. Generación")
    ano_sel = st.number_input("Año", min_value=2024, max_value=2035, value=datetime.now().year)
    mes_sel = st.selectbox("Mes", range(1, 13), index=datetime.now().month-1)

if st.button("🚀 GENERAR CUADRO", type="primary", use_container_width=True):
    with st.spinner("Conectando con Google Sheets y balanceando..."):
        historial_leido = procesar_historial_empalme(archivo_previo)
        sugerencias_leidas = procesar_sugerencias(link_sheet)
        config_leida = procesar_configuracion(link_config)
        
        resultado = generar_cuadro_equitativo(mes_sel, ano_sel, historial_leido, sugerencias_leidas, config_leida)
        
        cols_dias = [str(d) for d in range(1, calendar.monthrange(ano_sel, mes_sel)[1] + 1)]
        st.dataframe(resultado.style.applymap(aplicar_colores, subset=cols_dias), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resultado.to_excel(writer, index=True)
        st.download_button("📥 Descargar Nómina Excel", output.getvalue(), f"Turnos_{mes_sel}_{ano_sel}.xlsx", use_container_width=True)
