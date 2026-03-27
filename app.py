import streamlit as st
import pandas as pd
from datetime import datetime
import calendar
import random
import io
import holidays
import re
import xlsxwriter
import unicodedata

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Instrumentación", layout="wide")

INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES ARGUMEDO", 
    "KELLY CAUSIL", "JOAN CARMONA"
]

DIAS_MAP = {
    'LUNES': 0, 'MARTES': 1, 'MIERCOLES': 2, 'JUEVES': 3, 'VIERNES': 4, 'SABADO': 5, 'DOMINGO': 6
}

def es_festivo(dia, mes, ano):
    co_holidays = holidays.CO(years=ano)
    return datetime(ano, mes, dia).date() in co_holidays

def aplicar_colores(v):
    if v in ['L', 'D']: return 'background-color: #d9ead3; color: #000000;' 
    if v == 'P': return 'background-color: #f4cccc; color: #000000;'      
    if 'N' in str(v): return 'background-color: #cfe2f3; color: #000000;' 
    if 'C' in str(v): return 'background-color: #fff2cc; color: #000000;' 
    if v == 'V': return 'background-color: #e4d7f5; color: #000000;'      
    return ''

def normalizar_texto(texto):
    if pd.isna(texto): return ""
    texto = str(texto).strip()
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn').upper()

# --- LECTORES ---
def procesar_historial_empalme(file):
    historial = {p: ["", "", ""] for p in INTEGRANTES} 
    if not file: return historial
    try:
        df = pd.read_excel(file) if file.name.endswith('.xlsx') else pd.read_csv(file)
        if not any(str(c).isdigit() for c in df.columns):
            df = pd.read_excel(file, skiprows=9) if file.name.endswith('.xlsx') else pd.read_csv(file, skiprows=9)
        df.columns = [normalizar_texto(c) for c in df.columns]
        col_nombre = next((c for c in df.columns if 'NOMBRE' in c or 'UNNAMED' in c), None)
        cols_dias = [c for c in df.columns if c.isdigit()]
        if len(cols_dias) >= 3:
            ultimas_3 = sorted(cols_dias, key=int)[-3:]
            for _, r in df.iterrows():
                nom = normalizar_texto(r[col_nombre]) if col_nombre else normalizar_texto(r.name)
                if "GINELAP" in nom: nom = "GINELAP"
                if nom in INTEGRANTES: historial[nom] = [normalizar_texto(r[c]) for c in ultimas_3]
    except: pass
    return historial

def procesar_sugerencias(link):
    sugerencias = {p: {} for p in INTEGRANTES}
    if not link or not link.startswith("http"): return sugerencias
    try:
        base_url = link.split('/edit')[0]
        csv_link = f"{base_url}/export?format=csv"
        match = re.search(r'gid=(\d+)', link)
        if match: csv_link += f"&gid={match.group(1)}"
        df_sug = pd.read_csv(csv_link)
        df_sug.columns = [normalizar_texto(c) for c in df_sug.columns]
        col_nom = next((c for c in df_sug.columns if 'NOMBRE' in c), None)
        col_fec = next((c for c in df_sug.columns if 'FECHA' in c), None)
        col_sol = next((c for c in df_sug.columns if 'SOLICITUD' in c or 'TURNO' in c), None)
        if col_nom and col_fec and col_sol:
            for _, r in df_sug.iterrows():
                nom = normalizar_texto(r[col_nom])
                fecha = ''.join(filter(str.isdigit, str(r[col_fec])))
                sol = normalizar_texto(r[col_sol])
                if nom in INTEGRANTES and fecha and sol: sugerencias[nom][fecha] = sol
    except: pass
    return sugerencias

def procesar_configuracion(link):
    libres_fijos = {p: [] for p in INTEGRANTES}
    vacaciones = {p: [] for p in INTEGRANTES}
    if not link or not link.startswith("http"): return libres_fijos, vacaciones
    try:
        base_url = link.split('/edit')[0]
        csv_link = f"{base_url}/export?format=csv"
        match = re.search(r'gid=(\d+)', link)
        if match: csv_link += f"&gid={match.group(1)}"
        df_conf = pd.read_csv(csv_link)
        df_conf.columns = [normalizar_texto(c) for c in df_conf.columns]
        col_nom = next((c for c in df_conf.columns if 'NOMBRE' in c), None)
        col_dias = next((c for c in df_conf.columns if 'LIBRE' in c), None)
        col_vac = next((c for c in df_conf.columns if 'VACACION' in c), None)
        if col_nom:
            for _, r in df_conf.iterrows():
                nom_sheet = normalizar_texto(r[col_nom])
                nom_real = next((p for p in INTEGRANTES if nom_sheet in p or p in nom_sheet), None)
                if nom_real:
                    if col_dias and str(r[col_dias]) != 'nan':
                        nums = [int(x) for x in re.findall(r'\d+', str(r[col_dias]))]
                        libres_fijos[nom_real] = [n for n in nums if 0 <= n <= 6]
                    if col_vac and str(r[col_vac]) != 'nan':
                        vac_str = str(r[col_vac]).replace(' Y ', ',').replace('&', ',')
                        for parte in vac_str.split(','):
                            parte = parte.strip()
                            if '-' in parte:
                                rng_v = parte.split('-')
                                if len(rng_v) == 2: vacaciones[nom_real].extend(list(range(int(rng_v[0]), int(rng_v[1]) + 1)))
                            elif parte.isdigit(): vacaciones[nom_real].append(int(parte))
    except: pass
    return libres_fijos, vacaciones

# --- MOTOR PRINCIPAL ---
def generar_cuadro_equitativo(mes, ano, historial_previo, sugerencias_dict, config_dict, vacaciones_dict, semilla):
    rng = random.Random(semilla)
    dias_mes = calendar.monthrange(ano, mes)[1]
    df = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    turnos_totales = {p: 0 for p in INTEGRANTES}
    noches_totales = {p: 0 for p in INTEGRANTES}
    festivos_trabajados = {p: 0 for p in INTEGRANTES}

    def turno_en_dia(persona, dia_req):
        if dia_req > 0: return str(df.at[persona, str(dia_req)])
        idx = 2 + dia_req
        hist = historial_previo.get(persona, ["", "", ""])
        return str(hist[idx]) if 0 <= idx < 3 else ""

    def trabajo_mismo_dia_semana_anterior(persona, d):
        if d > 7:
            t_ant = str(df.at[persona, str(d-7)])
            return any(t in t_ant for t in ['C', 'N'])
        return False

    def no_puede_hacer_noche(p, d):
        if d < dias_mes:
            manana_wd = datetime(ano, mes, d+1).weekday()
            t_manana = str(df.at[p, str(d+1)])
            if any(x in t_manana for x in ['V', 'P']): return True
            if manana_wd in config_dict.get(p, []):
                if p == "JUAN CAMILO PEREZ" and manana_wd == 3: return False
                return True
        return False

    # FASE 1: PRE-LLENADO
    for d in range(1, dias_mes + 1):
        ds, wd = str(d), datetime(ano, mes, d).weekday()
        for p in INTEGRANTES:
            if d in vacaciones_dict.get(p, []): df.at[p, ds] = 'V'; continue
            req = sugerencias_dict.get(p, {}).get(ds)
            if wd in config_dict.get(p, []):
                if p == "JUAN CAMILO PEREZ" and wd == 3: # Jueves flexible
                    if req:
                        tl = 'C' if 'C' in req else ('N' if 'N' in req else req)
                        df.at[p, ds] = tl
                        if tl == 'N': noches_totales[p] += 1; turnos_totales[p] += 1
                        if tl == 'C': turnos_totales[p] += 1
                    continue
                df.at[p, ds] = "L"; continue
            if req:
                tl = 'C' if 'C' in req else ('N' if 'N' in req else req)
                df.at[p, ds] = tl
                if tl == 'N': noches_totales[p] += 1; turnos_totales[p] += 1
                if tl == 'C': turnos_totales[p] += 1
                if es_festivo(d, mes, ano): festivos_trabajados[p] += 1

    # FASE 2: REPARTICIÓN CON ROTACIÓN
    for d in range(1, dias_mes + 1):
        ds, wd = str(d), datetime(ano, mes, d).weekday()
        es_fest = es_festivo(d, mes, ano)
        es_finde = wd >= 5
        
        cuota_n = 2 - sum(1 for p in INTEGRANTES if 'N' in df.at[p, ds])
        cuota_c_base = 2 if (es_fest or wd == 6) else (5 if wd == 5 else 6)
        cuota_c = cuota_c_base - sum(1 for p in INTEGRANTES if 'C' in df.at[p, ds])

        for p in INTEGRANTES:
            if 'N' in turno_en_dia(p, d-1) and df.at[p, ds] == "": df.at[p, ds] = 'P'
            if df.at[p, ds] == "" and sum(1 for x in range(d-3, d) if any(t in turno_en_dia(p, x) for t in ['C', 'N'])) >= 3:
                df.at[p, ds] = 'D'

        for _ in range(max(0, cuota_n)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == "" and 'N' not in turno_en_dia(p, d-2) and not no_puede_hacer_noche(p, d)]
            if es_finde:
                filt = [p for p in disp if not trabajo_mismo_dia_semana_anterior(p, d)]
                if filt: disp = filt
            if es_fest: disp.sort(key=lambda x: festivos_trabajados[x])
            
            if disp:
                rng.shuffle(disp)
                disp.sort(key=lambda x: (noches_totales[x], turnos_totales[x]))
                df.at[disp[0], ds] = "N"; turnos_totales[disp[0]] += 1; noches_totales[disp[0]] += 1
                if es_fest: festivos_trabajados[disp[0]] += 1

        for _ in range(max(0, cuota_c)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if es_finde:
                filt = [p for p in disp if not trabajo_mismo_dia_semana_anterior(p, d)]
                if filt: disp = filt
            if es_fest: disp.sort(key=lambda x: festivos_trabajados[x])
            
            if disp:
                rng.shuffle(disp)
                disp.sort(key=lambda x: turnos_totales[x])
                df.at[disp[0], ds] = "C"; turnos_totales[disp[0]] += 1
                if es_fest: festivos_trabajados[disp[0]] += 1

        for p in INTEGRANTES:
            if df.at[p, ds] == "": df.at[p, ds] = "D"

    # Conteos finales
    df['TOTAL CORRIDOS'] = df.apply(lambda r: sum(1 for c in df.columns if 'C' in str(r[c]) and c.isdigit()), axis=1)
    df['TOTAL NOCHES'] = df.apply(lambda r: sum(1 for c in df.columns if 'N' in str(r[c]) and c.isdigit()), axis=1)
    df['TOTAL TURNOS'] = df['TOTAL CORRIDOS'] + df['TOTAL NOCHES']
    df['FINES DE SEMANA'] = df.apply(lambda r: sum(1 for d in range(1, dias_mes+1) if (datetime(ano, mes, d).weekday() >= 5 or es_festivo(d, mes, ano)) and any(t in str(r[str(d)]) for t in ['C', 'N'])), axis=1)
    return df

# --- FUNCIÓN DE OPTIMIZACIÓN (TU NUEVA FUNCIÓN) ---
def generar_mejor_escenario(n, mes, ano, hist, sug, conf, vacs):
    resultados = []
    dias_mes = calendar.monthrange(ano, mes)[1]
    barra_progreso = st.progress(0)

    for s in range(n):
        df = generar_cuadro_equitativo(mes, ano, hist, sug, conf, vacs, s)

        cargas_relativas = []

        for p in INTEGRANTES:
            dias_trabajables = 0

            for d in range(1, dias_mes + 1):
                val = df.at[p, str(d)]
                if val not in ['L', 'V']:  # Solo indisponibilidad real
                    dias_trabajables += 1

            if dias_trabajables > 0:
                ocupacion = df.at[p, "TOTAL TURNOS"] / dias_trabajables
                cargas_relativas.append(ocupacion)

        import numpy as np
        score = (
            np.std(cargas_relativas) * 2.0 +
            df["TOTAL NOCHES"].std() * 2.0
        )

        resultados.append((score, df))
        barra_progreso.progress((s + 1) / n)

    resultados.sort(key=lambda x: x[0])
    return resultados[0][1]

# --- INTERFAZ ---
st.title("🏥 Gestor de Turnos (Optimizador de Equidad)")
with st.sidebar:
    st.header("1. Carga de Datos")
    archivo_previo = st.file_uploader("Excel Mes Anterior:", type=['xlsx', 'csv'])
    link_sheet = st.text_input("Link Sugerencias:", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?pli=1&gid=0#gid=0")
    link_config = st.text_input("Link Configuración:", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?pli=1&gid=1679804429#gid=1679804429")
    
    st.header("2. Parámetros")
    ano_sel = st.number_input("Año:", min_value=2024, value=2026)
    mes_sel = st.selectbox("Mes:", range(1, 13), index=2) # Marzo
    iteraciones = st.slider("Número de simulaciones (n):", 10, 100, 30)
    mostrar_rx = st.checkbox("🔍 Rayos X")

if st.button("🚀 GENERAR MEJOR ESCENARIO", type="primary", use_container_width=True):
    hist, sug = procesar_historial_empalme(archivo_previo), procesar_sugerencias(link_sheet)
    conf, vacs = procesar_configuracion(link_config)
    
    if mostrar_rx:
        st.write("🏖️ Vacaciones:", vacs); st.write("🛑 Libres:", conf); st.stop()
    
    # EJECUCIÓN DEL OPTIMIZADOR
    with st.spinner('Simulando múltiples escenarios para encontrar el más equitativo...'):
        res = generar_mejor_escenario(iteraciones, mes_sel, ano_sel, hist, sug, conf, vacs)
    
    # Mostrar resultados
    dias_en_mes = calendar.monthrange(ano_sel, mes_sel)[1]
    st.success(f"✅ Se han analizado {iteraciones} versiones. ¡Aquí tienes la más equilibrada!")
    st.dataframe(res.style.map(aplicar_colores, subset=[str(d) for d in range(1, dias_en_mes + 1)]), use_container_width=True)
    
    # Botón de Descarga
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        res.to_excel(writer, sheet_name='Turnos')
    st.download_button("📥 Descargar Excel", output.getvalue(), f"Turnos_{mes_sel}_Optimizado.xlsx", use_container_width=True)
    
