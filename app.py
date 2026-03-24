import streamlit as st
import pandas as pd
from datetime import datetime
import calendar
import random
import io
import holidays
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Instrumentación", layout="wide")

INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES ARGUMEDO", 
    "KELLY CAUSIL", "JOAN CARMONA"
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

# --- LECTORES ---
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
    except Exception: st.sidebar.error("Error en historial.")
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
        for _, r in df_sug.iterrows():
            nom = str(r.get('NOMBRE','')).strip().upper()
            if "GINELAP" in nom: nom = "GINELAP"
            fecha = ''.join(filter(str.isdigit, str(r.get('FECHA',''))))
            sol = str(r.get('SOLICITUD','')).strip().upper()
            if nom in INTEGRANTES and fecha and sol != 'NAN': sugerencias[nom][fecha] = sol
    except: st.sidebar.warning("Link Sugerencias no detectado.")
    return sugerencias

def procesar_configuracion(link):
    libres_fijos = {p: [] for p in INTEGRANTES}
    if not link or not link.startswith("http"): return libres_fijos
    try:
        base_url = link.split('/edit')[0]
        csv_link = f"{base_url}/export?format=csv"
        match = re.search(r'gid=(\d+)', link)
        if match: csv_link += f"&gid={match.group(1)}"
        df_conf = pd.read_csv(csv_link)
        for _, r in df_conf.iterrows():
            nom = str(r.get('NOMBRE','')).strip().upper()
            if "GINELAP" in nom: nom = "GINELAP"
            dias_str = str(r.get('DIAS_LIBRES',''))
            if nom in INTEGRANTES and dias_str and dias_str.lower() != 'nan':
                libres_fijos[nom] = [int(x.strip()) for x in dias_str.split(',') if x.strip().isdigit()]
    except: st.sidebar.warning("Link Configuración no detectado.")
    return libres_fijos

# --- MOTOR ---
def generar_cuadro_equitativo(mes, ano, historial_previo, sugerencias_dict, config_dict):
    dias_mes = calendar.monthrange(ano, mes)[1]
    df = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    turnos_totales = {p: 0 for p in INTEGRANTES}; noches_totales = {p: 0 for p in INTEGRANTES}; finde_totales = {p: 0 for p in INTEGRANTES}

    def turno_en_dia(persona, dia_req):
        if dia_req > 0: return str(df.at[persona, str(dia_req)])
        else:
            idx = 2 + dia_req
            hist = historial_previo.get(persona, ["", "", ""])
            if 0 <= idx < 3: return str(hist[idx])
            return ""

    for d in range(1, dias_mes + 1):
        ds = str(d); wd = datetime(ano, mes, d).weekday(); es_finde_o_festivo = wd >= 5 or es_festivo(d, mes, ano)
        
        def racha_actual(persona):
            streak = 0
            for past_d in range(d - 1, d - 10, -1):
                t_p = turno_en_dia(persona, past_d)
                if any(t in t_p for t in ['C', 'N']) and 'P' not in t_p: streak += 1
                else: break
            return streak
        def necesita_descanso(persona): return racha_actual(persona) >= 3

        cuota_n = 2
        if es_finde_o_festivo and (wd == 6 or es_festivo(d, mes, ano)): cuota_c = 2
        elif wd == 5: cuota_c = 5
        else: cuota_c = 6

        # 1. POSTURNOS (PRIORIDAD 1)
        for p in INTEGRANTES:
            if 'N' in turno_en_dia(p, d-1): df.at[p, ds] = 'P'

        # 2. ASIGNACIONES FIJAS BLINDADAS (PRIORIDAD 2)
        if wd == 1 and df.at["JHON RIOS", ds] == "":
            df.at["JHON RIOS", ds] = "N"
            cuota_n -= 1; turnos_totales["JHON RIOS"] += 1; noches_totales["JHON RIOS"] += 1
            if es_finde_o_festivo: finde_totales["JHON RIOS"] += 1 # Suma si un martes es festivo
            
        if wd == 2 and df.at["ANGIE BERNAL", ds] == "":
            df.at["ANGIE BERNAL", ds] = "C"
            cuota_c -= 1; turnos_totales["ANGIE BERNAL"] += 1
            if es_finde_o_festivo: finde_totales["ANGIE BERNAL"] += 1 # Suma si un miércoles es festivo

        # 3. SUGERENCIAS (PRIORIDAD 3)
        for p in INTEGRANTES:
            if df.at[p, ds] != "": continue
            req = sugerencias_dict.get(p, {}).get(ds)
            if req:
                tl = 'C' if ('C' in req and 'P' not in req) else ('N' if ('N' in req and 'P' not in req) else req)
                df.at[p, ds] = tl
                if tl == 'N': cuota_n -= 1; noches_totales[p] += 1; turnos_totales[p] += 1
                if tl == 'C': cuota_c -= 1; turnos_totales[p] += 1
                if tl in ['C', 'N'] and es_finde_o_festivo: finde_totales[p] += 1 # Suma si pide finde

        # 4. LIBRES FIJOS (PRIORIDAD 4)
        for p in INTEGRANTES:
            if df.at[p, ds] == "" and wd in config_dict.get(p, []):
                if not (p == "JUAN CAMILO PEREZ" and wd == 3): 
                    df.at[p, ds] = "L"

        # 5. REPARTIR NOCHES
        for _ in range(max(0, cuota_n)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == "" and not necesita_descanso(p)]
            disp = [p for p in disp if 'N' not in turno_en_dia(p, d-2)]
            if d < dias_mes:
                m_wd = (wd + 1) % 7
                disp = [p for p in disp if m_wd not in config_dict.get(p, [])]
            if not disp: disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if disp:
                random.shuffle(disp)
                # Ojo aquí: Si es finde, ordena a los de menos findes primero
                if es_finde_o_festivo: disp.sort(key=lambda x: (noches_totales[x], finde_totales[x], racha_actual(x) >= 2, turnos_totales[x]))
                else: disp.sort(key=lambda x: (noches_totales[x], racha_actual(x) >= 2, turnos_totales[x]))
                
                elegido = disp[0]; df.at[elegido, ds] = "N"
                turnos_totales[elegido] += 1; noches_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1 # ¡Acá estaba el error, ya suma!

        # 6. REPARTIR CORRIDOS
        for _ in range(max(0, cuota_c)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == "" and not necesita_descanso(p)]
            if not disp: disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if disp:
                random.shuffle(disp)
                # Ojo aquí: Si es finde, ordena a los de menos findes primero
                if es_finde_o_festivo: disp.sort(key=lambda x: (racha_actual(x) >= 2, finde_totales[x], turnos_totales[x]))
                else: disp.sort(key=lambda x: (racha_actual(x) >= 2, turnos_totales[x], racha_actual(x)))
                
                elegido = disp[0]; df.at[elegido, ds] = "C"; turnos_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1 # ¡Acá estaba el error, ya suma!

        # 7. RELLENAR CON DESCANSO
        for p in INTEGRANTES:
            if df.at[p, ds] == "": df.at[p, ds] = "D"

    df['TOTAL TURNOS'] = df.apply(lambda r: sum(1 for c in df.columns if any(t in str(r[c]) for t in ['C', 'N'])), axis=1)
    df['TOTAL NOCHES'] = df.apply(lambda r: noches_totales[r.name], axis=1)
    df['FINES DE SEMANA'] = df.apply(lambda r: finde_totales[r.name], axis=1)
    return df

# --- INTERFAZ ---
st.title("🏥 Gestor de Turnos (Fines de Semana y Sugerencias)")

with st.sidebar:
    st.header("1. Cargar Datos")
    archivo_previo = st.file_uploader("Excel Mes Anterior:", type=['xlsx', 'csv'])
    link_sheet = st.text_input("Link Sugerencias:", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?pli=1&gid=0#gid=0")
    if link_sheet.startswith("http"): st.link_button("📝 Abrir Sugerencias", link_sheet, use_container_width=True)
    link_config = st.text_input("Link Configuración:", "")
    ano_sel = st.number_input("Año:", min_value=2024, value=datetime.now().year)
    mes_sel = st.selectbox("Mes:", range(1, 13), index=datetime.now().month-1)

if st.button("🚀 GENERAR CUADRO", type="primary", use_container_width=True):
    hist = procesar_historial_empalme(archivo_previo)
    sug = procesar_sugerencias(link_sheet); conf = procesar_configuracion(link_config)
    res = generar_cuadro_equitativo(mes_sel, ano_sel, hist, sug, conf)
    
    dias_en_mes = calendar.monthrange(ano_sel, mes_sel)[1]
    st.dataframe(res.style.applymap(aplicar_colores, subset=[str(d) for d in range(1, dias_en_mes + 1)]), use_container_width=True)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        res.to_excel(writer, index=True, sheet_name='Turnos')
        wb = writer.book; ws = writer.sheets['Turnos']
        ws.freeze_panes(1, 1)
        fmt_v = wb.add_format({'bg_color': '#d9ead3'}); fmt_r = wb.add_format({'bg_color': '#f4cccc'})
        fmt_az = wb.add_format({'bg_color': '#cfe2f3'}); fmt_am = wb.add_format({'bg_color': '#fff2cc'})
        ws.set_column(0, 0, 25); ws.set_column(1, dias_en_mes, 4)
        for r_idx, (idx, row) in enumerate(res.iterrows()):
            for c_idx, val in enumerate(row):
                f = None
                if val in ['L', 'D']: f = fmt_v
                elif val == 'P': f = fmt_r
                elif 'N' in str(val): f = fmt_az
                elif 'C' in str(val): f = fmt_am
                ws.write(r_idx + 1, c_idx + 1, val, f)
    st.download_button("📥 Descargar Excel", output.getvalue(), f"Turnos_{mes_sel}.xlsx", use_container_width=True)
