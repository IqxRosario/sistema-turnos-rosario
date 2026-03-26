import streamlit as st
import pandas as pd
from datetime import datetime
import calendar
import random
import io
import holidays
import re
import xlsxwriter

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Gestor de Turnos Pro", layout="wide")

# --- BLOQUEO VISUAL ---
st.markdown(
    """
    <style>
    [data-testid="stElementToolbar"] {display: none;}
    </style>
    """,
    unsafe_allow_html=True
)

def es_festivo(dia, mes, ano):
    co_holidays = holidays.CO(years=ano)
    return datetime(ano, mes, dia).date() in co_holidays

def aplicar_colores(v):
    if v in ['L', 'D']: return 'background-color: #d9ead3; color: #000000;' 
    if v == 'P': return 'background-color: #f4cccc; color: #000000;'      
    if 'N' in str(v): return 'background-color: #cfe2f3; color: #000000;' 
    if 'C' in str(v): return 'background-color: #fff2cc; color: #000000;' 
    return ''

# --- LECTORES DINÁMICOS ---

@st.cache_data(ttl=60)
def procesar_personal(link):
    """Lee la lista de integrantes desde la pestaña PERSONAL"""
    # Lista de respaldo por si el link falla o está vacío
    default = ["GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO", "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR", "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES ARGUMEDO", "KELLY CAUSIL", "JOAN CARMONA"]
    if not link or not link.startswith("http"): return default
    try:
        base_url = link.split('/edit')[0]
        # Buscamos específicamente la pestaña llamada PERSONAL
        csv_link = f"{base_url}/gviz/tq?tqx=out:csv&sheet=PERSONAL"
        df_per = pd.read_csv(csv_link)
        nombres = df_per['NOMBRE'].dropna().astype(str).str.strip().str.upper().tolist()
        return nombres if nombres else default
    except:
        return default

@st.cache_data(ttl=60)
def procesar_historial_empalme(file, lista_integrantes):
    historial = {p: ["", "", ""] for p in lista_integrantes} 
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
                if nom in lista_integrantes: historial[nom] = [str(r[c]).upper() for c in ultimas_3]
    except: pass
    return historial

@st.cache_data(ttl=60)
def procesar_sugerencias(link, lista_integrantes):
    sugerencias = {p: {} for p in lista_integrantes}
    if not link or not link.startswith("http"): return sugerencias
    try:
        base_url = link.split('/edit')[0]
        csv_link = f"{base_url}/gviz/tq?tqx=out:csv&sheet=Sheet1" # O el nombre de tu pestaña de sugerencias
        df_sug = pd.read_csv(csv_link)
        for _, r in df_sug.iterrows():
            nom = str(r.get('NOMBRE','')).strip().upper()
            if "GINELAP" in nom: nom = "GINELAP"
            fecha = ''.join(filter(str.isdigit, str(r.get('FECHA',''))))
            sol = str(r.get('SOLICITUD','')).strip().upper()
            if nom in lista_integrantes and fecha and sol != 'NAN': sugerencias[nom][fecha] = sol
    except: pass
    return sugerencias

@st.cache_data(ttl=60)
def procesar_configuracion(link, lista_integrantes):
    libres_fijos = {p: [] for p in lista_integrantes}
    if not link or not link.startswith("http"): return libres_fijos
    try:
        base_url = link.split('/edit')[0]
        csv_link = f"{base_url}/gviz/tq?tqx=out:csv&sheet=CONFIG" # Nombre de pestaña de config
        df_conf = pd.read_csv(csv_link)
        for _, r in df_conf.iterrows():
            nom = str(r.get('NOMBRE','')).strip().upper()
            if "GINELAP" in nom: nom = "GINELAP"
            dias_str = str(r.get('DIAS_LIBRES',''))
            if nom in lista_integrantes and dias_str and dias_str.lower() != 'nan':
                libres_fijos[nom] = [int(x.strip()) for x in dias_str.split(',') if x.strip().isdigit()]
    except: pass
    return libres_fijos

# --- MOTOR LÓGICO ---
def generar_cuadro_equitativo(mes, ano, historial_previo, sugerencias_dict, config_dict, semilla, lista_integrantes):
    random.seed(semilla)
    dias_mes = calendar.monthrange(ano, mes)[1]
    df = pd.DataFrame(index=lista_integrantes, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    turnos_totales = {p: 0 for p in lista_integrantes}
    noches_totales = {p: 0 for p in lista_integrantes}
    finde_totales = {p: 0 for p in lista_integrantes}

    def turno_en_dia(persona, dia_req):
        if dia_req > 0: return str(df.at[persona, str(dia_req)])
        else:
            idx = 2 + dia_req
            hist = historial_previo.get(persona, ["", "", ""])
            return str(hist[idx]) if 0 <= idx < 3 else ""

    for d in range(1, dias_mes + 1):
        ds = str(d); wd = datetime(ano, mes, d).weekday(); es_finde_o_festivo = wd >= 5 or es_festivo(d, mes, ano)
        
        def racha_actual(persona):
            streak = 0
            for past_d in range(d - 1, d - 10, -1):
                t_p = turno_en_dia(persona, past_d)
                if any(t in t_p for t in ['C', 'N']) and 'P' not in t_p: streak += 1
                else: break
            return streak

        cuota_n = 2
        cuota_c = 2 if (wd == 6 or es_festivo(d, mes, ano)) else (5 if wd == 5 else 6)

        # 1. Posturnos
        for p in lista_integrantes:
            if 'N' in turno_en_dia(p, d-1): df.at[p, ds] = 'P'
        # 2. Antifatiga
        for p in lista_integrantes:
            if df.at[p, ds] == "" and racha_actual(p) >= 3: df.at[p, ds] = 'D'
        # 3. Reglas Fijas (Solo si existen en la lista actual)
        if "JHON RIOS" in lista_integrantes and wd == 1 and df.at["JHON RIOS", ds] == "":
            df.at["JHON RIOS", ds] = "N"; cuota_n -= 1; turnos_totales["JHON RIOS"] += 1; noches_totales["JHON RIOS"] += 1
            if es_finde_o_festivo: finde_totales["JHON RIOS"] += 1
        if "ANGIE BERNAL" in lista_integrantes and wd == 2 and df.at["ANGIE BERNAL", ds] == "":
            df.at["ANGIE BERNAL", ds] = "C"; cuota_c -= 1; turnos_totales["ANGIE BERNAL"] += 1
            if es_finde_o_festivo: finde_totales["ANGIE BERNAL"] += 1

        # 4. Sugerencias
        for p in lista_integrantes:
            if df.at[p, ds] != "": continue
            req = sugerencias_dict.get(p, {}).get(ds)
            if req:
                tl = 'C' if 'C' in req else ('N' if 'N' in req else req)
                df.at[p, ds] = tl
                if tl == 'N': cuota_n -= 1; noches_totales[p] += 1; turnos_totales[p] += 1
                if tl == 'C': cuota_c -= 1; turnos_totales[p] += 1
                if tl in ['C', 'N'] and es_finde_o_festivo: finde_totales[p] += 1

        # 5. Repartir Noches
        for _ in range(max(0, cuota_n)):
            disp = [p for p in lista_integrantes if df.at[p, ds] == ""]
            if disp:
                random.shuffle(disp)
                disp.sort(key=lambda x: (noches_totales[x], finde_totales[x]))
                elegido = disp[0]; df.at[elegido, ds] = "N"; turnos_totales[elegido] += 1; noches_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1
        # 6. Repartir Corridos
        for _ in range(max(0, cuota_c)):
            disp = [p for p in lista_integrantes if df.at[p, ds] == ""]
            if disp:
                random.shuffle(disp)
                disp.sort(key=lambda x: (finde_totales[x], turnos_totales[x]))
                elegido = disp[0]; df.at[elegido, ds] = "C"; turnos_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1
        # 7. Rellenar
        for p in lista_integrantes:
            if df.at[p, ds] == "": df.at[p, ds] = "D"

    df['TOTAL TURNOS'] = df.apply(lambda r: sum(1 for c in df.columns if any(t in str(r[c]) for t in ['C', 'N'])), axis=1)
    df['TOTAL NOCHES'] = df.apply(lambda r: noches_totales[r.name], axis=1)
    df['FINES DE SEMANA'] = df.apply(lambda r: finde_totales[r.name], axis=1)
    return df

# --- INTERFAZ ---
with st.sidebar:
    st.header("Configuración de Negocio")
    link_principal = st.text_input("Enlace Google Sheets (Sugerencias/Personal):", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit")
    archivo_previo = st.file_uploader("Excel Mes Anterior:", type=['xlsx', 'csv'])
    ano_sel = st.number_input("Año:", min_value=2024, value=datetime.now().year)
    mes_sel = st.selectbox("Mes:", range(1, 13), index=datetime.now().month-1)
    semilla = st.number_input("🎲 Semilla de Reproducción:", value=42)

st.title("🏥 Sistema de Turnos Inteligente")

if st.button("🚀 GENERAR CUADRO", type="primary", use_container_width=True):
    # Paso 1: Leer el personal desde la pestaña PERSONAL
    lista_actual = procesar_personal(link_principal)
    
    # Paso 2: Procesar el resto con la lista dinámica
    hist = procesar_historial_empalme(archivo_previo, lista_actual)
    sug = procesar_sugerencias(link_principal, lista_actual)
    conf = procesar_configuracion(link_principal, lista_actual)
    
    # Paso 3: Motor
    res = generar_cuadro_equitativo(mes_sel, ano_sel, hist, sug, conf, semilla, lista_actual)
    
    dias_en_mes = calendar.monthrange(ano_sel, mes_sel)[1]
    st.dataframe(res.style.map(aplicar_colores, subset=[str(d) for d in range(1, dias_en_mes + 1)]), use_container_width=True)
    
    # Exportación Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        ws = wb.add_worksheet('Turnos')
        ws.freeze_panes(1, 1)
        base_fmt = {'border': 1, 'align': 'center', 'valign': 'vcenter'}
        fmt_v = wb.add_format({**base_fmt, 'bg_color': '#d9ead3'}); fmt_r = wb.add_format({**base_fmt, 'bg_color': '#f4cccc'})
        fmt_az = wb.add_format({**base_fmt, 'bg_color': '#cfe2f3'}); fmt_am = wb.add_format({**base_fmt, 'bg_color': '#fff2cc'})
        fmt_default = wb.add_format(base_fmt)
        fmt_headers = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#e0e0e0', 'align': 'center'})
        fmt_names = wb.add_format({'bold': True, 'border': 1, 'align': 'left'})
        
        ws.set_column(0, 0, 25); ws.set_column(1, dias_en_mes, 4); ws.set_column(dias_en_mes+1, dias_en_mes+3, 16)
        ws.write(0, 0, "INTEGRANTES", fmt_headers)
        for col_num, col_name in enumerate(res.columns): ws.write(0, col_num + 1, str(col_name), fmt_headers)
        for r_idx, (idx, row) in enumerate(res.iterrows()):
            ws.write(r_idx + 1, 0, str(idx), fmt_names)
            for c_idx, val in enumerate(row):
                f = fmt_default
                if val in ['L', 'D']: f = fmt_v
                elif val == 'P': f = fmt_r
                elif 'N' in str(val): f = fmt_az
                elif 'C' in str(val): f = fmt_am
                ws.write(r_idx + 1, c_idx + 1, val, f)
    
    st.download_button("📥 Descargar Excel", output.getvalue(), f"Turnos_{mes_sel}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
