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
st.set_page_config(page_title="Gestor de Instrumentación Pro", layout="wide")

# --- BLOQUEO VISUAL (OCULTAR HERRAMIENTAS NATIVAS) ---
st.markdown(
    """
    <style>
    [data-testid="stElementToolbar"] {display: none;}
    </style>
    """,
    unsafe_allow_html=True
)

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

# --- LECTORES CON CACHÉ ---
@st.cache_data(ttl=60)
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
    except: pass
    return historial

@st.cache_data(ttl=60)
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
    except: pass
    return sugerencias

@st.cache_data(ttl=60)
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
    except: pass
    return libres_fijos

# --- MOTOR LÓGICO CON SEMILLA ---
def generar_cuadro_equitativo(mes, ano, historial_previo, sugerencias_dict, config_dict, semilla):
    # LA CLAVE: Fijar la aleatoriedad para repetibilidad
    random.seed(semilla)
    
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

        def arruina_noche_futura(persona, dia_actual):
            if dia_actual + 1 <= dias_mes:
                s_man = sugerencias_dict.get(persona, {}).get(str(dia_actual + 1))
                if s_man and any(t in s_man for t in ['C', 'N', 'L']): return True 
            if dia_actual + 2 <= dias_mes:
                s_pas = sugerencias_dict.get(persona, {}).get(str(dia_actual + 2))
                if s_pas and 'N' in s_pas: return True 
            return False

        cuota_n = 2
        cuota_c = 2 if (wd == 6 or es_festivo(d, mes, ano)) else (5 if wd == 5 else 6)

        # 1. POSTURNOS
        for p in INTEGRANTES:
            if 'N' in turno_en_dia(p, d-1): df.at[p, ds] = 'P'

        # 2. DESCANSOS OBLIGATORIOS (Antifatiga)
        for p in INTEGRANTES:
            if df.at[p, ds] == "" and necesita_descanso(p): df.at[p, ds] = 'D'

        # 3. REGLAS FIJAS (Jhon y Angie)
        if wd == 1 and df.at["JHON RIOS", ds] == "":
            df.at["JHON RIOS", ds] = "N"; cuota_n -= 1; turnos_totales["JHON RIOS"] += 1; noches_totales["JHON RIOS"] += 1
            if es_finde_o_festivo: finde_totales["JHON RIOS"] += 1
        if wd == 2 and df.at["ANGIE BERNAL", ds] == "":
            df.at["ANGIE BERNAL", ds] = "C"; cuota_c -= 1; turnos_totales["ANGIE BERNAL"] += 1
            if es_finde_o_festivo: finde_totales["ANGIE BERNAL"] += 1

        # 4. SUGERENCIAS Y LIBRES FIJOS
        for p in INTEGRANTES:
            if df.at[p, ds] != "": continue
            req = sugerencias_dict.get(p, {}).get(ds)
            if req:
                tl = 'C' if ('C' in req and 'P' not in req) else ('N' if ('N' in req and 'P' not in req) else req)
                df.at[p, ds] = tl
                if tl == 'N': cuota_n -= 1; noches_totales[p] += 1; turnos_totales[p] += 1
                if tl == 'C': cuota_c -= 1; turnos_totales[p] += 1
                if tl in ['C', 'N'] and es_finde_o_festivo: finde_totales[p] += 1
            elif wd in config_dict.get(p, []) and not (p == "JUAN CAMILO PEREZ" and wd == 3):
                df.at[p, ds] = "L"

        # 5. REPARTIR NOCHES (Priorizando equidad)
        for _ in range(max(0, cuota_n)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == "" and 'N' not in turno_en_dia(p, d-2) and not arruina_noche_futura(p, d)]
            if not disp: disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if disp:
                random.shuffle(disp)
                disp.sort(key=lambda x: (noches_totales[x], finde_totales[x], turnos_totales[x]))
                elegido = disp[0]; df.at[elegido, ds] = "N"; turnos_totales[elegido] += 1; noches_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 6. REPARTIR CORRIDOS
        for _ in range(max(0, cuota_c)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if disp:
                random.shuffle(disp)
                disp.sort(key=lambda x: (finde_totales[x], turnos_totales[x]))
                elegido = disp[0]; df.at[elegido, ds] = "C"; turnos_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 7. RELLENAR
        for p in INTEGRANTES:
            if df.at[p, ds] == "": df.at[p, ds] = "D"

    # Estadísticas finales
    df['TOTAL TURNOS'] = df.apply(lambda r: sum(1 for c in df.columns if any(t in str(r[c]) for t in ['C', 'N'])), axis=1)
    df['TOTAL NOCHES'] = df.apply(lambda r: noches_totales[r.name], axis=1)
    df['FINES DE SEMANA'] = df.apply(lambda r: finde_totales[r.name], axis=1)
    return df

# --- INTERFAZ USUARIO ---
with st.sidebar:
    st.header("Configuración")
    archivo_previo = st.file_uploader("Excel Mes Anterior:", type=['xlsx', 'csv'])
    link_sheet = st.text_input("Link Sugerencias:", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?pli=1&gid=0#gid=0")
    link_config = st.text_input("Link Días Libres:", "")
    ano_sel = st.number_input("Año:", min_value=2024, value=datetime.now().year)
    mes_sel = st.selectbox("Mes:", range(1, 13), index=datetime.now().month-1)
    
    st.markdown("---")
    semilla = st.number_input("🎲 Semilla (Código de Cuadro):", value=42, help="Usa el mismo número para repetir el mismo cuadro exacto.")

st.title("🏥 Motor de Turnos Blindado")

if st.button("🚀 GENERAR CUADRO", type="primary", use_container_width=True):
    hist = procesar_historial_empalme(archivo_previo)
    sug = procesar_sugerencias(link_sheet); conf = procesar_configuracion(link_config)
    
    # Llamada al motor pasando la SEMILLA
    res = generar_cuadro_equitativo(mes_sel, ano_sel, hist, sug, conf, semilla)
    
    dias_en_mes = calendar.monthrange(ano_sel, mes_sel)[1]
    
    # Visualización con colores
    st.dataframe(res.style.map(aplicar_colores, subset=[str(d) for d in range(1, dias_en_mes + 1)]), use_container_width=True)
    
    # Exportación a Excel con xlsxwriter
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        ws = wb.add_worksheet('Turnos')
        ws.freeze_panes(1, 1)
        
        # Formatos
        base_fmt = {'border': 1, 'align': 'center', 'valign': 'vcenter'}
        fmt_v = wb.add_format({**base_fmt, 'bg_color': '#d9ead3'}) 
        fmt_r = wb.add_format({**base_fmt, 'bg_color': '#f4cccc'}) 
        fmt_az = wb.add_format({**base_fmt, 'bg_color': '#cfe2f3'}) 
        fmt_am = wb.add_format({**base_fmt, 'bg_color': '#fff2cc'}) 
        fmt_default = wb.add_format(base_fmt)
        fmt_headers = wb.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#e0e0e0'})
        fmt_names = wb.add_format({'bold': True, 'border': 1, 'align': 'left'})
        
        ws.set_column(0, 0, 25); ws.set_column(1, dias_en_mes, 4); ws.set_column(dias_en_mes + 1, dias_en_mes + 3, 16)
        
        ws.write(0, 0, "INTEGRANTES", fmt_headers)
        for col_num, col_name in enumerate(res.columns):
            ws.write(0, col_num + 1, str(col_name), fmt_headers)
            
        for r_idx, (idx, row) in enumerate(res.iterrows()):
            ws.write(r_idx + 1, 0, str(idx), fmt_names)
            for c_idx, val in enumerate(row):
                f = fmt_default
                if val in ['L', 'D']: f = fmt_v
                elif val == 'P': f = fmt_r
                elif 'N' in str(val): f = fmt_az
                elif 'C' in str(val): f = fmt_am
                ws.write(r_idx + 1, c_idx + 1, val, f)
    
    st.download_button(
        label="📥 Descargar Excel", 
        data=output.getvalue(), 
        file_name=f"Turnos_{mes_sel}_Semilla_{semilla}.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
