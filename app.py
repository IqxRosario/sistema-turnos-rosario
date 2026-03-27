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
    'LUN': 0, 'LUNES': 0, 
    'MAR': 1, 'MARTES': 1, 
    'MIE': 2, 'MIERCOLES': 2, 
    'JUE': 3, 'JUEVES': 3, 
    'VIE': 4, 'VIERNES': 4, 
    'SAB': 5, 'SABADO': 5, 
    'DOM': 6, 'DOMINGO': 6
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
    except Exception as e: 
        st.sidebar.error(f"Error en historial: {e}")
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
                if "GINELAP" in nom: nom = "GINELAP"
                fecha = ''.join(filter(str.isdigit, str(r[col_fec])))
                sol = normalizar_texto(r[col_sol])
                if nom in INTEGRANTES and fecha and sol not in ['NAN', 'NONE', '']: 
                    sugerencias[nom][fecha] = sol
    except: 
        st.sidebar.warning("Link Sugerencias no detectado o error de lectura.")
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
        col_dias = next((c for c in df_conf.columns if 'DIA' in c or 'LIBRE' in c or 'FIJO' in c), None)
        col_vac = next((c for c in df_conf.columns if 'VACACION' in c or 'VACA' in c), None)
        
        if col_nom:
            for _, r in df_conf.iterrows():
                nom_sheet = normalizar_texto(r[col_nom])
                
                # BUSCADOR INTELIGENTE DE NOMBRES
                nom_real = None
                for p in INTEGRANTES:
                    # Si el nombre del excel está dentro del oficial (Ej: "CAMILO" en "JUAN CAMILO PEREZ")
                    if nom_sheet in p or p in nom_sheet: 
                        nom_real = p
                        break
                if "GINELAP" in nom_sheet: nom_real = "GINELAP"
                
                if nom_real:
                    # 1. Leer libres fijos
                    if col_dias and str(r[col_dias]).lower() not in ['nan', 'none', '']:
                        dias_str = normalizar_texto(r[col_dias])
                        dias_asignados = []
                        
                        for nombre_dia, num_dia in DIAS_MAP.items():
                            if nombre_dia in dias_str:
                                dias_asignados.append(num_dia)
                        
                        if not dias_asignados:
                            nums = [int(x) for x in re.findall(r'\d+', dias_str)]
                            for n in nums:
                                if n == 7: dias_asignados.append(6)
                                elif 1 <= n <= 6: dias_asignados.append(n - 1)
                                elif n == 0: dias_asignados.append(0)
                        
                        libres_fijos[nom_real] = list(set(dias_asignados))
                    
                    # 2. Leer vacaciones
                    if col_vac and str(r[col_vac]).lower() not in ['nan', 'none', '']:
                        vac_str = str(r[col_vac])
                        rango = vac_str.split('-')
                        if len(rango) == 2 and rango[0].strip().isdigit() and rango[1].strip().isdigit():
                            inicio, fin = int(rango[0].strip()), int(rango[1].strip())
                            vacaciones[nom_real] = list(range(inicio, fin + 1))
                        else:
                            vacaciones[nom_real].extend([int(x) for x in re.findall(r'\d+', vac_str)])
                            
    except Exception as e: 
        st.sidebar.error(f"Error CRÍTICO leyendo la Configuración: {e}. Revisa el link y permisos.")
    return libres_fijos, vacaciones
    
# --- MOTOR CORREGIDO Y SIN ERRORES DE INDENTACIÓN ---
def generar_cuadro_equitativo(mes, ano, historial_previo, sugerencias_dict, config_dict, vacaciones_dict, semilla):
    rng = random.Random(semilla)
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

    def racha_actual(persona, d):
        streak = 0
        for past_d in range(d - 1, d - 10, -1):
            t_p = turno_en_dia(persona, past_d)
            if any(t in t_p for t in ['C', 'N']) and 'P' not in t_p: streak += 1
            else: break
        return streak

    # --- FUNCIÓN DE BLOQUEO DE NOCHES (Corregida e Identada) ---
    def no_puede_hacer_noche(persona, dia_actual):
        if dia_actual < dias_mes:
            # Calculamos qué día de la semana es mañana
            manana_dt = datetime(ano, mes, dia_actual + 1)
            wd_manana = manana_dt.weekday()
            turno_manana = str(df.at[persona, str(dia_actual + 1)])
            
            # Bloqueo por Vacaciones o Posturno (Aplica a todos)
            if any(x in turno_manana for x in ['V', 'P']): return True
            
            # Bloqueo por Libre Fijo
            if 'L' in turno_manana:
                # EXCEPCIÓN: Si mañana es el Jueves flexible de Camilo (wd 3), 
                # permitimos que haga Noche hoy Miércoles.
                if persona == "JUAN CAMILO PEREZ" and wd_manana == 3:
                    return False
                return True
        return False

    # --- FASE 1: PRE-LLENADO ---
    for d in range(1, dias_mes + 1):
        ds = str(d); wd = datetime(ano, mes, d).weekday()
        es_finde_o_festivo = wd >= 5 or es_festivo(d, mes, ano)

        for p in INTEGRANTES:
            if d in vacaciones_dict.get(p, []):
                df.at[p, ds] = 'V'
                continue

            req = sugerencias_dict.get(p, {}).get(ds)

         # 3. Libres Fijos (Jerarquía con excepción para Camilo)
            if wd in config_dict.get(p, []):
                if p == "JUAN CAMILO PEREZ":
                    # EL MARTES (wd 1) ES SAGRADO: Se pone 'L' de una vez
                    if wd == 1:
                        df.at[p, ds] = "L"
                        continue
                    # EL JUEVES (wd 3) ES FLEXIBLE: No ponemos 'L' todavía 
                    # para que la Fase 2 pueda asignarle N, C o P si es necesario.
                    elif wd == 3:
                        if req: # Si hay sugerencia expresa para el jueves
                            tl = 'C' if ('C' in req and 'P' not in req) else ('N' if ('N' in req and 'P' not in req) else req)
                            df.at[p, ds] = tl
                            if tl == 'N': noches_totales[p] += 1; turnos_totales[p] += 1
                            if tl == 'C': turnos_totales[p] += 1
                        continue
                else:
                    # Regla normal para el resto del equipo
                    df.at[p, ds] = "L"
                    continue
    
            # Sugerencias (Si no hay libre fijo)
            if req:
                tl = 'C' if ('C' in req and 'P' not in req) else ('N' if ('N' in req and 'P' not in req) else req)
                df.at[p, ds] = tl
                if tl == 'N': noches_totales[p] += 1; turnos_totales[p] += 1
                if tl == 'C': turnos_totales[p] += 1
                if tl in ['C', 'N'] and es_finde_o_festivo: finde_totales[p] += 1
                continue

            # Asignaciones Fijas Jhon/Angie
            if wd == 1 and p == "JHON RIOS" and df.at[p, ds] == "":
                df.at[p, ds] = "N"
                turnos_totales[p] += 1; noches_totales[p] += 1
            if wd == 2 and p == "ANGIE BERNAL" and df.at[p, ds] == "":
                df.at[p, ds] = "C"
                turnos_totales[p] += 1

    # --- FASE 2: REPARTICIÓN ---
    for d in range(1, dias_mes + 1):
        ds = str(d); wd = datetime(ano, mes, d).weekday()
        es_finde_o_festivo = wd >= 5 or es_festivo(d, mes, ano)

        cuota_n = 2 - sum(1 for p in INTEGRANTES if 'N' in df.at[p, ds])
        cuota_c_base = 2 if (es_finde_o_festivo and (wd == 6 or es_festivo(d, mes, ano))) else (5 if wd == 5 else 6)
        cuota_c = cuota_c_base - sum(1 for p in INTEGRANTES if 'C' in df.at[p, ds])

        for p in INTEGRANTES:
            if 'N' in turno_en_dia(p, d-1) and df.at[p, ds] == "": df.at[p, ds] = 'P'
            if df.at[p, ds] == "" and racha_actual(p, d) >= 3: df.at[p, ds] = 'D'

        # Repartir Noches
        for _ in range(max(0, cuota_n)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == "" and 'N' not in turno_en_dia(p, d-2) and not no_puede_hacer_noche(p, d)]
            if disp:
                rng.shuffle(disp)
                disp.sort(key=lambda x: (noches_totales[x], turnos_totales[x]))
                elegido = disp[0]
                df.at[elegido, ds] = "N"
                turnos_totales[elegido] += 1; noches_totales[elegido] += 1

        # Repartir Corridos
        for _ in range(max(0, cuota_c)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if disp:
                rng.shuffle(disp)
                disp.sort(key=lambda x: (turnos_totales[x], racha_actual(x, d)))
                elegido = disp[0]
                df.at[elegido, ds] = "C"; turnos_totales[elegido] += 1

        for p in INTEGRANTES:
            if df.at[p, ds] == "": df.at[p, ds] = "D"

    df['TOTAL CORRIDOS'] = df.apply(lambda r: sum(1 for c in df.columns if 'C' in str(r[c]) and c.isdigit()), axis=1)
    df['TOTAL NOCHES'] = df.apply(lambda r: noches_totales[r.name], axis=1)
    df['TOTAL TURNOS'] = df['TOTAL CORRIDOS'] + df['TOTAL NOCHES']
    df['FINES DE SEMANA'] = df.apply(lambda r: finde_totales[r.name], axis=1)
    return df

# --- INTERFAZ ---
st.title("🏥 Gestor de Turnos (Rápido y Blindado)")

with st.sidebar:
    st.header("1. Cargar Datos")
    archivo_previo = st.file_uploader("Excel Mes Anterior:", type=['xlsx', 'csv'])
    link_sheet = st.text_input("Link Sugerencias:", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?pli=1&gid=0#gid=0")
    if link_sheet.startswith("http"): st.link_button("📝 Abrir Sugerencias", link_sheet, use_container_width=True)
    link_config = st.text_input("Link Configuración:", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?pli=1&gid=1679804429#gid=1679804429")
    if link_sheet.startswith("http"): st.link_button("📝 Abrir Sugerencias", link_sheet, use_container_width=True)
    ano_sel = st.number_input("Año:", min_value=2024, value=datetime.now().year)
    mes_sel = st.selectbox("Mes:", range(1, 13), index=datetime.now().month-1)
    semilla = st.number_input("Semilla:", min_value=0, value=0)

if st.button("🚀 GENERAR CUADRO", type="primary", use_container_width=True):
    hist = procesar_historial_empalme(archivo_previo)
    sug = procesar_sugerencias(link_sheet)
    conf, vacs = procesar_configuracion(link_config)
    
    res = generar_cuadro_equitativo(mes_sel, ano_sel, hist, sug, conf, vacs, semilla)
    dias_en_mes = calendar.monthrange(ano_sel, mes_sel)[1]
    
    st.dataframe(res.style.map(aplicar_colores, subset=[str(d) for d in range(1, dias_en_mes + 1)]), use_container_width=True)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        ws = wb.add_worksheet('Turnos')
        ws.freeze_panes(1, 1) 
        
        base_fmt = {'border': 1, 'align': 'center', 'valign': 'vcenter'}
        fmt_v = wb.add_format({**base_fmt, 'bg_color': '#d9ead3'}) 
        fmt_r = wb.add_format({**base_fmt, 'bg_color': '#f4cccc'}) 
        fmt_az = wb.add_format({**base_fmt, 'bg_color': '#cfe2f3'}) 
        fmt_am = wb.add_format({**base_fmt, 'bg_color': '#fff2cc'}) 
        fmt_vac = wb.add_format({**base_fmt, 'bg_color': '#e4d7f5'}) 
        fmt_default = wb.add_format(base_fmt)
        fmt_headers = wb.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#e0e0e0'})
        fmt_names = wb.add_format({'bold': True, 'border': 1, 'align': 'left', 'valign': 'vcenter'})
        
        ws.set_column(0, 0, 25) 
        ws.set_column(1, dias_en_mes, 4) 
        ws.set_column(dias_en_mes + 1, dias_en_mes + 4, 16) 
        
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
                elif val == 'V': f = fmt_vac 
                ws.write(r_idx + 1, c_idx + 1, val, f)
    
    st.download_button(
        label="📥 Descargar Excel", 
        data=output.getvalue(), 
        file_name=f"Turnos_{mes_sel}.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
