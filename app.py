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

# Diccionario inteligente para detectar qué día escriben en el Excel
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
        # Busca columnas que digan DIA, LIBRE o FIJO
        col_dias = next((c for c in df_conf.columns if 'DIA' in c or 'LIBRE' in c or 'FIJO' in c), None)
        col_vac = next((c for c in df_conf.columns if 'VACACION' in c), None)
        
        if col_nom:
            for _, r in df_conf.iterrows():
                nom = normalizar_texto(r[col_nom])
                if "GINELAP" in nom: nom = "GINELAP"
                
                if nom in INTEGRANTES:
                    # 1. Leer libres fijos semanales (Ahora detecta texto como "LUNES" o números "1")
                    if col_dias and str(r[col_dias]).lower() not in ['nan', 'none', '']:
                        dias_str = normalizar_texto(r[col_dias])
                        dias_asignados = []
                        
                        # Buscar si escribieron la palabra (Ej: "LUNES")
                        for nombre_dia, num_dia in DIAS_MAP.items():
                            if nombre_dia in dias_str:
                                dias_asignados.append(num_dia)
                        
                        # Si no escribieron palabras, revisar si escribieron números (1=Lunes, 7=Domingo)
                        if not dias_asignados:
                            nums = [int(x) for x in re.findall(r'\d+', dias_str)]
                            for n in nums:
                                if n == 7: dias_asignados.append(6) # Domingo
                                elif 1 <= n <= 6: dias_asignados.append(n - 1) # Lunes a Sabado
                                elif n == 0: dias_asignados.append(0) # Por si alguien sabe usar el cero
                        
                        libres_fijos[nom] = list(set(dias_asignados))
                    
                    # 2. Leer vacaciones
                    if col_vac and str(r[col_vac]).lower() not in ['nan', 'none', '']:
                        vac_str = str(r[col_vac])
                        rango = vac_str.split('-')
                        if len(rango) == 2 and rango[0].strip().isdigit() and rango[1].strip().isdigit():
                            inicio, fin = int(rango[0].strip()), int(rango[1].strip())
                            vacaciones[nom] = list(range(inicio, fin + 1))
                        else:
                            vacaciones[nom].extend([int(x) for x in re.findall(r'\d+', vac_str)])
                            
    except Exception as e: 
        st.sidebar.warning(f"Error leyendo la Configuración: {e}")
    return libres_fijos, vacaciones

# --- MOTOR ---
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
            
        def arruina_corrido_futuro(persona, dia_actual):
            if racha_actual(persona) == 2 and dia_actual + 1 <= dias_mes:
                s_man = sugerencias_dict.get(persona, {}).get(str(dia_actual + 1))
                if s_man and any(t in s_man for t in ['C', 'N']): return True 
            return False

        cuota_n = 2
        if es_finde_o_festivo and (wd == 6 or es_festivo(d, mes, ano)): cuota_c = 2
        elif wd == 5: cuota_c = 5
        else: cuota_c = 6

        # 0. CANDADO ABSOLUTO: VACACIONES
        for p in INTEGRANTES:
            if d in vacaciones_dict.get(p, []):
                df.at[p, ds] = 'V'

        # 1. POSTURNOS
        for p in INTEGRANTES:
            if 'N' in turno_en_dia(p, d-1) and df.at[p, ds] == "": 
                df.at[p, ds] = 'P'

        # CANDADO ACTIVO: Forzar Descanso
        for p in INTEGRANTES:
            if df.at[p, ds] == "" and necesita_descanso(p):
                df.at[p, ds] = 'D'

        # 2. ASIGNACIONES FIJAS
        if wd == 1 and df.at["JHON RIOS", ds] == "":
            df.at["JHON RIOS", ds] = "N"
            cuota_n -= 1; turnos_totales["JHON RIOS"] += 1; noches_totales["JHON RIOS"] += 1
            if es_finde_o_festivo: finde_totales["JHON RIOS"] += 1
            
        if wd == 2 and df.at["ANGIE BERNAL", ds] == "":
            df.at["ANGIE BERNAL", ds] = "C"
            cuota_c -= 1; turnos_totales["ANGIE BERNAL"] += 1
            if es_finde_o_festivo: finde_totales["ANGIE BERNAL"] += 1

        # 3. SUGERENCIAS
        for p in INTEGRANTES:
            if df.at[p, ds] != "": continue
            req = sugerencias_dict.get(p, {}).get(ds)
            if req:
                tl = 'C' if ('C' in req and 'P' not in req) else ('N' if ('N' in req and 'P' not in req) else req)
                df.at[p, ds] = tl
                if tl == 'N': cuota_n -= 1; noches_totales[p] += 1; turnos_totales[p] += 1
                if tl == 'C': cuota_c -= 1; turnos_totales[p] += 1
                if tl in ['C', 'N'] and es_finde_o_festivo: finde_totales[p] += 1

        # 4. LIBRES FIJOS (CORREGIDO EL BLOQUEO A JUAN CAMILO)
        for p in INTEGRANTES:
            if df.at[p, ds] == "" and wd in config_dict.get(p, []):
                df.at[p, ds] = "L"

        # 5. REPARTIR NOCHES
        for _ in range(max(0, cuota_n)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            disp = [p for p in disp if 'N' not in turno_en_dia(p, d-2)]
            
            disp_segura = [p for p in disp if not arruina_noche_futura(p, d)]
            
            if d < dias_mes:
                m_wd = (wd + 1) % 7
                disp_segura = [p for p in disp_segura if m_wd not in config_dict.get(p, [])]

            if not disp_segura: 
                disp_segura = [p for p in disp if m_wd not in config_dict.get(p, [])] if d < dias_mes else disp
                if not disp_segura: disp_segura = disp 
            
            if disp_segura:
                rng.shuffle(disp_segura)
                if es_finde_o_festivo: disp_segura.sort(key=lambda x: (noches_totales[x], finde_totales[x], turnos_totales[x]))
                else: disp_segura.sort(key=lambda x: (noches_totales[x], turnos_totales[x]))
                
                elegido = disp_segura[0]; df.at[elegido, ds] = "N"
                turnos_totales[elegido] += 1; noches_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 6. REPARTIR CORRIDOS
        for _ in range(max(0, cuota_c)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            
            disp_segura = [p for p in disp if not arruina_corrido_futuro(p, d)]
            if not disp_segura: disp_segura = disp
            
            if disp_segura:
                rng.shuffle(disp_segura)
                if es_finde_o_festivo: disp_segura.sort(key=lambda x: (finde_totales[x], turnos_totales[x]))
                else: disp_segura.sort(key=lambda x: (turnos_totales[x], racha_actual(x)))
                
                elegido = disp_segura[0]; df.at[elegido, ds] = "C"; turnos_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 7. RELLENAR CON DESCANSO
        for p in INTEGRANTES:
            if df.at[p, ds] == "": df.at[p, ds] = "D"

    df['TOTAL CORRIDOS'] = df.apply(lambda r: sum(1 for c in df.columns if 'C' in str(r[c])), axis=1)
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
    link_config = st.text_input("Link Configuración:", "")
    ano_sel = st.number_input("Año:", min_value=2024, value=datetime.now().year)
    mes_sel = st.selectbox("Mes:", range(1, 13), index=datetime.now().month-1)
    semilla = st.number_input("Semilla:", min_value=0, value=42)

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
