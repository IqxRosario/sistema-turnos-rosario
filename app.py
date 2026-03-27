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
                if "GINELAP" in nom: nom = "GINELAP"
                fecha = ''.join(filter(str.isdigit, str(r[col_fec])))
                sol = normalizar_texto(r[col_sol])
                if nom in INTEGRANTES and fecha and sol not in ['NAN', 'NONE', '']: 
                    sugerencias[nom][fecha] = sol
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
                if "GINELAP" in nom_sheet: nom_real = "GINELAP"
                if nom_real:
                    if col_dias and str(r[col_dias]).lower() not in ['nan', 'none', '']:
                        dias_str = normalizar_texto(r[col_dias])
                        dias_asignados = [num for dia, num in DIAS_MAP.items() if dia in dias_str]
                        if not dias_asignados:
                            nums = [int(x) for x in re.findall(r'\d+', dias_str)]
                            for n in nums:
                                if n == 7: dias_asignados.append(6)
                                elif 1 <= n <= 6: dias_asignados.append(n - 1)
                        libres_fijos[nom_real] = list(set(dias_asignados))
                    if col_vac and str(r[col_vac]).lower() not in ['nan', 'none', '']:
                        vac_str = str(r[col_vac]).replace(' Y ', ',').replace('&', ',')
                        for parte in vac_str.split(','):
                            parte = parte.strip()
                            if '-' in parte:
                                rango = parte.split('-')
                                if len(rango) == 2: vacaciones[nom_real].extend(list(range(int(rango[0]), int(rango[1]) + 1)))
                            elif parte.isdigit(): vacaciones[nom_real].append(int(parte))
    except: pass
    return libres_fijos, vacaciones

# --- MOTOR ---
def generar_cuadro_equitativo(mes, ano, historial_previo, sugerencias_dict, config_dict, vacaciones_dict, semilla):
    rng = random.Random(semilla)
    dias_mes = calendar.monthrange(ano, mes)[1]
    df = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    turnos_totales, noches_totales = {p: 0 for p in INTEGRANTES}, {p: 0 for p in INTEGRANTES}

    def turno_en_dia(persona, dia_req):
        if dia_req > 0: return str(df.at[persona, str(dia_req)])
        idx = 2 + dia_req
        hist = historial_previo.get(persona, ["", "", ""])
        return str(hist[idx]) if 0 <= idx < 3 else ""

    def racha_actual(persona, d):
        streak = 0
        for past_d in range(d - 1, d - 10, -1):
            t_p = turno_en_dia(persona, past_d)
            if any(t in t_p for t in ['C', 'N']) and 'P' not in t_p: streak += 1
            else: break
        return streak

    def no_puede_hacer_noche(persona, dia_actual):
        if dia_actual < dias_mes:
            manana_dt = datetime(ano, mes, dia_actual + 1)
            wd_manana = manana_dt.weekday()
            turno_manana = str(df.at[persona, str(dia_actual + 1)])
            if any(x in turno_manana for x in ['V', 'P']): return True
            if 'L' in turno_manana:
                if persona == "JUAN CAMILO PEREZ" and wd_manana == 3: return False # Jueves flexible
                return True
        return False

    # FASE 1: PRE-LLENADO
    for d in range(1, dias_mes + 1):
        ds, wd = str(d), datetime(ano, mes, d).weekday()
        for p in INTEGRANTES:
            if d in vacaciones_dict.get(p, []):
                df.at[p, ds] = 'V'
                continue
            
            req = sugerencias_dict.get(p, {}).get(ds)
            # Prioridad Libres Fijos
            if wd in config_dict.get(p, []):
                if p == "JUAN CAMILO PEREZ":
                    if wd == 1: # Martes sagrado
                        df.at[p, ds] = "L"
                        continue
                    elif wd == 3: # Jueves flexible
                        if req:
                            tl = 'C' if ('C' in req and 'P' not in req) else ('N' if ('N' in req and 'P' not in req) else req)
                            df.at[p, ds] = tl
                            if tl == 'N': noches_totales[p] += 1; turnos_totales[p] += 1
                            if tl == 'C': turnos_totales[p] += 1
                        continue
                else:
                    df.at[p, ds] = "L"; continue

            if req:
                tl = 'C' if ('C' in req and 'P' not in req) else ('N' if ('N' in req and 'P' not in req) else req)
                df.at[p, ds] = tl
                if tl == 'N': noches_totales[p] += 1; turnos_totales[p] += 1
                if tl == 'C': turnos_totales[p] += 1
                continue

            if wd == 1 and p == "JHON RIOS" and df.at[p, ds] == "":
                df.at[p, ds] = "N"; turnos_totales[p] += 1; noches_totales[p] += 1
            if wd == 2 and p == "ANGIE BERNAL" and df.at[p, ds] == "":
                df.at[p, ds] = "C"; turnos_totales[p] += 1

    # FASE 2: REPARTICIÓN
    for d in range(1, dias_mes + 1):
        ds, wd = str(d), datetime(ano, mes, d).weekday()
        es_festejo = wd >= 5 or es_festivo(d, mes, ano)
        cuota_n = 2 - sum(1 for p in INTEGRANTES if 'N' in df.at[p, ds])
        cuota_c_base = 2 if (es_festejo and (wd == 6 or es_festivo(d, mes, ano))) else (5 if wd == 5 else 6)
        cuota_c = cuota_c_base - sum(1 for p in INTEGRANTES if 'C' in df.at[p, ds])

        for p in INTEGRANTES:
            if 'N' in turno_en_dia(p, d-1) and df.at[p, ds] == "": df.at[p, ds] = 'P'
            if df.at[p, ds] == "" and racha_actual(p, d) >= 3: df.at[p, ds] = 'D'

        for _ in range(max(0, cuota_n)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == "" and 'N' not in turno_en_dia(p, d-2) and not no_puede_hacer_noche(p, d)]
            if disp:
                rng.shuffle(disp)
                disp.sort(key=lambda x: (noches_totales[x], turnos_totales[x]))
                df.at[disp[0], ds] = "N"; turnos_totales[disp[0]] += 1; noches_totales[disp[0]] += 1

        for _ in range(max(0, cuota_c)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            if disp:
