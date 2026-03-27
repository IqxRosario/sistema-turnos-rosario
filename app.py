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
    'LUN': 0, 'LUNES': 0, 'MAR': 1, 'MARTES': 1, 'MIE': 2, 'MIERCOLES': 2, 
    'JUE': 3, 'JUEVES': 3, 'VIE': 4, 'VIERNES': 4, 'SAB': 5, 'SABADO': 5, 'DOM': 6, 'DOMINGO': 6
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
        csv_link = f"{base
