import streamlit as st
import pandas as pd
from datetime import datetime
import calendar
import random
import io
import holidays
import re
import xlsxwriter

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Instrumentación", layout="wide")

# --- MAGIA ANTI-ERRORES ---
# Este código oculta la "flechita tramposa" de Streamlit para obligar al usuario a usar el botón gigante.
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

# --- LECTORES CON CACHÉ (RÁPIDOS) ---
@st.cache_data(ttl=60)
def procesar_historial_empalme(file):
    historial = {p: ["", "", ""] for p in INTEGRANTES} 
    if not file: return historial
    try:
        df = pd.read_excel(file) if file.name.endswith('.xlsx') else pd.read_csv(file)
        if not any(str(c).isdigit() for c in df.columns):
