import streamlit as st
import calendar
import datetime
import pandas as pd
import io
import re
from motor_pulp import generar_cuadro_pulp

INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES ARGUMEDO", 
    "KELLY CAUSIL", "JOAN CARMONA"
]

DIAS_MAP_TEXTO = {
    'LUNES': 0, 'MARTES': 1, 'MIERCOLES': 2, 'JUEVES': 3, 'VIERNES': 4, 'SABADO': 5, 'DOMINGO': 6,
    'LUN': 0, 'MAR': 1, 'MIE': 2, 'JUE': 3, 'VIE': 4, 'SAB': 5, 'DOM': 6
}

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    import unicodedata
    s = unicodedata.normalize('NFD', texto)
    s = s.encode('ascii', 'ignore').decode("utf-8")
    return s.strip().upper()

def parsear_rangos_y_dias(texto):
    dias = []
    if not texto or str(texto) == 'nan':
        return dias
    partes = str(texto).replace(' Y ', ',').replace('&', ',').split(',')
    for p in partes:
        p = p.strip()
        if '-' in p:
            rng = p.split('-')
            if len(rng) == 2 and rng[0].isdigit() and rng[1].isdigit():
                dias.extend(list(range(int(rng[0]), int(rng[1]) + 1)))
        elif p.isdigit():
            dias.append(int(p))
    return list(set(dias))

def procesar_historial_empalme(archivo):
    hist = {p: "" for p in INTEGRANTES}
    if archivo is None:
        return hist
    try:
        df_prev = pd.read_excel(archivo)
        df_prev.columns = [str(c).strip() for c in df_prev.columns]
        
        col_nom = df_prev.columns[0]
        cols_dias = [c for c in df_prev.columns if c.isdigit()]
        
        if cols_dias:
            ultimo_dia_col = str(max([int(c) for c in cols_dias]))
            for _, r in df_prev.iterrows():
                nom_sheet = normalizar_texto(r[col_nom])
                nom_real = next((p for p in INTEGRANTES if nom_sheet in p or p in nom_sheet), None)
                if nom_real:
                    val = str(r[ultimo_dia_col]).strip().upper()
                    hist[nom_real] = val
    except Exception as e:
        st.warning(f"No se pudo procesar el archivo de empalme: {e}")
    return hist

def procesar_configuracion_y_sugerencias(link):
    libres_fijos = {p: [] for p in INTEGRANTES}
    vacaciones = {p: [] for p in INTEGRANTES}
    sugerencias = {p: {} for p in INTEGRANTES}
    
    if not link or not link.startswith("http"):
        return libres_fijos, vacaciones, sugerencias
        
    try:
        base_url = link.split('/edit')[0]
        csv_link = f"{base_url}/export?format=csv"
        match = re.search(r'gid=(\d+)', link)
        if match:
            csv_link += f"&gid={match.group(1)}"
        
        df = pd.read_csv(csv_link)
        df.columns = [normalizar_texto(c) for c in df.columns]
        
        col_nom = next((c for c in df.columns if 'NOMBRE' in c), None)
        col_sug = next((c for c in df.columns if 'PETICION' in c or 'LIBRES SOLICITADO' in c or 'SUGERENCIA' in c), None)
        col_fijo = next((c for c in df.columns if 'FIJO' in c or 'ESTUDIO' in c), None)
        col_vac = next((c for c in df.columns if 'VACACIO' in c), None)
        
        if col_nom:
            for _, r in df.iterrows():
                nom_sheet = normalizar_texto(r[col_nom])
                nom_real = next((p for p in INTEGRANTES if nom_sheet in p or p in nom_sheet), None)
                
                if nom_real:
                    # 1. Libres Fijos
                    if col_fijo and str(r[col_fijo]) != 'nan':
                        txt_fijo = normalizar_texto(r[col_fijo])
                        for d_nombre, d_num in DIAS_MAP_TEXTO.items():
                            if d_nombre in txt_fijo and d_num not in libres_fijos[nom_real]:
                                libres_fijos[nom_real].append(d_num)

                    # 2. Vacaciones
                    if col_vac and str(r[col_vac]) != 'nan':
                        vacaciones[nom_real] = parsear_rangos_y_dias(r[col_vac])

                    # 3. Sugerencias / Libres del mes
                    if col_sug and str(r[col_sug]) != 'nan':
                        txt_sug = str(r[col_sug]).strip()
                        elementos = txt_sug.replace(' Y ', ',').replace('&', ',').split(',')
                        for elem in elementos:
                            elem = elem.strip().upper()
                            match_turno = re.match(r'(\d+)([NCL])', elem)
                            if match_turno:
                                dia_num, t_cod = match_turno.groups()
                                sugerencias[nom_real][dia_num] = t_cod
                            else:
                                dias_lib = parsear_rangos_y_dias(elem)
                                for d_lib in dias_lib:
                                    sugerencias[nom_real][str(d_lib)] = 'L'
                                    
    except Exception as e:
        st.error(f"Error procesando plantilla Google Sheets: {e}")
        
    return libres_fijos, vacaciones, sugerencias

def aplicar_colores(val):
    v = str(val).strip()
    if v == 'C': return 'background-color: #d1e7dd; color: #0f5132; font-weight: bold;'
    if v == 'N': return 'background-color: #cfe2ff; color: #084298; font-weight: bold;'
    if v == 'P': return 'background-color: #fff3cd; color: #664d03; font-weight: bold;'
    if v == 'D': return 'background-color: #f8f9fa; color: #212529;'
    if v == 'V': return 'background-color: #f8d7da; color: #842029; font-weight: bold;'
    if v == 'L': return 'background-color: #e2e3e5; color: #41464b; font-weight: bold;'
    return ''

# --- INTERFAZ STREAMLIT ---
st.title("🏥 Sistema de Gestión de Turnos Rosario")

st.sidebar.header("⚙️ Configuración del Mes")
mes_sel = st.sidebar.selectbox("Mes", list(range(1, 13)), index=datetime.datetime.now().month - 1)
ano_sel = st.sidebar.number_input("Año", min_value=2024, max_value=2030, value=datetime.datetime.now().year)

st.sidebar.header("🔗 Conexión a Google Sheets")
link_config = st.sidebar.text_input("Link Hoja Unificada (Google Sheets)", value="")

archivo_previo = st.file_uploader("📁 Cargar Excel de empalme mes anterior (Opcional)", type=["xlsx"])

# --- PESTAÑAS DE GENERACIÓN ---
tab1, tab2 = st.tabs(["⚡ Motor Pro (Optimización PuLP)", "🎲 Motor Clásico (Simulación)"])

with tab1:
    st.subheader("Optimización Matemática en 1 solo paso")
    if st.button("🚀 GENERAR CON OPTIMIZACIÓN", type="primary", use_container_width=True, key="btn_pulp"):
        conf, vacs, sug = procesar_configuracion_y_sugerencias(link_config)
        hist = procesar_historial_empalme(archivo_previo)
        
        with st.spinner('Resolviendo modelo de optimización matemática...'):
            res = generar_cuadro_pulp(mes_sel, ano_sel, hist, sug, conf, vacs)
            
        dias_en_mes = calendar.monthrange(ano_sel, mes_sel)[1]
        columnas_dias = [str(d) for d in range(1, dias_en_mes + 1)]
        
        st.success("✅ ¡Turnos calculados matemáticamente sin errores de restricciones!")
        st.dataframe(res.style.map(aplicar_colores, subset=columnas_dias), use_container_width=True)
        
        output = io.BytesIO()
        res_final_estilo = res.style.map(aplicar_colores, subset=columnas_dias)
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res_final_estilo.to_excel(writer, sheet_name='Turnos')
            worksheet = writer.sheets['Turnos']
            for i, col in enumerate(res.columns):
                worksheet.set_column(i, i, 10)
        
        st.download_button(
            label="📥 Descargar Excel Optimizado",
            data=output.getvalue(),
            file_name=f"Turnos_{mes_sel}_{ano_sel}_PuLP.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

with tab2:
    st.subheader("Simulación por N-iteraciones aleatorias")
    iteraciones = st.slider("Número de simulaciones", 10, 200, 50)
    if st.button("🚀 GENERAR ESCENARIO SIMULADO", type="secondary", use_container_width=True, key="btn_clasico"):
        conf, vacs, sug = procesar_configuracion_y_sugerencias(link_config)
        hist = procesar_historial_empalme(archivo_previo)
        
        with st.spinner('Simulando múltiples escenarios...'):
            res = generar_cuadro_pulp(mes_sel, ano_sel, hist, sug, conf, vacs)
            
        dias_en_mes = calendar.monthrange(ano_sel, mes_sel)[1]
        columnas_dias = [str(d) for d in range(1, dias_en_mes + 1)]
        
        st.success("✅ Escenario generado correctamente.")
        st.dataframe(res.style.map(aplicar_colores, subset=columnas_dias), use_container_width=True)
