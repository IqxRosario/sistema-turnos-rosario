import calendar
import datetime
import io
import re
import unicodedata

import pandas as pd
import streamlit as st

from motor_pulp import generar_cuadro_pulp, ErrorCuadroInviable


# -----------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# -----------------------------------------------------------------------------
INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES ARGUMEDO",
    "KELLY CAUSIL", "JOAN CARMONA",
]

LINK_SHEETS_DEFECTO = (
    "https://docs.google.com/spreadsheets/d/"
    "1PZwvv0XQtSEDfC5GO6OLG7Fn8HqJNQUBZ1RNSRgBsss/edit?pli=1&gid=0#gid=0"
)

DIAS_MAP_TEXTO = {
    "LUNES": 0,
    "MARTES": 1,
    "MIERCOLES": 2,
    "JUEVES": 3,
    "VIERNES": 4,
    "SABADO": 5,
    "DOMINGO": 6,
}

CODIGOS_COLOR = {
    "C": "#F4D35E",
    "N": "#5DADE2",
    "P": "#F5B7B1",
    "D": "#D5D8DC",
    "V": "#BB8FCE",
    "L": "#82E0AA",
}


# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES (LÓGICA EXISTENTE, SIN CAMBIOS)
# -----------------------------------------------------------------------------
def normalizar_texto(texto):
    """Quita tildes y normaliza texto para comparar encabezados y nombres."""
    if texto is None:
        return ""
    texto = str(texto).strip().upper()
    return "".join(
        caracter
        for caracter in unicodedata.normalize("NFD", texto)
        if unicodedata.category(caracter) != "Mn"
    )


def parsear_rangos_y_dias(texto):
    """Convierte expresiones como '1-4, 7 y 10' en [1, 2, 3, 4, 7, 10]."""
    if texto is None or pd.isna(texto):
        return []

    texto = str(texto).strip()
    if not texto:
        return []

    texto = re.sub(r"\s+Y\s+|\s+&\s+", ",", texto, flags=re.IGNORECASE)
    partes = re.split(r"[,;\n]+", texto)
    resultado = []

    for parte in partes:
        parte = parte.strip()
        if not parte:
            continue

        coincidencia = re.fullmatch(r"(\d+)\s*-\s*(\d+)", parte)
        if coincidencia:
            inicio, fin = map(int, coincidencia.groups())
            if inicio <= fin:
                resultado.extend(range(inicio, fin + 1))
            else:
                resultado.extend(range(fin, inicio + 1))
            continue

        for numero in re.findall(r"\d+", parte):
            resultado.append(int(numero))

    return sorted(set(resultado))


def _extraer_gid(url):
    coincidencia = re.search(r"[?&#]gid=(\d+)", str(url))
    return coincidencia.group(1) if coincidencia else "0"


def _url_csv_sheets(url):
    """Transforma una URL de Google Sheets en una URL de exportación CSV."""
    url = str(url).strip()
    coincidencia = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not coincidencia:
        raise ValueError("La URL no parece ser una hoja válida de Google Sheets.")
    sheet_id = coincidencia.group(1)
    gid = _extraer_gid(url)
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def _buscar_columna(columnas, patrones):
    """Devuelve la primera columna cuyo nombre normalizado contiene un patrón."""
    for columna in columnas:
        nombre = normalizar_texto(columna)
        if any(patron in nombre for patron in patrones):
            return columna
    return None


def _parsear_dias_fijos(texto):
    """Interpreta nombres de días de la semana como días fijos de descanso."""
    if texto is None or pd.isna(texto):
        return []
    texto = normalizar_texto(texto)
    encontrados = []
    for nombre, numero in DIAS_MAP_TEXTO.items():
        if re.search(rf"\b{nombre}\b", texto):
            encontrados.append(numero)
    return sorted(set(encontrados))


def _nombre_integrante(valor):
    """Busca una coincidencia segura del nombre en la lista del equipo."""
    nombre = normalizar_texto(valor)
    if not nombre:
        return None

    exactos = [p for p in INTEGRANTES if normalizar_texto(p) == nombre]
    if len(exactos) == 1:
        return exactos[0]

    coincidencias = [
        p for p in INTEGRANTES
        if nombre in normalizar_texto(p) or normalizar_texto(p) in nombre
    ]
    return coincidencias[0] if len(coincidencias) == 1 else None


def procesar_configuracion_y_sugerencias(link):
    """Lee la hoja seleccionada y devuelve libres fijos, vacaciones, solicitudes y prioridades."""
    libres_fijos = {p: [] for p in INTEGRANTES}
    vacaciones = {p: [] for p in INTEGRANTES}
    sugerencias = {p: {} for p in INTEGRANTES}
    prioridades = {p: 99 for p in INTEGRANTES}

    url_csv = _url_csv_sheets(link)
    try:
        df = pd.read_csv(url_csv)
    except Exception as exc:
        raise RuntimeError(
            "No se pudo leer Google Sheets. Verifica que la hoja tenga permiso "
            "de lectura mediante enlace y que la URL corresponda a la pestaña correcta. "
            f"Detalle: {exc}"
        ) from exc

    if df.empty:
        raise RuntimeError("La hoja de Google Sheets está vacía.")

    df.columns = [str(c).strip() for c in df.columns]
    col_nombre = _buscar_columna(df.columns, ["NOMBRE", "INTEGRANTE", "EMPLEADO", "PERSONA"])
    col_peticion = _buscar_columna(
        df.columns, ["PETICION", "LIBRES SOLICITADO", "SUGERENCIA", "SOLICITUD", "TURNO SOLICITADO"]
    )
    col_fijo = _buscar_columna(df.columns, ["FIJO", "ESTUDIO", "DIA FIJO", "DIAS FIJOS"])
    col_vacaciones = _buscar_columna(df.columns, ["VACACIO", "VACACIONES", "VACACION"])
    col_prioridad = _buscar_columna(df.columns, ["ORDEN", "PRIORIDAD", "LLEGADA"])

    if col_nombre is None:
        raise RuntimeError(
            "No encontré una columna de nombres. Usa un encabezado como 'NOMBRE' o 'INTEGRANTE'."
        )

    for _, fila in df.iterrows():
        integrante = _nombre_integrante(fila.get(col_nombre, ""))
        if integrante is None:
            continue

        if col_prioridad is not None:
            valor = fila.get(col_prioridad)
            try:
                if pd.notna(valor) and str(valor).strip():
                    prioridades[integrante] = int(float(valor))
            except (TypeError, ValueError):
                pass

        if col_fijo is not None:
            libres_fijos[integrante] = _parsear_dias_fijos(fila.get(col_fijo, ""))

        if col_vacaciones is not None:
            vacaciones[integrante] = parsear_rangos_y_dias(fila.get(col_vacaciones, ""))

        if col_peticion is not None:
            texto_peticion = fila.get(col_peticion, "")
            if pd.isna(texto_peticion):
                continue
            texto_peticion = normalizar_texto(texto_peticion)
            if not texto_peticion:
                continue

            encontrados = list(re.finditer(r"(\d+)\s*([NCL])", texto_peticion))
            dias_con_codigo = set()
            for coincidencia in encontrados:
                dia = int(coincidencia.group(1))
                codigo = coincidencia.group(2)
                if dia >= 1:
                    if dia in sugerencias[integrante] and sugerencias[integrante][dia] != codigo:
                        raise RuntimeError(
                            f"Solicitudes incompatibles para {integrante}, día {dia}: "
                            f"{sugerencias[integrante][dia]} y {codigo}. Corrige la hoja antes de continuar."
                        )
                    sugerencias[integrante][dia] = codigo
                    dias_con_codigo.add(dia)

            texto_sin_codigos = re.sub(r"\d+\s*[NCL]", " ", texto_peticion)
            for dia in parsear_rangos_y_dias(texto_sin_codigos):
                if dia not in dias_con_codigo and dia >= 1:
                    if dia in sugerencias[integrante] and sugerencias[integrante][dia] != "L":
                        raise RuntimeError(
                            f"Solicitudes incompatibles para {integrante}, día {dia}: "
                            f"{sugerencias[integrante][dia]} y L. Corrige la hoja antes de continuar."
                        )
                    sugerencias[integrante][dia] = "L"

    return libres_fijos, vacaciones, sugerencias, prioridades


def procesar_historial_empalme(archivo):
    """Lee el mes anterior para calcular la noche final y la racha de turnos."""
    hist_noche = {p: False for p in INTEGRANTES}
    hist_consecutivos = {p: 0 for p in INTEGRANTES}

    if archivo is None:
        return hist_noche, hist_consecutivos

    try:
        if hasattr(archivo, "seek"):
            archivo.seek(0)
        df = pd.read_excel(archivo)
    except Exception as exc:
        raise RuntimeError(f"No se pudo leer el archivo del mes anterior: {exc}") from exc

    if df.empty or len(df.columns) < 2:
        return hist_noche, hist_consecutivos

    columna_nombre = df.columns[0]
    columnas_dias = []
    for columna in df.columns[1:]:
        coincidencia = re.fullmatch(r"\s*(\d+)\s*", str(columna))
        if coincidencia:
            columnas_dias.append((int(coincidencia.group(1)), columna))
    columnas_dias.sort(key=lambda elemento: elemento[0])
    if not columnas_dias:
        return hist_noche, hist_consecutivos

    for _, fila in df.iterrows():
        integrante = _nombre_integrante(fila.get(columna_nombre, ""))
        if integrante is None:
            continue

        valores = [normalizar_texto(fila.get(columna, "")) for _, columna in columnas_dias]
        valores = [v if v in {"C", "N", "P", "D", "L", "V"} else "" for v in valores]
        if valores:
            hist_noche[integrante] = valores[-1] == "N"

        racha = 0
        for valor in reversed(valores):
            if valor in {"C", "N"}:
                racha += 1
            else:
                break
        hist_consecutivos[integrante] = racha

    return hist_noche, hist_consecutivos


def aplicar_colores(df):
    """Aplica colores a los códigos de turno, sin colorear las columnas de totales."""
    def color_celda(valor):
        codigo = str(valor).strip().upper()
        if codigo in CODIGOS_COLOR:
            return f"background-color: {CODIGOS_COLOR[codigo]}; color: #202020; text-align: center;"
        return "text-align: center;"

    columnas_dias = [c for c in df.columns if str(c).isdigit()]
    return df.style.map(color_celda, subset=columnas_dias)


def mostrar_reportes_pulp(df_res):
    diagnostico = df_res.attrs.get("diagnostico", {})
    incumplimientos = df_res.attrs.get("incumplimientos", [])
    patrones = df_res.attrs.get("patrones_n_p_n_p", [])

    st.caption(
        f"Validación PuLP: {diagnostico.get('estado', 'sin estado')} · "
        f"Solicitudes no satisfechas: {diagnostico.get('solicitudes_incumplidas', 0)} · "
        f"Patrones N-P-N-P: {diagnostico.get('patrones_n_p_n_p', 0)}"
    )

    if incumplimientos:
        st.warning(
            "El cuadro cumple las restricciones obligatorias, pero algunas solicitudes no pudieron "
            "atenderse por conflictos con la cobertura o con solicitudes de mayor prioridad. "
            "Revisa el detalle antes de publicar el horario."
        )
        st.dataframe(pd.DataFrame(incumplimientos), use_container_width=True, hide_index=True)
    else:
        st.success("Todas las solicitudes puntuales se cumplieron.")

    if patrones:
        with st.expander("Ver patrones N-P-N-P encontrados"):
            st.dataframe(pd.DataFrame(patrones), use_container_width=True, hide_index=True)


def crear_excel(df_res, mes, ano):
    """Exporta el cuadro y los reportes del optimizador en hojas separadas."""
    salida = io.BytesIO()
    incumplimientos = df_res.attrs.get("incumplimientos", [])
    patrones = df_res.attrs.get("patrones_n_p_n_p", [])
    diagnostico = df_res.attrs.get("diagnostico", {})

    df_turnos = df_res.copy()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df_turnos.to_excel(writer, sheet_name="Turnos")
        pd.DataFrame([diagnostico or {"estado": "sin datos"}]).to_excel(
            writer, sheet_name="Diagnostico", index=False
        )
        if incumplimientos:
            pd.DataFrame(incumplimientos).to_excel(
                writer, sheet_name="Solicitudes no cumplidas", index=False
            )
        if patrones:
            pd.DataFrame(patrones).to_excel(
                writer, sheet_name="Patrones N-P-N-P", index=False
            )

    salida.seek(0)
    return salida.getvalue(), f"Turnos_{mes:02d}_{ano}.xlsx"


def generar_y_mostrar(mes_sel, ano_sel, link_config, archivo_previo):
    """Carga la configuración, genera, valida y muestra el horario con sus reportes."""
    if not st.button("⚙️ Generar cuadro de turnos", key="generar_cuadro_principal", type="primary"):
        return

    with st.spinner("Leyendo configuración y optimizando los turnos..."):
        try:
            conf, vacs, sug, prio = procesar_configuracion_y_sugerencias(link_config)
            hist_noche, hist_racha = procesar_historial_empalme(archivo_previo)

            dias_mes = calendar.monthrange(ano_sel, mes_sel)[1]
            errores_dias = []
            for persona, dias_persona in vacs.items():
                invalidos = [d for d in dias_persona if d < 1 or d > dias_mes]
                if invalidos:
                    errores_dias.append(f"{persona}: vacaciones en día(s) fuera del mes {invalidos}")
            for persona, solicitudes in sug.items():
                invalidos = [d for d in solicitudes if int(d) < 1 or int(d) > dias_mes]
                if invalidos:
                    errores_dias.append(f"{persona}: solicitudes en día(s) fuera del mes {invalidos}")
            if errores_dias:
                raise RuntimeError("Hay días inválidos en la hoja:\n- " + "\n- ".join(errores_dias))

            df_res = generar_cuadro_pulp(
                mes_sel,
                ano_sel,
                hist_noche,
                hist_racha,
                sug,
                conf,
                vacs,
                prio,
            )

        except (ErrorCuadroInviable, RuntimeError, ValueError) as exc:
            st.error(f"No se generó el cuadro. {exc}")
            st.info(
                "Corrige la configuración indicada y vuelve a intentarlo. No se mostrará ni descargará "
                "un cuadro si el optimizador no puede validar una solución."
            )
            return
        except Exception as exc:
            st.error(f"Ocurrió un error inesperado al generar el cuadro: {exc}")
            st.exception(exc)
            return

    st.success("Cuadro generado y validado correctamente.")
    mostrar_reportes_pulp(df_res)
    st.dataframe(aplicar_colores(df_res), use_container_width=True, height=600)

    try:
        excel_bytes, nombre_archivo = crear_excel(df_res, mes_sel, ano_sel)
        st.download_button(
            label="📥 Descargar cuadro e informes en Excel",
            data=excel_bytes,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="descargar_cuadro_principal",
        )
    except Exception as exc:
        st.error(f"El cuadro se generó, pero no se pudo preparar el Excel: {exc}")


# -----------------------------------------------------------------------------
# INTERFAZ STREAMLIT (CAMBIOS SOLO VISUALES / DE NAVEGACIÓN)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Gestor de Instrumentación",
    page_icon="🏥",
    layout="wide",
)

st.title("🏥 Sistema de Gestión de Turnos Rosario")
st.caption("Generación y validación de turnos mediante optimización PuLP.")

# El acceso a la hoja habitual queda siempre visible; no hay que copiar su URL.
st.link_button(
    "📄 Abrir hoja de configuración en Google Sheets",
    LINK_SHEETS_DEFECTO,
    type="secondary",
)
st.caption(
    "Desde esa hoja puedes consultar o actualizar nombres, solicitudes, días fijos, vacaciones y prioridades. "
    "La aplicación leerá automáticamente la hoja predeterminada al generar el cuadro."
)

with st.sidebar:
    st.header("⚙️ Configuración del cuadro")
    hoy = datetime.date.today()
    meses_es = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    ]
    mes_sel = st.selectbox(
        "Mes a programar",
        options=list(range(1, 13)),
        format_func=lambda m: meses_es[m - 1],
        index=hoy.month - 1,
    )
    ano_sel = st.selectbox(
        "Año",
        options=list(range(hoy.year, hoy.year + 5)),
        index=0,
    )

    archivo_previo = st.file_uploader(
        "Excel del mes anterior (opcional)",
        type=["xlsx", "xls"],
        help="Se utiliza para aplicar las reglas de empalme entre meses.",
    )

    # La opción avanzada es opcional. El usuario normal no tiene que copiar una URL.
    with st.expander("Configuración avanzada: usar otra hoja"):
        link_personalizado = st.text_input(
            "URL alternativa de Google Sheets",
            value="",
            placeholder="Deja vacío para usar la hoja habitual",
            help="Solo rellena esto si necesitas utilizar otra hoja o pestaña. La hoja habitual está disponible en el botón de arriba.",
        )
        if link_personalizado.strip():
            st.link_button("Abrir hoja alternativa", link_personalizado.strip())
            st.caption("Al generar el cuadro se utilizará esta URL alternativa.")
        else:
            st.caption("Se utilizará automáticamente la hoja habitual.")

link_config = link_personalizado.strip() if link_personalizado.strip() else LINK_SHEETS_DEFECTO

st.divider()
st.subheader("Generar horario")
st.info(
    "Las vacaciones, la cobertura diaria y las restricciones obligatorias no se relajan. "
    "Las solicitudes puntuales se optimizan por prioridad; los conflictos se muestran en un informe. "
    "Los días fijos son preferencias fuertes que pueden ceder si la cobertura lo exige."
)

st.markdown("**Motor de generación:** Optimización PuLP")
st.write(
    "Se utiliza un único motor para evitar confusiones entre pestañas que ejecutaban el mismo algoritmo. "
    "El resultado incluye el cuadro, las solicitudes no satisfechas y los patrones N-P-N-P detectados."
)
generar_y_mostrar(mes_sel, ano_sel, link_config, archivo_previo)
