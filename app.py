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
    "1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?gid=0#gid=0"
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
# FUNCIONES AUXILIARES
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
    reporte_carga = df_res.attrs.get("reporte_carga", [])

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
        st.dataframe
