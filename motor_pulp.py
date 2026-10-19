import calendar
import datetime
import pandas as pd
from pulp import (
    LpProblem,
    LpMaximize,
    LpMinimize,
    LpVariable,
    LpBinary,
    LpStatus,
    lpSum,
    value,
)

INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES ARGUMEDO",
    "KELLY CAUSIL", "JOAN CARMONA"
]


class ErrorCuadroInviable(RuntimeError):
    """El modelo no puede satisfacer simultáneamente sus reglas obligatorias."""


def _domingo_pascua(ano):
    """Calcula el domingo de Pascua con el algoritmo gregoriano de Meeus/Jones/Butcher."""
    a = ano % 19
    b = ano // 100
    c = ano % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return datetime.date(ano, mes, dia)


def _lunes_siguiente(fecha):
    """Aplica traslado de Ley Emiliani al lunes siguiente, salvo si ya es lunes."""
    return fecha + datetime.timedelta(days=(7 - fecha.weekday()) % 7)


def es_festivo_co(dia, mes, ano):
    """Calendario de festivos nacionales colombianos, sin depender de paquetes externos."""
    fecha = datetime.date(ano, mes, dia)
    pascua = _domingo_pascua(ano)
    fijos = {
        datetime.date(ano, 1, 1), datetime.date(ano, 5, 1),
        datetime.date(ano, 7, 20), datetime.date(ano, 8, 7),
        datetime.date(ano, 12, 8), datetime.date(ano, 12, 25),
        pascua - datetime.timedelta(days=3),  # Jueves Santo
        pascua - datetime.timedelta(days=2),  # Viernes Santo
    }
    trasladables = [
        datetime.date(ano, 1, 6), datetime.date(ano, 3, 19),
        datetime.date(ano, 6, 29), datetime.date(ano, 8, 15),
        datetime.date(ano, 10, 12), datetime.date(ano, 11, 1),
        datetime.date(ano, 11, 11),
        pascua + datetime.timedelta(days=39),  # Ascensión
        pascua + datetime.timedelta(days=60),  # Corpus Christi
        pascua + datetime.timedelta(days=68),  # Sagrado Corazón
    ]
    fijos.update(_lunes_siguiente(f) for f in trasladables)
    return fecha in fijos


def _prioridad_entera(valor):
    try:
        if pd.isna(valor):
            return 99
        return int(float(valor))
    except (TypeError, ValueError):
        return 99


def _obtener_solicitud(sugerencias_dict, persona, dia):
    """Acepta claves de día int o str y normaliza la letra de solicitud."""
    solicitudes = sugerencias_dict.get(persona, {}) or {}
    valor = solicitudes.get(str(dia), solicitudes.get(dia, ""))
    if valor is None:
        return ""
    valor = str(valor).strip().upper()
    return valor if valor in {"L", "C", "N"} else ""


def generar_cuadro_pulp(
    mes,
    ano,
    hist_noche,
    hist_racha,
    sugerencias_dict,
    config_dict,
    vacaciones_dict,
    prioridades_dict=None,
):
                """Genera turnos con reglas obligatorias y solicitudes por prioridad lexicográfica.

    Orden de optimización:
      1. Restricciones obligatorias.
      2. Minimizar solicitudes incumplidas según prioridad.
      3. Equilibrar fines de semana y festivos.
    """

    Orden de optimización:
      1. Restricciones obligatorias.
      2. Minimizar solicitudes incumplidas según prioridad.
      3. Equilibrar fines de semana y festivos.
    """

    Orden de optimización:
      1. Restricciones obligatorias.
      2. Minimizar solicitudes incumplidas según prioridad.
      3. Equilibrar fines de semana y festivos.
    """

    Orden de optimización:
      1. Restricciones obligatorias (cobertura, vacaciones, descansos, etc.).
      2. Minimizar solicitudes incumplidas, por nivel de prioridad ascendente
         (un número menor representa mayor prioridad).
      3. Equilibrar proporcionalmente los fines de semana/festivos según días disponibles.
      4. Prefer
