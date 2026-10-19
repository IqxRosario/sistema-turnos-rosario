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
      1. Restricciones obligatorias (cobertura, vacaciones, descansos, etc.).
      2. Minimizar solicitudes incumplidas, por nivel de prioridad ascendente
         (un número menor representa mayor prioridad).
      3. Equilibrar proporcionalmente los fines de semana/festivos según días disponibles.
      4. Preferencias flexibles: evitar N-P-N-P, respetar días fijos siempre que
         la cobertura lo permita y conservar las preferencias históricas.

    El DataFrame devuelto conserva el formato anterior. Los reportes se adjuntan
    en df.attrs['incumplimientos'] y df.attrs['diagnostico'].
    """
    if prioridades_dict is None:
        prioridades_dict = {p: 99 for p in INTEGRANTES}

    sugerencias_dict = sugerencias_dict or {}
    config_dict = config_dict or {}
    vacaciones_dict = vacaciones_dict or {}
    hist_noche = hist_noche or {}
    hist_racha = hist_racha or {}

    dias_mes = calendar.monthrange(ano, mes)[1]
    dias = list(range(1, dias_mes + 1))
    turnos = ("C", "N")
    prob = LpProblem("Optimizacion_Turnos_Hospital", LpMaximize)

    # Variables de asignación.
    x = {
        p: {
            d: {
                t: LpVariable(f"x_{p.replace(' ', '_')}_{d}_{t}", cat=LpBinary)
                for t in turnos
            }
            for d in dias
        }
        for p in INTEGRANTES
    }

    # Cada solicitud tiene una variable de incumplimiento. Estas variables no
    # permiten romper las reglas duras; solo hacen explícita la petición no atendida.
    incumple = {}
    prioridad_persona = {
        p: _prioridad_entera(prioridades_dict.get(p, 99)) for p in INTEGRANTES
    }
    solicitudes_por_prioridad = {}
    for p in INTEGRANTES:
        for d in dias:
            req = _obtener_solicitud(sugerencias_dict, p, d)
            if not req:
                continue
            y = LpVariable(f"incumple_{p.replace(' ', '_')}_{d}", cat=LpBinary)
            incumple[p, d] = y
            nivel = prioridad_persona[p]
            solicitudes_por_prioridad.setdefault(nivel, []).append(y)

            if req == "L":
                # Si se trabaja, la solicitud de libre queda incumplida.
                prob += x[p][d]["C"] + x[p][d]["N"] <= y, f"solicitud_libre_{p.replace(' ', '_')}_{d}"
            else:
                # Se pide el turno exacto; si no se asigna, la petición se incumple.
                prob += x[p][d][req] + y >= 1, f"solicitud_turno_{p.replace(' ', '_')}_{d}"

    # Reglas obligatorias de empalme entre meses.
    for p in INTEGRANTES:
        vacaciones_p = {int(v) for v in (vacaciones_dict.get(p, []) or []) if str(v).strip().isdigit()}
        if hist_noche.get(p, False) and 1 in vacaciones_p:
            raise ErrorCuadroInviable(
                f"Conflicto de empalme: {p} termina el mes anterior con N y tiene vacaciones el día 1. "
                "Debe resolverse en la configuración, porque P y V no pueden representarse a la vez."
            )
        if hist_noche.get(p, False):
            prob += x[p][1]["C"] == 0, f"empalme_noche_C_{p.replace(' ', '_')}"
            prob += x[p][1]["N"] == 0, f"empalme_noche_N_{p.replace(' ', '_')}"

        racha = max(0, int(hist_racha.get(p, 0) or 0))
        if racha >= 3:
            prob += x[p][1]["C"] + x[p][1]["N"] == 0, f"empalme_racha3_{p.replace(' ', '_')}"
        elif racha == 2:
            prob += (
                x[p][1]["C"] + x[p][1]["N"] + x[p][2]["C"] + x[p][2]["N"] <= 1,
                f"empalme_racha2_{p.replace(' ', '_')}"
            )
        elif racha == 1 and dias_mes >= 3:
            # Un turno el último día anterior + máximo 2 en los primeros 3 días.
            prob += lpSum(
                x[p][d][t] for d in range(1, 4) for t in turnos
            ) <= 2, f"empalme_racha1_{p.replace(' ', '_')}"

    # Cobertura, vacaciones y restricciones individuales.
    for d in dias:
        wd = datetime.date(ano, mes, d).weekday()
        es_festivo = es_festivo_co(d, mes, ano)
        cuota_n = 2
        cuota_c = 2 if (es_festivo or wd == 6) else (5 if wd == 5 else 6)

        prob += lpSum(x[p][d]["N"] for p in INTEGRANTES) == cuota_n, f"cobertura_N_{d}"
        prob += lpSum(x[p][d]["C"] for p in INTEGRANTES) == cuota_c, f"cobertura_C_{d}"

        # Vacaciones siempre intocables.
        for p in INTEGRANTES:
            vacaciones_p = {int(v) for v in (vacaciones_dict.get(p, []) or []) if str(v).strip().isdigit()}
            if d in vacaciones_p:
                prob += x[p][d]["C"] == 0, f"vacacion_C_{p.replace(' ', '_')}_{d}"
                prob += x[p][d]["N"] == 0, f"vacacion_N_{p.replace(' ', '_')}_{d}"

            # Máximo un turno diario.
            prob += x[p][d]["C"] + x[p][d]["N"] <= 1, f"max_un_turno_{p.replace(' ', '_')}_{d}"

            # Después de una noche, no puede trabajar C ni N al día siguiente.
            if d < dias_mes:
                prob += (
                    x[p][d]["N"] + x[p][d + 1]["C"] + x[p][d + 1]["N"] <= 1,
                    f"descanso_post_noche_{p.replace(' ', '_')}_{d}"
                )

            # Máximo 3 turnos C/N en cualquier ventana de 4 días.
            if d <= dias_mes - 3:
                prob += lpSum(
                    x[p][d + k][t] for k in range(4) for t in turnos
                ) <= 3, f"max_3_en_4_{p.replace(' ', '_')}_{d}"

        # Restricciones individuales confirmadas por la usuaria.
        # Jhon: lunes sin N; martes solo L o N (nunca C).
        if wd == 0:
            prob += x["JHON RIOS"][d]["N"] == 0, f"jhon_sin_noche_lunes_{d}"
        if wd == 1:
            prob += x["JHON RIOS"][d]["C"] == 0, f"jhon_martes_solo_L_o_N_{d}"

        # Ginelap: miércoles sin N; jueves y viernes intocables.
        if wd == 2:
            prob += x["GINELAP"][d]["N"] == 0, f"ginelap_sin_noche_miercoles_{d}"
        if wd in (3, 4):
            prob += x["GINELAP"][d]["C"] + x["GINELAP"][d]["N"] == 0, f"ginelap_libre_intocable_{d}"

        # Gerlis y Zariana: no noches los domingos. Lunes C permitido, no obligatorio.
        if wd == 6:
            for persona in ("GERLIS DOMINGUEZ", "ZARIANA REYES"):
                prob += x[persona][d]["N"] == 0, f"sin_noche_domingo_{persona.replace(' ', '_')}_{d}"

        # Ivette: lunes libre intocable.
        if wd == 0:
            prob += x["IVETTE VALENCIA"][d]["C"] + x["IVETTE VALENCIA"][d]["N"] == 0, f"ivette_lunes_intocable_{d}"

        # Angie: miércoles libre por defecto, puede cubrir únicamente con C; nunca N.
        if wd == 2:
            prob += x["ANGIE BERNAL"][d]["N"] == 0, f"angie_sin_noche_miercoles_{d}"

        # Días fijos cargados en Sheets son intocables, excepto las excepciones explícitas.
        for persona in INTEGRANTES:
            if wd not in set(config_dict.get(persona, []) or []):
                continue
            if persona == "JUAN CAMILO PEREZ" and wd in (1, 3):
                # Excepción limitada: C únicamente, hasta cuatro veces al mes en martes/jueves.
                prob += x[persona][d]["N"] == 0, f"juan_sin_noche_dia_fijo_{d}"
            elif persona == "ANGIE BERNAL" and wd == 2:
                prob += x[persona][d]["N"] == 0, f"angie_excepcion_C_miercoles_{d}"
            elif persona in ("IVETTE VALENCIA", "ERNESTO MUSKUS") and wd == 5:
                # Sábado libre por defecto, pero C/N permitido como excepción de cobertura.
                pass
            elif persona in ("GERLIS DOMINGUEZ", "ZARIANA REYES") and wd == 0:
                # Lunes de cirujano: C permitido, no obligatorio; no se abre N por esta excepción.
                prob += x[persona][d]["N"] == 0, f"lunes_cirujano_sin_N_{persona.replace(' ', '_')}_{d}"
            else:
                prob += x[persona][d]["C"] + x[persona][d]["N"] == 0, f"libre_fijo_intocable_{persona.replace(' ', '_')}_{d}"

    # Juan Camilo: máximo cuatro C en total entre martes y jueves; N prohibida allí.
    dias_jc = [d for d in dias if datetime.date(ano, mes, d).weekday() in (1, 3)]
    prob += lpSum(x["JUAN CAMILO PEREZ"][d]["C"] for d in dias_jc) <= 4, "juan_max_4_C_martes_jueves"
    for d in dias_jc:
        prob += x["JUAN CAMILO PEREZ"][d]["N"] == 0, f"juan_sin_N_martes_jueves_{d}"

    # Carga mensual: objetivos suaves ajustados por días realmente disponibles.
    # Se guardan desviaciones para que cobertura y reglas duras siempre tengan prioridad.
    desviaciones_carga = []
    objetivos_carga = {}
    for persona in INTEGRANTES:
        vacaciones_p = {int(v) for v in (vacaciones_dict.get(persona, []) or []) if str(v).strip().isdigit()}
        dias_disponibles = []
        dias_disponibles_noche = []
        for d in dias:
            wd = datetime.date(ano, mes, d).weekday()
            no_disponible = d in vacaciones_p or wd in set(config_dict.get(persona, []) or [])
            if persona == "GINELAP" and wd in (3, 4):
                no_disponible = True
            if persona == "IVETTE VALENCIA" and wd == 0:
                no_disponible = True
            if not no_disponible:
                dias_disponibles.append(d)

            no_disponible_noche = no_disponible
            if persona == "JHON RIOS" and wd == 0:
                no_disponible_noche = True
            if persona == "GINELAP" and wd == 2:
                no_disponible_noche = True
            if persona in ("GERLIS DOMINGUEZ", "ZARIANA REYES") and wd == 6:
                no_disponible_noche = True
            if persona == "ANGIE BERNAL" and wd == 2:
                no_disponible_noche = True
            if persona == "JUAN CAMILO PEREZ" and wd in (1, 3):
                no_disponible_noche = True
            if persona in ("IVETTE VALENCIA", "ERNESTO MUSKUS") and wd == 5:
                no_disponible_noche = True
            if not no_disponible_noche:
                dias_disponibles_noche.append(d)
        proporcion = len(dias_disponibles) / max(1, dias_mes)
        proporcion_noches = len(dias_disponibles_noche) / max(1, dias_mes)
        objetivo_turnos = 17 * proporcion
        objetivo_noches = 4.5 * proporcion_noches
        objetivos_carga[persona] = (
            objetivo_turnos, objetivo_noches, len(dias_disponibles), len(dias_disponibles_noche)
        )
        total_turnos_p = lpSum(x[persona][d][t] for d in dias for t in turnos)
        total_noches_p = lpSum(x[persona][d]["N"] for d in dias)
        des_turnos = LpVariable(f"desv_turnos_{persona.replace(' ', '_')}", lowBound=0)
        des_noches = LpVariable(f"desv_noches_{persona.replace(' ', '_')}", lowBound=0)
        prob += des_turnos >= total_turnos_p - objetivo_turnos, f"desv_turnos_pos_{persona.replace(' ', '_')}"
        prob += des_turnos >= objetivo_turnos - total_turnos_p, f"desv_turnos_neg_{persona.replace(' ', '_')}"
        prob += des_noches >= total_noches_p - objetivo_noches, f"desv_noches_pos_{persona.replace(' ', '_')}"
        prob += des_noches >= objetivo_noches - total_noches_p, f"desv_noches_neg_{persona.replace(' ', '_')}"
        desviaciones_carga.append((des_turnos, des_noches))

    # 1) Minimizar primero las excepciones de cobertura (solo se usan si hacen falta).
    # Esto evita que una solicitud personal active por sí sola un sábado/martes/miércoles excepcional.
    excepciones_cobertura = []
    for d in dias:
        wd = datetime.date(ano, mes, d).weekday()
        if wd == 5:
            for persona in ("IVETTE VALENCIA", "ERNESTO MUSKUS"):
                excepciones_cobertura.extend([x[persona][d]["C"], x[persona][d]["N"]])
        if wd in (1, 3):
            excepciones_cobertura.append(x["JUAN CAMILO PEREZ"][d]["C"])
        if wd == 2:
            excepciones_cobertura.append(x["ANGIE BERNAL"][d]["C"])
        if wd == 0 and wd in set(config_dict.get("GERLIS DOMINGUEZ", []) or []):
            excepciones_cobertura.append(x["GERLIS DOMINGUEZ"][d]["C"])
        if wd == 0 and wd in set(config_dict.get("ZARIANA REYES", []) or []):
            excepciones_cobertura.append(x["ZARIANA REYES"][d]["C"])
    if excepciones_cobertura:
        objetivo_excepciones = lpSum(excepciones_cobertura)
        prob.sense = LpMinimize
        prob.setObjective(objetivo_excepciones)
        prob.solve()
        estado_excepciones = LpStatus.get(prob.status, str(prob.status))
        if estado_excepciones != "Optimal":
            raise ErrorCuadroInviable(
                f"No existe una solución que respete cobertura y restricciones obligatorias para {mes:02d}/{ano}. "
                f"Estado de PuLP: {estado_excepciones}."
            )
        minimo_excepciones = int(round(value(objetivo_excepciones) or 0))
        prob += objetivo_excepciones == minimo_excepciones, "fijar_minimo_excepciones_cobertura"

    # 2) Optimización lexicográfica de solicitudes: se fija cada nivel de prioridad.
    # Un número menor representa mayor prioridad.
    prioridades_ordenadas = sorted(solicitudes_por_prioridad)
    if prioridades_ordenadas:
        for indice, nivel in enumerate(prioridades_ordenadas):
            objetivo_nivel = lpSum(solicitudes_por_prioridad[nivel])
            prob.sense = LpMinimize
            prob.setObjective(objetivo_nivel)
            prob.solve()
            estado = LpStatus.get(prob.status, str(prob.status))
            if estado != "Optimal":
                raise ErrorCuadroInviable(
                    f"No existe un cuadro que cumpla las restricciones obligatorias para {mes:02d}/{ano}. "
                    f"Estado de PuLP: {estado}. Revisa cobertura, vacaciones, descansos y reglas individuales."
                )
            mejor_valor = int(round(value(objetivo_nivel) or 0))
            prob += objetivo_nivel == mejor_valor, f"fijar_prioridad_{indice}_{nivel}"

    # 2) Equidad de fines de semana/festivos.
    # Se compara la proporción de días de fin de semana trabajados respecto a los
    # días de fin de semana disponibles de cada persona. Así, las vacaciones y
    # los descansos obligatorios de Ginelap no penalizan injustamente.
    dias_fin_semana = [
        d for d in dias
        if datetime.date(ano, mes, d).weekday() >= 5 or es_festivo_co(d, mes, ano)
    ]
    disponibilidad_fin_semana = {}
    for p in INTEGRANTES:
        vacaciones_p = {
            int(v) for v in (vacaciones_dict.get(p, []) or [])
            if str(v).strip().isdigit()
        }
        disponibilidad_fin_semana[p] = [
            d for d in dias_fin_semana
            if d not in vacaciones_p
            and datetime.date(ano, mes, d).weekday() not in set(config_dict.get(p, []) or [])
            and not (p == "GINELAP" and datetime.date(ano, mes, d).weekday() in (3, 4))
            and not (p == "IVETTE VALENCIA" and datetime.date(ano, mes, d).weekday() == 0)
        ]

    tasa_fin_semana_max = LpVariable("tasa_fin_semana_max", lowBound=0, upBound=1)
    tasa_fin_semana_min = LpVariable("tasa_fin_semana_min", lowBound=0, upBound=1)
    for p in INTEGRANTES:
        disponibles = disponibilidad_fin_semana[p]
        if not disponibles:
            continue
        total_fin_semana_p = lpSum(
            x[p][d][t] for d in disponibles for t in turnos
        )
        prob += (
            total_fin_semana_p <= tasa_fin_semana_max * len(disponibles),
            f"limite_tasa_fin_semana_max_{p.replace(' ', '_')}"
        )
        prob += (
            total_fin_semana_p >= tasa_fin_semana_min * len(disponibles),
            f"limite_tasa_fin_semana_min_{p.replace(' ', '_')}"
        )

    # Se prioriza reducir la diferencia entre las tasas de carga de fin de semana.
    # Se deja una tolerancia pequeña después para conservar margen de optimización
    # de patrones nocturnos y preferencias flexibles.
    brecha_fin_semana = tasa_fin_semana_max - tasa_fin_semana_min
    prob.sense = LpMinimize
    prob.setObjective(brecha_fin_semana)
    prob.solve()
    estado_equidad = LpStatus.get(prob.status, str(prob.status))
    if estado_equidad != "Optimal":
        raise ErrorCuadroInviable(
            f"No se pudo equilibrar la carga de fines de semana para {mes:02d}/{ano}. "
            f"Estado de PuLP: {estado_equidad}."
        )
    mejor_brecha_fin_semana = value(brecha_fin_semana)
    prob += (
        brecha_fin_semana <= (mejor_brecha_fin_semana or 0) + 0.02,
        "mantener_equidad_fin_semana"
    )

    # 3) Preferencias flexibles: evitar N-P-N-P y conservar días fijos libres.
    # El patrón N-P-N-P equivale a dos noches separadas por dos días.
    penaliza_patron = {}
    for p in INTEGRANTES:
        # Patrón dentro del mes: N(d) -> P(d+1) -> N(d+2).
        for d in range(1, dias_mes - 1):
            z = LpVariable(f"patron_NPNP_{p.replace(' ', '_')}_{d}", cat=LpBinary)
            penaliza_patron[p, d] = z
            prob += z >= x[p][d]["N"] + x[p][d + 2]["N"] - 1, f"detectar_patron_NPNP_{p.replace(' ', '_')}_{d}"
        # Patrón que cruza de mes: N al cierre previo -> P el día 1 -> N el día 2.
        if hist_noche.get(p, False) and dias_mes >= 2:
            z = LpVariable(f"patron_NPNP_empalme_{p.replace(' ', '_')}", cat=LpBinary)
            penaliza_patron[p, 0] = z
            prob += z >= x[p][2]["N"], f"detectar_patron_NPNP_empalme_{p.replace(' ', '_')}"

    pesos_flexibles = []
    for p in INTEGRANTES:
        dias_fijos = set(config_dict.get(p, []) or [])
        for d in dias:
            wd = datetime.date(ano, mes, d).weekday()
            req = _obtener_solicitud(sugerencias_dict, p, d)
            for t in turnos:
                # Una solicitud se optimiza en la fase lexicográfica; aquí no se
                # vuelve a premiar ni castigar para no alterar la prioridad.
                if req:
                    peso = 0
                elif (p == "JUAN CAMILO PEREZ" and wd in (1, 3)) or (
                    p in ("ERNESTO MUSKUS", "IVETTE VALENCIA") and wd == 5
                ):
                    peso = -80
                elif p == "ANGIE BERNAL" and wd == 2:
                    peso = -80
                else:
                    peso = 1
                pesos_flexibles.append(peso * x[p][d][t])

    # Minimizar patrones tiene más importancia que las preferencias flexibles.
    # Se maximiza el resto y se resta una penalización grande por patrón.
    penalizacion_carga = lpSum(2 * d_turnos + 5 * d_noches for d_turnos, d_noches in desviaciones_carga)
    objetivo_final = (
        lpSum(pesos_flexibles)
        - 200 * lpSum(penaliza_patron.values())
        - 10 * penalizacion_carga
    )
    prob.sense = LpMaximize
    prob.setObjective(objetivo_final)
    prob.solve()
    estado = LpStatus.get(prob.status, str(prob.status))
    if estado != "Optimal":
        raise ErrorCuadroInviable(
            f"No se pudo obtener una solución final válida para {mes:02d}/{ano}. Estado de PuLP: {estado}."
        )

    # Validación independiente de las asignaciones antes de construir la tabla.
    asignaciones = {}
    errores_validacion = []
    for p in INTEGRANTES:
        for d in dias:
            asignaciones[p, d] = next(
                (t for t in turnos if (value(x[p][d][t]) or 0) > 0.5), ""
            )
            vacs_p = {int(v) for v in (vacaciones_dict.get(p, []) or []) if str(v).strip().isdigit()}
            if d in vacs_p and asignaciones[p, d]:
                errores_validacion.append(f"{p}, día {d}: turno asignado durante vacaciones")

    for d in dias:
        n_real = sum(asignaciones[p, d] == "N" for p in INTEGRANTES)
        c_real = sum(asignaciones[p, d] == "C" for p in INTEGRANTES)
        wd = datetime.date(ano, mes, d).weekday()
        c_esperado = 2 if (es_festivo_co(d, mes, ano) or wd == 6) else (5 if wd == 5 else 6)
        if n_real != 2 or c_real != c_esperado:
            errores_validacion.append(
                f"Día {d}: cobertura N={n_real}/2, C={c_real}/{c_esperado}"
            )
    for p in INTEGRANTES:
        noches = sum(asignaciones[p, d] == "N" for d in dias)
        for d in dias[:-1]:
            if asignaciones[p, d] == "N" and asignaciones[p, d + 1] in {"C", "N"}:
                errores_validacion.append(f"{p}: no descansa después de N en día {d}")
        for d in range(1, dias_mes - 2):
            total_4 = sum(bool(asignaciones[p, d + k]) for k in range(4))
            if total_4 > 3:
                errores_validacion.append(f"{p}: más de 3 turnos en los días {d}-{d + 3}")

    # Validación independiente de las reglas individuales y los días fijos.
    for d in dias:
        wd = datetime.date(ano, mes, d).weekday()
        if wd == 0 and asignaciones["JHON RIOS", d] == "N":
            errores_validacion.append(f"Jhon: noche prohibida el lunes {d}")
        if wd == 1 and asignaciones["JHON RIOS", d] == "C":
            errores_validacion.append(f"Jhon: el martes solo puede ser L o N (día {d})")
        if wd == 2 and asignaciones["GINELAP", d] == "N":
            errores_validacion.append(f"Ginelap: noche prohibida el miércoles {d}")
        if wd in (3, 4) and asignaciones["GINELAP", d]:
            errores_validacion.append(f"Ginelap: libre intocable incumplido el día {d}")
        if wd == 6:
            for persona in ("GERLIS DOMINGUEZ", "ZARIANA REYES"):
                if asignaciones[persona, d] == "N":
                    errores_validacion.append(f"{persona}: noche dominical prohibida, día {d}")
        if wd == 0 and asignaciones["IVETTE VALENCIA", d]:
            errores_validacion.append(f"Ivette: lunes libre intocable incumplido, día {d}")
        if wd == 2 and asignaciones["ANGIE BERNAL", d] == "N":
            errores_validacion.append(f"Angie: noche prohibida el miércoles {d}")
        for persona in INTEGRANTES:
            if wd in set(config_dict.get(persona, []) or []) and asignaciones[persona, d]:
                excepcion_permitida = (
                    (persona == "JUAN CAMILO PEREZ" and wd in (1, 3) and asignaciones[persona, d] == "C")
                    or (persona == "ANGIE BERNAL" and wd == 2 and asignaciones[persona, d] == "C")
                    or (persona in ("IVETTE VALENCIA", "ERNESTO MUSKUS") and wd == 5 and asignaciones[persona, d] in {"C", "N"})
                    or (persona in ("GERLIS DOMINGUEZ", "ZARIANA REYES") and wd == 0 and asignaciones[persona, d] == "C")
                )
                if not excepcion_permitida:
                    errores_validacion.append(f"{persona}: libre fijo incumplido el día {d}")

    dias_jc = [d for d in dias if datetime.date(ano, mes, d).weekday() in (1, 3)]
    c_jc = sum(asignaciones["JUAN CAMILO PEREZ", d] == "C" for d in dias_jc)
    if c_jc > 4:
        errores_validacion.append(f"Juan Camilo: {c_jc} C en martes/jueves (máximo 4)")
    for d in dias_jc:
        if asignaciones["JUAN CAMILO PEREZ", d] == "N":
            errores_validacion.append(f"Juan Camilo: noche prohibida martes/jueves, día {d}")

    if errores_validacion:
        raise RuntimeError("La validación independiente falló: " + "; ".join(errores_validacion[:12]))

    # Crear la tabla visible. P se muestra después de N; V conserva prioridad visual,
    # pero ya se comprobó que no oculta turnos asignados.
    df = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in dias]).fillna("")
    for p in INTEGRANTES:
        vacaciones_p = {int(v) for v in (vacaciones_dict.get(p, []) or []) if str(v).strip().isdigit()}
        for d in dias:
            ds = str(d)
            wd = datetime.date(ano, mes, d).weekday()
            if d in vacaciones_p:
                df.at[p, ds] = "V"
            elif d == 1 and hist_noche.get(p, False):
                df.at[p, ds] = "P"
            elif d > 1 and asignaciones[p, d - 1] == "N":
                df.at[p, ds] = "P"
            elif asignaciones[p, d]:
                df.at[p, ds] = asignaciones[p, d]
            else:
                req = _obtener_solicitud(sugerencias_dict, p, d)
                if wd in set(config_dict.get(p, []) or []) or req == "L":
                    df.at[p, ds] = "L"
                elif (p == "JUAN CAMILO PEREZ" and wd in (1, 3)) or (
                    p in ("ERNESTO MUSKUS", "IVETTE VALENCIA") and wd == 5
                ) or (p == "IVETTE VALENCIA" and wd == 0) or (
                    p == "GINELAP" and wd in (3, 4)
                ) or (p in ("GERLIS DOMINGUEZ", "ZARIANA REYES") and wd == 0) or (
                    p == "JHON RIOS" and wd == 1
                ) or (p == "ANGIE BERNAL" and wd == 2):
                    df.at[p, ds] = "L"
                else:
                    df.at[p, ds] = "D"

    # Reporte explícito de solicitudes no satisfechas.
    incumplimientos = []
    for (p, d), y in incumple.items():
        if (value(y) or 0) > 0.5:
            req = _obtener_solicitud(sugerencias_dict, p, d)
            asignado = asignaciones[p, d]
            if d in {int(v) for v in (vacaciones_dict.get(p, []) or []) if str(v).strip().isdigit()}:
                motivo = "Coincide con vacaciones obligatorias"
            elif req == "L" and asignado:
                motivo = f"Se necesitó cubrir un turno ({asignado})"
            elif req in {"C", "N"} and asignado and asignado != req:
                motivo = f"Se asignó {asignado} en lugar de {req}"
            elif req in {"C", "N"} and not asignado:
                motivo = "No fue posible asignar el turno solicitado sin romper reglas obligatorias"
            else:
                motivo = "Conflicto con cobertura u otras solicitudes de mayor prioridad"
            incumplimientos.append({
                "Prioridad": prioridad_persona[p],
                "Integrante": p,
                "Día": d,
                "Solicitud": req,
                "Asignación": asignado or df.at[p, str(d)],
                "Motivo": motivo,
            })
    incumplimientos.sort(key=lambda r: (r["Prioridad"], r["Día"], r["Integrante"]))

    df["TOTAL CORRIDOS"] = [sum(asignaciones[p, d] == "C" for d in dias) for p in INTEGRANTES]
    df["TOTAL NOCHES"] = [sum(asignaciones[p, d] == "N" for d in dias) for p in INTEGRANTES]
    df["TOTAL TURNOS"] = df["TOTAL CORRIDOS"] + df["TOTAL NOCHES"]
    df["FINES DE SEMANA"] = [
        sum(
            asignaciones[p, d] in turnos
            and (datetime.date(ano, mes, d).weekday() >= 5 or es_festivo_co(d, mes, ano))
            for d in dias
        )
        for p in INTEGRANTES
    ]

    patrones = [
        {"Integrante": p, "Día primera noche": d, "Día segunda noche": d + 2}
        for p in INTEGRANTES
        for d in range(1, dias_mes - 1)
        if asignaciones[p, d] == "N" and asignaciones[p, d + 2] == "N"
    ]
    for p in INTEGRANTES:
        if hist_noche.get(p, False) and dias_mes >= 2 and asignaciones[p, 2] == "N":
            patrones.append({"Integrante": p, "Día primera noche": "Mes anterior", "Día segunda noche": 2})
    df.attrs["incumplimientos"] = incumplimientos
    df.attrs["patrones_n_p_n_p"] = patrones
    reporte_carga = []
    for persona in INTEGRANTES:
        objetivo_turnos, objetivo_noches, disponibles, disponibles_noche = objetivos_carga[persona]
        total_turnos = int(df.at[persona, "TOTAL TURNOS"])
        total_noches = int(df.at[persona, "TOTAL NOCHES"])
        reporte_carga.append({
            "Integrante": persona,
            "Días disponibles para turnos": disponibles,
            "Días disponibles para noche": disponibles_noche,
            "Objetivo turnos aprox.": round(objetivo_turnos, 1),
            "Turnos asignados": total_turnos,
            "Objetivo noches aprox.": round(objetivo_noches, 1),
            "Noches asignadas": total_noches,
        })
    df.attrs["reporte_carga"] = reporte_carga
    df.attrs["diagnostico"] = {
        "estado": estado,
        "solicitudes_incumplidas": len(incumplimientos),
        "patrones_n_p_n_p": len(patrones),
        "regla_carga": "Objetivos proporcionales suaves; cobertura y reglas duras tienen prioridad",
    }
    return df
