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


def es_festivo_co(dia, mes, ano):
    try:
        import holidays
        return datetime.date(ano, mes, dia) in holidays.CO(years=ano)
    except Exception:
        return False


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

    Optimizado para evitar bloqueos mediante penalización directa de fines de semana en lugar de tasas continuas.
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

    # Cada solicitud tiene una variable de incumplimiento.
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
                prob += x[p][d]["C"] + x[p][d]["N"] <= y, f"solicitud_libre_{p.replace(' ', '_')}_{d}"
            else:
                prob += x[p][d][req] + y >= 1, f"solicitud_turno_{p.replace(' ', '_')}_{d}"

    # Reglas obligatorias de empalme entre meses.
    for p in INTEGRANTES:
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

        for p in INTEGRANTES:
            vacaciones_p = {int(v) for v in (vacaciones_dict.get(p, []) or []) if str(v).strip().isdigit()}
            if d in vacaciones_p:
                prob += x[p][d]["C"] == 0, f"vacacion_C_{p.replace(' ', '_')}_{d}"
                prob += x[p][d]["N"] == 0, f"vacacion_N_{p.replace(' ', '_')}_{d}"

            prob += x[p][d]["C"] + x[p][d]["N"] <= 1, f"max_un_turno_{p.replace(' ', '_')}_{d}"

            if d < dias_mes:
                prob += (
                    x[p][d]["N"] + x[p][d + 1]["C"] + x[p][d + 1]["N"] <= 1,
                    f"descanso_post_noche_{p.replace(' ', '_')}_{d}"
                )

            if d <= dias_mes - 3:
                prob += lpSum(
                    x[p][d + k][t] for k in range(4) for t in turnos
                ) <= 3, f"max_3_en_4_{p.replace(' ', '_')}_{d}"

        req_iv = _obtener_solicitud(sugerencias_dict, "IVETTE VALENCIA", d)
        if wd == 0 and not req_iv in {"C", "N"}:
            prob += x["IVETTE VALENCIA"][d]["C"] + x["IVETTE VALENCIA"][d]["N"] == 0, f"ivette_lunes_{d}"

        req_jh = _obtener_solicitud(sugerencias_dict, "JHON RIOS", d)
        if wd == 1 and req_jh != "C":
            prob += x["JHON RIOS"][d]["C"] == 0, f"jhon_martes_C_{d}"

        req_an = _obtener_solicitud(sugerencias_dict, "ANGIE BERNAL", d)
        if wd == 2 and req_an != "N":
            prob += x["ANGIE BERNAL"][d]["N"] == 0, f"angie_miercoles_N_{d}"

        if wd in (3, 4):
            prob += x["GINELAP"][d]["C"] + x["GINELAP"][d]["N"] == 0, f"ginelap_jueves_viernes_{d}"

    # Entre 4 y 5 noches por integrante al mes.
    for p in INTEGRANTES:
        total_n = lpSum(x[p][d]["N"] for d in dias)
        prob += total_n >= 4, f"min_noches_{p.replace(' ', '_')}"
        prob += total_n <= 5, f"max_noches_{p.replace(' ', '_')}"

    # Máximo 2 lunes trabajados para Gerlis y Zariana.
    lunes = [d for d in dias if datetime.date(ano, mes, d).weekday() == 0]
    for p in ("GERLIS DOMINGUEZ", "ZARIANA REYES"):
        prob += lpSum(x[p][d][t] for d in lunes for t in turnos) <= 2, f"max_dos_lunes_{p.replace(' ', '_')}"

    # 1) Optimización lexicográfica por prioridades de solicitudes.
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
                    f"Estado de PuLP: {estado}."
                )
            mejor_valor = int(round(value(objetivo_nivel) or 0))
            prob += objetivo_nivel == mejor_valor, f"fijar_prioridad_{indice}_{nivel}"

    # 2) Equidad ULTRARRÁPIDA de fines de semana y preferencias flexibles.
    dias_fin_semana = [
        d for d in dias
        if datetime.date(ano, mes, d).weekday() >= 5 or es_festivo_co(d, mes, ano)
    ]

    penaliza_patron = {}
    for p in INTEGRANTES:
        for d in range(1, dias_mes - 1):
            z = LpVariable(f"patron_NPNP_{p.replace(' ', '_')}_{d}", cat=LpBinary)
            penaliza_patron[p, d] = z
            prob += z >= x[p][d]["N"] + x[p][d + 2]["N"] - 1, f"detectar_patron_NPNP_{p.replace(' ', '_')}_{d}"

    pesos_totales = []
    for p in INTEGRANTES:
        dias_fijos = set(config_dict.get(p, []) or [])
        vacaciones_p = {int(v) for v in (vacaciones_dict.get(p, []) or []) if str(v).strip().isdigit()}
        
        for d in dias:
            wd = datetime.date(ano, mes, d).weekday()
            es_fs = d in dias_fin_semana
            req = _obtener_solicitud(sugerencias_dict, p, d)
            
            for t in turnos:
                if req:
                    peso = 0
                elif wd in dias_fijos:
                    peso = -10000
                elif (p == "JUAN CAMILO PEREZ" and wd in (1, 3)) or (
                    p in ("ERNESTO MUSKUS", "IVETTE VALENCIA") and wd == 5
                ):
                    peso = -50
                elif es_fs:
                    peso = -5
                else:
                    peso = 1
                pesos_totales.append(peso * x[p][d][t])

    objetivo_final = lpSum(pesos_totales) - 200 * lpSum(penaliza_patron.values())
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
        if not 4 <= noches <= 5:
            errores_validacion.append(f"{p}: {noches} noches (debe tener entre 4 y 5)")
        for d in dias[:-1]:
            if asignaciones[p, d] == "N" and asignaciones[p, d + 1] in {"C", "N"}:
                errores_validacion.append(f"{p}: no descansa después de N en día {d}")
        for d in range(1, dias_mes - 2):
            total_4 = sum(bool(asignaciones[p, d + k]) for k in range(4))
            if total_4 > 3:
                errores_validacion.append(f"{p}: más de 3 turnos en los días {d}-{d + 3}")

    for d in dias:
        wd = datetime.date(ano, mes, d).weekday()
        req_iv = _obtener_solicitud(sugerencias_dict, "IVETTE VALENCIA", d)
        if wd == 0 and req_iv not in {"C", "N"} and asignaciones["IVETTE VALENCIA", d]:
            errores_validacion.append(f"Ivette: turno prohibido el lunes {d}")
        req_jh = _obtener_solicitud(sugerencias_dict, "JHON RIOS", d)
        if wd == 1 and req_jh != "C" and asignaciones["JHON RIOS", d] == "C":
            errores_validacion.append(f"Jhon: C prohibido el martes {d}")
        req_an = _obtener_solicitud(sugerencias_dict, "ANGIE BERNAL", d)
        if wd == 2 and req_an != "N" and asignaciones["ANGIE BERNAL", d] == "N":
            errores_validacion.append(f"Angie: N prohibida el miércoles {d}")
        if wd in (3, 4) and asignaciones["GINELAP", d]:
            errores_validacion.append(f"Ginelap: turno prohibido el día {d} (jueves/viernes obligatorios libres)")

    for p in ("GERLIS DOMINGUEZ", "ZARIANA REYES"):
        lunes_trabajados = sum(
            bool(asignaciones[p, d]) for d in dias
            if datetime.date(ano, mes, d).weekday() == 0
        )
        if lunes_trabajados > 2:
            errores_validacion.append(f"{p}: {lunes_trabajados} lunes trabajados (máximo 2)")

    if errores_validacion:
        raise RuntimeError("La validación independiente falló: " + "; ".join(errores_validacion[:12]))

    # Crear la tabla visible.
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
                ) or (p in ("GERLIS DOMINGUEZ", "ZARIANA REYES") and wd == 0):
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
    df.attrs["incumplimientos"] = incumplimientos
    df.attrs["patrones_n_p_n_p"] = patrones
    df.attrs["diagnostico"] = {
        "estado": estado,
        "solicitudes_incumplidas": len(incumplimientos),
        "patrones_n_p_n_p": len(patrones),
    }
    return df
