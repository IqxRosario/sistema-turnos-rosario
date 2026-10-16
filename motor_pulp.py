import calendar
import datetime
import pandas as pd
import pulp

INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES ARGUMEDO", 
    "KELLY CAUSIL", "JOAN CARMONA"
]

def es_festivo_co(dia, mes, ano):
    try:
        import holidays
        co_holidays = holidays.CO(years=ano)
        return datetime.date(ano, mes, dia) in co_holidays
    except Exception:
        return False

def generar_cuadro_pulp(mes, ano, hist_noche, hist_racha, sugerencias_dict, config_dict, vacaciones_dict, prioridades_dict=None):
    if prioridades_dict is None:
        prioridades_dict = {p: 99 for p in INTEGRANTES}
        
    dias_mes = calendar.monthrange(ano, mes)[1]
    dias = list(range(1, dias_mes + 1))
    turnos_asig = ['C', 'N']
    
    # 1. Crear problema de optimización
    prob = pulp.LpProblem("Optimizacion_Turnos_Hospital", pulp.LpMaximize)
    
    # Diccionario explícito de variables binarias (evita errores de sintaxis y espacios)
    x = {
        p: {
            d: {
                t: pulp.LpVariable(f"x_{p.replace(' ', '_')}_{d}_{t}", cat='Binary')
                for t in turnos_asig
            }
            for d in dias
        }
        for p in INTEGRANTES
    }
    
    # 2. Matriz de Pesos/Preferencias (Escalonadas por Orden de Llegada)
    pesos = {}
    for p in INTEGRANTES:
        orden = prioridades_dict.get(p, 99)
        
        # Ponderación dinámica según número de llegada
        puntos_peticion_turno = max(100, 1000 - (orden * 10))
        penalizacion_peticion_libre = min(-300, -1000 + (orden * 10))
        
        for d in dias:
            ds = str(d)
            wd = datetime.date(ano, mes, d).weekday()
            
            es_vacacion = d in vacaciones_dict.get(p, [])
            req = sugerencias_dict.get(p, {}).get(ds, "")
            
            for t in turnos_asig:
                if es_vacacion:
                    pesos[p, d, t] = -10000
                elif req == 'L':
                    pesos[p, d, t] = penalizacion_peticion_libre
                elif req and t in req:
                    pesos[p, d, t] = puntos_peticion_turno
                else:
                    peso_base = 1
                    
                    # Comodines
                    if p == "JUAN CAMILO PEREZ" and wd in [1, 3]: # Martes y Jueves
                        peso_base = -50
                    elif p == "ERNESTO MUSKUS" and wd == 5:       # Sábado
                        peso_base = -50
                    elif p == "IVETTE VALENCIA" and wd == 5:      # Sábado
                        peso_base = -50
                    elif wd in config_dict.get(p, []):
                        peso_base = -10000
                        
                    pesos[p, d, t] = peso_base
                        
    prob += pulp.lpSum([x[p][d][t] * pesos[p, d, t] for p in INTEGRANTES for d in dias for t in turnos_asig])
    
    # 3. Restricciones de Empalme con el Mes Anterior
    for p in INTEGRANTES:
        if hist_noche.get(p, False):
            prob += x[p][1]['C'] == 0
            prob += x[p][1]['N'] == 0
            
        racha_prev = hist_racha.get(p, 0)
        if racha_prev >= 3:
            prob += x[p][1]['C'] == 0
            prob += x[p][1]['N'] == 0
        elif racha_prev == 2:
            prob += x[p][1]['C'] + x[p][1]['N'] + x[p][2]['C'] + x[p][2]['N'] <= 1

    # 4. Restricciones mensuales estándar y REGLAS FIJAS DEL EQUIPO
    for d in dias:
        ds = str(d)
        wd = datetime.date(ano, mes, d).weekday()
        es_fest = es_festivo_co(d, mes, ano)
        
        cuota_n = 2
        cuota_c = 2 if (es_fest or wd == 6) else (5 if wd == 5 else 6)
        
        prob += pulp.lpSum([x[p][d]['N'] for p in INTEGRANTES]) == cuota_n
        prob += pulp.lpSum([x[p][d]['C'] for p in INTEGRANTES]) == cuota_c
        
        # --- REGLAS ESTRICTAS DE PROHIBICIÓN ---
        req_iv = sugerencias_dict.get("IVETTE VALENCIA", {}).get(ds, "")
        if wd == 0 and not any(t in req_iv for t in turnos_asig):
            prob += x["IVETTE VALENCIA"][d]['C'] + x["IVETTE VALENCIA"][d]['N'] == 0
            
        req_jh = sugerencias_dict.get("JHON RIOS", {}).get(ds, "")
        if wd == 1 and not ('C' in req_jh):
            prob += x["JHON RIOS"][d]['C'] == 0
            
        req_an = sugerencias_dict.get("ANGIE BERNAL", {}).get(ds, "")
        if wd == 2 and not ('N' in req_an):
            prob += x["ANGIE BERNAL"][d]['N'] == 0
            
        req_gi = sugerencias_dict.get("GINELAP", {}).get(ds, "")
        if wd in [3, 4] and not any(t in req_gi for t in turnos_asig):
            prob += x["GINELAP"][d]['C'] + x["GINELAP"][d]['N'] == 0
        
        # Lógica general de turnos
        for p in INTEGRANTES:
            prob += pulp.lpSum([x[p][d][t] for t in turnos_asig]) <= 1
            
            if d < dias_mes:
                prob += x[p][d]['N'] + x[p][d+1]['C'] + x[p][d+1]['N'] <= 1
                
            if d <= dias_mes - 3:
                prob += (pulp.lpSum([x[p][d+k][t] for k in range(4) for t in turnos_asig]) <= 3)

    # Máximo 2 Lunes al mes para Gerlis y Zariana
    lunes_del_mes = [d for d in dias if datetime.date(ano, mes, d).weekday() == 0]
    for p in ["GERLIS DOMINGUEZ", "ZARIANA REYES"]:
        prob += pulp.lpSum([x[p][d][t] for d in lunes_del_mes for t in turnos_asig]) <= 2

    # 5. Resolver con manejo seguro de solver
    try:
        solver = pulp.PULP_CBC_CMD(msg=False)
        prob.solve(solver)
    except Exception:
        prob.solve()
    
    # 6. Mapeo a DataFrame
    df_res = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in dias]).fillna("")
    
    for p in INTEGRANTES:
        for d in dias:
            ds = str(d)
            wd = datetime.date(ano, mes, d).weekday()
            
            if d in vacaciones_dict.get(p, []):
                df_res.at[p, ds] = 'V'
            elif d == 1 and hist_noche.get(p, False):
                df_res.at[p, ds] = 'P'
            elif d > 1 and df_res.at[p, str(d-1)] == 'N':
                df_res.at[p, ds] = 'P'
            else:
                asignado = False
                for t in turnos_asig:
                    val = pulp.value(x[p][d][t])
                    if val is not None and val > 0.5:
                        df_res.at[p, ds] = t
                        asignado = True
                        break
                
                if not asignado:
                    req = sugerencias_dict.get(p, {}).get(ds, "")
                    
                    es_comodin = (p == "JUAN CAMILO PEREZ" and wd in [1, 3]) or \
                                 (p == "ERNESTO MUSKUS" and wd == 5) or \
                                 (p == "IVETTE VALENCIA" and wd == 5)
                                 
                    es_intocable = (p == "IVETTE VALENCIA" and wd == 0) or \
                                   (p == "GINELAP" and wd in [3, 4])
                                   
                    es_lunes_gz = (p in ["GERLIS DOMINGUEZ", "ZARIANA REYES"] and wd == 0)
                    
                    if wd in config_dict.get(p, []):
                        df_res.at[p, ds] = 'L'
                    elif req == 'L':
                        df_res.at[p, ds] = 'L'
                    elif es_comodin or es_intocable or es_lunes_gz:
                        df_res.at[p, ds] = 'L'
                    else:
                        df_res.at[p, ds] = 'D'
                        
    df_res['TOTAL CORRIDOS'] = df_res.apply(lambda r: sum(1 for c in df_res.columns if 'C' in str(r[c]) and c.isdigit()), axis=1)
    df_res['TOTAL NOCHES'] = df_res.apply(lambda r: sum(1 for c in df_res.columns if 'N' in str(r[c]) and c.isdigit()), axis=1)
    df_res['TOTAL TURNOS'] = df_res['TOTAL CORRIDOS'] + df_res['TOTAL NOCHES']
    df_res['FINES DE SEMANA'] = df_res.apply(
        lambda r: sum(1 for d in range(1, dias_mes+1) 
                      if (datetime.date(ano, mes, d).weekday() >= 5 or es_festivo_co(d, mes, ano)) 
                      and any(t in str(r[str(d)]) for t in ['C', 'N'])), axis=1
    )
    
    return df_res
