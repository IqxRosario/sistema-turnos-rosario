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
    import holidays
    co_holidays = holidays.CO(years=ano)
    return datetime.date(ano, mes, dia) in co_holidays

def generar_cuadro_pulp(mes, ano, historial_previo, sugerencias_dict, config_dict, vacaciones_dict):
    dias_mes = calendar.monthrange(ano, mes)[1]
    dias = list(range(1, dias_mes + 1))
    turnos_asig = ['C', 'N']
    
    # 1. Crear problema de optimización
    prob = pulp.LpProblem("Optimizacion_Turnos_Hospital", pulp.LpMaximize)
    
    # Variables binarias
    x = pulp.LpVariable.dicts("x", (INTEGRANTES, dias, turnos_asig), cat=pulp.LpBinary)
    
    # 2. Matriz de Pesos/Preferencias
    pesos = {}
    for p in INTEGRANTES:
        for d in dias:
            ds = str(d)
            wd = datetime.date(ano, mes, d).weekday()
            
            es_vacacion = d in vacaciones_dict.get(p, [])
            es_libre_fijo = (wd in config_dict.get(p, [])) and not (p == "JUAN CAMILO PEREZ" and wd == 3)
            
            for t in turnos_asig:
                if es_vacacion or es_libre_fijo:
                    pesos[p, d, t] = -10000
                else:
                    req = sugerencias_dict.get(p, {}).get(ds, "")
                    if req and t in req:
                        pesos[p, d, t] = 100
                    else:
                        pesos[p, d, t] = 1
                        
    prob += pulp.lpSum([x[p][d][t] * pesos[p, d, t] for p in INTEGRANTES for d in dias for t in turnos_asig])
    
    # 3. Restricciones
    for d in dias:
        wd = datetime.date(ano, mes, d).weekday()
        es_fest = es_festivo_co(d, mes, ano)
        
        cuota_n = 2
        cuota_c = 2 if (es_fest or wd == 6) else (5 if wd == 5 else 6)
        
        prob += pulp.lpSum([x[p][d]['N'] for p in INTEGRANTES]) == cuota_n
        prob += pulp.lpSum([x[p][d]['C'] for p in INTEGRANTES]) == cuota_c
        
        for p in INTEGRANTES:
            prob += pulp.lpSum([x[p][d][t] for t in turnos_asig]) <= 1
            
            if d < dias_mes:
                prob += x[p][d]['N'] + x[p][d+1]['C'] + x[p][d+1]['N'] <= 1
                
            if d <= dias_mes - 3:
                prob += (pulp.lpSum([x[p][d+k][t] for k in range(4) for t in turnos_asig]) <= 3)

    # 4. Resolver
    solver = pulp.PULP_CBC_CMD(msg=False)
    prob.solve(solver)
    
    # 5. Mapeo a DataFrame
    df_res = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in dias]).fillna("")
    
    for p in INTEGRANTES:
        for d in dias:
            ds = str(d)
            wd = datetime.date(ano, mes, d).weekday()
            
            if d in vacaciones_dict.get(p, []):
                df_res.at[p, ds] = 'V'
            elif (wd in config_dict.get(p, [])) and not (p == "JUAN CAMILO PEREZ" and wd == 3):
                df_res.at[p, ds] = 'L'
            else:
                if d > 1 and df_res.at[p, str(d-1)] == 'N':
                    df_res.at[p, ds] = 'P'
                else:
                    asignado = False
                    for t in turnos_asig:
                        if pulp.value(x[p][d][t]) == 1:
                            df_res.at[p, ds] = t
                            asignado = True
                            break
                    if not asignado:
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
