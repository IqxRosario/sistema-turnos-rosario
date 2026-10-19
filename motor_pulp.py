import pulp
import pandas as pd
import numpy as np
import calendar

INTEGRANTES = [
    "GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO",
    "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR",
    "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES ARGUMEDO", 
    "KELLY CAUSIL", "JOAN CARMONA"
]

def obtener_festivos_colombia(ano, mes):
    """
    Calendario oficial de días festivos en Colombia (Ley Emiliani).
    """
    festivos_map = {
        2024: {
            1: [1, 8], 3: [25, 28, 29], 5: [1, 13], 6: [3, 10],
            7: [1, 20], 8: [7, 19], 10: [14], 11: [4, 11], 12: [8, 25]
        },
        2025: {
            1: [1, 6], 3: [24], 4: [17, 18], 5: [1, 29], 6: [16, 23, 30],
            7: [20], 8: [7, 18], 10: [13], 11: [3, 17], 12: [8, 25]
        },
        2026: {
            1: [1, 12], 3: [23], 4: [2, 3], 5: [18], 6: [8, 15, 29],
            7: [20], 8: [7, 17], 10: [12], 11: [2, 16], 12: [8, 25]
        }
    }
    return festivos_map.get(ano, {}).get(mes, [])

def generar_cuadro_pulp(mes_sel, ano_sel, hist_noche, hist_racha, sug, conf, vacs, prio):
    """
    Motor de optimización matemática PuLP adaptado a app.py.
    Garantiza estricta separación semántica:
    - V: Solo en días de vacaciones.
    - L: Solo en días solicitados como libres o fijos.
    - D: Descansos asignados automáticamente por el sistema.
    """
    num_dias = calendar.monthrange(ano_sel, mes_sel)[1]
    dias = list(range(1, num_dias + 1))
    
    primer_dia_semana, _ = calendar.monthrange(ano_sel, mes_sel)
    festivos_mes = obtener_festivos_colombia(ano_sel, mes_sel)
    
    # Mapeo de días de la semana: 0=Lunes, 1=Martes, 2=Miércoles, 3=Jueves, 4=Viernes, 5=Sábado, 6=Domingo
    dias_semana_mapeo = {i: [] for i in range(7)}
    for d in dias:
        dia_idx = (primer_dia_semana + d - 1) % 7
        dias_semana_mapeo[dia_idx].append(d)

    modelo = pulp.LpProblem("Optimizacion_Turnos_Hospital", pulp.LpMinimize)
    tipos_turno = ['C', 'N', 'P', 'D', 'V', 'L']
    
    # Variables de decisión
    x = pulp.LpVariable.dicts("x", ((p, d, t) for p in INTEGRANTES for d in dias for t in tipos_turno), cat='Binary')
    
    penalizaciones = []
    
    # 1. Un solo turno por persona por día
    for p in INTEGRANTES:
        for d in dias:
            modelo += pulp.lpSum([x[p, d, t] for t in tipos_turno]) == 1, f"Un_turno_{p}_{d}"

    # 2. Empalme con el mes anterior (Día 1)
    for p in INTEGRANTES:
        if hist_noche and hist_noche.get(p, False):
            modelo += x[p, 1, 'P'] == 1, f"Empalme_Post_turno_{p}_1"
        else:
            modelo += x[p, 1, 'P'] == 0, f"No_Post_turno_{p}_1"

    # 3. Secuencia Noche -> Post-turno obligatorio (N en d implica P en d+1)
    for p in INTEGRANTES:
        for d in range(1, num_dias):
            modelo += x[p, d + 1, 'P'] == x[p, d, 'N'], f"Post_turno_obligatorio_{p}_{d}"

    # 4. Prohibición estricta del patrón N-P-N
    for p in INTEGRANTES:
        for d in range(2, num_dias + 1):
            modelo += x[p, d, 'N'] + x[p, d - 1, 'P'] <= 1, f"No_NPN_{p}_{d}"

    # 5. Cobertura con diferenciación de DOMINGOS Y FESTIVOS DE COLOMBIA
    for d in dias:
        dia_idx = (primer_dia_semana + d - 1) % 7
        es_domingo_o_festivo = (dia_idx == 6) or (d in festivos_mes)
        
        req_corridos = 2 if es_domingo_o_festivo else 4
        
        modelo += pulp.lpSum([x[p, d, 'N'] for p in INTEGRANTES]) == 2, f"Cobertura_Noche_{d}"
        modelo += pulp.lpSum([x[p, d, 'C'] for p in INTEGRANTES]) >= req_corridos, f"Cobertura_Corrido_{d}"

    # 6. Vacaciones ESTRICTAS (V=1 solo en días autorizados, V=0 en el resto)
    for p in INTEGRANTES:
        dias_vac = set(vacs.get(p, [])) if vacs else set()
        for d in dias:
            if d in dias_vac:
                modelo += x[p, d, 'V'] == 1, f"Vacaciones_Si_{p}_{d}"
            else:
                modelo += x[p, d, 'V'] == 0, f"Vacaciones_No_{p}_{d}"

    # 7. Restricción estricta de 'L' (L=1 solo si fue solicitado explícitamente, L=0 en el resto)
    for p in INTEGRANTES:
        dict_sug = sug.get(p, {}) if sug else {}
        dias_solicitados_libres = {int(d_str) for d_str, t_code in dict_sug.items() if str(d_str).isdigit() and t_code == 'L'}
        
        for d in dias:
            if d not in dias_solicitados_libres:
                modelo += x[p, d, 'L'] == 0, f"No_L_sin_solicitud_{p}_{d}"

    # 8. Noches (4 a 5) y Equidad de turnos totales (16 a 18) para personas sin vacaciones
    for p in INTEGRANTES:
        dias_vac = vacs.get(p, []) if vacs else []
        if not dias_vac:
            noches_total = pulp.lpSum([x[p, d, 'N'] for d in dias])
            modelo += noches_total >= 4, f"Min_noches_{p}"
            modelo += noches_total <= 5, f"Max_noches_{p}"
            
            turnos_totales = pulp.lpSum([x[p, d, 'C'] + x[p, d, 'N'] for d in dias])
            modelo += turnos_totales >= 16, f"Min_turnos_equitativos_{p}"
            modelo += turnos_totales <= 18, f"Max_turnos_equitativos_{p}"

    # 9. REGLAS DURAS DE SALUD LABORAL E INCOMPATIBILIDADES
    for p in INTEGRANTES:
        if p == "JHON RIOS":
            for d in dias_semana_mapeo[0]:  # Lunes
                modelo += x[p, d, 'N'] == 0, f"Jhon_no_noche_lunes_{d}"
                
        if p == "GINELAP":
            for d in dias_semana_mapeo[2]:  # Miércoles
                modelo += x[p, d, 'N'] == 0, f"Ginelap_no_noche_miercoles_{d}"
                
        if p == "GERLIS DOMINGUEZ":
            for d in dias_semana_mapeo[6]:  # Domingo
                modelo += x[p, d, 'N'] == 0, f"Gerlis_no_noche_domingo_{d}"

    # 10. LIBRES FIJOS Y EXCEPCIONES
    for p in INTEGRANTES:
        dias_fijos = conf.get(p, []) if conf else []
        for dia_wk in dias_fijos:
            fechas_mes = dias_semana_mapeo[dia_wk]
            excepciones = []
            for d in fechas_mes:
                excepciones.append(x[p, d, 'C'])
                excepciones.append(x[p, d, 'N'])
                if d > 1:
                    excepciones.append(x[p, d - 1, 'N'])
            if excepciones:
                max_exc = 2
                if p == "IVETTE VALENCIA" and dia_wk == 0:  # Lunes
                    max_exc = 0
                modelo += pulp.lpSum(excepciones) <= max_exc, f"Max_exc_fijo_{p}_{dia_wk}"

    # 11. SUGERENCIAS Y LIBRES SOLICITADOS CON PONDERACIÓN DE PRIORIDAD
    for p in INTEGRANTES:
        prio_val = float(prio.get(p, 99)) if prio else 99.0
        peso_prioridad = max(100, int((100.0 - prio_val) * 100))
        
        dict_sug = sug.get(p, {}) if sug else {}
        for dia_str, t_code in dict_sug.items():
            if str(dia_str).isdigit():
                d = int(dia_str)
                if 1 <= d <= num_dias:
                    if t_code == 'N':
                        penalizaciones.append(peso_prioridad * (1 - x[p, d, 'N']))
                    elif t_code == 'C':
                        penalizaciones.append(peso_prioridad * (1 - x[p, d, 'C']))
                    elif t_code == 'L':
                        penalizaciones.append(peso_prioridad * (1 - x[p, d, 'L']))
                        if d > 1:
                            penalizaciones.append((peso_prioridad // 2) * x[p, d - 1, 'N'])

    if penalizaciones:
        modelo += pulp.lpSum(penalizaciones)
    else:
        modelo += 0

    solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=15)
    estado = modelo.solve(solver)
    status_str = pulp.LpStatus[estado]
    
    matriz_res = []
    for p in INTEGRANTES:
        row_res = {'NOMBRE': p}
        for d in dias:
            asignado = 'D'
            for t in tipos_turno:
                val = pulp.value(x[p, d, t])
                if val is not None and val > 0.5:
                    asignado = t
                    break
            row_res[str(d)] = asignado
        matriz_res.append(row_res)
        
    df_res = pd.DataFrame(matriz_res)
    
    totales_c, totales_n, totales_t = [], [], []
    for idx, row in df_res.iterrows():
        c_c = sum(1 for d in dias if row[str(d)] == 'C')
        n_c = sum(1 for d in dias if row[str(d)] == 'N')
        totales_c.append(c_c)
        totales_n.append(n_c)
        totales_t.append(c_c + n_c)
        
    df_res['TOTAL CORRIDOS'] = totales_c
    df_res['TOTAL NOCHES'] = totales_n
    df_res['TOTAL TURNOS'] = totales_t
    
    return df_res
    
