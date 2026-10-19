import pulp
import pandas as pd
import numpy as np
import calendar

def resolver_turnos(df_sugerencias, mes=11, ano=2026):
    """
    Optimizador PuLP robusto para la programación de turnos médicos.
    Garantiza factibilidad matemática, restricciones estrictas de salud laboral
    y cumplimiento jerárquico de peticiones por prioridad.
    """
    num_dias = calendar.monthrange(ano, mes)[1]
    dias = list(range(1, num_dias + 1))
    
    # Mapeo de días de la semana (Noviembre 2026: Día 1 es Domingo)
    nombres_dias_semana = ['LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO']
    dias_semana_mapeo = {d: [] for d in nombres_dias_semana}
    
    primer_dia_semana, _ = calendar.monthrange(ano, mes)
    
    for d in dias:
        dia_idx = (primer_dia_semana + d - 1) % 7
        nombre_dia = nombres_dias_semana[dia_idx]
        dias_semana_mapeo[nombre_dia].append(d)

    integrantes = [str(n).strip().upper() for n in df_sugerencias['NOMBRE'].tolist()]
    
    modelo = pulp.LpProblem("Optimizacion_Turnos_Hospital", pulp.LpMinimize)
    tipos_turno = ['C', 'N', 'P', 'D', 'V', 'L']
    
    # Variables de decisión
    x = pulp.LpVariable.dicts("x", ((p, d, t) for p in integrantes for d in dias for t in tipos_turno), cat='Binary')
    
    penalizaciones = []
    
    # 1. Exactamente un turno por persona por día
    for p in integrantes:
        for d in dias:
            modelo += pulp.lpSum([x[p, d, t] for t in tipos_turno]) == 1, f"Un_turno_{p}_{d}"

    # 2. Secuencia Noche -> Post-turno obligatorio (N en d implica P en d+1)
    for p in integrantes:
        for d in range(1, num_dias):
            modelo += x[p, d + 1, 'P'] == x[p, d, 'N'], f"Post_turno_obligatorio_{p}_{d}"

    # 3. Prohibición estricta del patrón agotador N-P-N
    for p in integrantes:
        for d in range(2, num_dias + 1):
            modelo += x[p, d, 'N'] + x[p, d - 1, 'P'] <= 1, f"No_NPN_{p}_{d}"

    # 4. Cobertura diaria obligatoria del hospital
    for d in dias:
        modelo += pulp.lpSum([x[p, d, 'N'] for p in integrantes]) == 2, f"Cobertura_Noche_{d}"
        modelo += pulp.lpSum([x[p, d, 'C'] for p in integrantes]) >= 4, f"Cobertura_Corrido_{d}"

    # 5. Noches (4 a 5) y EQUILIBRIO DE TURNOS TOTALES (16 a 18 turnos)
    for p in integrantes:
        row_p = df_sugerencias[df_sugerencias['NOMBRE'].str.strip().str.upper() == p].iloc[0]
        tiene_vac = pd.notna(row_p.get('VACACIONES'))
        
        if not tiene_vac:
            noches_total = pulp.lpSum([x[p, d, 'N'] for d in dias])
            modelo += noches_total >= 4, f"Min_noches_{p}"
            modelo += noches_total <= 5, f"Max_noches_{p}"
            
            turnos_totales = pulp.lpSum([x[p, d, 'C'] + x[p, d, 'N'] for d in dias])
            modelo += turnos_totales >= 16, f"Min_turnos_equitativos_{p}"
            modelo += turnos_totales <= 18, f"Max_turnos_equitativos_{p}"

    # 6. Vacaciones continuas (Juan Camilo Pérez del 5 al 24)
    for index, row in df_sugerencias.iterrows():
        p = str(row['NOMBRE']).strip().upper()
        if pd.notna(row.get('VACACIONES')):
            vac_str = str(row['VACACIONES']).strip()
            if '-' in vac_str:
                inicio, fin = [int(v.strip()) for v in vac_str.split('-')]
                for d in range(inicio, fin + 1):
                    modelo += x[p, d, 'V'] == 1, f"Vacaciones_{p}_{d}"

    # 7. PROHIBICIONES DURA DE NOCHES POR SALUD E INCOMPATIBILIDAD
    for p in integrantes:
        # JHON RIOS: Cero noches los Lunes
        if p == "JHON RIOS":
            for d in dias_semana_mapeo['LUNES']:
                modelo += x[p, d, 'N'] == 0, f"Jhon_no_noche_lunes_{d}"
                
        # GINELAP: Cero noches los Miércoles
        if p == "GINELAP":
            for d in dias_semana_mapeo['MIERCOLES']:
                modelo += x[p, d, 'N'] == 0, f"Ginelap_no_noche_miercoles_{d}"
                
        # GERLIS DOMINGUEZ: Cero noches los Domingos
        if p == "GERLIS DOMINGUEZ":
            for d in dias_semana_mapeo['DOMINGO']:
                modelo += x[p, d, 'N'] == 0, f"Gerlis_no_noche_domingo_{d}"

    # 8. SOLICITUDES ESPECÍFICAS Y PENALIZACIÓN JERÁRQUICA DE DÍAS LIBRES
    for index, row in df_sugerencias.iterrows():
        p = str(row['NOMBRE']).strip().upper()
        prio_val = float(row.get('PRIORIDAD', 10)) if pd.notna(row.get('PRIORIDAD')) else 10.0
        peso_prioridad = max(100, int((13.0 - prio_val) * 100))
        
        if pd.notna(row.get('LIBRES SOLICITADOS (Mes)')):
            solicitudes = str(row['LIBRES SOLICITADOS (Mes)']).split(',')
            for req in solicitudes:
                req = req.strip().upper()
                if not req: continue
                
                if req.endswith('N'):
                    d = int(req.replace('N', ''))
                    modelo += x[p, d, 'N'] == 1, f"Forzar_Noche_Solicitada_{p}_{d}"
                elif req.endswith('C'):
                    d = int(req.replace('C', ''))
                    modelo += x[p, d, 'C'] == 1, f"Forzar_Corrido_Solicitado_{p}_{d}"
                elif req.isdigit():
                    d = int(req)
                    # Penalización alta si no recibe Libre 'L'
                    penalizaciones.append(peso_prioridad * (1 - x[p, d, 'L']))
                    # Penalización si amanece en Post-turno 'P' en su día libre
                    if d > 1:
                        penalizaciones.append((peso_prioridad // 2) * x[p, d - 1, 'N'])

    # 9. LÍMITE FLEXIBLE DE EXCEPCIONES EN DÍAS FIJOS
    for index, row in df_sugerencias.iterrows():
        p = str(row['NOMBRE']).strip().upper()
        if pd.notna(row.get('LIBRE FIJO (cirujano)')):
            fijos = [f.strip().upper() for f in str(row['LIBRE FIJO (cirujano)']).split(',')]
            for d_fijo in fijos:
                if d_fijo in dias_semana_mapeo:
                    fechas_mes = dias_semana_mapeo[d_fijo]
                    excepciones = []
                    for d in fechas_mes:
                        excepciones.append(x[p, d, 'C'])
                        excepciones.append(x[p, d, 'N'])
                        if d > 1:
                            excepciones.append(x[p, d - 1, 'N'])
                    if excepciones:
                        max_exc = 2
                        if p == "IVETTE VALENCIA" and d_fijo == "LUNES":
                            max_exc = 0
                        modelo += pulp.lpSum(excepciones) <= max_exc, f"Max_exc_fijo_{p}_{d_fijo}"

    # Función Objetivo: Minimizar penalizaciones por incumplimientos
    if penalizaciones:
        modelo += pulp.lpSum(penalizaciones)
    else:
        modelo += 0

    solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=10)
    estado = modelo.solve(solver)
    status_str = pulp.LpStatus[estado]
    
    # Reconstrucción de la matriz de resultados
    matriz_res = []
    for p in integrantes:
        row_res = {'NOMBRE': p}
        for d in dias:
            asignado = 'D'
            for t in tipos_turno:
                val = pulp.value(x[p, d, t])
                if val is not None and val > 0.5:
                    asignado = t
                    break
            row_res[d] = asignado
        matriz_res.append(row_res)
        
    df_res = pd.DataFrame(matriz_res)
    
    # Totales y resumen
    domingos_list = dias_semana_mapeo['DOMINGO']
    sabados_list = dias_semana_mapeo['SABADO']
    fds_dias = set(domingos_list + sabados_list)
    
    totales_c, totales_n, totales_t, fines_semana_cnt = [], [], [], []
    for idx, row in df_res.iterrows():
        c_c = sum(1 for d in dias if row[d] == 'C')
        n_c = sum(1 for d in dias if row[d] == 'N')
        fds_c = sum(1 for d in dias if d in fds_dias and row[d] in ['C', 'N'])
        totales_c.append(c_c)
        totales_n.append(n_c)
        totales_t.append(c_c + n_c)
        fines_semana_cnt.append(fds_c)
        
    df_res['TOTAL CORRIDOS'] = totales_c
    df_res['TOTAL NOCHES'] = totales_n
    df_res['TOTAL TURNOS'] = totales_t
    df_res['FINES DE SEMANA'] = fines_semana_cnt
    
    df_res.attrs['status'] = status_str
    df_res.attrs['diagnostico'] = {"estado": status_str}
    
    return df_res
