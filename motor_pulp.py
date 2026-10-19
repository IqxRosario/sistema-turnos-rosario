import pulp
import pandas as pd
import numpy as np
import calendar

def resolver_turnos(df_sugerencias, mes=11, ano=2026):
    """
    Optimizador PuLP para asignación de turnos.
    Incorpora reglas duras de cobertura, descanso, reglas de equipo,
    y límites de equidad manteniendo alta velocidad de ejecución (<2s).
    """
    num_dias = calendar.monthrange(ano, mes)[1]
    dias = list(range(1, num_dias + 1))
    
    # -----------------------------------------------------------------
    # MAPEO DE DÍAS DE LA SEMANA (Noviembre 2026: Día 1 es Domingo)
    # -----------------------------------------------------------------
    nombres_dias_semana = ['LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO']
    dias_semana_mapeo = {d: [] for d in nombres_dias_semana}
    
    primer_dia_semana, _ = calendar.monthrange(ano, mes) # 0: Lun, 6: Dom
    
    for d in dias:
        dia_idx = (primer_dia_semana + d - 1) % 7
        nombre_dia = nombres_dias_semana[dia_idx]
        dias_semana_mapeo[nombre_dia].append(d)

    # Mantener el orden original exacto de integrantes de la planilla de entrada
    integrantes = [str(n).strip().upper() for n in df_sugerencias['NOMBRE'].tolist()]
    
    modelo = pulp.LpProblem("Optimizacion_Turnos_Hospital", pulp.LpMinimize)
    tipos_turno = ['C', 'N', 'P', 'D', 'V', 'L']
    
    # Variable de decisión binaria x[persona, día, tipo_turno]
    x = pulp.LpVariable.dicts("x", ((p, d, t) for p in integrantes for d in dias for t in tipos_turno), cat='Binary')
    
    # -----------------------------------------------------------------
    # RESTRICCIONES DURA Y SECUENCIAS OBLIGATORIAS
    # -----------------------------------------------------------------
    # 1. Un solo turno por persona al día
    for p in integrantes:
        for d in dias:
            modelo += pulp.lpSum([x[p, d, t] for t in tipos_turno]) == 1, f"Un_turno_{p}_{d}"

    # 2. Secuencia Noche -> Post-turno (N en d obliga P en d+1)
    for p in integrantes:
        for d in range(1, num_dias):
            modelo += x[p, d + 1, 'P'] == x[p, d, 'N'], f"Post_turno_obligatorio_{p}_{d}"

    # 3. Prohibición estricta del patrón agotador N-P-N
    for p in integrantes:
        for d in range(2, num_dias + 1):
            modelo += x[p, d, 'N'] + x[p, d - 1, 'P'] <= 1, f"No_NPN_{p}_{d}"

    # 4. Cobertura mínima diaria del hospital
    for d in dias:
        modelo += pulp.lpSum([x[p, d, 'N'] for p in integrantes]) == 2, f"Cobertura_Noche_{d}"
        modelo += pulp.lpSum([x[p, d, 'C'] for p in integrantes]) >= 4, f"Cobertura_Corrido_{d}"

    # 5. Rango estricto de noches y EQUILIBRIO DE TURNOS TOTALES (16 a 18 turnos)
    for p in integrantes:
        row_p = df_sugerencias[df_sugerencias['NOMBRE'].str.strip().str.upper() == p].iloc[0]
        tiene_vac = pd.notna(row_p.get('VACACIONES'))
        
        if not tiene_vac:
            noches_total = pulp.lpSum([x[p, d, 'N'] for d in dias])
            modelo += noches_total >= 4, f"Min_noches_{p}"
            modelo += noches_total <= 5, f"Max_noches_{p}"
            
            # EQUILIBRIO STRICTO: 16 a 18 turnos de trabajo totales (C + N) al mes
            turnos_totales = pulp.lpSum([x[p, d, 'C'] + x[p, d, 'N'] for d in dias])
            modelo += turnos_totales >= 16, f"Min_turnos_equitativos_{p}"
            modelo += turnos_totales <= 18, f"Max_turnos_equitativos_{p}"

    # 6. Vacaciones continuas (ej. Juan Camilo Pérez del 5 al 24)
    for index, row in df_sugerencias.iterrows():
        p = str(row['NOMBRE']).strip().upper()
        if pd.notna(row.get('VACACIONES')):
            vac_str = str(row['VACACIONES']).strip()
            if '-' in vac_str:
                inicio, fin = [int(v.strip()) for v in vac_str.split('-')]
                for d in range(inicio, fin + 1):
                    modelo += x[p, d, 'V'] == 1, f"Vacaciones_{p}_{d}"

    # -----------------------------------------------------------------
    # REGLAS INDIVIDUALES DE EQUIPO (NOCHES PROHIBIDAS POR SALUD)
    # -----------------------------------------------------------------
    for p in integrantes:
        # JHON RIOS: Cero noches los Lunes (evita choque con Pineda Lunes y Martes)
        if p == "JHON RIOS":
            for d in dias_semana_mapeo['LUNES']:
                modelo += x[p, d, 'N'] == 0, f"Jhon_no_noche_lunes_{d}"
                
        # GINELAP: Cero noches los Miércoles (garantiza estar de día Jueves con su compañera)
        if p == "GINELAP":
            for d in dias_semana_mapeo['MIERCOLES']:
                modelo += x[p, d, 'N'] == 0, f"Ginelap_no_noche_miercoles_{d}"
                
        # GERLIS DOMINGUEZ: Cero noches los Domingos (evita cansancio en cirugías de Lunes)
        if p == "GERLIS DOMINGUEZ":
            for d in dias_semana_mapeo['DOMINGO']:
                modelo += x[p, d, 'N'] == 0, f"Gerlis_no_noche_domingo_{d}"

    # -----------------------------------------------------------------
    # PARSER DE SOLICITUDES ESPECÍFICAS (C, N, L Y LIBRES REALES)
    # -----------------------------------------------------------------
    for index, row in df_sugerencias.iterrows():
        p = str(row['NOMBRE']).strip().upper()
        if pd.notna(row.get('LIBRES SOLICITADOS (Mes)')):
            solicitudes = str(row['LIBRES SOLICITADOS (Mes)']).split(',')
            for req in solicitudes:
                req = req.strip().upper()
                if not req: continue
                
                # Solicitud de Noche explícita (ej. "9N", "23N")
                if req.endswith('N'):
                    d = int(req.replace('N', ''))
                    modelo += x[p, d, 'N'] == 1, f"Forzar_Noche_Solicitada_{p}_{d}"
                    
                # Solicitud de Corrido explícito (ej. "29C" de Ernesto Muskus)
                elif req.endswith('C'):
                    d = int(req.replace('C', ''))
                    modelo += x[p, d, 'C'] == 1, f"Forzar_Corrido_Solicitado_{p}_{d}"
                    
                # Solicitud de Día Libre puro (ej. "13", "21")
                elif req.isdigit():
                    d = int(req)
                    modelo += x[p, d, 'L'] == 1, f"Forzar_Libre_Solicitado_{p}_{d}"
                    # BLOQUEO CLAVE: Noche prohibida el día d-1 para que sea un Libre real sin Post-turno 'P'
                    if d > 1:
                        modelo += x[p, d - 1, 'N'] == 0, f"Bloqueo_Noche_Previa_Libre_{p}_{d}"

    # -----------------------------------------------------------------
    # LÍMITE FLEXIBLE DE EXCEPCIONES EN DÍAS FIJOS
    # -----------------------------------------------------------------
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
                        max_exc = 2  # Límite flexible base
                        
                        # REGLA INTOCABLE: Los Lunes de Ivette Valencia no tienen excepciones (0)
                        if p == "IVETTE VALENCIA" and d_fijo == "LUNES":
                            max_exc = 0
                            
                        modelo += pulp.lpSum(excepciones) <= max_exc, f"Max_exc_fijo_{p}_{d_fijo}"

    # Resolver con solver CBC
    solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=10)
    estado = modelo.solve(solver)
    
    # Reconstrucción del cuadro preservando el orden original
    matriz_res = []
    for p in integrantes:
        row_res = {'NOMBRE': p}
        for d in dias:
            for t in tipos_turno:
                if pulp.value(x[p, d, t]) is not None and pulp.value(x[p, d, t]) > 0.5:
                    row_res[d] = t
                    break
        matriz_res.append(row_res)
        
    df_res = pd.DataFrame(matriz_res)
    df_res.attrs['status'] = pulp.LpStatus[estado]
    return df_res
