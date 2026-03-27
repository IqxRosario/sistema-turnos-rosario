# --- MOTOR CORREGIDO ---
def generar_cuadro_equitativo(mes, ano, historial_previo, sugerencias_dict, config_dict, vacaciones_dict, semilla):
    rng = random.Random(semilla)
    dias_mes = calendar.monthrange(ano, mes)[1]
    df = pd.DataFrame(index=INTEGRANTES, columns=[str(d) for d in range(1, dias_mes + 1)]).fillna("")
    
    # Contadores
    turnos_totales = {p: 0 for p in INTEGRANTES}
    noches_totales = {p: 0 for p in INTEGRANTES}
    finde_totales = {p: 0 for p in INTEGRANTES}

    def turno_en_dia(persona, dia_req):
        if dia_req > 0: return str(df.at[persona, str(dia_req)])
        else:
            idx = 2 + dia_req
            hist = historial_previo.get(persona, ["", "", ""])
            if 0 <= idx < 3: return str(hist[idx])
            return ""

    def racha_actual(persona, d):
        streak = 0
        for past_d in range(d - 1, d - 10, -1):
            t_p = turno_en_dia(persona, past_d)
            if any(t in t_p for t in ['C', 'N']) and 'P' not in t_p: streak += 1
            else: break
        return streak

    # ==============================================================
    # FASE 1: PRE-LLENADO (BLINDAJE DURO EN TODO EL MES)
    # ==============================================================
    for d in range(1, dias_mes + 1):
        ds = str(d); wd = datetime(ano, mes, d).weekday()
        es_finde_o_festivo = wd >= 5 or es_festivo(d, mes, ano)

        for p in INTEGRANTES:
            # 1. Vacaciones (Candado Absoluto)
            if d in vacaciones_dict.get(p, []):
                df.at[p, ds] = 'V'
                continue

            # 2. Sugerencias y Excepciones (Ana se maneja aquí desde el Excel de Sugerencias)
            req = sugerencias_dict.get(p, {}).get(ds)
            if req:
                # Evitar sobrescribir si la sugerencia pide C o N pero el dia es de Vacaciones
                tl = 'C' if ('C' in req and 'P' not in req) else ('N' if ('N' in req and 'P' not in req) else req)
                df.at[p, ds] = tl
                if tl == 'N': noches_totales[p] += 1; turnos_totales[p] += 1
                if tl == 'C': turnos_totales[p] += 1
                if tl in ['C', 'N'] and es_finde_o_festivo: finde_totales[p] += 1
                continue

            # 3. Libres Fijos de Configuración (Gerlis, Zariana, Ginelap)
            if wd in config_dict.get(p, []):
                # Excepción Camilo: Si es martes o jueves libre, verificar si no sugirió C6 antes de poner 'L'
                if p == "JUAN CAMILO PEREZ" and req and 'C' in req:
                     pass # Ya se asignó arriba por sugerencia
                else:
                    df.at[p, ds] = "L"
                continue

            # 4. Asignaciones Fijas (Jhon y Angie)
            if wd == 1 and p == "JHON RIOS" and df.at[p, ds] == "":
                df.at[p, ds] = "N"
                turnos_totales[p] += 1; noches_totales[p] += 1
                if es_finde_o_festivo: finde_totales[p] += 1
            
            if wd == 2 and p == "ANGIE BERNAL" and df.at[p, ds] == "":
                df.at[p, ds] = "C"
                turnos_totales[p] += 1
                if es_finde_o_festivo: finde_totales[p] += 1

    # ==============================================================
    # FASE 2: REPARTICIÓN DIARIA (EVALUANDO POSTURNOS Y BLOQUEOS)
    # ==============================================================
    for d in range(1, dias_mes + 1):
        ds = str(d); wd = datetime(ano, mes, d).weekday()
        es_finde_o_festivo = wd >= 5 or es_festivo(d, mes, ano)

        # Determinar Cuotas Diarias
        cuota_n = 2 - sum(1 for p in INTEGRANTES if 'N' in df.at[p, ds])
        
        if es_finde_o_festivo and (wd == 6 or es_festivo(d, mes, ano)): cuota_c_base = 2
        elif wd == 5: cuota_c_base = 5
        else: cuota_c_base = 6
        cuota_c = cuota_c_base - sum(1 for p in INTEGRANTES if 'C' in df.at[p, ds])

        # 1. Asignar Posturnos ('P') obligatorios del día anterior
        for p in INTEGRANTES:
            if 'N' in turno_en_dia(p, d-1) and df.at[p, ds] == "": 
                df.at[p, ds] = 'P'

        # 2. Descanso obligatorio por racha (3 días de turno)
        for p in INTEGRANTES:
            if df.at[p, ds] == "" and racha_actual(p, d) >= 3:
                df.at[p, ds] = 'D'

        # Función para saber si asignar 'N' hoy arruina un día Libre o Vaca mañana
        def no_puede_hacer_noche(persona, dia_actual):
            if dia_actual < dias_mes:
                turno_manana = str(df.at[persona, str(dia_actual + 1)])
                # Bloqueo General: Nadie hace Noche si mañana está Libre, Vacaciones o sugirió turno C/N
                if any(x in turno_manana for x in ['L', 'V', 'C', 'N']): return True
                
                # Bloqueo Ivette y Marcela: No pueden ser Posturno (P) en sus días libres. 
                # Como ya pre-llenamos los 'L', el chequeo de arriba ('L' in turno_manana) ya las protege.
            return False

        # 3. Repartir Noches (Con Blindaje)
        for _ in range(max(0, cuota_n)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            disp = [p for p in disp if 'N' not in turno_en_dia(p, d-2)] # No dos noches muy seguidas
            disp = [p for p in disp if not no_puede_hacer_noche(p, d)] # BLINDAJE CRUZADO
            
            if disp:
                rng.shuffle(disp)
                if es_finde_o_festivo: disp.sort(key=lambda x: (noches_totales[x], finde_totales[x], turnos_totales[x]))
                else: disp.sort(key=lambda x: (noches_totales[x], turnos_totales[x]))
                
                elegido = disp[0]
                df.at[elegido, ds] = "N"
                turnos_totales[elegido] += 1; noches_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 4. Repartir Corridos
        for _ in range(max(0, cuota_c)):
            disp = [p for p in INTEGRANTES if df.at[p, ds] == ""]
            
            # Bloquear C si racha es 2 y mañana ya tiene un C o N fijo (para evitar 4 turnos seguidos)
            disp_segura = []
            for p in disp:
                if racha_actual(p, d) == 2 and d < dias_mes:
                    t_man = str(df.at[p, str(d+1)])
                    if any(x in t_man for x in ['C', 'N']): continue
                disp_segura.append(p)
            
            if not disp_segura: disp_segura = disp
            
            if disp_segura:
                rng.shuffle(disp_segura)
                if es_finde_o_festivo: disp_segura.sort(key=lambda x: (finde_totales[x], turnos_totales[x]))
                else: disp_segura.sort(key=lambda x: (turnos_totales[x], racha_actual(x, d)))
                
                elegido = disp_segura[0]
                df.at[elegido, ds] = "C"
                turnos_totales[elegido] += 1
                if es_finde_o_festivo: finde_totales[elegido] += 1

        # 5. Rellenar lo sobrante con Descanso ('D')
        for p in INTEGRANTES:
            if df.at[p, ds] == "": df.at[p, ds] = "D"

    # ==============================================================
    # CALCULOS FINALES DE COLUMNAS
    # ==============================================================
    df['TOTAL CORRIDOS'] = df.apply(lambda r: sum(1 for c in df.columns if 'C' in str(r[c]) and c.isdigit()), axis=1)
    df['TOTAL NOCHES'] = df.apply(lambda r: noches_totales[r.name], axis=1)
    df['TOTAL TURNOS'] = df['TOTAL CORRIDOS'] + df['TOTAL NOCHES']
    df['FINES DE SEMANA'] = df.apply(lambda r: finde_totales[r.name], axis=1)
    
    return df
