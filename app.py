import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import io

st.set_page_config(page_title="Gestor Pro", layout="wide")

PERS = ["GERLIS DOMINGUEZ", "ANGIE BERNAL", "JHON RIOS", "MARCELA CASTRO", "ZARIANA REYES", "IVETTE VALENCIA", "GINELAP", "ANA ESCOBAR", "JUAN CAMILO PEREZ", "ERNESTO MUSKUS", "RANCES OSPINO", "KELLY JOHANA JURADO", "JOAN SEBASTIAN AGUDELO"]
ALTA = ['C1', 'C2', 'C3', 'C4', 'N1']

def get_f(a, m):
    f = {1:[1,6], 3:[23], 4:[2,3,16,17], 5:[1,18], 6:[8,15,29], 7:[20], 8:[7,17], 10:[12], 11:[2,16], 12:[8,25]}
    return f.get(m, [])

def proc_h(file):
    if not file: return {}, {}
    df = pd.read_excel(file, skiprows=9) if file.name.endswith('.xlsx') else pd.read_csv(file, skiprows=9)
    df.columns = df.columns.str.strip().str.upper()
    h_c, h_e = {}, {}
    for _, r in df.iterrows():
        n = str(r.get('NOMBRE','')).strip().upper()
        if n not in PERS: continue
        d_c = [c for c in df.columns if str(c).isdigit() and int(c)>=21]
        h_c[n] = {'a': sum(1 for d in d_c if any(t in str(r[d]).upper() for t in ALTA) and 'P' not in str(r[d]).upper()),
                  't': sum(1 for d in d_c if any(t in str(r[d]).upper() for t in ['C','N']) and 'P' not in str(r[d]).upper())}
        ult = [str(r[c]).upper() for c in [x for x in df.columns if str(x).isdigit()][-3:]]
        h_e[n] = {'n': 'N' in ult[-1], 's': sum(1 for x in ult if any(t in x for t in ['C','N']) and 'P' not in x)}
    return h_c, h_e

def proc_s(link):
    s = {p: {} for p in PERS}
    if not link: return s
    try:
        url = link.split('/edit')[0]+'/export?format=csv' if "/edit" in link else link
        df = pd.read_csv(url)
        for _, r in df.iterrows():
            n, f, sol = str(r.get('NOMBRE','')).upper(), ''.join(filter(str.isdigit, str(r.get('FECHA','')))), str(r.get('SOLICITUD','')).upper()
            if n in PERS and f and sol != 'NAN': s[n][f] = sol
    except: pass
    return s

def clr(v):
    if v in ['L','D']: return 'background-color: #b6d7a8'
    if v == 'P': return 'background-color: #f4cccc'
    if 'N' in str(v): return 'background-color: #cfe2f3'
    if any(t in str(v) for t in ALTA): return 'background-color: #fff2cc'
    return ''

def gen_m(m, a, h_c, h_e, sug):
    d_m = calendar.monthrange(a, m)[1]
    fest = get_f(a, m)
    df = pd.DataFrame(index=PERS, columns=[str(d) for d in range(1, d_m+1)]).fillna("")
    c1_a, tot_n = {p: h_c.get(p,{}).get('a',0) for p in PERS}, {p: h_c.get(p,{}).get('t',0) for p in PERS}
    u_n, cons = {p: -5 for p in PERS}, {p: h_e.get(p,{}).get('s',0) for p in PERS}

    for d in range(1, d_m+1):
        ds, fch = str(d), datetime(a, m, d)
        wd, es_f = fch.weekday(), (d in fest or fch.weekday()==6)
        if d == 21:
            for p in PERS: c1_a[p], tot_n[p] = 0, 0
        for p in PERS:
            if d==1 and h_e.get(p,{}).get('n'): df.at[p,'1'], u_n[p], cons[p] = 'P', 0, 0
            elif d==1 and h_e.get(p,{}).get('s',0)>=3: df.at[p,'1'], cons[p] = 'D', 0
            if d>1 and 'N' in str(df.at[p, str(d-1)]): df.at[p, ds], cons[p] = 'P', 0
        if wd==2: df.at["ANGIE BERNAL", ds] = "C6"
        if wd==1: df.at["JHON RIOS", ds] = "N1"; tot_n["JHON RIOS"]+=1; c1_a["JHON RIOS"]+=1; u_n["JHON RIOS"]=d
        if wd in [0,1,5]: df.at["IVETTE VALENCIA", ds] = "L"
        if wd in [3,4]: df.at["GINELAP", ds] = "L"
        if wd==0: df.at["GERLIS DOMINGUEZ", ds]=df.at["ZARIANA REYES", ds]="L"
        if wd==3: df.at["MARCELA CASTRO", ds]=df.at["JUAN CAMILO PEREZ", ds]="L"
        for p in PERS:
            rq = sug.get(p,{}).get(ds)
            if rq and df.at[p, ds]=="":
                df.at[p, ds]=rq
                if any(t in rq for t in ['C','N']) and 'P' not in rq:
                    tot_n[p]+=1
                    if any(t in rq for t in ALTA): c1_a[p]+=1
                if 'N' in rq: u_n[p]=d
        t_h = ['N1','N2','C1','C2','C3','C4'] + (['C5','C6'] if not es_f and wd<5 else (['C5'] if not es_f else []))
        for t in t_h:
            if (df[ds]==t).any() and t not in ['C1','C2']: continue
            disp = [p for p in PERS if df.at[p, ds]=="" and cons[p]<3]
            if 'N' in t:
                disp = [p for p in disp if (d-u_n[p])>2]
                if disp:
                    disp.sort(key=lambda x: tot_n[x])
                    el=disp[0]; df.at[el, ds], tot_n[el], u_n[el], cons[el] = t, tot_n[el]+1, d, cons[el]+1
            elif t=='C1':
                if d>1: disp = [p for p in disp if str(df.at[p, str(d-1)])!='C1']
                if disp:
                    disp.sort(key=lambda x: (c1_a[x], tot_n[x]))
                    el=disp[0]; df.at[el, ds], c1_a[el], tot_n[el], cons[el] = t, c1_a[el]+1, tot_n[el]+1, cons[el]+1
            elif disp:
                disp.sort(key=lambda x: tot_n[x])
                el=disp[0]; df.at[el, ds], tot_n[el], cons[el] = t, tot_n[el]+1, cons[el]+1
    df['TOTAL MES'] = df.apply(lambda r: sum(1 for x in r if any(t in str(x) for t in ['C','N']) and 'P' not in str(x)), axis=1)
    df['NÓMINA (21-20)'] = df.apply(lambda r: h_c.get(r.name,{}).get('t',0) + sum(1 for d in range(1, 21) if any(t in str(r[str(d)]) for t in ['C','N']) and 'P' not in str(r[str(d)])), axis=1)
    return df.replace("", "D")

st.title("🏥 Optimizador de Turnos V3")
with st.sidebar:
    arc = st.file_uploader("Historial", type=['xlsx', 'csv'])
    lnk = st.text_input("Link Sugerencias", "https://docs.google.com/spreadsheets/d/1PZwvv0XQtSEDfC5GO6OlG7Fn8HqJNQUBZ1RNSRgBsss/edit?gid=0#gid=0")
    mes = st.selectbox("Mes", range(1, 13), index=datetime.now().month-1)

if st.button("🚀 GENERAR"):
    hc, he = proc_h(arc)
    res = gen_m(mes, 2026, hc, he, proc_s(lnk))
    st.dataframe(res.style.applymap(clr, subset=[c for c in res.columns if c.isdigit()]), use_container_width=True)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as w: res.to_excel(w, index=True)
    st.download_button("📥 Descargar", out.getvalue(), f"Cuadro_{mes}.xlsx")
