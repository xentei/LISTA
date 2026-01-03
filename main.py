import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
from thefuzz import fuzz
import re
import unicodedata
import openpyxl 

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Control PSA V6.4", layout="wide", page_icon="🕵️")

# --- ESTILOS CSS ---
st.markdown("""
<style>
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }
    div.stButton > button:first-child {
        width: 100%; border-radius: 4px; height: 2.5rem; font-weight: bold; border: none;
    }
    .stCode { font-family: sans-serif !important; font-size: 15px !important; font-weight: bold; }
    
    .success-box {
        padding: 5px; background-color: #28a745; color: white; border-radius: 4px;
        text-align: center; font-weight: bold; font-size: 14px; height: 38px;
        display: flex; align-items: center; justify-content: center;
    }
    .delete-box {
        padding: 5px; background-color: #f8d7da; color: #721c24; border-radius: 4px;
        border: 1px solid #f5c6cb; text-align: center; font-weight: bold; font-size: 14px;
        height: 38px; display: flex; align-items: center; justify-content: center;
    }
    .warning-box {
        padding: 10px; background-color: #fff3cd; color: #856404; border-radius: 6px;
        border: 1px solid #ffeeba; text-align: center; font-weight: bold; font-size: 15px;
        display: flex; align-items: center; justify-content: center; flex-direction: column;
    }
    
    .jerarquia-text { font-size: 15px; font-weight: 700; padding-top: 10px; color: #555; }
    
    .header-green { color: #28a745; border-bottom: 3px solid #28a745; padding-bottom: 5px; font-weight: 800; font-size: 1.2rem;}
    .header-red { color: #800020; border-bottom: 3px solid #800020; padding-bottom: 5px; font-weight: 800; font-size: 1.2rem;}
    .header-yellow { color: #d39e00; border-bottom: 3px solid #d39e00; padding-bottom: 5px; font-weight: 800; font-size: 1.2rem;}

    hr { margin: 0.3rem 0 !important; opacity: 0.2; }
</style>
""", unsafe_allow_html=True)

st.title("🛡️ CONTROL DE PERSONAL - V6.4")

# --- 1. CONFIGURACIÓN Y EQUIVALENCIAS ---
EQUIVALENCIAS = {
    "oficial ayudante": "OFICIAL AYUDANTE", "of ayte": "OFICIAL AYUDANTE", "of. ayte": "OFICIAL AYUDANTE", "ayte": "OFICIAL AYUDANTE",
    "oficial principal": "OFICIAL PRINCIPAL", "of ppal": "OFICIAL PRINCIPAL", "of. ppal": "OFICIAL PRINCIPAL", "ppal": "OFICIAL PRINCIPAL",
    "oficial mayor": "OFICIAL MAYOR", "of mayor": "OFICIAL MAYOR", "of. mayor": "OFICIAL MAYOR",
    "oficial jefe": "OFICIAL JEFE", "of jefe": "OFICIAL JEFE", "of. jefe": "OFICIAL JEFE",
    "subinspector": "SUBINSPECTOR", "sub inspector": "SUBINSPECTOR", "subinsp": "SUBINSPECTOR",
    "inspector": "INSPECTOR", "insp": "INSPECTOR",
    "comisionado mayor": "COMISIONADO MAYOR", "cdo mayor": "COMISIONADO MAYOR", "cdo. mayor": "COMISIONADO MAYOR", "com mayor": "COMISIONADO MAYOR",
    "comisionado general": "COMISIONADO GENERAL", "cdo general": "COMISIONADO GENERAL", "cdo. general": "COMISIONADO GENERAL", "cdo gral": "COMISIONADO GENERAL",
    "psa": "PSA", "aux": "AUXILIAR", "auxiliar": "AUXILIAR"
}

# --- GESTIÓN DE ESTADO ---
if 'analisis_listo' not in st.session_state: st.session_state.analisis_listo = False
if 'df_faltan' not in st.session_state: st.session_state.df_faltan = []
if 'df_sobran' not in st.session_state: st.session_state.df_sobran = pd.DataFrame()
if 'total_parte' not in st.session_state: st.session_state.total_parte = 0
if 'total_lista' not in st.session_state: st.session_state.total_lista = 0
if 'checked_items' not in st.session_state: st.session_state.checked_items = set()
if 'confirmed_pairs' not in st.session_state: st.session_state.confirmed_pairs = set()
if 'rejected_pairs' not in st.session_state: st.session_state.rejected_pairs = set()

# --- FUNCIONES DE LIMPIEZA Y HELPERS ---
def normalizar_jerarquia(texto):
    if pd.isna(texto): return ""
    texto_limpio = str(texto).strip().lower()
    if texto_limpio in EQUIVALENCIAS: return EQUIVALENCIAS[texto_limpio]
    for key, value in EQUIVALENCIAS.items():
        if key in texto_limpio: return value
    return texto_limpio.upper()

def limpiar_nombre(texto):
    if pd.isna(texto): return ""
    texto = str(texto)
    texto = re.sub(r'\([^)]*\)', '', texto)
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^a-zA-Z\s]', '', texto)
    return texto.strip().upper()

def toggle_item(unique_id):
    if unique_id in st.session_state.checked_items:
        st.session_state.checked_items.remove(unique_id)
    else:
        st.session_state.checked_items.add(unique_id)

def confirmar_match(id_parte, id_lista):
    pair_id = f"{id_parte}|{id_lista}"
    st.session_state.confirmed_pairs.add(pair_id)

def rechazar_match(id_parte, id_lista):
    pair_id = f"{id_parte}|{id_lista}"
    st.session_state.rejected_pairs.add(pair_id)

# --- FUNCIÓN EXCEL ---
def borrar_sobrantes_excel(archivo_original, lista_nombres_borrar):
    try:
        wb = openpyxl.load_workbook(archivo_original)
        sheet_name = 'LISTA' if 'LISTA' in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        col_jerarquia = -1; col_nombre = -1; max_matches = 0
        for col in range(1, 20):
            matches = 0
            for row in range(1, 30):
                val = str(ws.cell(row=row, column=col).value).lower()
                if any(k in val for k in EQUIVALENCIAS.keys()): matches += 1
            if matches > max_matches: max_matches = matches; col_jerarquia = col; col_nombre = col + 1 
        if col_jerarquia == -1: return None 
        nombres_a_borrar_limpios = set([limpiar_nombre(n) for n in lista_nombres_borrar])
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            cell_nombre = row[col_nombre - 1] 
            cell_jerarquia = ws.cell(row=cell_nombre.row, column=col_jerarquia)
            cell_nombre_obj = ws.cell(row=cell_nombre.row, column=col_nombre)
            val_nombre_limpio = limpiar_nombre(str(cell_nombre_obj.value))
            if val_nombre_limpio in nombres_a_borrar_limpios:
                cell_jerarquia.value = None; cell_nombre_obj.value = None
        output = BytesIO(); wb.save(output); output.seek(0)
        return output
    except Exception as e: return None

# --- LECTURA PANDAS ---
def leer_excel_inteligente(archivo):
    try:
        xls = pd.ExcelFile(archivo)
        sheet_name = 'LISTA' if 'LISTA' in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        best_col_idx = -1; max_matches = 0
        for col_idx in range(len(df.columns)):
            col_data = df.iloc[:, col_idx].astype(str).str.lower()
            matches = col_data.apply(lambda x: any(k in x for k in EQUIVALENCIAS.keys())).sum()
            if matches > max_matches: max_matches = matches; best_col_idx = col_idx
        if best_col_idx != -1 and max_matches > 0 and best_col_idx + 1 < len(df.columns):
            subset = df.iloc[:, [best_col_idx, best_col_idx+1]].copy()
            subset.columns = ['Jerarquia', 'Nombre']
            def es_fila_valida(row):
                t = str(row['Jerarquia']).lower()
                return any(k in t for k in EQUIVALENCIAS.keys())
            subset = subset[subset.apply(es_fila_valida, axis=1)]
            return subset
        return None
    except: return None

def procesar_generico(texto_input, archivo_input):
    df = None
    if archivo_input:
        try:
            if archivo_input.name.endswith('csv'): df = pd.read_csv(archivo_input)
            else: df = pd.read_excel(archivo_input)
            if len(df.columns) >= 2: df = df.iloc[:, :2]; df.columns = ['Jerarquia', 'Nombre']
        except: return None
    elif texto_input:
        try:
            df = pd.read_csv(StringIO(texto_input), sep='\t', header=None, engine='python')
            if any(x in str(df.iloc[0, 0]).lower() for x in ['jerarquia', 'grado']):
                 df = pd.read_csv(StringIO(texto_input), sep='\t', engine='python')
            df = df.iloc[:, :2]; df.columns = ['Jerarquia', 'Nombre']
        except: return None
    return df

# --- ESTADOS DE INPUT ---
if 'p_txt' not in st.session_state: st.session_state.p_txt = ""
if 'l_txt' not in st.session_state: st.session_state.l_txt = ""
if 'p_key' not in st.session_state: st.session_state.p_key = 0
if 'l_key' not in st.session_state: st.session_state.l_key = 0

# --- FUNCIONES LIMPIEZA SEPARADAS ---
def limpiar_parte():
    st.session_state.p_txt = ""
    st.session_state.p_key += 1
    st.session_state.analisis_listo = False
    st.session_state.confirmed_pairs = set()
    st.session_state.rejected_pairs = set()

def limpiar_lista():
    st.session_state.l_txt = ""
    st.session_state.l_key += 1
    st.session_state.analisis_listo = False
    st.session_state.confirmed_pairs = set()
    st.session_state.rejected_pairs = set()

# --- CARGA DE DATOS ---
c1, c2 = st.columns(2)
with c1:
    with st.container(border=True):
        st.subheader("📋 1. EL PARTE")
        if st.button("🗑️ Limpiar Parte", on_click=limpiar_parte, key="cl_p"): pass
        p_txt = st.text_area("Parte", height=68, key="p_txt", label_visibility="collapsed", placeholder="Pegar Parte...")
        p_file = st.file_uploader("Archivo", type=["xlsx", "csv"], key=f"p_file_{st.session_state.p_key}", label_visibility="collapsed")
with c2:
    with st.container(border=True):
        st.subheader("📝 2. LISTA GUARDIA")
        if st.button("🗑️ Limpiar Lista", on_click=limpiar_lista, key="cl_l"): pass
        l_txt = st.text_area("Lista", height=68, key="l_txt", label_visibility="collapsed", placeholder="Pegar Lista...")
        l_file = st.file_uploader("Archivo", type=["xlsx"], key=f"l_file_{st.session_state.l_key}", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)
with st.expander("⚙️ Ajustes de Precisión"):
    # --- CAMBIO AQUI: DEFAULT 95 ---
    umbral = st.slider("Exigencia Estricta", 50, 100, 95)

# --- LÓGICA DE ANÁLISIS CENTRALIZADA ---
def ejecutar_analisis():
    """Ejecuta el análisis y guarda resultados en Session State"""
    df_p = procesar_generico(p_txt, p_file)
    df_l = leer_excel_inteligente(l_file) if l_file else procesar_generico(l_txt, None)

    if df_p is not None and df_l is not None and not df_p.empty and not df_l.empty:
        try:
            df_p['j_norm'] = df_p['Jerarquia'].apply(normalizar_jerarquia)
            df_l['j_norm'] = df_l['Jerarquia'].apply(normalizar_jerarquia)
            df_p['n_clean'] = df_p['Nombre'].apply(limpiar_nombre)
            df_l['n_clean'] = df_l['Nombre'].apply(limpiar_nombre)
            
            df_p['unique_id'] = df_p['Nombre'] + "_" + df_p.index.astype(str)
            df_l['unique_id'] = df_l['Nombre'] + "_" + df_l.index.astype(str)

            sobran = df_l.copy(); sobran['found'] = False
            faltan_temp = [] 

            # Comparación Estricta + Matches Manuales
            for idx_p, row_p in df_p.iterrows():
                candidatos = sobran[sobran['j_norm'] == row_p['j_norm']]
                encontrado = False
                for idx_l, row_l in candidatos.iterrows():
                    if row_l['found']: continue
                    if fuzz.token_set_ratio(row_p['n_clean'], row_l['n_clean']) >= umbral:
                        encontrado = True; sobran.at[idx_l, 'found'] = True; break
                
                if not encontrado:
                    for idx_l, row_l in sobran.iterrows():
                        if row_l['found']: continue
                        pair_id = f"{row_p['unique_id']}|{row_l['unique_id']}"
                        if pair_id in st.session_state.confirmed_pairs:
                            encontrado = True; sobran.at[idx_l, 'found'] = True; break

                if not encontrado:
                    faltan_temp.append(row_p)

            # Lógica Detective
            detective_matches = [] 
            df_sobran_reales = sobran[~sobran['found']]
            
            for f in faltan_temp:
                best_match = None; best_score = 0
                for idx_s, s in df_sobran_reales.iterrows():
                    pair_id = f"{f['unique_id']}|{s['unique_id']}"
                    if pair_id in st.session_state.rejected_pairs: continue 
                    
                    score = fuzz.token_sort_ratio(f['n_clean'], s['n_clean'])
                    if score > 50 and score < umbral: 
                        if score > best_score: best_score = score; best_match = s
                
                if best_match is not None:
                    detective_matches.append({'falta': f, 'sobra': best_match})

            st.session_state.df_faltan = faltan_temp
            st.session_state.df_sobran = df_sobran_reales
            st.session_state.detective_candidates = detective_matches
            st.session_state.total_parte = len(df_p)
            st.session_state.total_lista = len(df_l)
            st.session_state.analisis_listo = True
        except Exception as e:
            st.error(f"Error: {e}")
            st.session_state.analisis_listo = False

# --- INTERACCIONES DETECTIVE ---
def confirmar_y_recargar(id_parte, id_lista):
    confirmar_match(id_parte, id_lista)
    ejecutar_analisis() 

def rechazar_y_recargar(id_parte, id_lista):
    rechazar_match(id_parte, id_lista)
    ejecutar_analisis() 

# --- BOTÓN ANÁLISIS PRINCIPAL ---
if st.button("🔍 ANALIZAR AHORA", type="primary", use_container_width=True):
    ejecutar_analisis()

# --- VISUALIZACIÓN ---
if st.session_state.analisis_listo:
    st.divider()
    
    ids_en_detective_falta = [m['falta']['unique_id'] for m in st.session_state.detective_candidates]
    ids_en_detective_sobra = [m['sobra']['unique_id'] for m in st.session_state.detective_candidates]

    lista_verde_final = [f for f in st.session_state.df_faltan if f['unique_id'] not in ids_en_detective_falta]
    df_rojo_final = st.session_state.df_sobran[~st.session_state.df_sobran['unique_id'].isin(ids_en_detective_sobra)]
    
    # Dashboard
    m1, m2, m3, m4 = st.columns(4)
    with m1: st.metric("Parte", st.session_state.total_parte)
    with m2: st.metric("Lista", st.session_state.total_lista)
    with m3: st.metric("Faltan", len(lista_verde_final), delta=len(lista_verde_final) if len(lista_verde_final)>0 else None)
    with m4: st.metric("Sobran", len(df_rojo_final), delta=-len(df_rojo_final) if len(df_rojo_final)>0 else None)
    
    # --- ZONA DETECTIVE (AMARILLA) ---
    if st.session_state.detective_candidates:
        st.markdown("---")
        st.markdown('<div class="header-yellow">🕵️ ZONA DETECTIVE (Conflictos)</div>', unsafe_allow_html=True)
        st.info("Similitudes detectadas. Confirma si son la misma persona.")
        
        for match in st.session_state.detective_candidates:
            f = match['falta']
            s = match['sobra']
            
            c_izq, c_flecha, c_der, c_btn = st.columns([3, 1, 3, 2])
            
            with c_izq:
                st.caption("FALTA")
                st.markdown(f"**{f['Jerarquia']}**")
                st.markdown(f"<div class='warning-box'>{f['Nombre']}</div>", unsafe_allow_html=True)
            with c_flecha:
                st.markdown("<h2 style='text-align: center; color: #999;'>?</h2>", unsafe_allow_html=True)
            with c_der:
                st.caption("SOBRA")
                st.markdown(f"**{s['Jerarquia']}**")
                st.markdown(f"<div class='warning-box'>{s['Nombre']}</div>", unsafe_allow_html=True)
            with c_btn:
                st.caption("DECISIÓN")
                if st.button("✅ Son el mismo", key=f"yes_{f['unique_id']}", type="primary"):
                    confirmar_y_recargar(f['unique_id'], s['unique_id'])
                    st.rerun()
                if st.button("❌ No son", key=f"no_{f['unique_id']}"):
                    rechazar_y_recargar(f['unique_id'], s['unique_id'])
                    st.rerun()
            st.divider()

    st.markdown("---")
    col_res1, col_res2 = st.columns(2)
    
    # --- COLUMNA 1: AGREGAR (VERDE) ---
    with col_res1:
        st.markdown('<div class="header-green">FALTA AGREGAR</div>', unsafe_allow_html=True)
        
        if not lista_verde_final:
            st.markdown('<div class="green-msg">NO HACE FALTA AGREGAR A NADIE</div>', unsafe_allow_html=True)
        else:
            h1, h2, h3 = st.columns([1.2, 3, 0.8])
            h1.markdown("**JERARQUÍA**"); h2.markdown("**NOMBRE**"); h3.markdown("**LISTO**")
            st.markdown("---")
            
            for p in lista_verde_final:
                r1, r2, r3 = st.columns([1.2, 3, 0.8])
                nombre_upper = str(p['Nombre']).upper()
                jerarquia_upper = str(p['Jerarquia']).upper()
                is_checked = p['unique_id'] in st.session_state.checked_items

                with r1: st.markdown(f'<div class="jerarquia-text">{jerarquia_upper}</div>', unsafe_allow_html=True)
                with r2:
                    if is_checked: st.markdown(f'<div class="success-box">YA AGREGADO</div>', unsafe_allow_html=True)
                    else: st.code(nombre_upper, language="text")
                with r3:
                    label = "↩" if is_checked else "✔"
                    type_btn = "secondary" if is_checked else "primary"
                    st.button(label, key=f"btn_{p['unique_id']}", type=type_btn, on_click=toggle_item, args=(p['unique_id'],))
                st.markdown("<hr>", unsafe_allow_html=True)

    # --- COLUMNA 2: BORRAR (ROJO) ---
    with col_res2:
        st.markdown('<div class="header-red">SOBRA / BORRAR</div>', unsafe_allow_html=True)
        if df_rojo_final.empty:
             st.markdown('<div class="bordo-msg">NO HACE FALTA BORRAR A NADIE</div>', unsafe_allow_html=True)
        else:
            st.dataframe(df_rojo_final[['Jerarquia', 'Nombre']], hide_index=True, use_container_width=True, height=400)
            st.markdown("---")
            if l_file is not None:
                lista_nombres_borrar = df_rojo_final['Nombre'].tolist()
                l_file.seek(0)
                excel_limpio = borrar_sobrantes_excel(l_file, lista_nombres_borrar)
                if excel_limpio:
                    st.download_button(label="💾 DESCARGAR LISTA LIMPIA", data=excel_limpio, file_name=l_file.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
            else: st.info("ℹ️ Sube un Excel para habilitar el borrado automático.")
