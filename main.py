import streamlit as st
import pandas as pd
from io import StringIO
from thefuzz import fuzz
import re
import unicodedata

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Control PSA V4.1", layout="wide", page_icon="👮‍♂️")

# --- ESTILOS CSS ---
st.markdown("""
<style>
    /* Ajustes generales */
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }
    
    /* BOTONES */
    div.stButton > button:first-child {
        width: 100%;
        border-radius: 4px;
        height: 2.5rem;
        font-weight: bold;
        border: none;
    }
    
    /* CAJA DE COPIAR (Gris - Nombre) */
    .stCode {
        font-family: sans-serif !important;
        font-size: 15px !important;
        font-weight: bold;
    }
    
    /* CAJA VERDE (YA AGREGADO) */
    .success-box {
        padding: 5px;
        background-color: #28a745;
        color: white;
        border-radius: 4px;
        text-align: center;
        font-weight: bold;
        font-size: 14px;
        height: 38px;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 1px 2px rgba(0,0,0,0.1);
    }

    /* MENSAJES DE ESTADO */
    .bordo-msg {
        background-color: #800020;
        color: white;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        font-weight: bold;
        font-size: 16px;
        margin-top: 10px;
        border: 2px solid #5a0016;
    }
    
    .green-msg {
        background-color: #28a745;
        color: white;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        font-weight: bold;
        font-size: 16px;
        margin-top: 10px;
        border: 2px solid #1e7e34;
    }

    /* TEXTO JERARQUÍA (Grande) */
    .jerarquia-text {
        font-size: 15px;
        font-weight: 700;
        padding-top: 10px; 
        color: #555;
    }
    
    /* ENCABEZADOS DE COLUMNA */
    .header-green { color: #28a745; border-bottom: 3px solid #28a745; padding-bottom: 5px; font-weight: 800; font-size: 1.2rem;}
    .header-red { color: #800020; border-bottom: 3px solid #800020; padding-bottom: 5px; font-weight: 800; font-size: 1.2rem;}

    hr { margin: 0.3rem 0 !important; opacity: 0.2; }
</style>
""", unsafe_allow_html=True)

st.title("🛡️ CONTROL DE PERSONAL - V4.1")

# --- 1. CONFIGURACIÓN Y EQUIVALENCIAS ---
EQUIVALENCIAS = {
    "of ayte": "oficial ayudante",
    "of jefe": "oficial jefe",
    "of mayor": "oficial mayor",
    "of ppal": "oficial principal",
    "oficial ayudante": "oficial ayudante",
    "oficial jefe": "oficial jefe",
    "oficial mayor": "oficial mayor",
    "oficial principal": "oficial principal",
    "inspector": "inspector",
    "cabo 1": "cabo primero",
    "cabo": "cabo",
    "aux": "auxiliar",
    "ayte": "ayudante",
    "psa": "psa"
}

# --- GESTIÓN DE ESTADO ---
if 'analisis_listo' not in st.session_state: st.session_state.analisis_listo = False
if 'df_faltan' not in st.session_state: st.session_state.df_faltan = []
if 'df_sobran' not in st.session_state: st.session_state.df_sobran = pd.DataFrame()
if 'total_parte' not in st.session_state: st.session_state.total_parte = 0
if 'total_lista' not in st.session_state: st.session_state.total_lista = 0
if 'checked_items' not in st.session_state: st.session_state.checked_items = set()

def toggle_item(unique_id):
    if unique_id in st.session_state.checked_items:
        st.session_state.checked_items.remove(unique_id)
    else:
        st.session_state.checked_items.add(unique_id)

# --- FUNCIONES DE LIMPIEZA ---
def normalizar_jerarquia(texto):
    if pd.isna(texto): return ""
    texto_limpio = str(texto).strip().lower()
    if texto_limpio in EQUIVALENCIAS:
        return EQUIVALENCIAS[texto_limpio]
    for key, value in EQUIVALENCIAS.items():
        if key in texto_limpio:
            return value
    return texto_limpio

def limpiar_nombre(texto):
    if pd.isna(texto): return ""
    texto = str(texto)
    texto = re.sub(r'\([^)]*\)', '', texto)
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^a-zA-Z\s]', '', texto)
    return texto.strip().upper()

# --- LECTURA INTELIGENTE ---
def leer_excel_inteligente(archivo):
    try:
        xls = pd.ExcelFile(archivo)
        sheet_name = 'LISTA' if 'LISTA' in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        best_col_idx = -1
        max_matches = 0
        
        for col_idx in range(len(df.columns)):
            col_data = df.iloc[:, col_idx].astype(str).str.lower()
            matches = col_data.apply(lambda x: any(k in x for k in EQUIVALENCIAS.keys())).sum()
            if matches > max_matches:
                max_matches = matches
                best_col_idx = col_idx
        
        if best_col_idx != -1 and max_matches > 0:
            if best_col_idx + 1 < len(df.columns):
                subset = df.iloc[:, [best_col_idx, best_col_idx+1]].copy()
                subset.columns = ['Jerarquia', 'Nombre']
                def es_fila_valida(row):
                    t = str(row['Jerarquia']).lower()
                    return any(k in t for k in EQUIVALENCIAS.keys())
                subset = subset[subset.apply(es_fila_valida, axis=1)]
                if not subset.empty:
                    st.toast(f"✅ Columnas detectadas en hoja '{sheet_name}'")
                    return subset
        st.error("❌ No detecté columnas automáticamente.")
        return None
    except Exception as e:
        st.error(f"Error leyendo archivo: {e}")
        return None

def procesar_generico(texto_input, archivo_input):
    df = None
    if archivo_input:
        try:
            if archivo_input.name.endswith('csv'): df = pd.read_csv(archivo_input)
            else: df = pd.read_excel(archivo_input)
            if len(df.columns) >= 2:
                df = df.iloc[:, :2]
                df.columns = ['Jerarquia', 'Nombre']
        except: return None
    elif texto_input:
        try:
            df = pd.read_csv(StringIO(texto_input), sep='\t', header=None, engine='python')
            if any(x in str(df.iloc[0, 0]).lower() for x in ['jerarquia', 'grado']):
                 df = pd.read_csv(StringIO(texto_input), sep='\t', engine='python')
            df = df.iloc[:, :2]
            df.columns = ['Jerarquia', 'Nombre']
        except: return None
    return df

# --- LIMPIEZA DE INPUTS ---
if 'p_txt' not in st.session_state: st.session_state.p_txt = ""
if 'l_txt' not in st.session_state: st.session_state.l_txt = ""

def limpiar_parte():
    st.session_state.p_txt = ""
    st.session_state.p_key = st.session_state.get('p_key', 0) + 1
    st.session_state.checked_items = set() 
    st.session_state.analisis_listo = False

def limpiar_lista():
    st.session_state.l_txt = ""
    st.session_state.l_key = st.session_state.get('l_key', 0) + 1
    st.session_state.checked_items = set()
    st.session_state.analisis_listo = False

if 'p_key' not in st.session_state: st.session_state.p_key = 0
if 'l_key' not in st.session_state: st.session_state.l_key = 0

# --- CARGA DE DATOS ---
col_carga1, col_carga2 = st.columns(2)
with col_carga1:
    with st.container(border=True):
        st.subheader("📋 1. EL PARTE")
        if st.button("🗑️ Limpiar", on_click=limpiar_parte, key="btn_limpiar_parte"): pass
        p_txt = st.text_area("Parte", height=68, key="p_txt", label_visibility="collapsed", placeholder="Pegar Parte...")
        p_file = st.file_uploader("Archivo", type=["xlsx", "csv"], key=f"p_file_{st.session_state.p_key}", label_visibility="collapsed")

with col_carga2:
    with st.container(border=True):
        st.subheader("📝 2. LISTA GUARDIA")
        if st.button("🗑️ Limpiar", on_click=limpiar_lista, key="btn_limpiar_lista"): pass
        l_txt = st.text_area("Lista", height=68, key="l_txt", label_visibility="collapsed", placeholder="Pegar Lista...")
        l_file = st.file_uploader("Archivo", type=["xlsx"], key=f"l_file_{st.session_state.l_key}", label_visibility="collapsed")

# --- AJUSTES DE PRECISIÓN (RECUPERADO) ---
st.markdown("<br>", unsafe_allow_html=True)
with st.expander("⚙️ Ajustes de Precisión (Opcional)"):
    umbral = st.slider("Nivel de Exigencia (85 recomendado)", 50, 100, 85)

# --- BOTÓN ANÁLISIS ---
if st.button("🔍 ANALIZAR AHORA", type="primary", use_container_width=True):
    df_p = procesar_generico(p_txt, p_file)
    df_l = leer_excel_inteligente(l_file) if l_file else procesar_generico(l_txt, None)

    if df_p is not None and df_l is not None and not df_p.empty and not df_l.empty:
        try:
            df_p['j_norm'] = df_p['Jerarquia'].apply(normalizar_jerarquia)
            df_l['j_norm'] = df_l['Jerarquia'].apply(normalizar_jerarquia)
            df_p['n_clean'] = df_p['Nombre'].apply(limpiar_nombre)
            df_l['n_clean'] = df_l['Nombre'].apply(limpiar_nombre)

            sobran = df_l.copy()
            sobran['found'] = False
            faltan_temp = []
            
            # Usamos el umbral del slider
            
            for idx_p, row_p in df_p.iterrows():
                candidatos = sobran[sobran['j_norm'] == row_p['j_norm']]
                encontrado = False
                for idx_l, row_l in candidatos.iterrows():
                    if row_l['found']: continue
                    if fuzz.token_set_ratio(row_p['n_clean'], row_l['n_clean']) >= umbral:
                        encontrado = True
                        sobran.at[idx_l, 'found'] = True
                        break
                if not encontrado:
                    unique_id = f"{row_p['Nombre']}_{idx_p}" 
                    faltan_temp.append({
                        'Jerarquia': row_p['Jerarquia'], 
                        'Nombre': row_p['Nombre'],
                        'ID': unique_id
                    })

            st.session_state.df_faltan = faltan_temp
            st.session_state.df_sobran = sobran[~sobran['found']]
            st.session_state.total_parte = len(df_p)
            st.session_state.total_lista = len(df_l)
            st.session_state.analisis_listo = True
            st.session_state.checked_items = set()
            
        except Exception as e:
            st.error(f"Error: {e}")
            st.session_state.analisis_listo = False
    else:
        st.warning("⚠️ Carga ambos datos.")

# --- RESULTADOS ---
if st.session_state.analisis_listo:
    st.divider()
    
    # Dashboard
    m1, m2, m3, m4 = st.columns(4)
    with m1: st.metric("Parte", st.session_state.total_parte)
    with m2: st.metric("Lista", st.session_state.total_lista)
    with m3: 
        c_faltan = len(st.session_state.df_faltan)
        st.metric("Faltan", c_faltan, delta=c_faltan if c_faltan > 0 else None)
    with m4: 
        c_sobran = len(st.session_state.df_sobran)
        st.metric("Sobran", c_sobran, delta=-c_sobran if c_sobran > 0 else None)
    
    st.divider()

    col_res1, col_res2 = st.columns(2)
    
    # --- COLUMNA 1: AGREGAR (INTERACTIVO) ---
    with col_res1:
        st.markdown('<div class="header-green">FALTA AGREGAR</div>', unsafe_allow_html=True)
        faltan_lista = st.session_state.df_faltan
        
        if not faltan_lista:
            st.markdown('<div class="green-msg">NO HACE FALTA AGREGAR A NADIE</div>', unsafe_allow_html=True)
            st.balloons()
        else:
            h1, h2, h3 = st.columns([1.2, 3, 0.8])
            h1.markdown("**JERARQUÍA**")
            h2.markdown("**NOMBRE**")
            h3.markdown("**LISTO**")
            st.markdown("---")
            
            for p in faltan_lista:
                r1, r2, r3 = st.columns([1.2, 3, 0.8])
                nombre_upper = str(p['Nombre']).upper()
                jerarquia_upper = str(p['Jerarquia']).upper()
                is_checked = p['ID'] in st.session_state.checked_items

                with r1: 
                    # JERARQUÍA EN MAYÚSCULAS Y GRANDE
                    st.markdown(f'<div class="jerarquia-text">{jerarquia_upper}</div>', unsafe_allow_html=True)

                with r2:
                    if is_checked:
                        st.markdown(f'<div class="success-box">YA AGREGADO</div>', unsafe_allow_html=True)
                    else:
                        st.code(nombre_upper, language="text")

                with r3:
                    label = "↩" if is_checked else "✔"
                    type_btn = "secondary" if is_checked else "primary"
                    st.button(label, key=f"btn_{p['ID']}", type=type_btn, on_click=toggle_item, args=(p['ID'],))
                
                st.markdown("<hr style='margin: 2px 0 !important; opacity: 0.2;'>", unsafe_allow_html=True)

    # --- COLUMNA 2: BORRAR (TABLA DE EXCEL) ---
    with col_res2:
        st.markdown('<div class="header-red">SOBRA / BORRAR</div>', unsafe_allow_html=True)
        sobran_df = st.session_state.df_sobran
        
        if sobran_df.empty:
             st.markdown('<div class="bordo-msg">NO HACE FALTA BORRAR A NADIE</div>', unsafe_allow_html=True)
        else:
            # TABLA DE EXCEL (st.dataframe)
            st.dataframe(
                sobran_df[['Jerarquia', 'Nombre']], 
                hide_index=True, 
                use_container_width=True,
                height=500
            )
