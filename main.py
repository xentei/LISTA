import streamlit as st
import pandas as pd
from io import StringIO
from thefuzz import fuzz
import re
import unicodedata

# Configuración de página
st.set_page_config(page_title="Control de Personal", layout="wide")

st.title("👮‍♂️ Comparador de Listas de Guardia")
st.markdown("""
<style>
.small-font { font-size:14px !important; color: #666; }
</style>
<p class="small-font">Versión Especializada: Lee hoja 'LISTA' (Columnas D y E).</p>
""", unsafe_allow_html=True)

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
}

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
    # 1. Quitar paréntesis y números
    texto = re.sub(r'\([^)]*\)', '', texto)
    # 2. Quitar tildes
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    # 3. Dejar SOLO letras y espacios
    texto = re.sub(r'[^a-zA-Z\s]', '', texto)
    return texto.strip().upper()

def leer_excel_guardia(archivo):
    """
    Lee hoja 'LISTA', columnas D (Jerarquia) y E (Nombre).
    """
    try:
        df = pd.read_excel(archivo, sheet_name='LISTA', usecols="D:E", header=None)
        df.columns = ['Jerarquia', 'Nombre']
        
        def es_valida(row):
            t = str(row['Jerarquia']).lower()
            return any(k in t for k in EQUIVALENCIAS.keys())
        
        df = df[df.apply(es_valida, axis=1)]
        return df
    except Exception as e:
        st.error(f"Error leyendo la hoja 'LISTA': {e}")
        return None

def procesar_generico(texto_input, archivo_input):
    """Procesador para el PARTE"""
    df = None
    if archivo_input:
        try:
            if archivo_input.name.endswith('csv'):
                df = pd.read_csv(archivo_input)
            else:
                df = pd.read_excel(archivo_input)
            if len(df.columns) >= 2:
                df = df.iloc[:, :2]
                df.columns = ['Jerarquia', 'Nombre']
        except Exception:
            return None
    elif texto_input:
        try:
            df = pd.read_csv(StringIO(texto_input), sep='\t', header=None, engine='python')
            if any(x in str(df.iloc[0, 0]).lower() for x in ['jerarquia', 'grado']):
                 df = pd.read_csv(StringIO(texto_input), sep='\t', engine='python')
            df = df.iloc[:, :2]
            df.columns = ['Jerarquia', 'Nombre']
        except:
            return None
    return df

# --- 2. BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración")
    umbral = st.slider("Exigencia", 50, 100, 85)

# --- 3. INTERFAZ ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("📋 1. EL PARTE")
    p_txt = st.text_area("Pegar Parte", height=150)
    p_file = st.file_uploader("Subir Excel Parte", type=["xlsx", "csv"])

with col2:
    st.subheader("📝 2. LISTA DE GUARDIA")
    l_txt = st.text_area("Pegar (Opción B)", height=150)
    l_file = st.file_uploader("Subir Excel Guardia", type=["xlsx"])

# --- 4. LÓGICA ---
if st.button("🔍 COMPARAR", type="primary"):
    df_p = procesar_generico(p_txt, p_file)
    df_l = leer_excel_guardia(l_file) if l_file else procesar_generico(l_txt, None)

    if df_p is not None and df_l is not None and not df_p.empty and not df_l.empty:
        try:
            # Normalizar
            df_p['j_norm'] = df_p['Jerarquia'].apply(normalizar_jerarquia)
            df_l['j_norm'] = df_l['Jerarquia'].apply(normalizar_jerarquia)
            df_p['n_clean'] = df_p['Nombre'].apply(limpiar_nombre)
            df_l['n_clean'] = df_l['Nombre'].apply(limpiar_nombre)

            sobran = df_l.copy()
            sobran['found'] = False
            faltan = []

            # Comparar
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
                    faltan.append({'Jerarquia': row_p['Jerarquia'], 'Nombre': row_p['Nombre']})

            # Resultados
            st.divider()
            c1, c2 = st.columns(2)
            
            with c1:
                # AQUÍ ESTÁ LA MAGIA DE LOS GLOBOS 🎈
                if not faltan:
                    st.balloons()
                    st.success("✨ ¡PERFECTO! NO FALTA NADIE ✨")
                else:
                    st.warning(f"⚠️ FALTA AGREGAR ({len(faltan)})")
                    st.dataframe(pd.DataFrame(faltan), hide_index=True, use_container_width=True)
            
            with c2:
                sobra_df = sobran[~sobran['found']]
                if sobra_df.empty:
                     st.success("✅ LISTA LIMPIA (No sobra nadie)")
                else:
                    st.error(f"❌ SOBRA / BORRAR ({len(sobra_df)})")
                    st.dataframe(sobra_df[['Jerarquia', 'Nombre']], hide_index=True, use_container_width=True)

        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("⚠️ Faltan datos o el Excel no tiene la hoja 'LISTA'.")
