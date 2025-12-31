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
<p class="small-font">Versión Especializada: Lee columnas D y E de la hoja 'LISTA'.</p>
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
    # 1. Quitar paréntesis y números (10), (30)
    texto = re.sub(r'\([^)]*\)', '', texto)
    # 2. Quitar tildes
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    # 3. Dejar SOLO letras y espacios
    texto = re.sub(r'[^a-zA-Z\s]', '', texto)
    return texto.strip().upper()

def leer_excel_guardia(archivo):
    """
    Lógica ESPECÍFICA para la Lista de Guardia:
    - Hoja: 'LISTA'
    - Columnas: D (Jerarquía) y E (Nombre)
    """
    try:
        # Intentamos leer solo la hoja LISTA y las columnas D:E (usecols="D:E")
        # header=None para leer desde la primera fila y nosotros filtramos
        df = pd.read_excel(archivo, sheet_name='LISTA', usecols="D:E", header=None)
        
        # Asignamos nombres fijos
        df.columns = ['Jerarquia', 'Nombre']
        
        # FILTRO DE BASURA:
        # Eliminamos filas donde la Jerarquía no sea válida (borra títulos, encabezados, etc.)
        def es_fila_valida(row):
            jer = str(row['Jerarquia']).lower()
            # Si la celda contiene alguna de nuestras palabras clave, es válida
            return any(k in jer for k in EQUIVALENCIAS.keys())

        # Aplicamos filtro
        df = df[df.apply(es_fila_valida, axis=1)]
        
        return df
    except ValueError:
        st.error("❌ No encontré la hoja llamada 'LISTA' en el Excel.")
        return None
    except Exception as e:
        st.error(f"Error leyendo la lista: {e}")
        return None

def procesar_generico(texto_input, archivo_input):
    """
    Procesador genérico para el PARTE (o si pegan texto en la lista)
    """
    df = None
    if archivo_input:
        try:
            if archivo_input.name.endswith('csv'):
                df = pd.read_csv(archivo_input)
            else:
                df = pd.read_excel(archivo_input)
            
            # Normalizamos a 2 columnas
            if len(df.columns) >= 2:
                df = df.iloc[:, :2]
                df.columns = ['Jerarquia', 'Nombre']
        except Exception as e:
            st.error(f"Error: {e}")
            return None

    elif texto_input:
        try:
            df = pd.read_csv(StringIO(texto_input), sep='\t', header=None, engine='python')
            # Detectar encabezados
            if any(x in str(df.iloc[0, 0]).lower() for x in ['jerarquia', 'grado']):
                 df = pd.read_csv(StringIO(texto_input), sep='\t', engine='python')
            else:
                df.columns = ['Jerarquia', 'Nombre'] + [f'Col_{i}' for i in range(2, len(df.columns))]
            
            df = df.iloc[:, :2]
            df.columns = ['Jerarquia', 'Nombre']
        except:
            return None
    
    return df

# --- 2. BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración")
    umbral = st.slider("Exigencia de coincidencia", 50, 100, 85)

# --- 3. INTERFAZ PRINCIPAL ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📋 1. EL PARTE (Regla)")
    st.info("Sube el Parte Oficial (Excel normal o texto).")
    parte_txt = st.text_area("Pegar Parte", height=150, key="p_txt")
    parte_file = st.file_uploader("Subir Excel Parte", type=["xlsx", "csv"], key="p_file")

with col2:
    st.subheader("📝 2. LISTA DE GUARDIA")
    st.info("Busca hoja 'LISTA', columnas D y E.")
    lista_txt = st.text_area("Pegar (Opción B)", height=150, key="l_txt")
    lista_file = st.file_uploader("Subir Excel Guardia", type=["xlsx"], key="l_file")

# --- 4. LÓGICA DE COMPARACIÓN ---
if st.button("🔍 COMPARAR AHORA", type="primary"):
    
    # 1. Procesar PARTE (Genérico)
    df_parte = procesar_generico(parte_txt, parte_file)

    # 2. Procesar LISTA (Específico Excel Guardia o Genérico Texto)
    df_lista = None
    if lista_file:
        # AQUÍ APLICAMOS LA LÓGICA ESPECIAL
        df_lista = leer_excel_guardia(lista_file)
    else:
        df_lista = procesar_generico(lista_txt, None)

    # Validar
    if df_parte is not None and df_lista is not None and not df_parte.empty and not df_lista.empty:
        try:
            # --- A. LIMPIEZA ---
            df_parte['jerarquia_norm'] = df_parte['Jerarquia'].apply(normalizar_jerarquia)
            df_lista['jerarquia_norm'] = df_lista['Jerarquia'].apply(normalizar_jerarquia)
            
            df_parte['nombre_limpio'] = df_parte['Nombre'].apply(limpiar_nombre)
            df_lista['nombre_limpio'] = df_lista['Nombre'].apply(limpiar_nombre)

            # --- B. COMPARACIÓN ---
            sobran = df_lista.copy()
            sobran['match_encontrado'] = False
            ausentes_data = []

            # Recorremos el PARTE buscando en la LISTA
            for idx_p, row_p in df_parte.iterrows():
                jerarquia_obj = row_p['jerarquia_norm']
                nombre_obj = row_p['nombre_limpio']
                encontrado = False
                
                # FILTRO: Solo comparamos con gente de la misma jerarquía
                candidatos = sobran[sobran['jerarquia_norm'] == jerarquia_obj]

                for idx_l, row_l in candidatos.iterrows():
                    if row_l['match_encontrado']: continue 
                    
                    # COMPARACIÓN FUZZY
                    ratio = fuzz.token_set_ratio(nombre_obj, row_l['nombre_limpio'])
                    
                    if ratio >= umbral:
                        encontrado = True
                        sobran.at[idx_l, 'match_encontrado'] = True
                        break
                
                if not encontrado:
                    # Guardamos los datos para la tabla de faltantes
                    ausentes_data.append({
                        'Jerarquia': row_p['Jerarquia'],
                        'Nombre': row_p['Nombre']
                    })

            # --- C. RESULTADOS EN TABLAS ---
            st.divider()
            r_col1, r_col2 = st.columns(2)
            
            # TABLA 1: FALTA AGREGAR
            with r_col1:
                st.success(f"✅ FALTA AGREGAR A LA LISTA ({len(ausentes_data)})")
                if ausentes_data:
                    df_ausentes = pd.DataFrame(ausentes_data)
                    # Mostramos tabla interactiva (copiable)
                    st.dataframe(df_ausentes, hide_index=True, use_container_width=True)
            
            # TABLA 2: SOBRA / BORRAR
            with r_col2:
                df_sobran = sobran[sobran['match_encontrado'] == False]
                st.error(f"❌ SOBRA / BORRAR DE LA LISTA ({len(df_sobran)})")
                if not df_sobran.empty:
                    # Mostramos solo columnas relevantes
                    st.dataframe(df_sobran[['Jerarquia', 'Nombre']], hide_index=True, use_container_width=True)

        except Exception as e:
             st.error(f"Error procesando: {e}")
    else:
        st.warning("⚠️ Faltan datos o no se encontró la hoja 'LISTA' en el Excel.")
