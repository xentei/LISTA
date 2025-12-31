import os
import time

import streamlit as st
import pandas as pd
from io import StringIO
from thefuzz import fuzz
import re
import unicodedata

st.set_page_config(page_title="Control de Personal", layout="wide")

st.title("👮‍♂️ Comparador de Listas de Servicio")
st.markdown("""
<style>
.small-font { font-size:14px !important; color: #666; }
</style>
<p class="small-font">Modo Acción: Te dice qué agregar y qué borrar para que tu lista quede perfecta.</p>
""", unsafe_allow_html=True)

# --- CONFIGURACIÓN ---
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
    texto = re.sub(r'\([^)]*\)', '', texto)
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^a-zA-Z\s]', '', texto)
    return texto.strip().upper()

def procesar_texto_pegado(texto_raw):
    try:
        if not texto_raw: return None
        df = pd.read_csv(StringIO(texto_raw), sep='\t', header=None)
        
        primera_celda = str(df.iloc[0, 0]).lower()
        palabras_clave = ['jerarquia', 'grado', 'jerarquía', 'apellido', 'nombre']
        
        if any(x in primera_celda for x in palabras_clave):
            df = pd.read_csv(StringIO(texto_raw), sep='\t')
        else:
            nuevas_cols = ['Jerarquia', 'Nombre']
            if len(df.columns) > 2:
                nuevas_cols += [f'Col_{i}' for i in range(2, len(df.columns))]
            df.columns = nuevas_cols

        df.columns = df.columns.str.strip().str.lower()
        return df
    except:
        return None

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración")
    umbral = st.slider("Exigencia de coincidencia", 50, 100, 85)

# --- INTERFAZ (CAMBIOS APLICADOS AQUÍ) ---
col1, col2 = st.columns(2)
with col1:
    # Cambio de título solicitado
    st.subheader("📋 PEGA ACA EL PARTE")
    parte_input = st.text_area("Copia del Excel oficial", height=300, key="parte")
with col2:
    # Cambio de título solicitado
    st.subheader("📝 PEGA ACA LA LISTA")
    lista_input = st.text_area("Copia lo que escribiste", height=300, key="lista")

if st.button("🔍 COMPARAR AHORA", type="primary"):
    if parte_input and lista_input:
        df_parte = procesar_texto_pegado(parte_input)
        df_lista = procesar_texto_pegado(lista_input)

        if df_parte is not None and df_lista is not None and not df_parte.empty and not df_lista.empty:
            try:
                col_jer_p, col_nom_p = df_parte.columns[0], df_parte.columns[1]
                col_jer_l, col_nom_l = df_lista.columns[0], df_lista.columns[1]

                ausentes = []
                sobran = df_lista.copy()
                sobran['match_encontrado'] = False

                # 1. LIMPIEZA
                df_parte['jerarquia_norm'] = df_parte[col_jer_p].apply(normalizar_jerarquia)
                sobran['jerarquia_norm'] = sobran[col_jer_l].apply(normalizar_jerarquia)
                
                df_parte['nombre_limpio'] = df_parte[col_nom_p].apply(limpiar_nombre)
                sobran['nombre_limpio'] = sobran[col_nom_l].apply(limpiar_nombre)

                # 2. COMPARACIÓN
                for idx_p, row_p in df_parte.iterrows():
                    jerarquia_obj = row_p['jerarquia_norm']
                    nombre_obj = row_p['nombre_limpio']
                    encontrado = False
                    
                    candidatos = sobran[sobran['jerarquia_norm'] == jerarquia_obj]

                    for idx_l, row_l in candidatos.iterrows():
                        if row_l['match_encontrado']: continue 
                        nombre_candidato = row_l['nombre_limpio']
                        ratio = fuzz.token_set_ratio(nombre_obj, nombre_candidato)
                        
                        if ratio >= umbral:
                            encontrado = True
                            sobran.at[idx_l, 'match_encontrado'] = True
                            break
                    
                    if not encontrado:
                        ausentes.append(f"{row_p[col_jer_p]} - {row_p[col_nom_p]}")

                # 3. RESULTADOS (COLORES Y TEXTOS CAMBIADOS)
                st.divider()
                r_col1, r_col2 = st.columns(2)
                
                with r_col1:
                    # LÓGICA VERDE: AGREGA A LA LISTA
                    if len(ausentes) > 0:
                        st.success(f"✅ AGREGA A LA LISTA ESTOS {len(ausentes)}")
                        for p in ausentes:
                            st.write(f"- {p}")
                    else:
                        st.balloons()
                        st.success("✨ ¡Perfecto! No falta nadie.")
                
                with r_col2:
                    # LÓGICA ROJA: BORRAR DE LA LISTA
                    sobran_final = sobran[sobran['match_encontrado'] == False]
                    if len(sobran_final) > 0:
                        st.error(f"❌ BORRAR DE LA LISTA ({len(sobran_final)})")
                        for index, row in sobran_final.iterrows():
                            st.write(f"- {row[col_jer_l]} - {row[col_nom_l]}")
                    else:
                        st.success("✨ ¡Limpio! No sobra nadie.")
            
            except Exception as e:
                 st.error(f"Error: {e}")
        else:
            st.error("Error leyendo datos.")
    else:
        st.info("Pega las listas para comenzar.")

def main():
    print("Iniciando servicio Nur...")
    # Tu lógica aquí
    while True:
        # Ejemplo de loop para que el contenedor no se apague si es un bot
        print("Nur está corriendo...")
        time.sleep(60) 

if __name__ == "__main__":
    main()
