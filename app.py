import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
from PIL import Image
import unicodedata

# Configuración de la página web
st.set_page_config(page_title="Cotizador Express - VGM SpA", layout="wide")
st.title("Cotizador Express - VGM SpA 🚀")

# Función para limpiar texto (quitar acentos, minúsculas y caracteres raros)
def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    for char in ['( ', ' )', '(', ')', '-', ',', '.', '/', '"', "'", '+']:
        texto = texto.replace(char, ' ')
    return texto

# Función inteligente para remover el plural en español ('s') y mejorar búsquedas
def limpiar_plurales(texto):
    palabras = texto.split()
    limpias = []
    for p in palabras:
        if len(p) > 3 and p.endswith('s'):
            limpias.append(p[:-1])  # Quita la 's' al final
        else:
            limpias.append(p)
    return " ".join(limpias)

# Función para procesar y limpiar precios chilenos de forma segura
def limpiar_precio(valor):
    if pd.isna(valor):
        return 0.0
    val_str = str(valor).strip().replace("$", "").replace(" ", "")
    if not val_str:
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        if "," in val_str:
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(".", "")
        try:
            return float(val_str)
        except:
            return 0.0

# Inicialización de la barra lateral
with st.sidebar:
    st.subheader("⚙️ Configuración Básica")
    api_key = st.text_input("Ingresa tu Gemini API Key:", type="password")
    st.markdown("---")
    st.subheader("📊 Verificación de Columnas")

# Columnas de la interfaz principal
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Sube tu lista de precios (Excel .xlsx)")
    archivo_excel = st.file_uploader("Selecciona tu archivo Excel optimizado", type=["xlsx"])

with col2:
    st.subheader("2. Sube el pantallazo del pedido (WhatsApp / Correo)")
    imagen_pedido = st.file_uploader("Selecciona la imagen del pedido", type=["png", "jpg", "jpeg"])

# LEER EXCEL: Ignorando la fila de logos superiores
if archivo_excel:
    try:
        df_preview = pd.read_excel(archivo_excel, header=1, nrows=5)
        columnas_disponibles = [str(c).strip() for c in df_preview.columns]
        
        idx_cod = next((i for i, c in enumerate(columnas_disponibles) if 'cod' in c.lower() or 'id' in c.lower()), 0)
        idx_desc = next((i for i, c in enumerate(columnas_disponibles) if 'desc' in c.lower() or 'nom' in c.lower() or 'art' in c.lower() or 'prod' in c.lower() or 'det' in c.lower()), 1)
        idx_precio = next((i for i, c in enumerate(columnas_disponibles) if 'prec' in c.lower() or 'val' in c.lower() or 'neto' in c.lower() or 'unit' in c.lower()), len(columnas_disponibles) - 1)
        
        with st.sidebar:
            st.write("Confirma que los títulos correspondan a tu Excel:")
            col_codigo = st.selectbox("Columna de Código:", columnas_disponibles, index=idx_cod)
            col_desc = st.selectbox("Columna de Descripción:", columnas_disponibles, index=idx_desc)
            col_precio = st.selectbox("Columna de Precio Neto:", columnas_disponibles, index=idx_precio)
            st.success("✅ Títulos de columnas cargados con éxito")
    except Exception as e:
        st.sidebar.error(f"Error al analizar la estructura del Excel: {e}")

# Ejecución del motor inteligente híbrido
if archivo_excel and imagen_pedido and api_key:
    if st.button("🔥 Generar Cotización Automática"):
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.5-flash')
            
            st.info("🔄 Buscando y cruzando coincidencias en tu catálogo...")
            
            df = pd.read_excel(archivo_excel, header=1)
            df.columns = [str(c).strip() for c in df.columns]
            
            # Pre-procesar el catálogo completo quitando acentos y minúsculas
            df['__desc_clean'] = df[col_desc].apply(normalizar_texto)
            df['__cod_clean'] = df[col_codigo].apply(normalizar_texto)
            
            imagen_lista = Image.open(imagen_pedido)
            
            prompt_extraccion = """
            Analiza detalladamente esta imagen de pedido. Extrae cada producto solicitado y su cantidad.
            Devuelve el resultado ÚNICAMENTE en un formato JSON puro, sin textos introductorios, usando exactamente esta estructura:
            {
                "productos": [
                    {"busqueda": "nombre o marca clave del producto 1", "cantidad": 2},
                    {"busqueda": "nombre o marca clave del producto 2", "cantidad": 5}
                ]
            }
            No uses marcas de bloque markdown tipo ```json ni nada extra, solo entrega el texto del JSON directo.
            """
            
            response = model.generate_content([prompt_extraccion, imagen_lista])
            texto_limpio = response.text.strip().replace("```json", "").replace("```", "")
            
            datos_pedido = json.loads(texto_limpio)
            lista_productos = datos_pedido.get("productos", [])
            
            cotizacion_final = []
            
            for item in lista_productos:
                termino_busqueda = item.get("busqueda", "")
                cantidad = item.get("cantidad", 1)
                
                if not termino_busqueda:
                    continue
                
                # Normalizar y quitar plurales tanto al pedido como al inventario
                termino_norm = normalizar_texto(termino_busqueda)
                termino_sin_plural = limpiar_plurales(termino_norm)
                palabras_busqueda = [p for p in termino_sin_plural.split() if len(p) > 2]
                
                if not palabras_busqueda:
                    palabras_busqueda = [termino_norm]
                
                # Calcular cuántas palabras del pedido calzan con el artículo del inventario
                def calcular_coincidencias(row):
                    texto_inventario = limpiar_plurales(str(row['__desc_clean']) + " " + str(row['__cod_clean']))
                    coincidencias = sum(1 for p in palabras_busqueda if p in texto_inventario)
                    return coincidencias
                
                df['__score'] = df.apply(calcular_coincidencias, axis=1)
                max_score = df['__score'].max()
                
                # Si encontró algún grado de coincidencia (al menos 1 palabra clave)
                if max_score > 0:
                    mejor_coincidencia = df[df['__score'] == max_score].iloc[0]
                    precio_final = limpiar_precio(mejor_coincidencia[col_precio])
                    
                    # Si el match es parcial (ej: menos de la mitad de las palabras calzan), ponemos alerta visual
                    score_relativo = max_score / len(palabras_busqueda)
                    alerta_visual = ""
                    if score_relativo < 0.5:
                        alerta_visual = "⚠️ (Revisar coincidencia) "
                    
                    cotizacion_final.append({
                        "Código": mejor_coincidencia[col_codigo],
                        "Descripción Catálogo": f"{alerta_visual}{mejor_coincidencia[col_desc]}",
                        "Cantidad": cantidad,
                        "Precio Unitario": precio_final,
                        "Total": precio_final * cantidad
                    })
                else:
                    cotizacion_final.append({
                        "Código": "MANUAL",
                        "Descripción Catálogo": f"❌ NO ENCONTRADO: '{termino_busqueda}'",
                        "Cantidad": cantidad,
                        "Precio Unitario": 0.0,
                        "Total": 0.0
                    })
            
            if cotizacion_final:
                df_resultado = pd.DataFrame(cotizacion_final)
                st.success("¡Cotización procesada con éxito!")
                st.dataframe(df_resultado, use_container_width=True)
                
                total_neto = df_resultado["Total"].sum()
                st.metric(label="Total Neto Cotizado (CLP)", value=f"${total_neto:,.0f}")
            else:
                st.warning("No se procesaron productos interpretables.")
                
        except Exception as e:
            st.error(f"Error durante el proceso: {e}")
else:
    st.info("Introduce tu Gemini API Key a la izquierda y sube tus dos archivos para comenzar.")
