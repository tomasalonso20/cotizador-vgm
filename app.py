import streamlit as st
import google.generativeai as genai
import pandas as pd
import json

# Configuración de la página web
st.set_page_config(page_title="Cotizador Express - VGM SpA", layout="wide")
st.title("Cotizador Express - VGM SpA 🚀")

# Barra lateral para la API Key
with st.sidebar:
    st.subheader("Configuración")
    api_key = st.text_input("Ingresa tu Gemini API Key:", type="password")
    st.markdown("---")
    st.markdown("Desarrollado para VGM SpA")

# Columnas de la interfaz principal
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Sube tu lista de precios (Excel .xlsx)")
    archivo_excel = st.file_uploader("Selecciona tu archivo Excel optimizado", type=["xlsx"])

with col2:
    st.subheader("2. Sube el pantallazo del pedido (WhatsApp / Correo)")
    imagen_pedido = st.file_uploader("Selecciona la imagen del pedido", type=["png", "jpg", "jpeg"])

# Procesamiento cuando todo está cargado
if archivo_excel and imagen_pedido and api_key:
    try:
        # 1. Configurar la API de Gemini de forma segura
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.5-flash')
        
        st.info("🔄 Leyendo el catálogo y extrayendo los datos del pedido...")
        
        # 2. Cargar el Excel gigante localmente en la memoria (Rápido y Gratis)
        df = pd.read_excel(archivo_excel)
        df.columns = [str(c).strip() for c in df.columns] # Limpiar espacios en los nombres de columnas
        
        # Detectar de forma automática las columnas clave por palabras parecidas
        col_codigo = next((c for c in df.columns if 'cod' in c.lower() or 'id' in c.lower()), df.columns[0])
        col_desc = next((c for c in df.columns if 'desc' in c.lower() or 'nom' in c.lower() or 'art' in c.lower() or 'prod' in c.lower()), df.columns[1] if len(df.columns) > 1 else df.columns[0])
        col_precio = next((c for c in df.columns if 'prec' in c.lower() or 'val' in c.lower() or 'neto' in c.lower() or 'unit' in c.lower()), df.columns[-1])

        # 3. Pedirle a Gemini únicamente que extraiga el texto de la imagen (Ahorra 99% de tokens)
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
        
        response = model.generate_content([prompt_extraccion, imagen_pedido])
        texto_limpio = response.text.strip().replace("```json", "").replace("```", "")
        
        # Cargar los productos que Gemini descubrió en la foto
        datos_pedido = json.loads(texto_limpio)
        lista_productos = datos_pedido.get("productos", [])
        
        # 4. Motor de búsqueda interno en Python (Busca en los 2.000 productos al instante)
        cotizacion_final = []
        
        for item in lista_productos:
            termino_busqueda = str(item.get("busqueda", "")).lower().strip()
            cantidad = item.get("cantidad", 1)
            
            if not termino_busqueda:
                continue
                
            # Separar los términos por palabras para una búsqueda ultra flexible
            palabras = termino_busqueda.split()
            coincidencias = df.copy()
            
            # Filtrar las filas del Excel que contengan todas las palabras que Gemini leyó
            for palabra in palabras:
                coincidencias = coincidencias[
                    coincidencias[col_desc].astype(str).str.lower().str.contains(palabra, na=False) |
                    coincidencias[col_codigo].astype(str).str.lower().str.contains(palabra, na=False)
                ]
            
            if not coincidencias.empty:
                # Tomamos el primer resultado más preciso del catálogo
                mejor_match = coincidencias.iloc[0]
                try:
                    precio_unitario = float(mejor_match[col_precio])
                except:
                    precio_unitario = 0.0
                    
                cotizacion_final.append({
                    "Código": mejor_match[col_codigo],
                    "Descripción Catálogo": mejor_match[col_desc],
                    "Cantidad": cantidad,
                    "Precio Unitario": precio_unitario,
                    "Total": precio_unitario * cantidad
                })
            else:
                # Si el producto de la foto no coincide con nada de tu Excel
                cotizacion_final.append({
                    "Código": "N/A",
                    "Descripción Catálogo": f"⚠️ REVISAR: No se halló en Excel como '{item.get('busqueda')}'",
                    "Cantidad": cantidad,
                    "Precio Unitario": 0.0,
                    "Total": 0.0
                })
        
        # 5. Dibujar los resultados finales en pantalla
        if cotizacion_final:
            df_resultado = pd.DataFrame(cotizacion_final)
            st.success("¡Cotización generada exitosamente con tu catálogo completo!")
            st.dataframe(df_resultado, use_container_width=True)
            
            # Calcular e imprimir el total monetario
            total_neto = df_resultado["Total"].sum()
            st.metric(label="Total Neto Cotizado (CLP)", value=f"${total_neto:,.0f}")
        else:
            st.warning("No se detectaron productos procesables en la imagen.")
            
    except Exception as e:
        st.error(f"Error técnico durante el proceso: {e}")
        st.info("Consejo: Asegúrate de que las columnas de tu Excel tengan títulos claros como 'Código', 'Descripción' y 'Precio'.")
else:
    st.info("Por favor, introduce tu Gemini API Key en la izquierda y sube los dos archivos para cotizar.")
