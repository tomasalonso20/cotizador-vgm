import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
from PIL import Image
import unicodedata

# Configuración de la página web
st.set_page_config(page_title="Cotizador Express - VGM SpA", layout="wide")
st.title("Cotizador Express - VGM SpA 🚀")

# Lista de palabras vacías en español para limpiar búsquedas
STOP_WORDS = {'de', 'para', 'con', 'un', 'una', 'el', 'la', 'los', 'las', 'del', 'al', 'en', 'y', 'por', 'sobre', 'kit', 'juego', 'set'}

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    for char in ['( ', ' )', '(', ')', '-', ',', '.', '/', '"', "'", '+']:
        texto = texto.replace(char, ' ')
    return texto

def limpiar_plurales(texto):
    palabras = texto.split()
    limpias = []
    for p in palabras:
        if len(p) > 3 and p.endswith('s'):
            limpias.append(p[:-1])
        else:
            limpias.append(p)
    return " ".join(limpias)

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

# Procesamiento de columnas del Excel
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

# Ejecución del motor híbrido con IA integrada
if archivo_excel and imagen_pedido and api_key:
    if st.button("🔥 Generar Cotización Inteligente"):
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.5-flash')
            
            st.info("🔄 Fase 1: Leyendo imagen del pedido con Visión Artificial...")
            
            df = pd.read_excel(archivo_excel, header=1)
            df.columns = [str(c).strip() for c in df.columns]
            
            # Optimización de índices de búsqueda acelerada
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
            
            st.info("🔄 Fase 2: Pre-filtrando candidatos semánticos del catálogo...")
            
            # Mapear cantidades indexadas por la búsqueda original
            cantidades_dict = {item.get("busqueda", ""): item.get("cantidad", 1) for item in lista_productos}
            candidates_rag = {}
            
            for item in lista_productos:
                termino = item.get("busqueda", "")
                if not termino:
                    continue
                
                termino_norm = normalizar_texto(termino)
                termino_sin_plural = limpiar_plurales(termino_norm)
                palabras_clave = [p for p in termino_sin_plural.split() if len(p) > 2 and p not in STOP_WORDS]
                
                if not palabras_clave:
                    palabras_clave = [termino_norm]
                
                def score_prefiltrado(row):
                    txt_inv = str(row['__desc_clean']) + " " + str(row['__cod_clean'])
                    return sum(2 if p in txt_inv else 0 for p in palabras_clave)
                
                df['__tmp_score'] = df.apply(score_prefiltrado, axis=1)
                df_filtrado = df[df['__tmp_score'] > 0].sort_values(by='__tmp_score', ascending=False).head(10) # Subimos a 10 candidatos para no dejar fuera los códigos correctos
                
                lista_candidatos = []
                for _, r in df_filtrado.iterrows():
                    lista_candidatos.append({
                        "codigo": str(r[col_codigo]),
                        "descripcion": str(r[col_desc]),
                        "precio": float(limpiar_precio(r[col_precio]))
                    })
                candidates_rag[termino] = lista_candidatos

            st.info("🔄 Fase 3: Resolviendo ambigüedades con el cerebro analítico de Gemini...")
            
            prompt_resolucion = f"""
            Actúas como un expertísimo en repuestos y herramientas industriales para VGM SpA. 
            Cruza los productos solicitados por el cliente con las mejores opciones de candidatos encontradas en nuestro catálogo Excel.
            
            ⚠️ REGLAS INQUEBRANTABLES DE NEGOCIO (Si fallas en esto, la cotización queda mal hecha):
            
            1. REGLA CRÍTICA PARA PISTOLAS NEUMÁTICAS / DE IMPACTO:
               - "Pistola de impacto", "Pistola neumática" o "Llave de impacto" son términos equivalentes en el taller.
               - Si la búsqueda original es genérica (ej: "Pistola Neumática" o "Pistola de impacto") sin especificar torque o tamaño especial, DEBES ELEGIR EL CÓDIGO 'YT09511' (Llave de impacto std 1/2).
               - PROHIBIDO: No confundas esto con pistolas para inflar neumáticos (como el código YT2370). Jamás asignes una pistola de inflar si el cliente pide una pistola neumática general de taller.
            
            2. REGLA CRÍTICA PARA LINTERNAS IMANTADAS:
               - Si el cliente solicita una "linterna imantada" o "linterna", el modelo exacto que corresponde a la imagen de nuestro correo de referencia es la 'YT08518'. 
               - Si el código 'YT08518' aparece listado entre tus candidatos, DEBES ELEGIRLO OBLIGATORIAMENTE.
               - PROHIBIDO: No elijas lámparas UV para detección de fugas (como YT08500) a menos que el cliente use explícitamente el término "UV" o "fugas".
            
            3. FILTRO DE COINCIDENCIA LOGICA ESTRICTA:
               - Si los candidatos NO tienen ninguna relación real con el producto pedido (ej: el cliente pide un "Kit montaje" y el catálogo solo arroja herramientas sueltas o nada que ver), debes marcarlo como NO ENCONTRADO. Pon obligatoriamente 'codigo_elegido': 'MANUAL' y 'descripcion_elegida': '❌ NO ENCONTRADO'. No inventes cruces erróneos.
            
            4. PRECIOS PARA ITEMS MANUALES:
               - Si decides dejar un producto como 'MANUAL', el campo 'precio_elegido' DEBE SER obligatoriamente 0.0.
            
            Analiza el siguiente diccionario de búsquedas y candidatos:
            {json.dumps(candidates_rag, ensure_ascii=False, indent=2)}
            
            Devuelve ÚNICAMENTE un JSON estructurado de la siguiente forma, sin bloques markdown ni texto adicional:
            {{
                "resultados": [
                    {{
                        "busqueda_original": "nombre exacto de la busqueda original",
                        "codigo_elegido": "código real del catálogo o 'MANUAL'",
                        "descripcion_elegida": "descripción exacta del catálogo o '❌ NO ENCONTRADO'",
                        "precio_elegido": 12345.0,
                        "coincidencia_exacta": true o false
                    }}
                ]
            }}
            """
            
            response_resolucion = model.generate_content(prompt_resolucion)
            texto_resolucion = response_resolucion.text.strip().replace("```json", "").replace("```", "")
            
            datos_finales = json.loads(texto_resolucion)
            resultados_lista = datos_finales.get("resultados", [])
            
            cotizacion_final = []
            for res in resultados_lista:
                origen = res.get("busqueda_original", "")
                
                # --- ESCUDO PROTECTOR DE TIPOS DE DATOS ---
                cant = cantidades_dict.get(origen, 1)
                if cant is None:
                    cant = 1
                try:
                    cant = int(cant)
                except:
                    cant = 1
                
                px = res.get("precio_elegido", 0.0)
                if px is None:
                    px = 0.0
                try:
                    px = float(px)
                except:
                    px = 0.0
                # ------------------------------------------

                desc = res.get("descripcion_elegida", "")
                
                if not res.get("coincidencia_exacta", True) and "❌" not in desc:
                    desc = f"⚠️ (Match aproximado) {desc}"
                    
                cotizacion_final.append({
                    "Código": res.get("codigo_elegido", "MANUAL"),
                    "Descripción Catálogo": desc,
                    "Cantidad": cant,
                    "Precio Unitario": px,
                    "Total": px * cant
                })
                
            if cotizacion_final:
                df_resultado = pd.DataFrame(cotizacion_final)
                st.success("¡Cotización inteligente procesada con éxito!")
                st.dataframe(df_resultado, use_container_width=True)
                
                total_neto = df_resultado["Total"].sum()
                st.metric(label="Total Neto Cotizado (CLP)", value=f"${total_neto:,.0f}")
            else:
                st.warning("No se pudieron asociar los productos de forma lógica.")
                
        except Exception as e:
            st.error(f"Error crítico en el motor de IA: {e}")
else:
    st.info("Introduce tu Gemini API Key a la izquierda y sube tus dos archivos para comenzar.")
