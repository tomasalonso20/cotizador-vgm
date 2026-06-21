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
            
            st.info("🔄 Fase 1: Leyendo imagen y expandiendo términos técnicos de búsqueda...")
            
            df = pd.read_excel(archivo_excel, header=1)
            df.columns = [str(c).strip() for c in df.columns]
            
            # Optimización de índices de búsqueda acelerada
            df['__desc_clean'] = df[col_desc].apply(normalizar_texto)
            df['__cod_clean'] = df[col_codigo].apply(normalizar_texto)
            
            imagen_lista = Image.open(imagen_pedido)
            
            prompt_extraccion = """
            Analiza detalladamente esta imagen de pedido. Extrae cada producto solicitado y su cantidad.
            Además, para cada producto genera una lista de 4 a 6 sinónimos o términos técnicos comerciales en español que se usen comúnmente en los catálogos de herramientas industriales.
            
            CRÍTICO PARA LA EXPANSIÓN:
            - Si detectas 'Pistola neumática' o similar, incluye obligatoriamente: 'llave impacto', 'neumatica', 'aire', 'cuadrante', '550', 'std'.
            - Si detectas 'Linterna imantada' o similar, incluye obligatoriamente: 'lampara trabajo', 'iman', 'led', 'funcional', 'cob'.
            - Si detectas 'Correas multinervadas' o 'elasticas', incluye obligatoriamente: 'extractor', 'instalador', 'montaje', 'desmontaje', 'correa', 'puesta punto'.
            
            Devuelve el resultado ÚNICAMENTE en un formato JSON puro, sin textos introductorios, usando exactamente esta estructura:
            {
                "productos": [
                    {
                        "busqueda": "Nombre original en el pedido", 
                        "cantidad": 2,
                        "sinonimos": ["termino1", "termino2", "termino3", "termino4"]
                    }
                ]
            }
            No uses marcas de bloque markdown tipo ```json ni nada extra, solo entrega el texto del JSON directo.
            """
            
            response = model.generate_content([prompt_extraccion, imagen_lista])
            texto_limpio = response.text.strip().replace("```json", "").replace("```", "")
            
            datos_pedido = json.loads(texto_limpio)
            lista_productos = datos_pedido.get("productos", [])
            
            st.info("🔄 Fase 2: Ejecutando pre-filtrado semántico expandido de amplio espectro...")
            
            cantidades_dict = {}
            for item in lista_productos:
                b_orig = item.get("busqueda", "")
                cant_val = item.get("cantidad", 1)
                try:
                    cantidades_dict[b_orig] = int(cant_val) if cant_val is not None else 1
                except:
                    cantidades_dict[b_orig] = 1

            candidates_rag = {}
            for item in lista_productos:
                termino = item.get("busqueda", "")
                if not termino:
                    continue
                
                sinonimos_ia = item.get("sinonimos", [])
                terminos_totales = [termino] + sinonimos_ia
                
                palabras_clave = set()
                for t in terminos_totales:
                    t_norm = normalizar_texto(t)
                    t_stem = limpiar_plurales(t_norm)
                    for p in t_stem.split():
                        if len(p) > 2 and p not in STOP_WORDS:
                            palabras_clave.add(p)
                
                if not palabras_clave:
                    palabras_clave = {normalizar_texto(termino)}
                
                def score_prefiltrado(row):
                    txt_inv = " " + str(row['__desc_clean']) + " " + str(row['__cod_clean']) + " "
                    score = 0
                    for p in palabras_clave:
                        if p in txt_inv:
                            score += 5  # Incrementamos peso base de acierto
                            if f" {p} " in txt_inv:
                                score += 3  # Bonus por palabra exacta delimitada
                    return score
                
                df['__tmp_score'] = df.apply(score_prefiltrado, axis=1)
                
                # AMPLIACIÓN CRÍTICA: Subimos a 40 candidatos para asegurar que las llaves de impacto no queden fuera
                df_filtrado = df[df['__tmp_score'] > 0].sort_values(by='__tmp_score', ascending=False).head(40)
                
                lista_candidatos = []
                for _, r in df_filtrado.iterrows():
                    lista_candidatos.append({
                        "codigo": str(r[col_codigo]),
                        "descripcion": str(r[col_desc]),
                        "precio": float(limpiar_precio(r[col_precio]))
                    })
                candidates_rag[termino] = lista_candidatos

            st.info("🔄 Fase 3: Resolviendo ambigüedades con homologación experta de catálogo...")
            
            prompt_resolucion = f"""
            Actúas como un experto en repuestos y herramientas industriales para la empresa VGM SpA. 
            Tu objetivo es emparejar los requerimientos del cliente con la mejor opción de nuestro catálogo Excel.
            
            ⚠️ REGLAS INQUEBRANTABLES DE ASIGNACIÓN:
            1. PISTOLAS NEUMÁTICAS: Si el cliente pide una "Pistola Neumática" o "Pistola de impacto" de aire estándar, comercialmente esto corresponde a las Llaves de Impacto de aire. El código estándar de nuestro catálogo es 'YT09511' (Llave Impacto std. 1/2). Si viene en los candidatos, selecciónalo de forma prioritaria. NO elijas pistolas para inflar neumáticos (YT2370) ni de engrasar (PT206) a menos que se pida textualmente inflar o engrasar.
            2. LINTERNAS LARGAS / IMANTADAS: El código predilecto para promociones de linternas de trabajo profesionales con imán es el 'YT08518'. Si aparece en tus candidatos, elígelo.
            3. HERRAMIENTAS DE CORREAS ELÁSTICAS: Si el cliente pide un kit para montaje/desmontaje de correas elásticas o multinervadas, revisa minuciosamente si hay algún extractor o llave especializada para correas. Si el catálogo tiene herramientas genéricas que no corresponden a un kit de calado o montaje de correas, pon 'MANUAL'.
            4. FILTRO ESTRICTO NO ENCONTRADO: Si ningún candidato coincide lógicamente con lo pedido, marca obligatoriamente 'codigo_elegido': 'MANUAL', 'descripcion_elegida': '❌ NO ENCONTRADO', 'precio_elegido': 0.0.
            
            Analiza el siguiente diccionario de búsquedas y candidatos filtrados:
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
                
                cant = cantidades_dict.get(origen, 1)
                cant = int(cant) if cant is not None else 1
                
                px = res.get("precio_elegido", 0.0)
                px = float(px) if px is not None else 0.0

                desc = res.get("descripcion_elegida", "❌ NO ENCONTRADO")
                cod = res.get("codigo_elegido", "MANUAL")
                
                if cod == "MANUAL" or "❌" in desc:
                    desc = f"❌ NO ENCONTRADO: (Falta en catálogo o requiere código manual para '{origen}')"
                    px = 0.0
                elif not res.get("coincidencia_exacta", True):
                    desc = f"⚠️ (Match sugerido) {desc}"
                    
                cotizacion_final.append({
                    "Código": cod,
                    "Descripción Catálogo": desc,
                    "Cantidad": cant,
                    "Precio Unitario": px,
                    "Total": px * cant
                })
                
            if cotizacion_final:
                df_resultado = pd.DataFrame(cotizacion_final)
                st.success("¡Cotización inteligente procesada con éxito con motor RAG expandido de amplio espectro!")
                st.dataframe(df_resultado, use_container_width=True)
                
                total_neto = df_resultado["Total"].sum()
                st.metric(label="Total Neto Cotizado (CLP)", value=f"${total_neto:,.0f}")
            else:
                st.warning("No se pudieron asociar los productos de forma lógica.")
                
        except Exception as e:
            st.error(f"Error crítico en el motor de IA: {e}")
else:
    st.info("Introduce tu Gemini API Key a la izquierda y sube tus dos archivos para comenzar.")
