import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
from PIL import Image
import unicodedata
import io

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

def generar_excel_comercial(df_cotiz, cliente, empresa, desc_porcentaje, total_neto, iva, total_bruto):
    output = io.BytesIO()
    df_excel = df_cotiz.copy()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_excel.to_excel(writer, sheet_name="Cotización", index=False, startrow=7)
        
        workbook = writer.book
        worksheet = writer.sheets["Cotización"]
        
        # Inyección de metadatos comerciales en la cabecera del Excel
        worksheet["A1"] = "COTIZACIÓN COMERCIAL - VGM SpA"
        worksheet["A3"] = f"Empresa: {empresa if empresa else 'No especificada'}"
        worksheet["A4"] = f"Cliente: {cliente if cliente else 'No especificado'}"
        worksheet["A5"] = f"Descuento Comercial: {desc_porcentaje}%"
        worksheet["E3"] = f"Fecha: {pd.Timestamp.now().strftime('%d-%m-%Y')}"
        
        # Añadir bloque de totales al final del archivo
        start_row = len(df_excel) + 9
        worksheet[f"D{start_row}"] = "Total Neto:"
        worksheet[f"E{start_row}"] = total_neto
        worksheet[f"D{start_row+1}"] = "IVA (19%):"
        worksheet[f"E{start_row+1}"] = iva
        worksheet[f"D{start_row+2}"] = "Total Bruto:"
        worksheet[f"E{start_row+2}"] = total_bruto
        
    return output.getvalue()

# Inicialización de la barra lateral
with st.sidebar:
    st.subheader("💼 Datos de la Cotización")
    nombre_cliente = st.text_input("Nombre del Cliente:", placeholder="Ej: Claudia Araya")
    empresa_cliente = st.text_input("Empresa / Entidad:", placeholder="Ej: Tattersall")
    descuento_aplicar = st.number_input("Descuento a aplicar (%)", min_value=0, max_value=100, value=0, step=1)
    
    st.markdown("---")
    
    # Configuración técnica congelada/oculta por defecto
    with st.expander("⚙️ Configuración del Sistema (API / Columnas)", expanded=False):
        api_key = st.text_input("Ingresa tu Gemini API Key:", type="password")
        st.subheader("📊 Verificación de Columnas")
        col_codigo = "CODIGO"
        col_desc = "DESCRIPCION"
        col_precio = "PRECIO UNITARIO NETO"

# Columnas de la interfaz principal
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Sube tu lista de precios (Excel .xlsx)")
    archivo_excel = st.file_uploader("Selecciona tu archivo Excel optimizado", type=["xlsx"])

with col2:
    st.subheader("2. Sube el pantallazo del pedido (WhatsApp / Correo)")
    imagen_pedido = st.file_uploader("Selecciona la imagen del pedido", type=["png", "jpg", "jpeg"])

# Procesamiento interno de columnas del Excel (mantiene lógica de asignación automática)
if archivo_excel:
    try:
        df_preview = pd.read_excel(archivo_excel, header=1, nrows=5)
        columnas_disponibles = [str(c).strip() for c in df_preview.columns]
        
        idx_cod = next((i for i, c in enumerate(columnas_disponibles) if 'cod' in c.lower() or 'id' in c.lower()), 0)
        idx_desc = next((i for i, c in enumerate(columnas_disponibles) if 'desc' in c.lower() or 'nom' in c.lower() or 'art' in c.lower() or 'prod' in c.lower() or 'det' in c.lower()), 1)
        idx_precio = next((i for i, c in enumerate(columnas_disponibles) if 'prec' in c.lower() or 'val' in c.lower() or 'neto' in c.lower() or 'unit' in c.lower()), len(columnas_disponibles) - 1)
        
        col_codigo = columnas_disponibles[idx_cod]
        col_desc = columnas_disponibles[idx_desc]
        col_precio = columnas_disponibles[idx_precio]
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
            
            df['__desc_clean'] = df[col_desc].apply(normalizar_texto)
            df['__cod_clean'] = df[col_codigo].apply(normalizar_texto)
            
            imagen_lista = Image.open(imagen_pedido)
            
            prompt_extraccion = """
            Analiza detalladamente esta imagen de pedido. Extrae cada producto solicitado y su cantidad.
            Además, para cada producto genera una lista de 4 a 6 sinónimos o términos técnicos comerciales en español que se usen comúnmente en los catálogos de herramientas industriales.
            
            CRÍTICO PARA LA EXPANSIÓN:
            - Si detectas 'Pistola neumática' o similar, incluye obligatoriamente: 'llave impacto', 'neumatica', 'aire', 'cuadrante', '550', 'std'.
            - Si detectas 'Linterna imantada' o similar, incluye obligatoriamente: 'lampara trabajo', 'iman', 'led', 'funcional', 'cob'.
            - Si detectas 'Linterna led de cabeza', 'linterna de cabeza', 'lampara de cabeza' o 'frontal', incluye obligatoriamente: 'l-head-1', 'l_head_1', 'cabeza', 'frontal', 'cintillo'. NO agregues el término 'imantada'.
            - Si detectas 'Correas multinervadas' o 'elasticas', incluye obligatoriamente: 'extractor', 'instalador', 'montaje', 'desmontaje', 'correa', 'puesta punto'.
            - Si detectas 'Llave de impacto inalámbrica' o similar, incluye obligatoriamente: 'yt8277935', 'yt8277925', 'bateria', 'inalambrico', 'torque'.
            
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
                            score += 5
                            if f" {p} " in txt_inv:
                                score += 3
                    return score
                
                df['__tmp_score'] = df.apply(score_prefiltrado, axis=1)
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
            2. LINTERNAS LARGAS / IMANTADAS: El código predilecto para promociones de linternas de trabajo profesionales con imán es el 'YT08518'. Si aparece en tus candidatos, elígelo únicamente si piden linternas magnéticas, imantadas o lámparas de taller portátiles.
            3. LINTERNAS/LÁMPARAS DE CABEZA (FRONTALES): Si el cliente pide una "linterna led de cabeza", "linterna de cabeza", "lámpara de cabeza" o "frontal", bajo ninguna circunstancia elijas una linterna imantada o lámpara de taller (como YT08518). El código exacto asignado a este producto es 'L-HEAD-1'. Si este código aparece en tus candidatos, elígelo de forma obligatoria y prioritaria.
            4. HERRAMIENTAS DE CORREAS ELÁSTICAS: Si el cliente pide un kit para montaje/desmontaje de correas elásticas o multinervadas, revisa minuciosamente si hay algún extractor o llave especializada para correas. Si el catálogo tiene herramientas genéricas que no corresponden a un kit de calado o montaje de correas, pon 'MANUAL'.
            5. LLAVES DE IMPACTO INALÁMBRICAS: Cuando el pedido mencione una "llave de impacto inalámbrica" (o a batería, inalámbrico) sin detallar un código explícito, debes priorizar obligatoriamente estos dos códigos en este estricto orden de preferencia:
               - 1ra Opción (Preferencia Absoluta): Código 'YT8277935'
               - 2da Opción (Alternativa): Código 'YT8277925'
               NO elijas el modelo industrial YT828073 u otros códigos similares a menos que el cliente pida textualmente un torque gigante de 2400NM o un cuadrante de 3/4". De lo contrario, quédate siempre con las dos prioridades de arriba.
            6. FILTRO ESTRICTO NO ENCONTRADO: Si ningún candidato coincide lógicamente con lo pedido, marca obligatoriamente 'codigo_elegido': 'MANUAL', 'descripcion_elegida': '❌ NO ENCONTRADO', 'precio_elegido': 0.0.
            
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
                cant = int(cantidades_dict.get(origen, 1))
                px_lista = float(res.get("precio_elegido", 0.0))
                desc = res.get("descripcion_elegida", "❌ NO ENCONTRADO")
                cod = res.get("codigo_elegido", "MANUAL")
                
                if cod == "MANUAL" or "❌" in desc:
                    desc = f"❌ NO ENCONTRADO: (Falta en catálogo o requiere código manual para '{origen}')"
                    px_lista = 0.0
                elif not res.get("coincidencia_exacta", True):
                    desc = f"⚠️ (Match sugerido) {desc}"
                
                # Aplicación matemática del descuento comercial por ítem
                descuento_unidades = px_lista * (descuento_aplicar / 100)
                px_final_neto = px_lista - descuento_unidades
                total_item_neto = px_final_neto * cant
                    
                cotizacion_final.append({
                    "Código": cod,
                    "Descripción Catálogo": desc,
                    "Cantidad": cant,
                    "Precio Lista (Neto)": px_lista,
                    "Descuento Aplicado": f"{descuento_aplicar}%",
                    "Precio Final (Neto)": px_final_neto,
                    "Total Neto": total_item_neto
                })
                
            if cotizacion_final:
                df_resultado = pd.DataFrame(cotizacion_final)
                st.success("¡Cotización inteligente procesada con éxito!")
                
                # Mostrar tabla de resultados con los nuevos cálculos de cara al usuario
                st.dataframe(df_resultado, use_container_width=True)
                
                # Bloque de Cálculos Tributarios Estándar (Chile)
                subtotal_lista = (df_resultado["Precio Lista (Neto)"] * df_resultado["Cantidad"]).sum()
                total_neto_final = df_resultado["Total Neto"].sum()
                descuento_total_pesos = subtotal_lista - total_neto_final
                iva_calculado = total_neto_final * 0.19
                total_bruto = total_neto_final + iva_calculado
                
                st.markdown("### 📊 Desglose de Valores Comerciales (CLP)")
                c_neto, c_desc, c_neto_f, c_iva, c_bruto = st.columns(5)
                
                c_neto.metric(label="Total Lista Neto", value=f"${subtotal_lista:,.0f}")
                c_desc.metric(label="Descuento Total", value=f"-${descuento_total_pesos:,.0f}")
                c_neto_f.metric(label="Neto Final Cliente", value=f"${total_neto_final:,.0f}")
                c_iva.metric(label="IVA (19%)", value=f"${iva_calculado:,.0f}")
                c_bruto.metric(label="Total Bruto a Pagar", value=f"${total_bruto:,.0f}")
                
                st.markdown("---")
                st.subheader("📥 Descargar Documento Comercial")
                
                # Generación del archivo binario de Excel para descarga inmediata
                excel_binario = generar_excel_comercial(
                    df_resultado, nombre_cliente, empresa_cliente, 
                    descuento_aplicar, total_neto_final, iva_calculado, total_bruto
                )
                
                nombre_archivo_excel = f"Cotizacion_{empresa_cliente.replace(' ', '_') if empresa_cliente else 'Cliente'}_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx"
                
                st.download_button(
                    label="🟢 Descargar Cotización en Excel",
                    data=excel_binario,
                    file_name=nombre_archivo_excel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No se pudieron asociar los productos de forma lógica.")
                
        except Exception as e:
            st.error(f"Error crítico en el motor de IA: {e}")
else:
    st.info("Introduce tu Gemini API Key a la izquierda y sube tus dos archivos para comenzar.")
