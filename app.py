import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
from PIL import Image
import unicodedata
import io
import os

# Importaciones nativas de openpyxl para diseño y manejo de imágenes
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

# Configuración de la página web
st.set_page_config(page_title="Cotizador Express - VGM SpA", layout="wide")
st.title("Cotizador Express - VGM SpA 🚀 (Edición Móvil Ultra-Eficiente)")

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

# FUNCIÓN MAESTRA: Escanea y detecta la cabecera real saltándose filas vacías o títulos del ERP
def leer_csv_tolerante(ruta_archivo):
    if not os.path.exists(ruta_archivo):
        return None
    encodings_a_probar = ['utf-8', 'latin1', 'iso-8859-1', 'utf-8-sig', 'cp1252']
    delimitadores_a_probar = [';', ',']
    
    for enc in encodings_a_probar:
        for sep in delimitadores_a_probar:
            try:
                with open(ruta_archivo, 'r', encoding=enc) as f:
                    lineas = [f.readline() for _ in range(15)]
                
                fila_cabecera_idx = 0
                for idx, line in enumerate(lineas):
                    line_low = line.lower()
                    if 'cod' in line_low or 'desc' in line_low or 'prec' in line_low or 'art' in line_low:
                        fila_cabecera_idx = idx
                        break
                
                df_res = pd.read_csv(ruta_archivo, sep=sep, encoding=enc, skiprows=fila_cabecera_idx)
                df_res = df_res.dropna(how='all', axis=1)
                df_res = df_res.dropna(how='all', axis=0)
                
                if df_res.shape[1] >= 2:
                    return df_res
            except:
                continue
    return None

# FUNCIÓN: Generación de Excel con Calce Exacto e Inserción de Imágenes Dinámicas
def generar_excel_comercial(df_cotiz, cliente, empresa, nro_cotiz, total_neto, iva, total_bruto, logo_bytes=None, dict_imagenes=None):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cotización"
    
    ws.views.sheetView[0].showGridLines = True
    
    font_titulo = Font(name="Arial", size=12, bold=True, color="1F497D")
    font_cabecera_tabla = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_negrita = Font(name="Arial", size=10, bold=True)
    font_normal = Font(name="Arial", size=10)
    font_firma = Font(name="Arial", size=11, bold=True, italic=True)
    
    fill_azul_header = PatternFill(start_color="365F91", end_color="365F91", fill_type="solid")
    fill_totales = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
    
    borde_delgado = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
    )
    
    ws["G1"] = f"Fecha: {pd.Timestamp.now().strftime('%d-%m-%Y')}"
    ws["G1"].font = font_negrita
    ws["G1"].alignment = Alignment(horizontal="right")
    
    if logo_bytes:
        try:
            img_stream = io.BytesIO(logo_bytes)
            img = OpenpyxlImage(img_stream)
            img.width = 130
            img.height = 50
            ws.add_image(img, 'A1')
        except:
            pass
            
    ws["A3"] = f"COTIZACIÓN N°{nro_cotiz}"
    ws["A3"].font = font_titulo
    
    ws["A4"] = "76.834.968-1"
    ws["A4"].font = font_negrita
    
    ws["A5"] = "Chopin 2848. San Joaquín. Santiago"
    ws["A5"].font = font_normal
    
    ws["A6"] = f"Sr(a).: {cliente if cliente else 'No especificado'}"
    ws["A6"].font = font_negrita
    
    ws["A7"] = f"Empresa: {empresa if empresa else 'No especificada'}"
    ws["A7"].font = font_negrita
    
    ws["A8"] = "En atención a su gentil solicitud de cotización, tenemos el agrado de hacer llegar a usted nuestra propuesta:"
    ws["A8"].font = font_normal
    
    titulos_columnas = ["CODIGO", "MARCA", "DESCRIPCIÓN", "PRECIO UNITARIO NETO", "CANTIDAD", "PRECIO UNITARIO TOTAL", "IMAGEN REFERENCIAL"]
    fila_tabla_inicio = 10
    
    for col_idx, texto_col in enumerate(titulos_columnas, 1):
        celda = ws.cell(row=fila_tabla_inicio, column=col_idx, value=texto_col)
        celda.font = font_cabecera_tabla
        celda.fill = fill_azul_header
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = borde_delgado
    ws.row_dimensions[fila_tabla_inicio].height = 28
    
    fila_actual = fila_tabla_inicio + 1
    for _, fila in df_cotiz.iterrows():
        cod_original = str(fila["Código"])
        cod_limpio = cod_original.strip().lower()
        
        ws.cell(row=fila_actual, column=1, value=cod_original).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=fila_actual, column=2, value=str(fila["Marca"])).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=fila_actual, column=3, value=str(fila["Descripción Catálogo"])).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        c_neto = ws.cell(row=fila_actual, column=4, value=float(fila["Precio Final (Neto)"]))
        c_neto.number_format = '$#,##0'
        c_neto.alignment = Alignment(horizontal="right", vertical="center")
        
        ws.cell(row=fila_actual, column=5, value=int(fila["Cantidad"])).alignment = Alignment(horizontal="center", vertical="center")
        
        c_total = ws.cell(row=fila_actual, column=6, value=float(fila["Total Neto"]))
        c_total.number_format = '$#,##0'
        c_total.alignment = Alignment(horizontal="right", vertical="center")
        
        ws.cell(row=fila_actual, column=7, value="").alignment = Alignment(horizontal="center", vertical="center")
        if dict_imagenes and cod_limpio in dict_imagenes:
            try:
                img_prod_stream = io.BytesIO(dict_imagenes[cod_limpio])
                img_excel = OpenpyxlImage(img_prod_stream)
                img_excel.width = 115
                img_excel.height = 80
                ws.add_image(img_excel, f"G{fila_actual}")
            except:
                ws.cell(row=fila_actual, column=7, value="[Error de Imagen]").font = font_normal
        
        for c_idx in range(1, 8):
            celda_f = ws.cell(row=fila_actual, column=c_idx)
            celda_f.font = font_normal
            celda_f.border = borde_delgado
            
        ws.row_dimensions[fila_actual].height = 65  
        fila_actual += 1
        
    ws.cell(row=fila_actual, column=1, value="Observaciones:").font = font_negrita
    
    celda_lbl_sub = ws.cell(row=fila_actual, column=5, value="SUBTOTAL")
    celda_lbl_sub.font = font_negrita
    celda_lbl_sub.alignment = Alignment(horizontal="right", vertical="center")
    celda_lbl_sub.border = borde_delgado
    
    celda_val_sub = ws.cell(row=fila_actual, column=6, value=total_neto)
    celda_val_sub.font = font_negrita
    celda_val_sub.number_format = '$#,##0'
    celda_val_sub.alignment = Alignment(horizontal="right", vertical="center")
    celda_val_sub.border = borde_delgado
    fila_actual += 1
    
    celda_lbl_iva = ws.cell(row=fila_actual, column=5, value="IVA")
    celda_lbl_iva.font = font_negrita
    celda_lbl_iva.alignment = Alignment(horizontal="right", vertical="center")
    celda_lbl_iva.border = borde_delgado
    
    celda_val_iva = ws.cell(row=fila_actual, column=6, value=iva)
    celda_val_iva.font = font_negrita
    celda_val_iva.number_format = '$#,##0'
    celda_val_iva.alignment = Alignment(horizontal="right", vertical="center")
    celda_val_iva.border = borde_delgado
    fila_actual += 1
    
    ws.cell(row=fila_actual, column=1, value="Condiciones de Venta:").font = font_negrita
    
    celda_lbl_tot = ws.cell(row=fila_actual, column=5, value="TOTAL")
    celda_lbl_tot.font = font_negrita
    celda_lbl_tot.alignment = Alignment(horizontal="right", vertical="center")
    celda_lbl_tot.border = borde_delgado
    
    celda_val_tot = ws.cell(row=fila_actual, column=6, value=total_bruto)
    celda_val_tot.font = font_negrita
    celda_val_tot.fill = fill_totales
    celda_val_tot.number_format = '$#,##0'
    celda_val_tot.alignment = Alignment(horizontal="right", vertical="center")
    celda_val_tot.border = borde_delgado
    fila_actual += 1
    
    ws.cell(row=fila_actual, column=1, value="1: Plazo de entrega por confirmar").font = font_normal
    fila_actual += 1
    
    ws.cell(row=fila_actual, column=1, value="2. Validez de cotización: 7 días").font = font_normal
    ws.cell(row=fila_actual, column=7, value="Enrique Hernández P.").font = font_firma
    ws.cell(row=fila_actual, column=7).alignment = Alignment(horizontal="right")
    fila_actual += 1
    
    ws.cell(row=fila_actual, column=1, value="Condiciones de Pago: CONTADO").font = font_negrita
    ws.cell(row=fila_actual, column=7, value="VGM SpA").font = font_negrita
    ws.cell(row=fila_actual, column=7).alignment = Alignment(horizontal="right")
    fila_actual += 1
    
    ws.cell(row=fila_actual, column=1, value="A la espera de sus comentarios, le saluda atentamente :").font = font_normal
    
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 46
    ws.column_dimensions['D'].width = 24  
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 24  
    ws.column_dimensions['G'].width = 24
    
    wb.save(output)
    return output.getvalue()

# AUTODETECCIÓN DE LOGO CORPORATIVO EN EL SERVIDOR
logo_bytes = None
for ext in ["png", "jpg", "jpeg"]:
    if os.path.exists(f"logo.{ext}"):
        with open(f"logo.{ext}", "rb") as f:
            logo_bytes = f.read()
        break

# Barra lateral - Interfaz de Usuario
with st.sidebar:
    st.subheader("💼 Datos de la Cotización")
    nombre_cliente = st.text_input("Nombre del Cliente:", placeholder="Ej: Claudia Araya")
    empresa_cliente = st.text_input("Empresa / Entidad:", placeholder="Ej: Transporte Santa Alberta")
    numero_folio = st.text_input("Número de Cotización:", value="EHP-TSA-2026")
    descuento_aplicar = st.number_input("Descuento a aplicar (%)", min_value=0, max_value=100, value=0, step=1)
    precio_manual_input = st.text_input("Precio Neto Fijo Alternativo (Opcional):", placeholder="Ej: 500000")
    
    st.markdown("---")
    st.subheader("🖼️ Branding & Personalización")
    if logo_bytes:
        st.success("✅ Logo de la empresa cargado desde GitHub.")
        logo_empresa = st.file_uploader("Reemplazar logo actual (Opcional)", type=["png", "jpg", "jpeg"])
    else:
        st.warning("⚠️ Sin logo guardado en el servidor.")
        logo_empresa = st.file_uploader("Sube el Logo de tu Empresa (PNG/JPG)", type=["png", "jpg", "jpeg"])
    
    st.markdown("---")
    st.subheader("📦 Fotos de los Productos")
    fotos_productos = st.file_uploader(
        "Sube las fotos de ESTA cotización (Nómbrar las fotos con su código, ej: YT09511.jpg)", 
        type=["png", "jpg", "jpeg"], 
        accept_multiple_files=True
    )
    
    st.markdown("---")
    # MOTOR DE SEGURIDAD: INTENTA LEER LA LLAVE AUTOMÁTICAMENTE DE STREAMLIT SECRETS
    api_key = None
    if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"].strip():
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("🔑 Motor Gemini autenticado de forma permanente.")
    else:
        with st.expander("⚙️ Autenticación de Motor (Requerido)", expanded=True):
            api_key = st.text_input("Ingresa tu Gemini API Key:", type="password")

st.subheader("1. Sube el pantallazo de la solicitud del cliente")
imagen_pedido = st.file_uploader("Selecciona la imagen del pedido o correo", type=["png", "jpg", "jpeg"])

# MOTOR DE CARGA: LECTURA DEL CATÁLOGO BASE VIGENTE
df_catalogo = leer_csv_tolerante("lista_vigente.csv")

if df_catalogo is not None:
    st.success("✅ Cerebro comercial 'lista_vigente.csv' conectado y estructurado con éxito.")
else:
    st.error("❌ No se encontró o no se pudo mapear 'lista_vigente.csv' en GitHub. Por favor súbelo para activar la app.")
    st.stop()

# MOTOR DE CARGA: LECTURA INTELIGENTE DEL CEREBRO HISTÓRICO DE VENTAS (2025-2026)
df_historial = leer_csv_tolerante("historial_ventas.csv")

if df_historial is not None:
    st.info("🧠 Cerebro Histórico (ERP Ventas) conectado. El sistema auditará precios anteriores automáticamente.")
else:
    st.warning("⚠️ Historial 'historial_ventas.csv' no detectado en el servidor. Operando en modo catálogo estándar.")

# Procesamiento Inteligente
if df_catalogo is not None and imagen_pedido and api_key:
    if st.button("🔥 Generar Cotización Excel Inteligente"):
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.5-flash')
            
            st.info("🔄 Fase 1: Leyendo imagen de solicitud e interpretando requerimientos...")
            df = df_catalogo.copy()
            df.columns = [str(c).strip() for c in df.columns]
            
            columnas_disponibles = list(df.columns)
            n_cols = len(columnas_disponibles)
            
            idx_cod = next((i for i, c in enumerate(columnas_disponibles) if 'cod' in c.lower() or 'id' in c.lower()), 0)
            idx_desc = next((i for i, c in enumerate(columnas_disponibles) if 'desc' in c.lower() or 'nom' in c.lower() or 'art' in c.lower() or 'prod' in c.lower() or 'det' in c.lower()), 0 if n_cols == 1 else 1)
            idx_precio = next((i for i, c in enumerate(columnas_disponibles) if 'prec' in c.lower() or 'val' in c.lower() or 'neto' in c.lower() or 'unit' in c.lower()), n_cols - 1)
            idx_marca = next((i for i, c in enumerate(columnas_disponibles) if 'mar' in c.lower() or 'bra' in c.lower() or 'fab' in c.lower()), -1)
            
            col_codigo = columnas_disponibles[idx_cod]
            col_desc = columnas_disponibles[idx_desc]
            col_precio = columnas_disponibles[idx_precio]
            col_marca = columnas_disponibles[idx_marca] if idx_marca != -1 else None
            
            df['__desc_clean'] = df[col_desc].apply(normalizar_texto)
            df['__cod_clean'] = df[col_codigo].apply(normalizar_texto)
            
            imagen_lista = Image.open(imagen_pedido)
            
            prompt_extraccion = """
            Analiza detalladamente esta imagen de pedido o correo de solicitud. Extrae cada producto solicitado y su cantidad.
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
            
            st.info("🔄 Fase 2: Cruzando datos semánticos con la lista de precios...")
            cantidades_dict = {item.get("busqueda", ""): int(item.get("cantidad", 1)) if item.get("cantidad") is not None else 1 for item in lista_productos}

            candidates_rag = {}
            for item in lista_productos:
                termino = item.get("busqueda", "")
                if not termino: continue
                
                palabras_clave = set()
                for t in [termino] + item.get("sinonimos", []):
                    for p in limpiar_plurales(normalizar_texto(t)).split():
                        if len(p) > 2 and p not in STOP_WORDS: palabras_clave.add(p)
                
                if not palabras_clave: palabras_clave = {normalizar_texto(termino)}
                
                def score_prefiltrado(row):
                    txt_inv = " " + str(row['__desc_clean']) + " " + str(row['__cod_clean']) + " "
                    return sum(5 + (3 if f" {p} " in txt_inv else 0) for p in palabras_clave if p in txt_inv)
                
                df['__tmp_score'] = df.apply(score_prefiltrado, axis=1)
                df_filtrado = df[df['__tmp_score'] > 0].sort_values(by='__tmp_score', ascending=False).head(40)
                
                cand_list = []
                for _, r in df_filtrado.iterrows():
                    c_dict = {
                        "codigo": str(r[col_codigo]), 
                        "descripcion": str(r[col_desc]), 
                        "precio": float(limpiar_precio(r[col_precio]))
                    }
                    if col_marca and col_marca in df.columns:
                        c_dict["marca"] = str(r[col_marca])
                    cand_list.append(c_dict)
                    
                candidates_rag[termino] = cand_list

            st.info("🔄 Fase 3: Resolviendo códigos y auditando historial comercial...")
            prompt_resolucion = f"""
            Actúas como un experto en repuestos y herramientas industriales para la empresa VGM SpA. 
            Tu objetivo es emparejar los requerimientos del cliente con la mejor opción de nuestro catálogo Excel.
            
            ⚠️ REGLAS INQUEBRANTABLES DE ASIGNACIÓN:
            1. PISTOLAS NEUMÁTICAS: Código estándar es 'YT09511'. NO elijas pistolas para inflar neumáticos (YT2370).
            2. LINTERNAS LARGAS / IMANTADAS: Código predilecto es 'YT08518'.
            3. LINTERNAS/LÁMPARAS DE CABEZA (FRONTALES): El código exacto asignado es 'L-HEAD-1'. Queda prohibido elegir la imantada YT08518.
            4. LLAVES DE IMPACTO INALÁMBRICAS: Priorizar 1ra Opción: 'YT8277935'. 2da Opción: 'YT8277925'.
            5. FILTRO ESTRICTO NO ENCONTRADO: Si no calza, marca 'codigo_elegido': 'MANUAL', 'descripcion_elegida': '❌ NO ENCONTRADO', 'precio_elegido': 0.0.
            
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
                cant_val = cantidades_dict.get(origen, 1)
                cant = int(cant_val) if cant_val is not None else 1
                
                cod = str(res.get("codigo_elegido", "MANUAL")).strip()
                desc = str(res.get("descripcion_elegida", "❌ NO ENCONTRADO"))
                px_lista = float(res.get("precio_elegido", 0.0))
                marca = "YATO/VOREL"
                
                if cod != "MANUAL":
                    cod_norm = normalizar_texto(cod)
                    match_rows = df[df['__cod_clean'] == cod_norm]
                    if match_rows.empty:
                        match_rows = df[df[col_codigo].astype(str).str.strip().str.lower() == cod.lower()]
                        
                    if not match_rows.empty:
                        r_match = match_rows.iloc[0]
                        desc = str(r_match[col_desc])
                        px_lista = float(limpiar_precio(r_match[col_precio]))
                        if col_marca and col_marca in df.columns and not pd.isna(r_match[col_marca]):
                            marca = str(r_match[col_marca]).strip().upper()
                
                if cod.upper() == "L-HEAD-1" or "HEAD" in desc.upper() or "CABEZA" in desc.upper():
                    marca = "IRIMO"
                
                # AUDITORÍA MEDIANTE EL CEREBRO HISTÓRICO DE VENTAS
                if df_historial is not None and cod != "MANUAL" and empresa_cliente:
                    try:
                        columnas_h = list(df_historial.columns)
                        col_h_cli = next((c for c in columnas_h if 'cli' in c.lower() or 'emp' in c.lower() or 'raz' in c.lower() or 'nom' in c.lower()), columnas_h[0])
                        col_h_cod = next((c for c in columnas_h if 'cod' in c.lower() or 'art' in c.lower() or 'pro' in c.lower()), columnas_h[1] if len(columnas_h) > 1 else columnas_h[0])
                        col_h_px = next((c for c in columnas_h if 'prec' in c.lower() or 'net' in c.lower() or 'val' in c.lower() or 'vta' in c.lower() or 'tot' in c.lower()), columnas_h[-1])
                        
                        term_emp = normalizar_texto(empresa_cliente)
                        df_h_filtrado = df_historial[df_historial[col_h_cli].astype(str).apply(normalizar_texto).str.contains(term_emp, na=False, regex=False)]
                        match_hist_prod = df_h_filtrado[df_h_filtrado[col_h_cod].astype(str).str.strip().str.lower() == cod.lower()]
                        
                        if not match_hist_prod.empty:
                            ultimo_precio_cobrado = float(limpiar_precio(match_hist_prod.iloc[-1][col_h_px]))
                            desc = f"✨ [Historial ERP: Cobrado ${ultimo_precio_cobrado:,.0f} anteriormente] | " + desc
                    except:
                        pass
                
                if cod == "MANUAL" or "❌" in desc:
                    desc = f"❌ NO ENCONTRADO: (Falta en catálogo o requiere código manual para '{origen}')"
                    px_lista = 0.0
                    marca = "MANUAL"
                elif not res.get("coincidencia_exacta", True):
                    desc = f"⚠️ (Match sugerido) {desc}"
                
                precio_manual_val = limpiar_precio(precio_manual_input)
                if descuento_aplicar == 0 and precio_manual_val > 0 and cod != "MANUAL":
                    px_final_neto = precio_manual_val
                    texto_descuento_col = "Precio Especial"
                else:
                    px_final_neto = px_lista - (px_lista * (descuento_aplicar / 100))
                    texto_descuento_col = f"{descuento_aplicar}%"
                    
                cotizacion_final.append({
                    "Código": cod, "Marca": marca, "Descripción Catálogo": desc, "Cantidad": cant,
                    "Precio Lista (Neto)": px_lista, "Descuento Aplicado": texto_descuento_col,
                    "Precio Final (Neto)": px_final_neto, "Total Neto": px_final_neto * cant
                })
                
            if cotizacion_final:
                df_resultado = pd.DataFrame(cotizacion_final)
                st.success("¡Cotización construida exitosamente!")
                st.dataframe(df_resultado, use_container_width=True)
                
                subtotal_lista = (df_resultado["Precio Lista (Neto)"] * df_resultado["Cantidad"]).sum()
                total_neto_final = df_resultado["Total Neto"].sum()
                descuento_total_pesos = max(subtotal_lista - total_neto_final, 0.0)
                iva_calculado = total_neto_final * 0.19
                total_bruto = total_neto_final + iva_calculado
                
                st.markdown("### 📊 Desglose Económico de la Operación")
                c_neto, c_desc, c_neto_f, c_iva, c_bruto = st.columns(5)
                c_neto.metric(label="Total Lista Neto", value=f"${subtotal_lista:,.0f}")
                c_desc.metric(label="Descuento Otorgado", value=f"-${descuento_total_pesos:,.0f}")
                c_neto_f.metric(label="Neto Final Cliente", value=f"${total_neto_final:,.0f}")
                c_iva.metric(label="IVA (19%)", value=f"${iva_calculado:,.0f}")
                c_bruto.metric(label="Total Bruto", value=f"${total_bruto:,.0f}")
                
                st.markdown("---")
                
                dict_imagenes_procesadas = {}
                if fotos_productos:
                    for foto in fotos_productos:
                        nombre_id = os.path.splitext(foto.name)[0].strip().lower()
                        dict_imagenes_procesadas[nombre_id] = foto.getvalue()
                
                # Priorizar logo subido manualmente, si no, usar el logo automático de GitHub
                logo_data = logo_empresa.getvalue() if logo_empresa else logo_bytes
                
                excel_binario = generar_excel_comercial(
                    df_resultado, nombre_cliente, empresa_cliente, numero_folio, 
                    total_neto_final, iva_calculado, total_bruto, logo_data, dict_imagenes_procesadas
                )
                
                st.download_button(
                    label="🟢 Descargar Documento Excel Premium (.xlsx)", 
                    data=excel_binario, 
                    file_name=f"Cotizacion_{numero_folio}.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.warning("No se logró procesar ningún ítem.")
        except Exception as e:
            st.error(f"Error en el motor de automatización: {e}")
else:
    st.info("Por favor carga el pantallazo del pedido para operar.")
