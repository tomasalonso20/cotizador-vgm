import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
from PIL import Image
import unicodedata
import io

# Importaciones para estilización avanzada de Excel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Importaciones para la generación del PDF Corporativo
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

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

# FUNCIÓN: Generación de Excel con Estilo de Catálogo
def generar_excel_comercial(df_cotiz, cliente, empresa, nro_cotiz, total_neto, iva, total_bruto):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cotización"
    
    ws.views.sheetView[0].showGridLines = True
    
    font_titulo_central = Font(name="Arial", size=14, bold=True, underline="single")
    font_cabecera_tabla = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_negrita = Font(name="Arial", size=10, bold=True)
    font_normal = Font(name="Arial", size=10)
    font_firma = Font(name="Arial", size=11, bold=True, italic=True)
    
    fill_azul_vgm = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    borde_delgado = Border(
        left=Side(style='thin', color='A0A0A0'), right=Side(style='thin', color='A0A0A0'),
        top=Side(style='thin', color='A0A0A0'), bottom=Side(style='thin', color='A0A0A0')
    )
    
    ws["A1"] = "VGM SpA"
    ws["A1"].font = Font(name="Arial", size=13, bold=True)
    ws["A2"] = "76.834.968-1"
    ws["A2"].font = font_negrita
    ws["A3"] = "Chopin 2848. San Joaquín. Santiago"
    ws["A3"].font = font_normal
    
    ws["G1"] = f"Fecha: {pd.Timestamp.now().strftime('%d-%m-%Y')}"
    ws["G1"].font = font_negrita
    ws["G1"].alignment = Alignment(horizontal="right")
    
    ws["C5"] = f"COTIZACIÓN N° {nro_cotiz}"
    ws["C5"].font = font_titulo_central
    ws["C5"].alignment = Alignment(horizontal="center", vertical="center")
    
    ws["A7"] = f"Sr(a).: {cliente if cliente else 'No especificado'}"
    ws["A7"].font = font_negrita
    ws["A8"] = f"Empresa: {empresa if empresa else 'No especificada'}"
    ws["A8"].font = font_negrita
    ws["A9"] = "En atención a su gentil solicitud de cotización, tenemos el agrado de hacer llegar a usted nuestra propuesta:"
    ws["A9"].font = font_normal
    
    titulos_columnas = ["CÓDIGO", "MARCA", "DESCRIPCIÓN", "PRECIO NETO", "CANTIDAD", "TOTAL NETO", "IMAGEN REFERENCIAL"]
    fila_tabla_inicio = 12
    
    # CORREGIDO: Cambiado 'en' por 'in' para cumplir la sintaxis de Python
    for col_idx, texto_col in enumerate(titulos_columnas, 1):
        celda = ws.cell(row=fila_tabla_inicio, column=col_idx, value=texto_col)
        celda.font = font_cabecera_tabla
        celda.fill = fill_azul_vgm
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = borde_delgado
        
    fila_actual = fila_tabla_inicio + 1
    for _, fila in df_cotiz.iterrows():
        ws.cell(row=fila_actual, column=1, value=str(fila["Código"])).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=fila_actual, column=2, value="YATO/VOREL").alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=fila_actual, column=3, value=str(fila["Descripción Catálogo"])).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        c_neto = ws.cell(row=fila_actual, column=4, value=float(fila["Precio Final (Neto)"]))
        c_neto.number_format = '$#,##0'
        c_neto.alignment = Alignment(horizontal="right", vertical="center")
        
        ws.cell(row=fila_actual, column=5, value=int(fila["Cantidad"])).alignment = Alignment(horizontal="center", vertical="center")
        
        c_total = ws.cell(row=fila_actual, column=6, value=float(fila["Total Neto"]))
        c_total.number_format = '$#,##0'
        c_total.alignment = Alignment(horizontal="right", vertical="center")
        
        ws.cell(row=fila_actual, column=7, value="[Espacio Imagen]").alignment = Alignment(horizontal="center", vertical="center")
        
        for c_idx in range(1, 8):
            celda_f = ws.cell(row=fila_actual, column=c_idx)
            celda_f.font = font_normal
            celda_f.border = borde_delgado
            
        ws.row_dimensions[fila_actual].height = 45  
        fila_actual += 1
        
    totales_comerciales = [("SUBTOTAL", total_neto), ("IVA (19%)", iva), ("TOTAL BRUTO", total_bruto)]
    for etiqueta, valor in totales_comerciales:
        celda_lbl = ws.cell(row=fila_actual, column=5, value=etiqueta)
        celda_lbl.font = font_negrita
        celda_lbl.alignment = Alignment(horizontal="right", vertical="center")
        celda_lbl.border = borde_delgado
        
        celda_val = ws.cell(row=fila_actual, column=6, value=valor)
        celda_val.font = font_negrita
        celda_val.number_format = '$#,##0'
        celda_val.alignment = Alignment(horizontal="right", vertical="center")
        celda_val.border = borde_delgado
        fila_actual += 1
        
    fila_actual += 1
    ws.cell(row=fila_actual, column=1, value="Observaciones:").font = font_negrita
    ws.cell(row=fila_actual+1, column=1, value="Condiciones de Venta:").font = font_negrita
    ws.cell(row=fila_actual+2, column=1, value="1: Plazo de entrega por confirmar").font = font_normal
    ws.cell(row=fila_actual+3, column=1, value="2: Validez de cotización: 7 días").font = font_normal
    ws.cell(row=fila_actual+4, column=1, value="Condiciones de Pago: CONTADO").font = font_negrita
    
    ws.cell(row=fila_actual+3, column=6, value="Enrique Hernández P.").font = font_firma
    ws.cell(row=fila_actual+4, column=6, value="VGM SpA").font = font_negrita
    
    ws.column_dimensions['A'].width = 13
    ws.column_dimensions['B'].width = 13
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 11
    ws.column_dimensions['F'].width = 16
    ws.column_dimensions['G'].width = 20
    
    wb.save(output)
    return output.getvalue()

# FUNCIÓN: Generación de PDF Oficial de Alta Calidad
def generar_pdf_comercial(df_cotiz, cliente, empresa, nro_cotiz, total_neto, iva, total_bruto):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    
    styles = getSampleStyleSheet()
    style_normal = styles['Normal']
    
    style_empresa = ParagraphStyle('EmpresaStyle', parent=style_normal, fontName='Helvetica-Bold', fontSize=12, leading=14)
    style_rut = ParagraphStyle('RutStyle', parent=style_normal, fontName='Helvetica', fontSize=10, leading=12)
    style_derecha = ParagraphStyle('DerechaStyle', parent=style_normal, fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=2)
    style_titulo = ParagraphStyle('TituloStyle', parent=style_normal, fontName='Helvetica-Bold', fontSize=14, leading=16, alignment=1, underline=True)
    style_tabla_header = ParagraphStyle('TableHeader', parent=style_normal, fontName='Helvetica-Bold', fontSize=9, leading=11, textColor=colors.white, alignment=1)
    style_tabla_celda = ParagraphStyle('TableCell', parent=style_normal, fontName='Helvetica', fontSize=8, leading=10)
    style_tabla_celda_c = ParagraphStyle('TableCellC', parent=style_normal, fontName='Helvetica', fontSize=8, leading=10, alignment=1)
    style_tabla_celda_r = ParagraphStyle('TableCellR', parent=style_normal, fontName='Helvetica', fontSize=8, leading=10, alignment=2)
    
    p_empresa = Paragraph("VGM SpA", style_empresa)
    p_rut = Paragraph("RUT: 76.834.968-1<br/>Chopin 2848. San Joaquín. Santiago", style_rut)
    p_fecha = Paragraph(f"Fecha: {pd.Timestamp.now().strftime('%d-%m-%Y')}", style_derecha)
    
    header_table = Table([[p_empresa, p_fecha], [p_rut, Paragraph("", style_normal)]], colWidths=[270, 270])
    header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(header_table)
    story.append(Spacer(1, 15))
    
    story.append(Paragraph(f"COTIZACIÓN N° {nro_cotiz}", style_titulo))
    story.append(Spacer(1, 15))
    
    story.append(Paragraph(f"<b>Sr(a).:</b> {cliente if cliente else 'No especificado'}", style_rut))
    story.append(Paragraph(f"<b>Empresa:</b> {empresa if empresa else 'No especificada'}", style_rut))
    story.append(Spacer(1, 8))
    story.append(Paragraph("En atención a su gentil solicitud de cotización, tenemos el agrado de hacer llegar a usted nuestra propuesta:", style_rut))
    story.append(Spacer(1, 12))
    
    headers = [
        Paragraph("CÓDIGO", style_tabla_header), Paragraph("MARCA", style_tabla_header),
        Paragraph("DESCRIPCIÓN", style_tabla_header), Paragraph("P. NETO", style_tabla_header),
        Paragraph("CANT", style_tabla_header), Paragraph("TOTAL", style_tabla_header),
        Paragraph("IMAGEN", style_tabla_header)
    ]
    
    tabla_data = [headers]
    for _, fila in df_cotiz.iterrows():
        tabla_data.append([
            Paragraph(str(fila["Código"]), style_tabla_celda_c),
            Paragraph("YATO/VOREL", style_tabla_celda_c),
            Paragraph(str(fila["Descripción Catálogo"]), style_tabla_celda),
            Paragraph(f"${fila['Precio Final (Neto)']:,.0f}", style_tabla_celda_r),
            Paragraph(str(int(fila["Cantidad"])), style_tabla_celda_c),
            Paragraph(f"${fila['Total Neto']:,.0f}", style_tabla_celda_r),
            Paragraph("[Cuadro Ref]", style_tabla_celda_c)
        ])
        
    t_productos = Table(tabla_data, colWidths=[55, 55, 175, 55, 35, 60, 105])
    t_productos.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4F81BD")),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_productos)
    story.append(Spacer(1, 10))
    
    p_obs_titulo = Paragraph("<b>Observaciones:</b>", style_rut)
    p_cond1 = Paragraph("<b>Condiciones de Venta:</b>", style_rut)
    p_cond2 = Paragraph("1: Plazo de entrega por confirmar<br/>2: Validez de cotización: 7 días", style_rut)
    p_pago = Paragraph("<b>Condiciones de Pago: CONTADO</b>", style_rut)
    
    bloque_izq = [p_obs_titulo, Spacer(1,4), p_cond1, p_cond2, Spacer(1,4), p_pago]
    
    t_totales_data = [
        ["", Paragraph("SUBTOTAL:", style_derecha), Paragraph(f"${total_neto:,.0f}", style_derecha)],
        ["", Paragraph("IVA (19%):", style_derecha), Paragraph(f"${iva:,.0f}", style_derecha)],
        ["", Paragraph("TOTAL BRUTO:", style_derecha), Paragraph(f"${total_bruto:,.0f}", style_derecha)]
    ]
    t_totales = Table(t_totales_data, colWidths=[100, 100, 100])
    t_totales.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (1,0), (-1,-1), 0.5, colors.lightgrey),
        ('BACKGROUND', (1,2), (2,2), colors.HexColor("#E2EFDA"))
    ]))
    
    footer_table = Table([[bloque_izq, t_totales]], colWidths=[240, 300])
    footer_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(footer_table)
    story.append(Spacer(1, 30))
    
    p_firma = Paragraph("<i>Enrique Hernández P.</i><br/><b>VGM SpA</b>", style_derecha)
    firma_table = Table([["", p_firma]], colWidths=[340, 200])
    story.append(firma_table)
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# Barra lateral de configuración y datos comerciales
with st.sidebar:
    st.subheader("💼 Datos de la Cotización")
    nombre_cliente = st.text_input("Nombre del Cliente:", placeholder="Ej: Claudia Araya")
    empresa_cliente = st.text_input("Empresa / Entidad:", placeholder="Ej: Tattersall")
    numero_folio = st.text_input("Número de Cotización:", value="EHP-SKC-5299")
    descuento_aplicar = st.number_input("Descuento a aplicar (%)", min_value=0, max_value=100, value=0, step=1)
    precio_manual_input = st.text_input("Precio Neto Fijo Alternativo (Opcional):", placeholder="Ej: 500000")
    
    st.markdown("---")
    with st.expander("⚙️ Configuración del Sistema (API / Columnas)", expanded=False):
        api_key = st.text_input("Ingresa tu Gemini API Key:", type="password")
        st.subheader("📊 Verificación de Columnas")
        col_codigo = "CODIGO"
        col_desc = "DESCRIPCION"
        col_precio = "PRECIO UNITARIO NETO"

# Columnas de carga de archivos de la interfaz
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Sube tu lista de precios (Excel .xlsx)")
    archivo_excel = st.file_uploader("Selecciona tu archivo Excel optimizado", type=["xlsx"])
with col2:
    st.subheader("2. Sube el pantallazo del pedido (WhatsApp / Correo)")
    imagen_pedido = st.file_uploader("Selecciona la imagen del pedido", type=["png", "jpg", "jpeg"])

if archivo_excel:
    try:
        df_preview = pd.read_excel(archivo_excel, header=1, nrows=5)
        columnas_disponibles = [str(c).strip() for c in df_preview.columns]
        idx_cod = next((i for i, c in enumerate(columnas_disponibles) if 'cod' in c.lower() or 'id' in c.lower()), 0)
        idx_desc = next((i for i, c in enumerate(columnas_disponibles) if 'desc' in c.lower() or 'nom' in c.lower() or 'art' in c.lower() or 'prod' in c.lower() or 'det' in c.lower()), 1)
        idx_precio = next((i for i, c in enumerate(columnas_disponibles) if 'prec' in c.lower() or 'val' in c.lower() or 'neto' in c.lower() or 'unit' in c.lower()), len(columnas_disponibles) - 1)
        col_codigo, col_desc, col_precio = columnas_disponibles[idx_cod], columnas_disponibles[idx_desc], columnas_disponibles[idx_precio]
    except Exception as e:
        st.sidebar.error(f"Error al analizar las columnas del Excel: {e}")

# Ejecución del motor híbrido RAG
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
                
                candidates_rag[termino] = [{"codigo": str(r[col_codigo]), "descripcion": str(r[col_desc]), "precio": float(limpiar_precio(r[col_precio]))} for _, r in df_filtrado.iterrows()]

            st.info("🔄 Fase 3: Resolviendo ambigüedades con homologación experta de catálogo...")
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
                cant = int(cantidades_dict.get(origen, 1))
                px_lista = float(res.get("precio_elegido", 0.0))
                desc = res.get("descripcion_elegida", "❌ NO ENCONTRADO")
                cod = res.get("codigo_elegido", "MANUAL")
                
                if cod == "MANUAL" or "❌" in desc:
                    desc = f"❌ NO ENCONTRADO: (Falta en catálogo o requiere código manual para '{origen}')"
                    px_lista = 0.0
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
                    "Código": cod, "Descripción Catálogo": desc, "Cantidad": cant,
                    "Precio List (Neto)": px_lista, "Descuento Aplicado": texto_descuento_col,
                    "Precio Final (Neto)": px_final_neto, "Total Neto": px_final_neto * cant
                })
                
            if cotizacion_final:
                df_resultado = pd.DataFrame(cotizacion_final)
                st.success("¡Cotización inteligente procesada con éxito!")
                st.dataframe(df_resultado, use_container_width=True)
                
                subtotal_lista = (df_resultado["Precio List (Neto)"] * df_resultado["Cantidad"]).sum()
                total_neto_final = df_resultado["Total Neto"].sum()
                descuento_total_pesos = max(subtotal_lista - total_neto_final, 0.0)
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
                st.subheader("📥 Descargar Documentos Comerciales Oficiales")
                
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    excel_binario = generar_excel_comercial(df_resultado, nombre_cliente, empresa_cliente, numero_folio, total_neto_final, iva_calculado, total_bruto)
                    st.download_button(label="🟢 Descargar Cotización en Excel", data=excel_binario, file_name=f"Cotizacion_{numero_folio}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with col_d2:
                    pdf_binario = generar_pdf_comercial(df_resultado, nombre_cliente, empresa_cliente, numero_folio, total_neto_final, iva_calculado, total_bruto)
                    st.download_button(label="🔵 Descargar Cotización en PDF", data=pdf_binario, file_name=f"Cotizacion_{numero_folio}.pdf", mime="application/pdf")
            else:
                st.warning("No se pudieron asociar los productos.")
        except Exception as e:
            st.error(f"Error crítico en el motor de IA: {e}")
else:
    st.info("Introduce tu Gemini API Key a la izquierda y sube tus dos archivos para comenzar.")
