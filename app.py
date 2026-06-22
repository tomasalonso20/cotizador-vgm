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

# FUNCIÓN: Generación de Excel con Estilo de Catálogo Corporativo
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

# FUNCIÓN: Generación de PDF Oficial Corporativo Protegido contra Errores
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
    
    # Estilos específicos para evitar Spacers dentro de celdas
    style_cond_bold = ParagraphStyle('CondBold', parent=style_rut, fontName='Helvetica-Bold', spaceAfter=4)
    style_cond_normal = ParagraphStyle('CondNormal', parent=style_rut, spaceAfter=4)
    
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
        
    t_productos = Table(tabla_data, colWidths=
