import streamlit as st
import pandas as pd
import google.generativeai as genai
import json
from PIL import Image

# Configuración de la página web
st.set_page_config(page_title="Cotizador Inteligente - VGM SpA", layout="wide")

st.title("⚙️ Cotizador Express - VGM SpA")
st.write("Sube tu lista de precios y el pantallazo del pedido para generar la cotización al instante.")

# BARRA LATERAL: Configuración inicial
st.sidebar.header("🔑 Configuración")
api_key = st.sidebar.text_input("Ingresa tu Gemini API Key:", type="password")
st.sidebar.markdown("[¿Cómo obtener una API Key gratis?](https://aistudio.google.com/)")

# Subir archivos obligatorios
uploaded_excel = st.file_uploader("1. Sube tu lista de precios (Excel .xlsx)", type=["xlsx"])
uploaded_image = st.file_uploader("2. Sube el pantallazo del pedido (WhatsApp / Correo)", type=["png", "jpg", "jpeg"])

if uploaded_excel and uploaded_image and api_key:
    # Configurar la IA de Google
    genai.configure(api_key=api_key)
    
    # Leer el archivo Excel de precios
    df_precios = pd.read_excel(uploaded_excel)
    
    # Convertir los precios a texto para que la IA los pueda leer como base de conocimiento
    contexto_precios = df_precios.to_string(index=False)
    
    st.success("✅ Archivos cargados con éxito. Procesando cotización con IA...")
    
    # Cargar la imagen del pantallazo
    imagen_pedido = Image.open(uploaded_image)
    
    # Crear el modelo de IA con capacidades de visión
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    # Prompt maestro con tus reglas de negocio corporativas
    prompt = f"""
    Eres el asistente de cotizaciones automatizado de la empresa VGM SpA.
    Tu tarea es mirar el pantallazo adjunto (pedido del cliente) y buscar cada producto solicitado en la lista de precios que te proporciono abajo.
    
    REGLAS CRÍTICAS DE TRADUCCIÓN:
    - Si el cliente pide 'cuello de ganso', busca en la lista de precios como 'CORONA ACODADA'.
    - Si un producto tiene varias marcas o calidades disponibles, elige siempre la opción más económica de la lista.
    
    LISTA DE PRECIOS OFICIAL DE LA EMPRESA:
    {contexto_precios}
    
    Debes devolver estrictamente un código JSON (un arreglo de objetos) con el siguiente formato, sin texto adicional alrededor, sin bloques de código ```json:
    [
      {{
        "codigo": "Código exacto del Excel",
        "marca": "Marca del Excel",
        "descripcion": "Descripción del producto",
        "precio_unitario": 1500,
        "cantidad": 2,
        "url_imagen": "URL de la columna URL_IMAGEN si existe, si no déjala vacía"
      }}
    ]
    """
    
    try:
        # Ejecutar la Inteligencia Artificial
        response = model.generate_content([prompt, imagen_pedido])
        texto_limpio = response.text.strip().replace("
```json", "").replace("```", "")
        items_cotizados = json.loads(texto_limpio)
        
        # Cálculos de dinero
        subtotal = 0
        filas_html = ""
        
        for item in items_cotizados:
            total_item = item['precio_unitario'] * item['cantidad']
            subtotal += total_item
            
            # Formatear valores a moneda chilena
            p_unitario_formato = f"$ {item['precio_unitario']:,}".replace(",", ".")
            p_total_formato = f"$ {total_item:,}".replace(",", ".")
            
            # Validar si el item tiene imagen referencial
            img_html = f"<img src='{item['url_imagen']}' width='80'>" if item.get('url_imagen') else "Sin imagen"
            
            # Construir filas de la tabla imitando tu formato azul/amarillo/verde
            filas_html += f"""
            <tr>
                <td style='border: 1px solid #000; padding: 8px; font-weight: bold; background-color: #e2f0d9;'>{item['codigo']}</td>
                <td style='border: 1px solid #000; padding: 8px;'>{item['marca']}</td>
                <td style='border: 1px solid #000; padding: 8px; font-size: 11px;'>{item['descripcion']}</td>
                <td style='border: 1px solid #000; padding: 8px; text-align: right;'>{p_unitario_formato}</td>
                <td style='border: 1px solid #000; padding: 8px; text-align: center;'>{item['cantidad']}</td>
                <td style='border: 1px solid #000; padding: 8px; text-align: right; font-weight: bold;'>{p_total_formato}</td>
                <td style='border: 1px solid #000; padding: 8px; text-align: center;'>{img_html}</td>
            </tr>
            """
            
        iva = int(subtotal * 0.19)
        total_bruto = subtotal + iva
        
        # Formato de totales chilenos
        subtotal_f = f"$ {subtotal:,}".replace(",", ".")
        iva_f = f"$ {iva:,}".replace(",", ".")
        total_f = f"$ {total_bruto:,}".replace(",", ".")
        
        # PLANTILLA DE DISEÑO DE TU DOCUMENTO REAL (Imitando la imagen 37236_2.png)
        html_cotizacion = f"""
        <div style="font-family: Arial, sans-serif; max-width: 900px; margin: auto; padding: 20px; border: 1px solid #ccc; background-color: #fff;">
            <!-- Encabezado Corporativo VGM -->
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
                <tr>
                    <td style="width: 50%;">
                        <div style="font-weight: bold; font-size: 24px; color: #1f4e79;">VGM SpA</div>
                        <div style="font-size: 12px; color: #555;">76.834.960-1<br>Chiloé 2840, San Joaquín, Santiago</div>
                    </td>
                    <td style="width: 50%; text-align: right; vertical-align: top;">
                        <div style="font-weight: bold; font-size: 14px;">COTIZACIÓN AUTOMÁTICA</div>
                        <div style="font-size: 12px; color: #555;">Fecha: Hoy</div>
                    </td>
                </tr>
            </table>
            
            <!-- Cuerpo de la Tabla de Productos -->
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
                <thead>
                    <tr style="background-color: #2f5597; color: white; font-size: 12px;">
                        <th style="border: 1px solid #000; padding: 8px;">CÓDIGO</th>
                        <th style="border: 1px solid #000; padding: 8px;">MARCA</th>
                        <th style="border: 1px solid #000; padding: 8px;">DESCRIPCIÓN</th>
                        <th style="border: 1px solid #000; padding: 8px;">P. UNITARIO NETO</th>
                        <th style="border: 1px solid #000; padding: 8px;">CANTIDAD</th>
                        <th style="border: 1px solid #000; padding: 8px;">VALOR TOTAL NETO</th>
                        <th style="border: 1px solid #000; padding: 8px;">IMAGEN REFERENCIAL</th>
                    </tr>
                </thead>
                <tbody>
                    {filas_html}
                </tbody>
            </table>
            
            <!-- Cuadro de Totales -->
            <table style="width: 40%; margin-left: 60%; border-collapse: collapse; margin-top: 20px; font-size: 13px;">
                <tr>
                    <td style="padding: 5px; font-weight: bold; text-align: right;">SUBTOTAL:</td>
                    <td style="padding: 5px; text-align: right; border-bottom: 1px solid #000;">{subtotal_f}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; font-weight: bold; text-align: right;">IVA (19%):</td>
                    <td style="padding: 5px; text-align: right; border-bottom: 1px solid #000;">{iva_f}</td>
                </tr>
                <tr style="font-size: 15px; font-weight: bold;">
                    <td style="padding: 5px; text-align: right;">TOTAL:</td>
                    <td style="padding: 5px; text-align: right; color: #1f4e79;">{total_f}</td>
                </tr>
            </table>
            
            <!-- Condiciones al pie de página -->
            <div style="margin-top: 30px; font-size: 11px; color: #555; border-top: 1px solid #ccc; padding-top: 10px;">
                <strong>Condiciones de Venta:</strong><br>
                1. Plazo de entrega por confirmar.<br>
                2. Validez de la cotización: 7 días.<br>
                A la espera de sus comentarios, le saluda atentamente:<br><br>
                <strong>Enrique Hernández P.</strong><br>VGM SpA
            </div>
        </div>
        """
        
        # Mostrar el resultado visual impecable en la pantalla web
        st.markdown("### 📄 Vista Previa de tu Cotización")
        st.components.v1.html(html_cotizacion, height=800, scroller=True)
        st.info("💡 Para guardarlo como PDF: Haz clic derecho en la vista previa, selecciona 'Imprimir' y elige 'Guardar como PDF'.")
        
    except Exception as e:
        st.error(f"Hubo un problema al procesar la imagen. Asegúrate de que tu API Key sea correcta. Error: {e}")
else:
    st.info("A la espera de que ingreses tu API Key de Gemini y cargues ambos archivos en la barra lateral.")
