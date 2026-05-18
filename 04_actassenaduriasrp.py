import pandas as pd
import qrcode
import fitz  # PyMuPDF
import os
from datetime import datetime
from PIL import Image  # Para ajustar DPI de imágenes JPG
from rellenarnum import process_excel_file

# Rutas de archivos y directorios
SENADURIAS_LAYOUT_PATH = os.path.join('Archivos', 'Datos', 'Senadurias', 'senaduriasrp_layout.xlsx')
SENADURIAS_VOTOS_PATH = os.path.join('Archivos', 'Datos', 'Senadurias', 'senaduriasrp_layout_votos.xlsx')
SENADURIAS_PDF_FORMAT_PATH = os.path.join('Archivos', 'Formatos', '3ERP.pdf')
FONT_PATH = os.path.join('Archivos', 'Fonts', 'ARIAL.TTF')
SELLO_PATH = os.path.join('Archivos', 'Imagenes', 'sello.png')
SENADURIAS_PDF_DIR = os.path.join('Actas', 'Senadurias_rp','PDF')
SENADURIAS_JPG_DIR = os.path.join('Actas', 'Senadurias_rp','JPG')

# Generar el archivo de votos a partir del layout
process_excel_file(SENADURIAS_LAYOUT_PATH, SENADURIAS_VOTOS_PATH)

# Leer el archivo Excel
df = pd.read_excel(SENADURIAS_VOTOS_PATH, dtype={
    'BOLETAS_SOBRANTES': str,
    'TOTAL_VOTOS_ASENTADOS': str,
    'TOTAL_PERSONAS_VOTARON': str,
    'TOTAL_REP_PARTIDO_CI_VOTARON': str,
    'P1': str,'P2': str,'P3': str,'P4': str,'P5': str,'P6': str,'P7': str,'P8': str,'P9': str,'P10': str,'P11': str,'P12': str,'P13': str,'P14': str,'P15': str,'NO_REGISTRADAS': str,'NULOS': str,'TOTAL_VOTOS': str,
    # Otras columnas que requieran ceros a la izquierda
}).fillna('')  # Reemplaza NaN con cadenas vacías

# Ruta del PDF base (formato)
pdf_base_path = SENADURIAS_PDF_FORMAT_PATH

# Crear directorio para guardar los PDFs generados
os.makedirs(SENADURIAS_PDF_DIR, exist_ok=True)
os.makedirs(SENADURIAS_JPG_DIR, exist_ok=True)  # Para guardar las imágenes JPG

# Definir las áreas específicas para cada columna (x, y)
column_areas = {
    "NOMBRE_ESTADO": (70, 250), 
    "ID_DISTRITO": (365, 250),
    "SECCION": (75, 315),
    "DATOS_QR": (265, 18),  # Añadir posición para el QR
    "CASILLA": (365, 55),
    "LETRA1": (50, 461),
    "BOLETAS_SOBRANTES": (328, 461),
    "LETRA2": (50, 548),
    "TOTAL_PERSONAS_VOTARON": (328, 548),
    "LETRA5": (435, 408),
    "TOTAL_VOTOS_SACADOS": (730, 408),
    "LETRA6": (515, 105),
    "P1": (747, 105),
    "LETRA7": (515, 130),
    "P2": (747, 130),
    "LETRA8": (515, 155),
    "P3": (747, 155),
    "LETRA9": (515, 180),
    "P4": (747, 180),
    "LETRA10": (515, 205),
    "P5": (747, 205),
    "LETRA11": (515, 230),
    "P6": (747, 230),
    "LETRA12": (515, 255),
    "P7": (747, 255),
    "LETRA21": (515, 280),
    "NO_REGISTRADAS": (747, 280),
    "LETRA22": (515, 303),
    "NULOS": (747, 303),
    "LETRA23": (515, 326),
    "TOTAL_VOTOS": (730, 326),
    "SELLO": (900, 250),  # Área para insertar el sello
    "ID_ACTA": (915, 94),
    "USUARIO": (915, 98)
}

# Lista de columnas para agregar espacios
columnas_con_espacios = ["SECCION", "BOLETAS_SOBRANTES", "TOTAL_PERSONAS_VOTARON", "TOTAL_REP_PARTIDO_CI_VOTARON", "TOTAL_VOTOS_SACADOS", "TOTAL_VOTOS_ASENTADOS", "P1", "P2", "P3", "P4", "P5", "P6", "P7",  "NO_REGISTRADAS", "NULOS", "TOTAL_VOTOS"]

columnas_con_cuatro = ["SECCION", "BOLETAS_SOBRANTES", "TOTAL_PERSONAS_VOTARON", "TOTAL_REP_PARTIDO_CI_VOTARON", "TOTAL_VOTOS_SACADOS", "TOTAL_VOTOS_ASENTADOS", "TOTAL_VOTOS"]

# Asegurarse de que todas las columnas en columnas_con_espacios tengan ceros a la izquierda
for columna in columnas_con_espacios:
    df[columna] = df[columna].fillna('').astype(str).apply(lambda x: x.zfill(3))

# Asegurarse de que todas las columnas en columnas_con_espacios tengan ceros a la izquierda
for columna in columnas_con_cuatro:
    df[columna] = df[columna].fillna('').astype(str).apply(lambda x: x.zfill(4))


# Columnas para excluir de la impresión de texto
columnas_excluidas = ["DATOS_QR"]

# Función para agregar espacios entre caracteres
def agregar_espacios(texto):
    return '    '.join(list(texto))

# Obtener la fecha actual para usar en los nombres de archivos
fecha_actual = datetime.now().strftime("%Y%m%d_%H%M%S")  # Formato: YYYYMMDD_HHMMSS

# Para cada registro en el Excel, crea un PDF y JPG
for index, row in df.iterrows():
    try:
        pdf_document = fitz.open(pdf_base_path)
        pagina = pdf_document[0]
        pagina.insert_font(fontfile=FONT_PATH, fontname="F0")

        for col_name, value in row.items():
            if col_name in column_areas and col_name not in columnas_excluidas:
                x, y = column_areas[col_name]  # Solo (x, y)

                if col_name in columnas_con_espacios:
                    texto = agregar_espacios(str(value))
                else:
                    texto = str(value)

                pagina.insert_text(
                    (x, y),
                    texto,
                    fontsize=12,
                    fontname="F0",
                    rotate=0,
                    color=(0, 0, 0),  # Color negro
                )
                
            # Si la columna es "DATOS_QR", genera e inserta el código QR en la posición correcta
            if col_name == "DATOS_QR":
                x_qr, y_qr = column_areas[col_name]  # Posición del QR
                qr = qrcode.make(str(value))
                qr_path = f"qr_{index + 1}.png"
                qr.save(qr_path)
                
                # Ajustar la posición de la imagen QR
                rect = fitz.Rect(x_qr, y_qr, x_qr + 50, y_qr + 50)
                pagina.insert_image(rect, filename=qr_path)
        
        # Inserción del sello como imagen
        if os.path.exists(SELLO_PATH):
            x_sello, y_sello = column_areas["SELLO"]
            rect_sello = fitz.Rect(x_sello, y_sello, x_sello + 230, y_sello + 120)  # Ajustar tamaño según el sello
            pagina.insert_image(rect_sello, filename=SELLO_PATH)
        
        pdf_output_path = os.path.join(SENADURIAS_PDF_DIR, f"3ERP_{fecha_actual}_{index + 1}.pdf")
        pdf_document.save(pdf_output_path)
        print(f"Archivo generado: 3ERP_{fecha_actual}_{index + 1}.pdf")
        
        # Ajustar resolución y tamaño del pixmap
        width_px, height_px = 5103, 2550
        scale = (width_px / pagina.rect.width, height_px / pagina.rect.height)

        # Obtener el pixmap con la resolución y dimensiones correctas
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(*scale))
        
        # Guardar como JPG y ajustar DPI a 300
        jpg_path = os.path.join(SENADURIAS_JPG_DIR, f"0101-3ERP-{fecha_actual}_{index + 1}.jpg")
        pixmap.save(jpg_path, "jpg")  # Guardar como JPG
        print(f"Archivo generado: 0101-3ERP-{fecha_actual}_{index + 1}.jpg")

        # Usar PIL para cambiar DPI del JPG
        image = Image.open(jpg_path)
        image.save(jpg_path, dpi=(300, 300))  # Establecer DPI
        
        # Eliminar el archivo temporal del código QR
        if os.path.exists(qr_path):
            os.remove(qr_path)  # Eliminar el archivo temporal
    except Exception as e:
        print(f"Error al procesar el registro {index + 1}: {e}")

print("PDFs y JPGs generados y guardados en las carpetas  " + SENADURIAS_PDF_DIR + " y " + SENADURIAS_JPG_DIR + ".")
