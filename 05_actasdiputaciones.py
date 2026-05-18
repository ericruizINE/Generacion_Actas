import pandas as pd
import qrcode
import fitz  # PyMuPDF
import os
from datetime import datetime
from PIL import Image  # Para ajustar DPI de imágenes JPG
from rellenarnum import process_excel_file

# Rutas de archivos y directorios
DIPUTACIONES_LAYOUT_PATH = os.path.join('Archivos', 'Datos', 'Diputaciones', 'diputaciones_layout.xlsx')
DIPUTACIONES_VOTOS_PATH = os.path.join('Archivos', 'Datos', 'Diputaciones', 'diputaciones_layout_votos.xlsx')
DIPUTACIONES_PDF_FORMAT_PATH = os.path.join('Archivos', 'Formatos', '4V1.pdf')
FONT_PATH = os.path.join('Archivos', 'Fonts', 'ARIAL.TTF')
SELLO_PATH = os.path.join('Archivos', 'Imagenes', 'sello.png')
DIPUTACIONES_PDF_DIR = os.path.join('Actas', 'Diputaciones', 'PDF', )
DIPUTACIONES_JPG_DIR = os.path.join('Actas', 'Diputaciones', 'JPG', )

# Generar el archivo de votos a partir del layout
process_excel_file(DIPUTACIONES_LAYOUT_PATH, DIPUTACIONES_VOTOS_PATH)

# Leer el archivo Excel
df = pd.read_excel(DIPUTACIONES_VOTOS_PATH, dtype={
    'BOLETAS_SOBRANTES': str,
    'TOTAL_VOTOS_ASENTADOS': str,
    'TOTAL_PERSONAS_VOTARON': str,
    'TOTAL_REP_PARTIDO_CI_VOTARON': str,
    'P1': str,'P2': str,'P3': str,'P4': str,'P5': str,'P6': str,'P7': str,'P8': str,'P9': str,'P10': str,'P11': str,'P12': str,'P13': str,'P14': str,'P15': str,'NO_REGISTRADAS': str,'NULOS': str,'TOTAL_VOTOS': str,
    # Otras columnas que requieran ceros a la izquierda
}).fillna('')  # Reemplaza NaN con cadenas vacías

# Ruta del PDF base (formato)
pdf_base_path = DIPUTACIONES_PDF_FORMAT_PATH

# Crear directorio para guardar los PDFs generados
os.makedirs(DIPUTACIONES_PDF_DIR, exist_ok=True)
os.makedirs(DIPUTACIONES_JPG_DIR, exist_ok=True)  # Para guardar las imágenes JPG

# Definir las áreas específicas para cada columna (x, y)
column_areas = {
    "NOMBRE_ESTADO": (72, 269), 
    "ID_DISTRITO": (362, 269),
    "SECCION": (76, 325),
    "DATOS_QR": (253, 18),  # Añadir posición para el QR
    "CASILLA": (366, 55),
    "LETRA1": (40, 478),
    "BOLETAS_SOBRANTES": (338, 478),
    "LETRA2": (40, 564),
    "TOTAL_PERSONAS_VOTARON": (338, 564),
    "LETRA3": (40, 656),
    "TOTAL_REP_PARTIDO_CI_VOTARON": (338, 656),
    "LETRA4": (40, 731),
    "TOTAL_VOTOS_ASENTADOS": (338, 731),
    "LETRA5": (420, 628),
    "TOTAL_VOTOS_SACADOS": (739, 628),
    "LETRA6": (489, 120),
    "P1": (737, 120),
    "LETRA7": (489, 145),
    "P2": (737, 145),
    "LETRA8": (489, 170),
    "P3": (737, 172),
    "LETRA9": (489, 197),
    "P4": (737, 197),
    "LETRA10": (489, 220),
    "P5": (737, 220),
    "LETRA11": (489, 245),
    "P6": (737, 245),
    "LETRA12": (489, 270),
    "P7": (737, 270),
    "LETRA13": (489, 295),
    "P8": (737, 295),
    "LETRA14": (489, 320),
    "P9": (737, 320),
    "LETRA15": (489, 345),
    "P10": (737, 345),
    "LETRA16": (489, 375),
    "P11": (737, 375),
    "LETRA17": (489, 400),
    "P12": (737, 400),
    "LETRA18": (489, 425),
    "P13": (737, 425),
    "LETRA19": (489, 450),
    "P14": (737, 450),
    "LETRA20": (489, 477),
    "P15": (737, 477),
    "LETRA21": (489, 505),
    "NO_REGISTRADAS": (737, 505),
    "LETRA22": (489, 530),
    "NULOS": (737, 530),
    "LETRA23": (489, 555),
    "TOTAL_VOTOS": (737, 555),
    "SELLO": (850, 350),  # Área para insertar el sello
    "ID_ACTA": (900, 152),
    "USUARIO": (900, 170)
}

# Lista de columnas para agregar espacios
columnas_con_espacios = ["SECCION", "BOLETAS_SOBRANTES", "TOTAL_PERSONAS_VOTARON", "TOTAL_REP_PARTIDO_CI_VOTARON", "TOTAL_VOTOS_SACADOS", "TOTAL_VOTOS_ASENTADOS", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "P11", "P12", "P13", "P14", "P15", "P13", "NO_REGISTRADAS", "NULOS", "TOTAL_VOTOS"]

columnas_con_cuatro = ["SECCION"]

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
        
        pdf_output_path = os.path.join(DIPUTACIONES_PDF_DIR, f"4V1-{fecha_actual}_{index + 1}.pdf")
        pdf_document.save(pdf_output_path)
        print(f"Archivo generado: 4V1-{fecha_actual}_{index + 1}.pdf")
        
        # Ajustar resolución y tamaño del pixmap
        width_px, height_px = 5103, 3300
        scale = (width_px / pagina.rect.width, height_px / pagina.rect.height)

        # Obtener el pixmap con la resolución y dimensiones correctas
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(*scale))
        
        # Guardar como JPG y ajustar DPI a 300
        jpg_path = os.path.join(DIPUTACIONES_JPG_DIR, f"0101-4V1-{fecha_actual}_{index + 1}.jpg")
        pixmap.save(jpg_path, "jpg")  # Guardar como JPG
        print(f"Archivo generado: 0101-4V1-{fecha_actual}_{index + 1}.jpg")

        # Usar PIL para cambiar DPI del JPG
        image = Image.open(jpg_path)
        image.save(jpg_path, dpi=(300, 300))  # Establecer DPI
        
        # Eliminar el archivo temporal del código QR
        if os.path.exists(qr_path):
            os.remove(qr_path)  # Eliminar el archivo temporal
    except Exception as e:
        print(f"Error al procesar el registro {index + 1}: {e}")

print("PDFs y JPGs generados y guardados en las carpetas 'PDF/diputaciones' y 'JPG/diputaciones.")
