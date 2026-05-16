import pandas as pd
import qrcode
import fitz  # PyMuPDF
import os
import re
from datetime import datetime
from PIL import Image  # Para ajustar DPI de imágenes JPG
from rellenarnum import process_excel_file

# Rutas de archivos y directorios
PRESIDENCIA_LAYOUT_PATH = os.path.join('Archivos', 'Datos', 'Presidencia', 'presidencia_layout.xlsx')
PRESIDENCIA_VOTOS_PATH = os.path.join('Archivos', 'Datos', 'Presidencia', 'presidencia_layout_votos.xlsx')
PDF_FORMAT_PATH = os.path.join('Archivos', 'Formatos', '2.pdf')
FONT_PATH = os.path.join('Archivos', 'Fonts', 'LetraBrenda-Regular.otf')
SELLO_PATH = os.path.join('Archivos', 'Imagenes', 'sello.png')
PRESIDENCIA_PDF_DIR = os.path.join('Actas', 'Presidencia', 'PDF')
PRESIDENCIA_JPG_DIR = os.path.join('Actas', 'Presidencia', 'JPG')

# Procesar el archivo Excel
process_excel_file(PRESIDENCIA_LAYOUT_PATH)

# Leer el archivo Excel
df = pd.read_excel(PRESIDENCIA_VOTOS_PATH, dtype={
    'BOLETAS_SOBRANTES': str,
    'TOTAL_VOTOS_ASENTADOS': str,
    'TOTAL_PERSONAS_VOTARON': str,
    'TOTAL_REP_PARTIDO_CI_VOTARON': str,
    'P1': str,'P2': str,'P3': str,'P4': str,'P5': str,'P6': str,'P7': str,'P8': str,'P9': str,'P10': str,'P11': str,'P12': str,'P13': str,'P14': str,'P15': str,'NO_REGISTRADAS': str,'NULOS': str,'TOTAL_VOTOS': str,
    # Otras columnas que requieran ceros a la izquierda
}).fillna('')  # Reemplaza NaN con cadenas vacías

# Ruta del PDF base (formato)
pdf_base_path = PDF_FORMAT_PATH

# Crear directorio para guardar los PDFs generados
os.makedirs(PRESIDENCIA_PDF_DIR, exist_ok=True)
os.makedirs(PRESIDENCIA_JPG_DIR, exist_ok=True)  # Para guardar las imágenes JPG

# Definir las áreas específicas para cada columna (x, y)
column_areas = {
    "NOMBRE_ESTADO": (70, 261), 
    "ID_DISTRITO": (360, 260),
    "SECCION": (75, 319),
    "DATOS_QR": (260, 22),  # Añadir posición para el QR
    "CASILLA": (365, 60),
    "LETRA1": (40, 475),
    "BOLETAS_SOBRANTES": (338, 475),
    "LETRA2": (40, 560),
    "TOTAL_PERSONAS_VOTARON": (338, 560),
    "LETRA3": (40, 645),
    "TOTAL_REP_PARTIDO_CI_VOTARON": (338, 645),
    "LETRA4": (40, 725),
    "TOTAL_VOTOS_ASENTADOS": (338, 725),
    "LETRA5": (420, 630),
    "TOTAL_VOTOS_SACADOS": (740, 630),
    "LETRA6": (487, 120),
    "P1": (736, 120),
    "LETRA7": (487, 145),
    "P2": (736, 145),
    "LETRA8": (487, 170),
    "P3": (736, 170),
    "LETRA9": (487, 195),
    "P4": (736, 195),
    "LETRA10": (487, 220),
    "P5": (736, 220),
    "LETRA11": (487, 245),
    "P6": (736, 245),
    "LETRA12": (487, 270),
    "P7": (736, 270),
    "LETRA13": (487, 295),
    "P8": (736, 295),
    "LETRA14": (487, 320),
    "P9": (736, 320),
    "LETRA15": (487, 345),
    "P10": (736, 345),
    "LETRA16": (487, 373),
    "P11": (736, 373),
    "LETRA17": (487, 398),
    "P12": (736, 398),
    "LETRA18": (487, 423),
    "P13": (736, 423),
    "LETRA19": (487, 448),
    "P14": (736, 448),
    "LETRA20": (487, 475),
    "P15": (736, 475),
    "LETRA21": (487, 503),
    "NO_REGISTRADAS": (736, 503),
    "LETRA22": (487, 530),
    "NULOS": (736, 530),
    "LETRA23": (486, 555),
    "TOTAL_VOTOS": (736, 555),
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
    return '      '.join(list(texto))

# Obtener la fecha actual para usar en los nombres de archivos
fecha_actual = datetime.now().strftime("%Y%m%d_%H%M%S")  # Formato: YYYYMMDD_HHMMSS

def eliminar_espacios(cadena):
    return re.sub(r'\s+', '', cadena)

# Para cada registro en el Excel, crea un PDF y JPG
for index, row in df.iterrows():
    try:
        pdf_document = fitz.open(pdf_base_path)
        pagina = pdf_document[0]

        # Setting font
        pagina.insert_font(fontfile=FONT_PATH, fontname="F0")
        id1 = row['ID_ESTADO'] 
        id2 = row['ID_DISTRITO'] 
        id3 = row['SECCION'] 
        id4 = row['CASILLA'] 
        nombre_estado = row['NOMBRE_ESTADO'] 
        nombre_estado_lower = nombre_estado.lower()
        nombre_estado_final = eliminar_espacios(nombre_estado_lower)

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
                    fontsize=13,
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
        
        pdf_output_path = os.path.join(PRESIDENCIA_PDF_DIR, f"{nombre_estado_final}-{id2}-{id3}-{id4}.pdf")
        pdf_document.save(pdf_output_path)
        print(f"Archivo generado: {pdf_output_path} No: {index + 1}")
        
        # Ajustar resolución y tamaño del pixmap
        width_px, height_px = 5103, 3300
        scale = (width_px / pagina.rect.width, height_px / pagina.rect.height)

        # Obtener el pixmap con la resolución y dimensiones correctas
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(*scale))
        
        # Guardar como JPG y ajustar DPI a 300
        jpg_path = os.path.join(PRESIDENCIA_JPG_DIR, f"0101-{nombre_estado_final}-{id2}-{id3}-{id4}.jpg")
        print(f"Archivo generado: {jpg_path} No: {index + 1}")
        pixmap.save(jpg_path, "jpg")  # Guardar como JPG

        # Usar PIL para cambiar DPI del JPG
        image = Image.open(jpg_path)
        image.save(jpg_path, dpi=(300, 300))  # Establecer DPI
        
        # Eliminar el archivo temporal del código QR
        if os.path.exists(qr_path):
            os.remove(qr_path)  # Eliminar el archivo temporal
    except Exception as e:
        print(f"Error al procesar el registro {index + 1}: {e}")

print("PDFs y JPGs generados y guardados en las carpetas 'Archivos/presidencia/PDF/' y 'Archivos/presidencia/JPG/'.")


