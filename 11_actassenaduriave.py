import pandas as pd
import qrcode
import re
import fitz  # PyMuPDF
import os
from datetime import datetime
from PIL import Image  # Para ajustar DPI de imágenes JPG
from rellenarnum import process_excel_file

# Rutas de archivos y directorios
SENADURIAS_VE_LAYOUT_PATH = os.path.join('Archivos', 'Datos', 'Senadurias', 'senaduriasve_layout.xlsx')
SENADURIAS_VE_VOTOS_PATH = os.path.join('Archivos', 'Datos', 'Senadurias', 'senaduriasve_layout_votos.xlsx')
SENADURIAS_VE_PDF_FORMAT_PATH = os.path.join('Archivos', 'Formatos', 'VE_SEN.pdf')
FONT_PATH = os.path.join('Archivos', 'Fonts', 'LetraBrenda-Regular.otf')
SENADURIAS_VE_PDF_DIR = os.path.join('Actas', 'Senadurias_VE', 'PDF')
SENADURIAS_VE_JPG_DIR = os.path.join('Actas', 'Senadurias_VE', 'JPG')

# Generar el archivo de votos a partir del layout
process_excel_file(SENADURIAS_VE_LAYOUT_PATH, SENADURIAS_VE_VOTOS_PATH)

# Leer el archivo Excel
df = pd.read_excel(SENADURIAS_VE_VOTOS_PATH, dtype={
    'TOTAL DE BOLETAS SOBRANTES': str,
    'TOTAL_VOTOS_ASENTADO': str,
    'TOTAL_PERSONAS_VOTARON': str,
    'SOBRES_VOTO': str,
    'P1': str,'P2': str,'P3': str,'P4': str,'P5': str,'P6': str,'P7': str,'P8': str,'P9': str,'P10': str,'P11': str,'P12': str,'P13': str,'P14': str,'P15': str,'NO_REGISTRADAS': str,'NULOS': str,'TOTAL_VOTOS': str,
    # Otras columnas que requieran ceros a la izquierda
}).fillna('')  # Reemplaza NaN con cadenas vacías

# Ruta del PDF base (formato)
pdf_base_path = SENADURIAS_VE_PDF_FORMAT_PATH

# Crear directorio para guardar los archivos generados
os.makedirs(SENADURIAS_VE_PDF_DIR, exist_ok=True)
os.makedirs(SENADURIAS_VE_JPG_DIR, exist_ok=True)

# Definir las áreas específicas para cada columna (x, y)
column_areas = {
    "NOMBRE_ESTADO": (145, 287), 
    #"USUARIO": (250, 35),
    #"ID_DISTRITO": (250, 270),
    #"SECCION": (120, 290),
    #"CASILLA": (250, 50),
    #"ID_ACTA": (250, 65),
    #"LETRA1": (90, 420),
    #"TOTAL DE BOLETAS SOBRANTES": (26, 420),
    "LETRA2": (158, 508),
    "TOTAL_PERSONAS_VOTARON": (58, 508),
    "LETRA3": (158, 620),
    "SOBRES_VOTO": (58, 620),
    #"LETRA4": (165, 750),
    #"TOTAL_VOTOS_ASENTADO": (26, 750),
    #"LETRA4": (90, 585),
    #"TOTAL_VOTOS_ASENTADO": (26, 585),
    "LETRA5": (158, 710),
    "TOTAL_VOTOS_SACADOS": (345, 710),
    "LETRA6": (488, 130),
    "P1": (737, 130),
    "LETRA7": (488, 160),
    "P2": (737, 160),
    "LETRA8": (488, 190),
    "P3": (737, 190),
    "LETRA9": (488, 220),
    "P4": (737, 220),
    "LETRA10": (488, 250),
    "P5": (737, 250),
    "LETRA11": (488, 278),
    "P6": (737, 278),
    "LETRA12": (488, 307),
    "P7": (737, 307),
    "LETRA13": (488, 337),
    "P8": (737, 337),
    "LETRA14": (488, 367),
    "P9": (737, 367),
    "LETRA15": (488, 396),
    "P10": (737, 396),
    "LETRA16": (488, 423),
    "P11": (737, 423),
    "LETRA17": (488, 453),
    "P12": (737, 453),
    "LETRA18": (488, 485),
    "P13": (737, 485),
    "LETRA19": (488, 513),
    "P14": (737, 513),
    "LETRA20": (488, 540),
    "P15": (737, 540),
    "LETRA21": (488, 568),
    "NO_REGISTRADAS": (737, 568),
    "LETRA22": (488, 595),
    "NULOS": (737, 595),
    "LETRA23": (488, 623),
    "TOTAL_VOTOS": (737, 623)
    
}

# Lista de columnas para agregar espacios
columnas_con_espacios = ["SECCION", "TOTAL DE BOLETAS SOBRANTES", "TOTAL_PERSONAS_VOTARON", "SOBRES_VOTO", "TOTAL_VOTOS_SACADOS", "TOTAL_VOTOS_ASENTADO", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "P11", "P12", "P13", "P14", "P15", "NO_REGISTRADAS", "NULOS", "TOTAL_VOTOS"]

columnas_con_cuatro = ["SECCION"]

# Columnas para excluir de la impresión de texto
columnas_excluidas = ["DATOS_QR"]

# Asegurarse de que todas las columnas en columnas_con_espacios tengan ceros a la izquierda
for columna in columnas_con_espacios:
    df[columna] = df[columna].fillna('').astype(str).apply(lambda x: x.zfill(3))

# Asegurarse de que todas las columnas en columnas_con_espacios tengan ceros a la izquierda
for columna in columnas_con_cuatro:
    df[columna] = df[columna].fillna('').astype(str).apply(lambda x: x.zfill(4))


# Función para agregar espacios entre caracteres
def agregar_espacios(texto):
    return '      '.join(list(texto))

# Obtener la fecha actual para usar en los nombres de archivos
fecha_actual = datetime.now().strftime("%Y%m%d_%H%M%S")  # Formato: YYYYMMDD_HHMMSS

def eliminar_espacios(cadena):
    # Usamos una expresión regular para eliminar todos los espacios en blanco (incluyendo tabulaciones y espacios especiales)
    return re.sub(r'\s+', '', cadena)

# Para cada registro en el Excel, crea un PDF y JPG
for index, row in df.iterrows():
    try:
        nombre_estado = row['NOMBRE_ESTADO'] 
        nombre_estado_lower = nombre_estado.lower()
        nombre_estado_final = eliminar_espacios(nombre_estado_lower)
        nombre_archivo2 = row['ID_DISTRITO'] 
        os.makedirs(os.path.join(SENADURIAS_VE_JPG_DIR, nombre_estado_final), exist_ok=True)
        pdf_document = fitz.open(pdf_base_path)
        pagina = pdf_document[0]

        # Setting font
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
                    fontsize=13,
                    fontname="F0",
                    rotate=0,
                    color=(0, 0, 0),  # Color negro
                )
        
        #nombre_archivo = row['ID_ACTA']               
        #pdf_output_path = os.path.join(SENADURIAS_VE_PDF_DIR, f"{nombre_archivo}.pdf")
        #print(f"Archivo PDF generado acta:{nombre_archivo}.pdf {fecha_actual}_{index + 1}")
        #pdf_document.save(pdf_output_path)
        
        # Ajustar resolución y tamaño del pixmap
        width_px, height_px = 5103, 3300
        scale = (width_px / pagina.rect.width, height_px / pagina.rect.height)

        # Obtener el pixmap con la resolución y dimensiones correctas
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(*scale))
        
        # Guardar como JPG y ajustar DPI a 300
        jpg_path = os.path.join(os.path.join(SENADURIAS_VE_JPG_DIR, nombre_estado_final), f"ve_{nombre_estado_final}{nombre_archivo2}.jpg")
        print(f"Archivo JPG generado: {jpg_path} {fecha_actual}_{index + 1}")
        pixmap.save(jpg_path, "jpg")  # Guardar como JPG

        # Usar PIL para cambiar DPI del JPG
        image = Image.open(jpg_path)
        image.save(jpg_path, dpi=(300, 300))  # Establecer DPI
        
    except Exception as e:
        print(f"Error al procesar el registro {index + 1}: {e}")

print(f"PDFs y JPGs generados y guardados en '{SENADURIAS_VE_PDF_DIR}' y '{SENADURIAS_VE_JPG_DIR}'")
