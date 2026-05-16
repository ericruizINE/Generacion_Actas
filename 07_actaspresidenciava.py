import pandas as pd
import qrcode
import hashlib
import re
import fitz  # PyMuPDF
import os
from datetime import datetime
from PIL import Image  # Para ajustar DPI de imágenes JPG
from rellenarnum import process_excel_file

# Rutas de archivos y directorios
PRESIDENCIA_VA_LAYOUT_PATH = os.path.join('Archivos', 'Datos', 'Presidencia', 'presidenciava_layout.xlsx')
PRESIDENCIA_VA_VOTOS_PATH = os.path.join('Archivos', 'Datos', 'Presidencia', 'presidenciava_layout_votos.xlsx')
PRESIDENCIA_VA_PDF_FORMAT_PATH = os.path.join('Archivos', 'Formatos', 'VA.pdf')
FONT_PATH = os.path.join('Archivos', 'Fonts', 'LetraBrenda-Regular.otf')
PRESIDENCIA_VA_PDF_DIR = os.path.join('Actas', 'Presidencia_va', 'PDF')
PRESIDENCIA_VA_JPG_DIR = os.path.join('Actas', 'Presidencia_va', 'JPG')

# Generar el archivo de votos a partir del layout
process_excel_file(PRESIDENCIA_VA_LAYOUT_PATH, PRESIDENCIA_VA_VOTOS_PATH)

# Leer el archivo Excel
df = pd.read_excel(PRESIDENCIA_VA_VOTOS_PATH, dtype={
    'BOLETAS_SOBRANTES': str,
    'TOTAL_VOTOS_ASENTADOS': str,
    'TOTAL_PERSONAS_VOTARON': str,
    'SOBRES_VOTO': str,
    'P1': str,'P2': str,'P3': str,'P4': str,'P5': str,'P6': str,'P7': str,'P8': str,'P9': str,'P10': str,'P11': str,'P12': str,'P13': str,'P14': str,'P15': str,'NO_REGISTRADAS': str,'NULOS': str,'TOTAL_VOTOS': str,
    # Otras columnas que requieran ceros a la izquierda
}).fillna('')  # Reemplaza NaN con cadenas vacías

# Ruta del PDF base (formato)
pdf_base_path = PRESIDENCIA_VA_PDF_FORMAT_PATH

# Crear directorios para guardar los archivos generados
os.makedirs(PRESIDENCIA_VA_PDF_DIR, exist_ok=True)
os.makedirs(PRESIDENCIA_VA_JPG_DIR, exist_ok=True)

# Definir las áreas específicas para cada columna (x, y)
column_areas = {
    "NOMBRE_ESTADO": (80, 191), 
    #"ID_DISTRITO": (175, 191),
    #"SECCION": (190, 191),
    #"CASILLA": (190, 230),
    #"ID_ACTA": (190, 230),
    #"CLAVE_ACTA": (190, 240),
    #"LETRA1": (90, 420),
    #"BOLETAS_SOBRANTES": (26, 420),
    "LETRA2": (90, 420),
    "TOTAL_PERSONAS_VOTARON": (26, 420),
    "LETRA3": (90, 473),
    "SOBRES_VOTO": (26, 473),
    "LETRA4": (90, 565),
    "TOTAL_VOTOS_ASENTADOS": (26, 565),
    #"LETRA4": (90, 585),
    #"TOTAL_VOTOS_ASENTADOSS": (26, 585),
    #"LETRA5": (90, 590),
    #"TOTAL_VOTOS_SACADOS": (26, 590),
    "LETRA6": (380, 84),
    "P1": (565, 84),
    "LETRA7": (380, 108),
    "P2": (565, 108),
    "LETRA8": (380, 133),
    "P3": (565, 133),
    "LETRA9": (380, 158),
    "P4": (565, 158),
    "LETRA10": (380, 181),
    "P5": (565, 181),
    "LETRA11": (380, 203),
    "P6": (565, 203),
    "LETRA12": (380, 227),
    "P7": (565, 227),
    "LETRA13": (380, 252),
    "P8": (565, 252),
    "LETRA14": (380, 275),
    "P9": (565, 275),
    "LETRA15": (380, 300),
    "P10": (565, 300),
    "LETRA16": (380, 325),
    "P11": (565, 325),
    "LETRA17": (380, 350),
    "P12": (565, 350),
    "LETRA18": (380, 373),
    "P13": (565, 373),
    "LETRA19": (380, 395),
    "P14": (565, 395),
    "LETRA20": (380, 420),
    "P15": (565, 420),
    "LETRA21": (380, 445),
    "NO_REGISTRADAS": (565, 445),
    "LETRA22": (380, 470),
    "NULOS": (565, 470),
    "LETRA23": (380, 495),
    "TOTAL_VOTOS": (565, 495)    
}

# Lista de columnas para agregar espacios
columnas_con_espacios = ["BOLETAS_SOBRANTES", "TOTAL_PERSONAS_VOTARON", "SOBRES_VOTO", "TOTAL_VOTOS_SACADOS", "TOTAL_VOTOS_ASENTADOS", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "P11", "P12", "P13", "P14", "P15", "NO_REGISTRADAS", "NULOS", "TOTAL_VOTOS"]

#columnas_con_cuatro = ["SECCION"]

# Columnas para excluir de la impresión de texto
columnas_excluidas = ["DATOS_QR"]

# Asegurarse de que todas las columnas en columnas_con_espacios tengan ceros a la izquierda
for columna in columnas_con_espacios:
    df[columna] = df[columna].fillna('').astype(str).apply(lambda x: x.zfill(3))

# Asegurarse de que todas las columnas en columnas_con_espacios tengan ceros a la izquierda
#for columna in columnas_con_cuatro:
   # df[columna] = df[columna].fillna('').astype(str).apply(lambda x: x.zfill(4))


# Función para agregar espacios entre caracteres
def agregar_espacios(texto):
    return '     '.join(list(texto))

# Obtener la fecha actual para usar en los nombres de archivos
fecha_actual = datetime.now().strftime("%Y%m%d_%H%M%S")  # Formato: YYYYMMDD_HHMMSS

def eliminar_espacios(cadena):
    # Usamos una expresión regular para eliminar todos los espacios en blanco (incluyendo tabulaciones y espacios especiales)
    return re.sub(r'\s+', '', cadena)

# Para calcular el hash SHA-256 de un archivo
def calcular_sha256(file_path):
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

# Para cada registro en el Excel, crea un PDF y JPG
for index, row in df.iterrows():
    try:
        nombre_estado = row['NOMBRE_ESTADO']
        nombre_estado_lower = nombre_estado.lower()
        nombre_estado_final = eliminar_espacios(nombre_estado_lower)
        nombre_archivo2 = row['ID_DISTRITO']
        output_state_dir = os.path.join(PRESIDENCIA_VA_JPG_DIR, nombre_estado_final)
        os.makedirs(output_state_dir, exist_ok=True)
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
                    fontsize=12,
                    fontname="F0",
                    rotate=0,
                    color=(0, 0, 0),  # Color negro
                )
                                 
        pdf_output_path = os.path.join(PRESIDENCIA_VA_PDF_DIR, f"va_{nombre_estado_final}_{nombre_archivo2}.pdf")
        print(f"Archivo PDF generado: {nombre_archivo2}.pdf {fecha_actual}_{index + 1}")
        pdf_document.save(pdf_output_path)
        
        # Ajustar resolución y tamaño del pixmap
        width_px, height_px = 5103, 3300
        scale = (width_px / pagina.rect.width, height_px / pagina.rect.height)

        # Obtener el pixmap con la resolución y dimensiones correctas
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(*scale))
        
        # Guardar como JPG y ajustar DPI a 300
        jpg_path = os.path.join(output_state_dir, f"va_{nombre_estado_final}_{nombre_archivo2}.jpg")
        print(f"Archivo JPG generado: {jpg_path} {fecha_actual}_{index + 1}")
        pixmap.save(jpg_path, "jpg")  # Guardar como JPG

        # Usar PIL para cambiar DPI del JPG
        image = Image.open(jpg_path)
        image.save(jpg_path, dpi=(300, 300))  # Establecer DPI

        #sha256_hash = calcular_sha256(jpg_path)
        #jpg_path2 = os.path.join("diputaciones_va/JPG", f"{sha256_hash}.jpg")
        #os.rename(jpg_path, jpg_path2)
        
    except Exception as e:
        print(f"Error al procesar el registro {index + 1}: {e}")

print(f"Archivos generados y guardados en '{PRESIDENCIA_VA_PDF_DIR}' y '{PRESIDENCIA_VA_JPG_DIR}'.")
