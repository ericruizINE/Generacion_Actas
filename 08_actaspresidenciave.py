import pandas as pd
import qrcode
import re
import fitz  # PyMuPDF
import os
from datetime import datetime
from PIL import Image  # Para ajustar DPI de imágenes JPG
from rellenarnum import process_excel_file

# Rutas de archivos y directorios
PRESIDENCIA_VE_LAYOUT_PATH = os.path.join('Archivos', 'Datos', 'Presidencia', 'presidenciave_layout.xlsx')
PRESIDENCIA_VE_VOTOS_PATH = os.path.join('Archivos', 'Datos', 'Presidencia', 'presidenciave_layout_votos.xlsx')
PRESIDENCIA_VE_PDF_FORMAT_PATH = os.path.join('Archivos', 'Formatos', 'VE.pdf')
FONT_PATH = os.path.join('Archivos', 'Fonts', 'LetraBrenda-Regular.otf')
PRESIDENCIA_VE_PDF_DIR = os.path.join('Actas', 'Presidencia_ve', 'PDF')
PRESIDENCIA_VE_JPG_DIR = os.path.join('Actas', 'Presidencia_ve', 'JPG')

# Generar el archivo de votos a partir del layout
process_excel_file(PRESIDENCIA_VE_LAYOUT_PATH, PRESIDENCIA_VE_VOTOS_PATH)

# Leer el archivo Excel
df = pd.read_excel(PRESIDENCIA_VE_VOTOS_PATH, dtype={
    'BOLETAS_SOBRANTES': str,
    'TOTAL_VOTOS_ASENTADOS': str,
    'TOTAL_PERSONAS_VOTARON': str,
    'P1': str,'P2': str,'P3': str,'P4': str,'P5': str,'P6': str,'P7': str,'P8': str,'P9': str,'P10': str,'P11': str,'P12': str,'P13': str,'P14': str,'P15': str,'NO_REGISTRADAS': str,'NULOS': str,'TOTAL_VOTOS': str,
    # Otras columnas que requieran ceros a la izquierda
}).fillna('')  # Reemplaza NaN con cadenas vacías

# Ruta del PDF base (formato)
pdf_base_path = PRESIDENCIA_VE_PDF_FORMAT_PATH

# Crear directorios para guardar los archivos generados
os.makedirs(PRESIDENCIA_VE_PDF_DIR, exist_ok=True)
os.makedirs(PRESIDENCIA_VE_JPG_DIR, exist_ok=True)

# Definir las áreas específicas para cada columna (x, y)
column_areas = {
    "NOMBRE_ESTADO": (130, 270), 
    #"USUARIO": (250, 35),
    #"ID_DISTRITO": (250, 270),
    #"SECCION": (120, 290),
    #"CASILLA": (250, 50),
    #"ID_ACTA": (250, 65),
    #"LETRA1": (90, 420),
    #"BOLETAS_SOBRANTES": (26, 420),
    "LETRA2": (158, 580),
    "TOTAL_PERSONAS_VOTARON": (58, 580),
    "LETRA3": (158, 670),
    "TOTAL_VOTOS_ASENTADOS": (58, 670),
    #"LETRA4": (165, 750),
    #"SOBRES_VOTO": (26, 750),
    #"LETRA4": (90, 585),
    #"TOTAL_VOTOS_ASENTADOS": (26, 585),
    "LETRA5": (488, 69),
    "TOTAL_VOTOS_SACADOS": (731, 69),
    "LETRA6": (488, 187),
    "P1": (737, 187),
    "LETRA7": (488, 212),
    "P2": (737, 212),
    "LETRA8": (488, 235),
    "P3": (737, 235),
    "LETRA9": (488, 260),
    "P4": (737, 260),
    "LETRA10": (488, 286),
    "P5": (737, 286),
    "LETRA11": (488, 312),
    "P6": (737, 312),
    "LETRA12": (488, 337),
    "P7": (737, 337),
    "LETRA13": (488, 363),
    "P8": (737, 363),
    "LETRA14": (488, 389),
    "P9": (737, 389),
    "LETRA15": (488, 415),
    "P10": (737, 415),
    "LETRA16": (488, 441),
    "P11": (737, 441),
    "LETRA17": (488, 467),
    "P12": (737, 467),
    "LETRA18": (488, 493),
    "P13": (737, 493),
    "LETRA19": (488, 519),
    "P14": (737, 519),
    "LETRA20": (488, 545),
    "P15": (737, 545),
    "LETRA21": (488, 571),
    "NO_REGISTRADAS": (737, 571),
    "LETRA22": (488, 597),
    "NULOS": (737, 597),
    "LETRA23": (488, 623),
    "TOTAL_VOTOS": (717, 623)
    
}

# Lista de columnas para agregar espacios
columnas_con_espacios = ["SECCION", "BOLETAS_SOBRANTES", "TOTAL_PERSONAS_VOTARON", "TOTAL_VOTOS_SACADOS", "TOTAL_VOTOS_ASENTADOS", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "P11", "P12", "P13", "P14", "P15", "NO_REGISTRADAS", "NULOS", "TOTAL_VOTOS"]

columnas_con_cuatro = ["SECCION","TOTAL_VOTOS"]

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
        output_state_dir = os.path.join(PRESIDENCIA_VE_JPG_DIR, nombre_estado_final)
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
                    fontsize=13,
                    fontname="F0",
                    rotate=0,
                    color=(0, 0, 0),  # Color negro
                )
        
        nombre_archivo = row.get('ID_ACTA', f"{nombre_estado_final}_{nombre_archivo2}")
        pdf_output_path = os.path.join(PRESIDENCIA_VE_PDF_DIR, f"{nombre_archivo}.pdf")
        pdf_document.save(pdf_output_path)
        print(f"Archivo PDF generado: {pdf_output_path} {fecha_actual}_{index + 1}")

        # Ajustar resolución y tamaño del pixmap
        width_px, height_px = 5103, 3300
        scale = (width_px / pagina.rect.width, height_px / pagina.rect.height)

        # Obtener el pixmap con la resolución y dimensiones correctas
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(*scale))
        
        # Guardar como JPG y ajustar DPI a 300
        jpg_path = os.path.join(output_state_dir, f"ve_{nombre_estado_final}_{nombre_archivo2}.jpg")
        print(f"Archivo JPG generado: {jpg_path} {fecha_actual}_{index + 1}")
        pixmap.save(jpg_path, "jpg")  # Guardar como JPG

        # Usar PIL para cambiar DPI del JPG
        image = Image.open(jpg_path)
        image.save(jpg_path, dpi=(300, 300))  # Establecer DPI
        
    except Exception as e:
        print(f"Error al procesar el registro {index + 1}: {e}")

print(f"Archivos generados y guardados en '{PRESIDENCIA_VE_PDF_DIR}' y '{PRESIDENCIA_VE_JPG_DIR}'.")
