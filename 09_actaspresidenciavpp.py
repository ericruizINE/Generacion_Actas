import pandas as pd
import qrcode
import fitz  # PyMuPDF
import os
import re
from datetime import datetime
from PIL import Image  # Para ajustar DPI de imágenes JPG
from rellenarnum import process_excel_file

# Definir rutas constantes
PRESIDENCIA_VPP_LAYOUT_PATH = os.path.join('Archivos', 'Datos', 'Presidencia', 'presidenciavpp_layout.xlsx')
PRESIDENCIA_VPP_VOTOS_PATH = os.path.join('Archivos', 'Datos', 'Presidencia', 'presidenciavpp_layout_votos.xlsx')
PRESIDENCIA_VPP_PDF_FORMAT_PATH = os.path.join('Archivos', 'Formatos', 'VPP-1.pdf')
FONT_PATH = os.path.join('Archivos', 'Fonts', 'LetraBrenda-Regular.otf')
PRESIDENCIA_VPP_PDF_DIR = os.path.join('Actas', 'Presidencia_vpp', 'PDF')
PRESIDENCIA_VPP_JPG_DIR = os.path.join('Actas', 'Presidencia_vpp', 'JPG')

# Generar el archivo de votos desde el layout
process_excel_file(PRESIDENCIA_VPP_LAYOUT_PATH, PRESIDENCIA_VPP_VOTOS_PATH)

# Leer el archivo Excel
df = pd.read_excel(PRESIDENCIA_VPP_VOTOS_PATH, dtype={
    'TOTAL DE BOLETAS SOBRANTES': str,
    'TOTAL_VOTOS_ASENTADO': str,
    'TOTAL_PERSONAS_VOTARON': str,
    'SOBRES_VOTO': str,
    'P1': str,'P2': str,'P3': str,'P4': str,'P5': str,'P6': str,'P7': str,'P8': str,'P9': str,'P10': str,'P11': str,'P12': str,'P13': str,'P14': str,'P15': str,'NO_REGISTRADAS': str,'NULOS': str,'TOTAL_VOTOS': str,
    # Otras columnas que requieran ceros a la izquierda
}).fillna('')  # Reemplaza NaN con cadenas vacías

# Ruta del PDF base (formato)
pdf_base_path = PRESIDENCIA_VPP_PDF_FORMAT_PATH

# Crear directorio para guardar los PDFs generados
#os.makedirs("PDF/vpp", exist_ok=True)
#os.makedirs("JPG/vpp", exist_ok=True)  # Para guardar las imágenes JPG

# Definir las áreas específicas para cada columna (x, y)
column_areas = {
    "NOMBRE_ESTADO": (80, 191), 
    "ID_DISTRITO": (100, 205),
    #"SECCION": (120, 205),
    #"CASILLA": (190, 230),
    #"ID_ACTA": (190, 230),
    #"CLAVE_ACTA": (190, 240),
    #"LETRA1": (90, 420),
    #"TOTAL DE BOLETAS SOBRANTES": (26, 420),
    "LETRA2": (90, 420),
    "TOTAL_PERSONAS_VOTARON": (26, 420),
    "LETRA3": (90, 473),
    "SOBRES_VOTO": (26, 473),
    "LETRA4": (90, 565),
    "TOTAL_VOTOS_ASENTADO": (26, 565),
    #"LETRA4": (90, 585),
    #"TOTAL_VOTOS_ASENTADO": (26, 585),
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
    "LETRA21": (380, 443),
    "NO_REGISTRADAS": (565, 443),
    "LETRA22": (380, 468),
    "NULOS": (565, 468),
    "LETRA23": (380, 493),
    "TOTAL_VOTOS": (565, 493)   
}

# Lista de columnas para agregar espacios
columnas_con_espacios = ["TOTAL DE BOLETAS SOBRANTES", "TOTAL_PERSONAS_VOTARON", "SOBRES_VOTO", "TOTAL_VOTOS_SACADOS", "TOTAL_VOTOS_ASENTADO", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "P11", "P12", "P13", "P14", "P15", "NO_REGISTRADAS", "NULOS", "TOTAL_VOTOS"]

#columnas_con_cuatro = ["SECCION"]

# Columnas para excluir de la impresión de texto
columnas_excluidas = ["DATOS_QR"]

# Asegurarse de que todas las columnas en columnas_con_espacios tengan ceros a la izquierda
for columna in columnas_con_espacios:
    df[columna] = df[columna].fillna('').astype(str).apply(lambda x: x.zfill(3))

# Asegurarse de que todas las columnas en columnas_con_espacios tengan ceros a la izquierda
#for columna in columnas_con_cuatro:
    #df[columna] = df[columna].fillna('').astype(str).apply(lambda x: x.zfill(4))

# Función para agregar espacios entre caracteres
def agregar_espacios(texto):
    return '     '.join(list(texto))

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
        sec = row['SEC'] 
        os.makedirs(os.path.join(PRESIDENCIA_VPP_JPG_DIR, nombre_estado_final), exist_ok=True) 
        pdf_document = fitz.open(pdf_base_path)
        pagina = pdf_document[0]

        # Setting font
        fnt = FONT_PATH
        pagina.insert_font(fontfile=fnt, fontname="F0")

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
                    rotate=0,
                    fontname="F0",
                    color=(0, 0, 0),  # Color negro
                )
        
        #nombre_archivo = row['ID_ACTA']               
        #pdf_output_path = os.path.join(PRESIDENCIA_VPP_PDF_DIR, f"{nombre_archivo}.pdf")
        #print(f"Archivo PDF generado: {nombre_archivo}.pdf {fecha_actual}_{index + 1}")
        #pdf_document.save(pdf_output_path)
        
        # Ajustar resolución y tamaño del pixmap
        width_px, height_px = 5103, 3300
        scale = (width_px / pagina.rect.width, height_px / pagina.rect.height)

        # Obtener el pixmap con la resolución y dimensiones correctas
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(*scale))
        
        # Guardar como JPG y ajustar DPI a 300
        jpg_path = os.path.join(os.path.join(PRESIDENCIA_VPP_JPG_DIR, nombre_estado_final), f"vpp_{nombre_estado_final}{nombre_archivo2}-{sec}.jpg")
        print(f"Archivo JPG generado: {jpg_path} No: {index + 1}")
        pixmap.save(jpg_path, "jpg")  # Guardar como JPG

        # Usar PIL para cambiar DPI del JPG
        image = Image.open(jpg_path)
        image.save(jpg_path, dpi=(300, 300))  # Establecer DPI
        
    except Exception as e:
        print(f"Error al procesar el registro {index + 1}: {e}")

print(f"PDFs y JPGs generados y guardados en '{PRESIDENCIA_VPP_PDF_DIR}' y '{PRESIDENCIA_VPP_JPG_DIR}'.")
