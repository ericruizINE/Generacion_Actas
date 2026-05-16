import os
import pandas as pd
import numpy as np
from num2words import num2words
from random import randint


def generate_random_numbers(limit, num_columns):
    # Generar números aleatorios que sumen hasta 'limit'
    random_numbers = np.random.rand(num_columns)
    random_numbers /= random_numbers.sum()  # Normalizar para que sumen 1
    random_numbers *= limit  # Escalar para que sumen 'limit'
    return random_numbers.round().astype(int)  # Redondear a enteros


def process_excel_file(input_path, output_path=None):
    """Procesa un archivo Excel y genera un nuevo archivo con números aleatorios y sus letras."""
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"No se encontró el archivo: {input_path}")

    df = pd.read_excel(input_path)

    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_votos{ext}"

    columns_to_fill = ['P1','P2','P3','P4','P5','P6','P7','P8','P9','P10','P11','P12','P13','P14','P15','NO_REGISTRADAS','NULOS']
    mapeo_columnas = {
        "LETRA1": "BOLETAS_SOBRANTES",
        "LETRA2": "TOTAL_PERSONAS_VOTARON",
        "LETRA3": "TOTAL_REP_PARTIDO_CI_VOTARON",
        "LETRA4": "TOTAL_VOTOS_ASENTADOS",
        "LETRA5": "TOTAL_VOTOS_SACADOS",
        "LETRA6": "P1",
        "LETRA7": "P2",
        "LETRA8": "P3",
        "LETRA9": "P4",
        "LETRA10": "P5",
        "LETRA11": "P6",
        "LETRA12": "P7",
        "LETRA13": "P8",
        "LETRA14": "P9",
        "LETRA15": "P10",
        "LETRA16": "P11",
        "LETRA17": "P12",
        "LETRA18": "P13",
        "LETRA19": "P14",
        "LETRA20": "P15",
        "LETRA21": "NO_REGISTRADAS",
        "LETRA22": "NULOS",
        "LETRA23": "TOTAL_VOTOS",
    }
    columns_to_concatenate = ['ID_ESTADO','ID_DISTRITO','TIPO_ACTA','SECCION','ID_CASILLA','TIPO_CASILLA','EXT_CONTIGUA','ID_TIPO_CANDIDATURA']
    columns_to_concatenate2 = ['TIPO_CASILLA','ID_CASILLA','EXT_CONTIGUA']

    num_columns = len(columns_to_fill)
    if num_columns == 0:
        raise ValueError("La lista de columnas a rellenar está vacía.")

    for index, row in df.iterrows():
        n = randint(0, 350)
        repre = randint(0, 25)
        limit = row['LISTA_NOMINAL_CASILLA'] - n
        if limit < 0:
            limit = 0
        random_numbers = generate_random_numbers(limit, num_columns)
        total_votos = random_numbers.sum()
        df.loc[index, columns_to_fill] = random_numbers
        df.loc[index, 'BOLETAS_SOBRANTES'] = row['LISTA_NOMINAL_CASILLA'] - total_votos
        df.loc[index, 'TOTAL_PERSONAS_VOTARON'] = total_votos
        df.loc[index, 'TOTAL_REP_PARTIDO_CI_VOTARON'] = repre
        df.loc[index, 'TOTAL_VOTOS_ASENTADOS'] = total_votos - repre
        df.loc[index, 'TOTAL_VOTOS'] = total_votos
        df.loc[index, 'TOTAL_VOTOS_SACADOS'] = total_votos

    df['DATOS_QR'] = df[columns_to_concatenate].astype(str).agg('|'.join, axis=1)
    df['CASILLA'] = df[columns_to_concatenate2].astype(str).agg(' - '.join, axis=1)

    for columna_letra, columna_numero in mapeo_columnas.items():
        if columna_numero in df.columns:
            df[columna_letra] = df[columna_numero].apply(
                lambda x: (
                    'Ilegible' if '/' in str(x) else
                    '' if '*' in str(x) else
                    num2words(x, lang='es').capitalize() if not pd.isnull(x) else ''
                )
            )

    df.to_excel(output_path, index=False)
    print(f"Archivo {input_path} procesado. Salida: {output_path}")
    return output_path


if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("Uso: python rellenarnum.py <ruta_archivo_entrada> [ruta_archivo_salida]")
        sys.exit(1)

    entrada = sys.argv[1]
    salida = sys.argv[2] if len(sys.argv) > 2 else None
    process_excel_file(entrada, salida)
