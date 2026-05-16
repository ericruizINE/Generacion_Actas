# Actas Electorales - Sistema de Generación de Documentos

Sistema automatizado de Python para la generación de actas electorales en formatos PDF y JPG. Soporta múltiples variantes de actas (Presidencia, Senadurías, Diputaciones) con opciones especiales como RP (Representación Proporcional), VA (Voto Anticipado), VE (Voto en el Extranjero) y VPP (Voto en Prisión Preventiva).

## Funcionalidad General

El sistema realiza las siguientes operaciones:

1. **Procesamiento de Excel** - Lee layouts de Excel y genera archivos de votos procesados mediante `rellenarnum.process_excel_file()`
2. **Llenado de PDF** - Utiliza plantillas PDF base para insertar datos:
   - Texto con coordenadas específicas
   - Códigos QR con información del acta
   - Números con espacios para legibilidad
   - Fuentes personalizadas para campos específicos
3. **Exportación de Imágenes** - Convierte PDFs a JPGs a 300 DPI con resolución 5103x3300 px
4. **Organización por Estado** - Genera salida estructurada por estado en carpetas separadas

## Estructura de Directorios

```
Actas/
├── Presidencia_va/          # Actas de Presidencia variante VA
│   ├── PDF/
│   └── JPG/
├── Presidencia_ve/          # Actas de Presidencia variante VE
│   ├── PDF/
│   └── JPG/
├── Presidencia_vpp/         # Actas de Presidencia variante VPP
│   ├── PDF/
│   └── JPG/
├── Senadurias_va/           # Actas de Senadurías variante VA
│   ├── PDF/
│   └── JPG/
├── Senadurias_VE/           # Actas de Senadurías variante VE
│   ├── PDF/
│   └── JPG/
└── Diputaciones_va/         # Actas de Diputaciones variante VA
    ├── PDF/
    └── JPG/

Archivos/
├── Datos/
│   ├── Presidencia/         # Layouts y votos para Presidencia
│   ├── Senadurias/          # Layouts y votos para Senadurías
│   └── Diputaciones/        # Layouts y votos para Diputaciones
├── Fonts/
│   ├── ARIAL.TTF
│   ├── LetraBrenda-Regular.otf
│   └── Letras/              # Fuentes adicionales
├── Formatos/                # Plantillas PDF base (VA-DIP.pdf, VE-SEN.pdf, etc.)
└── VOTACIONES_ALTERNAS.csv

PDF/
├── presidencia2026/         # PDFs de presidencia 2026
└── ...

JPG/
├── presidencia2026/         # JPGs de presidencia 2026
└── ...
```

## Dependencias principales

- `pandas` - Procesamiento de archivos Excel
- `PyMuPDF` (`fitz`) - Manipulación y generación de PDFs
- `qrcode` - Generación de códigos QR
- `Pillow` - Procesamiento de imágenes JPG
- `num2words` - Conversión de números a letras (en `rellenarnum`)
- `selenium` - Automatización web (scripts auxiliares)
- `webdriver-manager` - Gestión de WebDriver para Selenium

Para instalar todas las dependencias:

```bash
pip install -r requirements.txt
```

## Módulos principales

### `rellenarnum.py`
Módulo central que implementa `process_excel_file(layout_path, votos_path)`:
- Lee un archivo Excel de layout con estructura de acta
- Procesa datos y genera un archivo de votos procesado
- Utilizado por todos los scripts de generación de actas antes de leer los datos

## Scripts de Generación de Actas

Los scripts siguen un patrón consistente:

1. **Importan `process_excel_file`** desde `rellenarnum`
2. **Definen constantes de rutas** usando `os.path.join()`:
   - `*_LAYOUT_PATH` - Ruta al archivo de layout de Excel
   - `*_VOTOS_PATH` - Ruta donde se generará el archivo de votos
   - `*_PDF_FORMAT_PATH` - Ruta a la plantilla PDF base
   - `FONT_PATH` - Ruta a la fuente para inserción de texto
   - `*_PDF_DIR` - Directorio de salida para PDFs
   - `*_JPG_DIR` - Directorio de salida para JPGs
3. **Llaman a `process_excel_file()`** antes de leer el Excel
4. **Registran la fuente** con `pagina.insert_font(fontfile=FONT_PATH, fontname="F0")`
5. **Insertan texto** en coordenadas específicas con espacios entre caracteres para valores numéricos
6. **Generan salida** en PDF y JPG a 300 DPI

### Scripts por Tipo de Acta

#### Presidencia

**`01_actaspresidencia.py`** - Actas base de Presidencia
- Layout: `Archivos/Datos/Presidencia/presidencia_layout.xlsx`
- Plantilla: `Archivos/Formatos/2.pdf`
- Salida: `Actas/Presidencia/PDF/` y `JPG/`

**`02_actaspresidencia_especial.py`** - Actas especiales de Presidencia
- Plantilla: `Archivos/Formatos/2E.pdf`
- Salida: `Actas/Presidencia_especial/PDF/` y `JPG/`

**`07_actaspresidenciava.py`** - Actas de Presidencia - Voto Anticipado (VA)
- Layout: `Archivos/Datos/Presidencia/presidenciava_layout.xlsx`
- Plantilla: `Archivos/Formatos/VA.pdf`
- Salida: `Actas/Presidencia_va/PDF/` y `JPG/`
- Fuente: `LetraBrenda-Regular.otf`

**`08_actaspresidenciave.py`** - Actas de Presidencia - Voto en el Extranjero (VE)
- Layout: `Archivos/Datos/Presidencia/presidenciave_layout.xlsx`
- Plantilla: `Archivos/Formatos/VE.pdf`
- Salida: `PDF/presidencia2026/` y `JPG/presidencia2026/`
- Fuente: `ARIAL.TTF`

**`09_actaspresidenciavpp.py`** - Actas de Presidencia - Voto en Prisión Preventiva (VPP)
- Layout: `Archivos/Datos/Presidencia/presidenciavpp_layout.xlsx`
- Plantilla: `Archivos/Formatos/VPP-1.pdf`
- Salida: `Actas/Presidencia_vpp/PDF/` y `JPG/`
- Fuente: `LetraBrenda-Regular.otf`

#### Senadurías

**`03_actassenadurias.py`** - Actas base de Senadurías
- Layout: `Archivos/Datos/Senadurias/senadurias_layout.xlsx`
- Plantilla: `Archivos/Formatos/3V1.pdf`
- Salida: `Actas/Senadurias/PDF/` y `JPG/`

**`04_actassenaduriasrp.py`** - Actas de Senadurías - Representación Proporcional (RP)
- Layout: `Archivos/Datos/Senadurias/senadorias_rp_layout.xlsx`
- Plantilla: `Archivos/Formatos/3ERP.pdf`
- Salida: `Actas/Senadurias_rp/PDF/` y `JPG/`

**`10_actasenaduriasva.py`** - Actas de Senadurías - Voto Anticipado (VA)
- Layout: `Archivos/Datos/Senadurias/senaduriasva_layout.xlsx`
- Plantilla: `Archivos/Formatos/VA-SEN.pdf.pdf`
- Salida: `Actas/Senadurias_va/PDF/` y `JPG/`
- Fuente: `LetraBrenda-Regular.otf`

**`11_actassenaduriave.py`** - Actas de Senadurías - Voto en el Extranjero (VE)
- Layout: `Archivos/Datos/Senadurias/senaduriasve_layout.xlsx`
- Plantilla: `Archivos/Formatos/VE-SEN.pdf.pdf`
- Salida: `Actas/Senadurias_VE/PDF/` y `JPG/`
- Fuente: `LetraBrenda-Regular.otf`

#### Diputaciones

**`05_actasdiputaciones.py`** - Actas base de Diputaciones
- Layout: `Archivos/Datos/Diputaciones/diputaciones_layout.xlsx`
- Plantilla: `Archivos/Formatos/4V1.pdf`
- Salida: `Actas/Diputaciones/PDF/` y `JPG/`

**`06_actasdiputacionesrp.py`** - Actas de Diputaciones - Representación Proporcional (RP)
- Layout: `Archivos/Datos/Diputaciones/diputaciones_rp_layout.xlsx`
- Plantilla: `Archivos/Formatos/4ERP.pdf`
- Salida: `Actas/Diputaciones_rp/PDF/` y `JPG/`

**`12_actasdiputacionesva.py`** - Actas de Diputaciones - Voto Anticipado (VA)
- Layout: `Archivos/Datos/Diputaciones/diputacionesva_layout.xlsx`
- Plantilla: `Archivos/Formatos/VA-DIP.pdf`
- Salida: `Actas/Diputaciones_va/PDF/` y `JPG/`
- Fuente: `LetraBrenda-Regular.otf`

**`13_actasdiputacionesve.py`** - Actas de Diputaciones - Voto en el Extranjero (VE) *[PENDIENTE]*
- Layout: `Archivos/Datos/Diputaciones/diputacionesve_layout.xlsx`
- Plantilla: `Archivos/Formatos/VE-DIP.pdf`
- Salida: `Actas/Diputaciones_ve/PDF/` y `JPG/`
- Fuente: `LetraBrenda-Regular.otf`

### Scripts de Utilidad

**`crop_generated_jpgs.py`**
- Procesa JPGs generados para detectar y recortar regiones específicas
- Utiliza thresholding binario y análisis de componentes conectados
- Detecta patrones de círculos con puntos y bloques verticales
- Recorta y guarda subimágenes en carpetas de salida

**`votos_alternas.py`**
- Procesa datos de votaciones alternas desde CSV
- Genera reportes y análisis de votos especiales

**`convertirnumaletra.py`**
- Utilidad para conversión de números a letras
- Soporta formato con espacios para legibilidad en actas

## Flujo de Procesamiento

1. Usuario ejecuta script de acta específico
2. Script define constantes de ruta
3. `process_excel_file()` lee layout y genera archivo de votos
4. Script lee archivo de votos con `pd.read_excel()`
5. Para cada registro en el Excel:
   - Abre plantilla PDF base
   - Registra fuente personalizada
   - Inserta texto en coordenadas específicas
   - Aplica espacios entre caracteres para números
   - Genera JPG a 300 DPI desde PDF
   - Ajusta DPI con PIL
6. Imprime confirmación de generación

## Ejecución

Para generar actas de un tipo específico, ejecute desde la raíz del repositorio:

### Presidencia
```bash
python 01_actaspresidencia.py             # Actas base
python 02_actaspresidencia_especial.py    # Actas especiales
python 07_actaspresidenciava.py           # Voto Anticipado (VA)
python 08_actaspresidenciave.py           # Voto en el Extranjero (VE)
python 09_actaspresidenciavpp.py          # Voto en Prisión Preventiva (VPP)
```

### Senadurías
```bash
python 03_actassenadurias.py              # Actas base
python 04_actassenaduriasrp.py            # Representación Proporcional (RP)
python 10_actasenaduriasva.py             # Voto Anticipado (VA)
python 11_actassenaduriave.py             # Voto en el Extranjero (VE)
```

### Diputaciones
```bash
python 05_actasdiputaciones.py            # Actas base
python 06_actasdiputacionesrp.py          # Representación Proporcional (RP)
python 12_actasdiputacionesva.py          # Voto Anticipado (VA)
python 13_actasdiputacionesve.py          # Voto en el Extranjero (VE) [PENDIENTE]
```

Los archivos generados se guardan automáticamente en las carpetas especificadas según el tipo y variante de acta.

## Notas Técnicas

### Patrones de Implementación

Todos los scripts de generación de actas siguen estos patrones:

**Constantes centralizadas:**
```python
SCRIPT_VA_LAYOUT_PATH = os.path.join('Archivos', 'Datos', 'Tipo', 'script_va_layout.xlsx')
SCRIPT_VA_VOTOS_PATH = os.path.join('Archivos', 'Datos', 'Tipo', 'script_va_votos.xlsx')
SCRIPT_VA_PDF_FORMAT_PATH = os.path.join('Archivos', 'Formatos', 'VA.pdf')
FONT_PATH = os.path.join('Archivos', 'Fonts', 'LetraBrenda-Regular.otf')
SCRIPT_VA_PDF_DIR = os.path.join('Actas', 'Tipo_va', 'PDF')
SCRIPT_VA_JPG_DIR = os.path.join('Actas', 'Tipo_va', 'JPG')
```

**Procesamiento de votos:**
```python
from rellenarnum import process_excel_file

process_excel_file(SCRIPT_VA_LAYOUT_PATH, SCRIPT_VA_VOTOS_PATH)
df = pd.read_excel(SCRIPT_VA_VOTOS_PATH, dtype={...})
```

**Inserción de texto:**
```python
pagina.insert_font(fontfile=FONT_PATH, fontname="F0")
pagina.insert_text((x, y), texto, fontsize=12, fontname="F0", color=(0, 0, 0))
```

### Características por Tipo de Acta

| Tipo | Scripts | Variantes | Ubicación |
|------|---------|-----------|-----------|
| **Presidencia** | 01, 02 | Base, Especial | `Actas/Presidencia/` |
| **Presidencia Variantes** | 07, 08, 09 | VA, VE, VPP | `Actas/Presidencia_*/` |
| **Senadurías** | 03, 04 | Base, RP | `Actas/Senadurias/`, `Actas/Senadurias_rp/` |
| **Senadurías Variantes** | 10, 11 | VA, VE | `Actas/Senadurias_va/`, `Actas/Senadurias_VE/` |
| **Diputaciones** | 05, 06 | Base, RP | `Actas/Diputaciones/`, `Actas/Diputaciones_rp/` |
| **Diputaciones Variantes** | 12, 13* | VA, VE | `Actas/Diputaciones_va/`, `Actas/Diputaciones_ve/` |

*Script 13 pendiente de implementación

### Nomenclatura de Variantes

- **VA** = Voto Anticipado - Permite a ciudadanos votar antes de la fecha de elecciones
- **VE** = Voto en el Extranjero - Para ciudadanos residentes fuera del país
- **VPP** = Voto en Prisión Preventiva - Para personas en detención preventiva
- **RP** = Representación Proporcional - Sistema electoral basado en proporcionalidad

### Scripts de Utilidad

**`rellenarnum.py`** - Función `process_excel_file(layout_path, votos_path)`
- Convierte números a letras en columnas `LETRA*`
- Genera archivo de votos procesado para inserción en PDF
- Esencial para flujo layout → votos → PDF/JPG

**`crop_generated_jpgs.py`** - Procesamiento de imágenes generadas
- Detecta regiones de interés en JPGs
- Utiliza binary thresholding y connected components
- Identifica patrones de círculo-con-punto y bloques verticales

**`votos_alternas.py`** - Procesamiento de votaciones especiales
- Lee `Archivos/VOTACIONES_ALTERNAS.csv`
- Genera análisis de votos alternativos

### Requisitos del Sistema

- Python 3.7+
- 100+ MB libres para archivos generados
- Chrome/Chromium (para scripts de Captura/)

### Mejoras Futuras

- Integración de validación de coordenadas
- Automatización de detección de plantillas
- API REST para generación bajo demanda
- Dashboard web para monitoreo

## Licencia

Este proyecto es parte del sistema electoral del INE. Consulte las políticas internas para distribución y uso.
