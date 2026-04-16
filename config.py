"""
Configuración central del proyecto cg-forms.

INSTRUCCIONES:
  1. Ajusta EXCEL_SETTINGS con la ruta, hoja y columna de estado de tu fuente de datos.
  2. En cada entrada de PDF_TEMPLATES, pega tu mapeo real:
       clave = nombre exacto del campo AcroForm en el PDF
       valor = nombre de la columna del CSV/Excel (en snake_case) o clave de DERIVED_FIELDS
  3. En cada entrada de WORD_TEMPLATES, pega tu mapeo real:
       clave = nombre de la variable {{ variable }} en el .docx (sin llaves)
       valor = nombre de la columna del CSV/Excel (en snake_case) o clave de DERIVED_FIELDS
  4. En DERIVED_FIELDS define campos calculados como funciones lambda
     que reciben el contexto base (dict) y devuelven un string.
  5. Ajusta OUTPUT_FOLDER_PATTERN para definir cómo se nombran las carpetas de salida.
"""

# ──────────────────────────────────────────────
# FUENTE DE DATOS
# ──────────────────────────────────────────────

EXCEL_SETTINGS = {
    "file": "data/datoscsv.csv",
    "sheet_name": "Hoja1",
    "status_column": "procesado",
    "pending_values": ["", "no", "NO", "No", None],
    "processed_value": "SI",
}

# ──────────────────────────────────────────────
# PLANTILLAS PDF
# ──────────────────────────────────────────────
# Cada entrada define:
#   template    -> ruta relativa al PDF rellenable original
#   output_name -> nombre del archivo generado
#   mapping     -> { campo_pdf: columna_csv_excel_o_derivado }

PDF_TEMPLATES = [
    {
        "template": "templates/pdf/formulario1.pdf",
        "output_name": "formulario1.pdf",
        "mapping": {
            "N de registro": "registro",
            "Provincia": "provincia",
            "Distrito": "distrito",
            "Corregimiento": "corregimiento",
            "Fecha oficio": "oficio_fecha",
            "Hora oficio": "oficio_hora",
            "Despacho Solicitante": "despacho_solicitante",
            "Autoridad Solicitante": "autoridad_solcitante",
            "N de Oficio": "oficio_exp",
            "N de noticia criminal": "carpetilla",
            "Delito Generico": "delito_generico",
            "Delito Especifico": "delito_especifico",
            "Lugar del hecho": "lugar_hecho",
            "Fase del proceso": "fase_proceso",
            "Fecha": "fecha_diligencia",
            "Hora": "hora_diligencia",
            "Servicio solicitado": "servicio_solicitado",
            "Profesional de psicologia": "profesional",
            "operador": "operador",
            "Nombre del usuario": "usuario_iniciales",
            "Sexo": "usuario_genero",
            "N de cedula": "usuario_cedula",
            "Fecha de nacimiento": "usuario_fecha_nacimiento",
            "Edad": "usuario_edad",
            "Nacionalidad": "usuario_nacionalidad",
            "Nombre investigado": "investigado_nombre",
            "Edad y sexo": "investigado_edad_sexo",
            "Parentesco": "investigado_parentesco",
            "Coordinadora": "coodinadora",
            "Tecnico operadora": "tecnico_operador",
        },
    },
    {
        "template": "templates/pdf/formulario2.pdf",
        "output_name": "formulario2.pdf",
        "mapping": {
            "N de registro": "registro",
            "Provincia": "provincia",
            "Distrito": "distrito",
            "Corregimiento": "corregimiento",
            "Fecha oficio": "oficio_fecha",
            "Hora oficio": "oficio_hora",
            "Despacho Solicitante": "despacho_solicitante",
            "Autoridad Solicitante": "autoridad_solcitante",
            "N de Oficio": "oficio_exp",
            "N de noticia criminal": "carpetilla",
            "Delito Generico": "delito_generico",
            "Delito Especifico": "delito_especifico",
            "Lugar del hecho": "lugar_hecho",
            "Fase del proceso": "fase_proceso",
            "Fecha": "fecha_diligencia",
            "Hora": "hora_diligencia",
            "Servicio solicitado": "servicio_solicitado",
            "Profesional de psicologia": "profesional",
            "operador": "operador",
            "Nombre del usuario": "usuario_iniciales",
            "Sexo": "usuario_genero",
            "N de cedula": "usuario_cedula",
            "Fecha de nacimiento": "usuario_fecha_nacimiento",
            "Edad": "usuario_edad",
            "Nacionalidad": "usuario_nacionalidad",
            "Nombre investigado": "investigado_nombre",
            "Edad y sexo": "investigado_edad_sexo",
            "Parentesco": "investigado_parentesco",
            "Coordinadora": "coodinadora",
            "Tecnico operadora": "tecnico_operador",
        },
    },
    {
        "template": "templates/pdf/formulario3.pdf",
        "output_name": "formulario3.pdf",
        "mapping": {
            "N de registro": "registro",
            "Provincia": "provincia",
            "Distrito": "distrito",
            "Corregimiento": "corregimiento",
            "Fecha oficio": "oficio_fecha",
            "Hora oficio": "oficio_hora",
            "Despacho Solicitante": "despacho_solicitante",
            "Autoridad Solicitante": "autoridad_solcitante",
            "N de Oficio": "oficio_exp",
            "N de noticia criminal": "carpetilla",
            "Delito Generico": "delito_generico",
            "Delito Especifico": "delito_especifico",
            "Lugar del hecho": "lugar_hecho",
            "Fase del proceso": "fase_proceso",
            "Fecha": "fecha_diligencia",
            "Hora": "hora_diligencia",
            "Servicio solicitado": "servicio_solicitado",
            "Profesional de psicologia": "profesional",
            "operador": "operador",
            "Nombre del usuario": "usuario_iniciales",
            "Sexo": "usuario_genero",
            "N de cedula": "usuario_cedula",
            "Fecha de nacimiento": "usuario_fecha_nacimiento",
            "Edad": "usuario_edad",
            "Nacionalidad": "usuario_nacionalidad",
            "Nombre investigado": "investigado_nombre",
            "Edad y sexo": "investigado_edad_sexo",
            "Parentesco": "investigado_parentesco",
            "Coordinadora": "coodinadora",
            "Tecnico operadora": "tecnico_operador",
        },
    },
]

# ──────────────────────────────────────────────
# PLANTILLAS WORD
# ──────────────────────────────────────────────
# Cada entrada define:
#   template    -> ruta relativa al .docx con variables {{ var }}
#   output_name -> nombre del archivo generado
#   mapping     -> { variable_docx: columna_csv_excel_o_derivado }
#
# Si mapping está vacío, se pasa el contexto completo a docxtpl
# (todas las columnas del CSV/Excel + campos derivados).

WORD_TEMPLATES = [
    {
        "template": "templates/word/plantilla1.docx",
        "output_name": "plantilla1.docx",
        "mapping": {
            # PEGAR_AQUI_MAPEO_WORD_1
            # Ejemplo:
            # "nombre": "nombre_completo",
            # "expediente": "numero_expediente",
        },
    },
    {
        "template": "templates/word/plantilla2.docx",
        "output_name": "plantilla2.docx",
        "mapping": {
            # PEGAR_AQUI_MAPEO_WORD_2
        },
    },
]

# ──────────────────────────────────────────────
# CAMPOS DERIVADOS
# ──────────────────────────────────────────────
# Funciones que reciben el contexto base (dict de la fila) y devuelven un valor.
# Se calculan ANTES de aplicar los mapeos de cada documento.
#
# Ejemplo:
#   "nombre_completo": lambda ctx: f"{ctx.get('nombre', '')} {ctx.get('apellido', '')}".strip(),
#   "fecha_formateada": lambda ctx: ctx.get("fecha", "")[:10] if ctx.get("fecha") else "",

DERIVED_FIELDS = {
    # PEGAR_AQUI_CAMPOS_DERIVADOS
}

# ──────────────────────────────────────────────
# NOMBRE DE CARPETA DE SALIDA
# ──────────────────────────────────────────────
# Función que recibe (contexto, indice_fila) y devuelve el nombre de la carpeta.
# Si los campos necesarios están vacíos, usa un fallback con el índice.


def output_folder_name(ctx: dict, row_index: int) -> str:
    """Genera el nombre de la carpeta de salida para un registro."""
    expediente = str(ctx.get("oficio_exp", "") or ctx.get("registro", "")).strip()
    nombre = str(ctx.get("usuario_iniciales", "") or ctx.get("investigado_nombre", "")).strip()

    if expediente and nombre:
        # Limpiar caracteres no válidos para nombres de carpeta
        safe_name = "".join(
            c if c.isalnum() or c in (" ", "-", "_") else "_" for c in nombre
        )
        safe_exp = "".join(
            c if c.isalnum() or c in ("-", "_") else "_" for c in expediente
        )
        return f"{safe_exp}_{safe_name}".strip("_")

    if expediente:
        return expediente
    if nombre:
        safe_name = "".join(
            c if c.isalnum() or c in (" ", "-", "_") else "_" for c in nombre
        )
        return safe_name.strip("_")

    return f"registro_{row_index}"


# ──────────────────────────────────────────────
# RUTAS DE SALIDA Y LOG
# ──────────────────────────────────────────────

OUTPUT_DIR = "output"
LOG_FILE = "output/cg-forms.log"
