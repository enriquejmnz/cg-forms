# cg-forms — Automatización de documentos PDF y Word desde Excel

Genera automáticamente documentos PDF rellenables y Word (.docx) a partir de un archivo Excel. Cada fila del Excel produce un juego completo de documentos en su propia carpeta.

## Estructura del proyecto

```
cg-forms/
├── main.py              # Script principal (lógica + CLI)
├── config.py            # Configuración y mapeos (EDITAR AQUÍ)
├── requirements.txt     # Dependencias Python
├── README.md
├── data/
│   ├── datos.xlsx       # Fuente de datos principal (Hoja1)
│   ├── datoscsv.csv     # Referencia auxiliar si necesitas revisar datos previos
│   ├── pdf1.csv         # Mapping PDF 1 usado para config.py
│   ├── pdf2.csv         # Mapping PDF 2 usado para config.py
│   └── pdf3.csv         # Mapping PDF 3 usado para config.py
├── templates/
│   ├── pdf/
│   │   ├── formulario1.pdf
│   │   ├── formulario2.pdf
│   │   └── formulario3.pdf
│   └── word/
│       ├── plantilla1.docx
│       └── plantilla2.docx
└── output/              # Aquí se generan los documentos
    ├── cg-forms.log
    └── <carpeta_por_registro>/
        ├── formulario1.pdf
        ├── formulario2.pdf
        ├── formulario3.pdf
        ├── plantilla1.docx
        └── plantilla2.docx
```

## Instalación

```bash
# 1. Crear entorno virtual (recomendado)
python3 -m venv .venv
source .venv/bin/activate   # Linux/macOS
# .venv\Scripts\activate    # Windows

# 2. Instalar dependencias
pip install -r requirements.txt
```

## Preparación de datos

### 1. Fuente de datos (`data/datos.xlsx`, hoja `Hoja1`)

Tu Excel debe tener:
- Una fila de encabezados en la primera fila.
- Una columna llamada **`procesado`** (o el nombre que configures en `config.py`).
- Las filas con `procesado` vacío, "no", "No" o "NO" serán procesadas.
- Cuando se procesan exitosamente, se marcan como "SI".

Ejemplo:

| nombre | apellido | numero_expediente | fecha | procesado |
|--------|----------|-------------------|-------|-----------|
| Juan   | Pérez    | EXP-001           | 2025-01-15 | |
| María  | López    | EXP-002           | 2025-01-16 | SI |
| Carlos | García   | EXP-003           | 2025-01-17 | no |

En este ejemplo se procesarían las filas pendientes y luego se actualizaría el mismo Excel.

### 2. Plantillas Word

Abre tus archivos `.docx` en Word/LibreOffice y coloca variables con la sintaxis de Jinja2:

```
Estimado/a {{ nombre }} {{ apellido }},

Su expediente número {{ numero_expediente }} ha sido registrado
con fecha {{ fecha }}.
```

Los nombres de variable deben coincidir con las columnas del Excel (en snake_case) o con los campos derivados definidos en `config.py`.

### 3. PDFs rellenables

Los PDFs deben tener **campos de formulario AcroForm** (campos rellenables). Para saber exactamente cómo se llaman los campos:

```bash
python main.py --list-pdf-fields
```

Esto muestra algo como:

```
📄 templates/pdf/formulario1.pdf
   3 campo(s) encontrado(s):
   - "Nombre Completo" (tipo: /Tx, valor actual: "")
   - "Fecha Solicitud" (tipo: /Tx, valor actual: "")
   - "Numero Expediente" (tipo: /Tx, valor actual: "")
```

Usa esos nombres **exactos** como claves en el mapping de `config.py`.

## Configuración de mapeos

Edita `config.py` para ajustar lo necesario. Los mapeos PDF ya fueron cargados desde `data/pdf1.csv`, `data/pdf2.csv` y `data/pdf3.csv`.

### Mapeo PDF

```python
PDF_TEMPLATES = [
    {
        "template": "templates/pdf/formulario1.pdf",
        "output_name": "formulario1.pdf",
        "mapping": {
            # clave = nombre del campo en el PDF (exacto)
            # valor = columna del Excel en snake_case o campo derivado
            "Nombre Completo": "nombre_completo",
            "Fecha Solicitud": "fecha_solicitud",
            "Numero Expediente": "numero_expediente",
        },
    },
    # ... más PDFs
]
```

### Mapeo Word

```python
WORD_TEMPLATES = [
    {
        "template": "templates/word/plantilla1.docx",
        "output_name": "plantilla1.docx",
        "mapping": {
            # clave = nombre de la variable {{ var }} en el docx (sin llaves)
            # valor = columna del Excel en snake_case o campo derivado
            "nombre": "nombre_completo",
            "expediente": "numero_expediente",
        },
    },
    # ... más plantillas
]
```

**Tip**: Si las variables del Word tienen los mismos nombres que las columnas del Excel, puedes dejar el mapping vacío `{}` y el sistema pasa todo el contexto directamente.

### Campos derivados

Para crear campos calculados a partir de las columnas existentes:

```python
DERIVED_FIELDS = {
    "nombre_completo": lambda ctx: f"{ctx.get('nombre', '')} {ctx.get('apellido', '')}".strip(),
    "fecha_corta": lambda ctx: ctx.get("fecha", "")[:10] if ctx.get("fecha") else "",
    "saludo": lambda ctx: "Sr." if ctx.get("genero") == "M" else "Sra.",
}
```

## Uso

### Listar campos de los PDFs

```bash
python main.py --list-pdf-fields
```

Ejecuta esto primero para conocer los nombres exactos de los campos de formulario.

### Validar variables de los Word

```bash
python main.py --validate-word-vars
```

Verifica que las variables `{{ ... }}` de tus plantillas Word tengan correspondencia en el mapping.

### Modo prueba (dry-run)

```bash
python main.py --dry-run
```

Simula el procesamiento completo **sin generar archivos ni modificar el Excel**. Útil para verificar que la configuración es correcta.

### Procesar una sola fila

```bash
python main.py --only-index 0 --dry-run    # Simular solo la primera fila
python main.py --only-index 0              # Procesar solo la primera fila
```

El índice es 0-based (la primera fila de datos es el índice 0).

### Procesar todo

```bash
python main.py
```

Procesa todas las filas pendientes, genera documentos y marca como "SI" las exitosas.

## Flujo de procesamiento

```
Excel → fila pendiente → dict normalizado (snake_case, sin NaN)
                        → + campos derivados = contexto enriquecido
                        → mapeo PDF → llenar formulario1.pdf, formulario2.pdf, formulario3.pdf
                        → mapeo Word → renderizar plantilla1.docx, plantilla2.docx
                        → si todo OK → marcar "procesado" = "SI" en el Excel
                        → si falla algo → log error, NO marcar, continuar con siguiente
```

## Logs

- Consola: muestra progreso en tiempo real (nivel INFO)
- Archivo: `output/cg-forms.log` (nivel DEBUG, incluye detalles completos)

Al final del procesamiento se muestra un resumen:

```
══════════════════════════════════════════════════
RESUMEN FINAL
══════════════════════════════════════════════════
Total filas en Excel:       50
Filas pendientes al inicio: 12
Exitosas:                   10
Fallidas:                    2
Modo:                       PRODUCCIÓN
══════════════════════════════════════════════════
```

## Notas técnicas

- **pypdf** se usa para llenar PDFs con campos AcroForm. Los PDFs deben ser formularios rellenables (no PDFs escaneados o planos).
- **docxtpl** usa Jinja2 para renderizar variables en Word. Soporta condicionales, bucles y filtros dentro del `.docx`.
- El Excel se lee con `pandas` (dtype=str para evitar conversiones automáticas) y se actualiza con `openpyxl`.
- Las columnas del Excel se normalizan a `snake_case` automáticamente (e.g., "Número Expediente" → "numero_expediente").

## Solución de problemas

| Problema | Solución |
|----------|----------|
| "No se encontraron campos AcroForm" | El PDF no tiene campos de formulario rellenables. Debe ser un PDF con campos editables, no un PDF plano. |
| "Campo X del mapping no existe en el PDF" | El nombre del campo no coincide exactamente. Usa `--list-pdf-fields` para ver los nombres reales. |
| "Variable X no tiene valor en el contexto" | La variable del Word no tiene columna correspondiente en el Excel ni campo derivado. Verifica el mapping. |
| El Excel no se actualiza | Cierra el archivo en Excel/LibreOffice antes de ejecutar el script. |
