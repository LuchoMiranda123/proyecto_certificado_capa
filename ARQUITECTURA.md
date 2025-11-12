# 📐 Arquitectura del Proyecto

## 📊 Resumen de la Refactorización

Se ha reorganizado el código para mejorar la mantenibilidad y separar responsabilidades.

### Antes:
- `app.py`: **1038 líneas** - Todo mezclado (UI + lógica de negocio)

### Después:
- `app.py`: **711 líneas** (-31%) - Solo interfaz de Streamlit
- `generador_archivos.py`: **421 líneas** - Lógica de generación
- **Total**: 1132 líneas (94 líneas adicionales por documentación)

---

## 📁 Estructura de Archivos

```
Certificados/
├── app.py                      # 🎨 Interfaz de usuario con Streamlit
├── generador_archivos.py       # 🔧 Lógica de generación de archivos
├── formato_excel.py            # 📊 Formateo de Excel
├── config_cursos.json          # ⚙️ Configuración de cursos
├── requirements.txt            # 📦 Dependencias
├── README.md                   # 📖 Documentación de uso
└── ARQUITECTURA.md            # 📐 Este archivo
```

---

## 🎯 Responsabilidades

### `app.py` - Interfaz de Usuario
**Responsabilidad**: Gestionar la interfaz de Streamlit y la interacción con el usuario.

**Funciones**:
- ✅ Configuración de la página de Streamlit
- ✅ Gestión del estado de sesión
- ✅ Carga de archivos (Personal Asignado, Maestro de Notas)
- ✅ Procesamiento y validación de DNIs
- ✅ Selección de cursos
- ✅ Configuración de detalles de cada curso
- ✅ Interfaz de descarga de archivos

**No contiene**: Lógica de generación de archivos, conversión a PDF, procesamiento de datos.

---

### `generador_archivos.py` - Lógica de Negocio
**Responsabilidad**: Generar los archivos de capacitación en los formatos solicitados.

**Funciones públicas**:

#### `get_nombre_completo_curso(nombre_truncado, config_cursos)`
Mapea nombres truncados de hojas Excel a nombres completos.
- **Entrada**: Nombre truncado (max 31 caracteres)
- **Salida**: Nombre completo del curso
- **Usa**: Detección inteligente de números de hoja y partes

#### `generar_zip_formatos(dnis_procesados, selected_courses, maestro_excel, course_configs, output_format, progress_callback=None)`
Función principal que genera el archivo ZIP con todos los formatos.
- **Entrada**: 
  - DataFrame con DNIs procesados
  - Lista de cursos seleccionados
  - Archivo Excel con notas
  - Configuraciones de cursos
  - Formato de salida ("Excel (.xlsx)", "PDF", "Ambos (Excel + PDF)")
  - Callback opcional para reportar progreso
- **Salida**: 
  - BytesIO con el ZIP generado
  - Nombre sugerido del archivo
  - Lista de warnings/errores

**Funciones internas**:

#### `buscar_nota_en_maestro(dni, maestro_curso)`
Busca la información de nota de un DNI en el maestro.

#### `procesar_curso(curso, dnis_procesados, maestro_excel, course_config)`
Procesa un curso individual y genera su DataFrame.

#### `convertir_excel_a_pdf(excel_data, base_filename, excel_app)`
Convierte un archivo Excel a PDF usando win32com.
- Maneja instancia compartida de Excel
- Limpia archivos temporales
- Manejo robusto de errores

---

### `formato_excel.py` - Formateo
**Responsabilidad**: Crear y formatear archivos Excel con estilos específicos.

**Función principal**:
- `create_formatted_excel(df, config)`: Genera Excel formateado con logos, firmas, estilos

---

## 🔄 Flujo de Datos

```
Usuario interactúa con Streamlit (app.py)
    ↓
Selecciona cursos y formato de salida
    ↓
app.py llama a generar_zip_formatos()
    ↓
generador_archivos.py:
    ├─ Procesa cada curso
    ├─ Busca notas en maestro
    ├─ Genera Excel (formato_excel.py)
    ├─ Convierte a PDF (si se solicita)
    └─ Empaqueta todo en ZIP
    ↓
Retorna ZIP a app.py
    ↓
Usuario descarga el archivo
```

---

## ✅ Ventajas de esta Arquitectura

### 1. **Separación de Responsabilidades**
- UI separada de lógica de negocio
- Cada módulo tiene un propósito claro

### 2. **Mantenibilidad**
- Archivos más pequeños y manejables
- Más fácil localizar y corregir errores
- Código más legible

### 3. **Reutilización**
- `generador_archivos.py` puede usarse en otros contextos
- Funciones independientes del framework UI

### 4. **Testing**
- Más fácil probar funciones aisladas
- Mock de dependencias más simple
- Tests unitarios sin necesidad de Streamlit

### 5. **Colaboración**
- Múltiples personas pueden trabajar sin conflictos
- Cambios en UI no afectan lógica de generación

### 6. **Escalabilidad**
- Fácil agregar nuevos formatos de exportación
- Fácil modificar lógica sin tocar la UI

---

## 🚀 Uso Programático

Ahora puedes usar la lógica de generación sin Streamlit:

```python
from generador_archivos import generar_zip_formatos
import pandas as pd
import json

# Cargar configuración
with open('config_cursos.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

# Preparar datos
dnis_df = pd.DataFrame([...])
maestro = pd.ExcelFile('maestro_notas.xlsx')

# Generar archivos
zip_data, filename, warnings = generar_zip_formatos(
    dnis_procesados=dnis_df,
    selected_courses=['Curso 1', 'Curso 2'],
    maestro_excel=maestro,
    course_configs={...},
    output_format="PDF"
)

# Guardar ZIP
with open(filename, 'wb') as f:
    f.write(zip_data.getvalue())
```

---

## 🔧 Mantenimiento

### Para modificar la UI:
- Edita `app.py`
- No necesitas tocar `generador_archivos.py`

### Para modificar la lógica de generación:
- Edita `generador_archivos.py`
- Los cambios se reflejan automáticamente en `app.py`

### Para agregar un nuevo formato de exportación:
1. Modifica `convertir_excel_a_pdf()` o crea nueva función
2. Actualiza `generar_zip_formatos()` para soportar el nuevo formato
3. Agrega opción en `app.py` (radio button)

---

## 📝 Historial de Cambios

### v2.0 (12 Nov 2025)
- ✅ Refactorización completa
- ✅ Separación de UI y lógica de negocio
- ✅ Creación de `generador_archivos.py`
- ✅ Reducción de `app.py` de 1038 a 711 líneas
- ✅ Mejora en mantenibilidad y testing

### v1.0 (Anterior)
- Todo el código en `app.py` (1038 líneas)
