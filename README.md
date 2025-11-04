# Generador de Formatos de Capacitación

Sistema para generar formatos de capacitación en Excel con datos personalizados para múltiples cursos.

## 📋 Características

- ✅ Gestión de hasta 25 cursos diferentes
- ✅ Configuración personalizada por curso
- ✅ Firmas únicas por capacitador
- ✅ Interfaz interactiva paso a paso
- ✅ Generación masiva de formatos en ZIP
- ✅ Edición directa de datos del personal

## 🚀 Instalación

1. Clona el repositorio:
```bash
git clone https://github.com/LuchoMiranda123/proyecto_certificado_capa.git
cd proyecto_certificado_capa
```

2. Crea un entorno virtual:
```bash
python -m venv .venv
```

3. Activa el entorno virtual:
- Windows:
  ```bash
  .venv\Scripts\activate
  ```
- Linux/Mac:
  ```bash
  source .venv/bin/activate
  ```

4. Instala las dependencias:
```bash
pip install -r requirements.txt
```

## ⚙️ Configuración de Cursos

### Archivo `config_cursos.json`

Este archivo contiene la configuración de todos los cursos. Estructura:

```json
{
  "cursos": {
    "Nombre del Curso": {
      "nombre": "Nombre del Curso",
      "tema_motivo": "Capacitación en...",
      "contenido_subtemas": "• Tema 1\n• Tema 2\n• Tema 3",
      "capacitador": "Nombre del Capacitador",
      "duracion": "02:00:00",
      "grabacion": "https://youtu.be/...",
      "firma": "firma_capacitador1.png"
    }
  },
  "configuracion_default": {
    "tema_motivo": "Capacitación en seguridad",
    "contenido_subtemas": "* Tema 1\n* Tema 2\n* Tema 3",
    "capacitador": "Jose Alvines",
    "duracion": "00:30:00",
    "grabacion": "https://youtu.be/ejemplo",
    "firma": "firma_capacitador.png"
  }
}
```

### Campos Configurables por Curso:

1. **nombre**: Nombre del curso (debe coincidir con el nombre de la hoja en el Excel maestro)
2. **tema_motivo**: Descripción del tema o motivo de la capacitación
3. **contenido_subtemas**: Lista de subtemas (usa \n para saltos de línea)
4. **capacitador**: Nombre completo del capacitador/entrenador
5. **duracion**: Duración en formato HH:MM:SS
6. **grabacion**: URL del material o grabación
7. **firma**: Nombre del archivo de firma (debe estar en `plantillas/firmas/`)

## 📁 Estructura de Archivos

```
proyecto_certificado_capa/
│
├── app.py                      # Aplicación principal Streamlit
├── formato_excel.py            # Generador de formatos Excel
├── config_cursos.json          # Configuración de cursos
├── requirements.txt            # Dependencias Python
│
├── plantillas/
│   ├── logo_liderman.png      # Logo de la empresa
│   ├── firma_capacitador.png  # Firma por defecto
│   └── firmas/                # Directorio para firmas
│       ├── firma_capacitador1.png
│       ├── firma_capacitador2.png
│       └── ...
│
├── contexto/                   # Archivos de contexto
├── docs/                       # Documentación
└── __pycache__/               # Cache de Python (ignorado en git)
```

## 📸 Gestión de Firmas

1. Coloca todas las firmas en el directorio `plantillas/firmas/`
2. Nombra los archivos de forma descriptiva (ej: `firma_jose_alvines.png`)
3. Referencia el nombre del archivo en `config_cursos.json`
4. Formatos soportados: PNG, JPG
5. Tamaño recomendado: 200x100 píxeles (se ajusta automáticamente)

## 🎯 Uso

1. Ejecuta la aplicación:
```bash
streamlit run app.py
```

2. Sigue los 5 pasos:
   - **Paso 1**: Cargar archivos base (Personal Asignado + Maestro de Notas)
   - **Paso 2**: Ingresar DNIs a procesar
   - **Paso 3**: Seleccionar cursos
   - **Paso 4**: Configurar detalles (se cargan automáticamente desde config_cursos.json)
   - **Paso 5**: Generar y descargar formatos

## 📝 Agregar Nuevos Cursos

1. Abre `config_cursos.json`
2. Agrega un nuevo curso en la sección `"cursos"`:
```json
"Mi Nuevo Curso": {
  "nombre": "Mi Nuevo Curso",
  "tema_motivo": "Descripción del curso",
  "contenido_subtemas": "• Tema A\n• Tema B\n• Tema C",
  "capacitador": "Nombre Capacitador",
  "duracion": "01:30:00",
  "grabacion": "https://youtu.be/...",
  "firma": "firma_nuevo_capacitador.png"
}
```
3. Guarda el archivo
4. En la aplicación, click en "🔄 Recargar configuración desde archivo"

## 🔧 Requisitos

- Python 3.8+
- streamlit
- pandas
- openpyxl

## 📦 Dependencias

Ver `requirements.txt` para la lista completa de dependencias.

## 🤝 Contribuir

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto es de uso interno.

## ✉️ Contacto

Lucho Miranda - [@LuchoMiranda123](https://github.com/LuchoMiranda123)

---

**Nota**: Asegúrate de que los nombres de los cursos en `config_cursos.json` coincidan exactamente con los nombres de las hojas en tu archivo Excel maestro.
