"""
Script de ejemplo: Uso programático de generador_archivos.py
Demuestra cómo usar el módulo sin Streamlit
"""

import pandas as pd
import json
from generador_archivos import get_nombre_completo_curso, generar_zip_formatos

def ejemplo_basico():
    """
    Ejemplo básico de uso del módulo generador_archivos
    """
    print("🔧 Ejemplo de uso programático de generador_archivos.py\n")
    
    # 1. Cargar configuración de cursos
    print("1️⃣ Cargando configuración de cursos...")
    with open('config_cursos.json', 'r', encoding='utf-8') as f:
        config_cursos = json.load(f)
    print(f"   ✅ {len(config_cursos['cursos'])} cursos cargados\n")
    
    # 2. Ejemplo de mapeo de nombres
    print("2️⃣ Probando mapeo de nombres truncados:")
    ejemplos_nombres = [
        "1_IPERC Linea Base - Identif",
        "14_IPERC Linea Base - Identi",
        "8_EPP - Equipos de Protecció"
    ]
    
    for nombre_truncado in ejemplos_nombres:
        nombre_completo = get_nombre_completo_curso(nombre_truncado, config_cursos)
        print(f"   📝 '{nombre_truncado}' → '{nombre_completo}'")
    print()
    
    # 3. Ejemplo de preparación de datos (simulado)
    print("3️⃣ Preparando datos de ejemplo...")
    dnis_ejemplo = pd.DataFrame({
        'DNI': ['12345678', '87654321', '11223344'],
        'Nombre': ['JUAN PEREZ GARCIA', 'MARIA LOPEZ SANTOS', 'CARLOS DIAZ RUIZ'],
        'Unidad': ['Unidad A', 'Unidad A', 'Unidad B']
    })
    print(f"   ✅ {len(dnis_ejemplo)} registros preparados\n")
    
    # 4. Ejemplo de generación (comentado para no ejecutar realmente)
    print("4️⃣ Para generar archivos:")
    print("""
    # Cargar maestro de notas
    maestro_excel = pd.ExcelFile('maestro_notas.xlsx')
    
    # Configurar cursos
    course_configs = {
        'Curso 1': {
            'Nombre Curso': 'Nombre Completo del Curso',
            'Tema/Motivo': 'Capacitación en seguridad',
            # ... más configuración
        }
    }
    
    # Generar ZIP
    zip_data, filename, warnings = generar_zip_formatos(
        dnis_procesados=dnis_ejemplo,
        selected_courses=['Curso 1', 'Curso 2'],
        maestro_excel=maestro_excel,
        course_configs=course_configs,
        output_format="PDF",
        progress_callback=lambda idx, total, curso: print(f"Procesando {idx}/{total}: {curso}")
    )
    
    # Guardar archivo
    with open(filename, 'wb') as f:
        f.write(zip_data.getvalue())
    
    print(f"✅ Archivo generado: {filename}")
    """)
    
    print("\n✨ El módulo está listo para usar de forma programática")
    print("📖 Ver ARQUITECTURA.md para más detalles")


if __name__ == "__main__":
    ejemplo_basico()
