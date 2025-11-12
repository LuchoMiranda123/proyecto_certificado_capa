"""
Módulo para generar archivos de capacitación en formato Excel y PDF.
Maneja la creación de ZIPs con los formatos solicitados.
"""

import pandas as pd
import os
from io import BytesIO
import zipfile
import tempfile
import time
from formato_excel import create_formatted_excel


def get_nombre_completo_curso(nombre_truncado, config_cursos):
    """
    Mapea un nombre truncado de hoja de Excel (máx 31 caracteres) 
    al nombre completo del curso en config_cursos.json
    
    Args:
        nombre_truncado (str): Nombre truncado de la hoja de Excel
        config_cursos (dict): Configuración de cursos desde JSON
    
    Returns:
        str: Nombre completo del curso o el nombre truncado si no se encuentra
    """
    # EXTRAER el número de hoja antes de limpiar (ej: "14_IPERC..." -> 14)
    numero_hoja = None
    if nombre_truncado and nombre_truncado[0].isdigit():
        # Extraer todos los dígitos del inicio
        digitos = ''
        for char in nombre_truncado:
            if char.isdigit():
                digitos += char
            elif char == '_':
                break
            else:
                break
        if digitos:
            numero_hoja = int(digitos)
    
    # Remover números y guiones bajos del inicio (ej: "1_IPERC..." -> "IPERC...")
    nombre_limpio = nombre_truncado.lstrip('0123456789_').strip()
    
    # Normalizar: convertir a minúsculas, remover espacios extras, puntuación
    nombre_limpio_norm = ' '.join(nombre_limpio.lower().replace(':', '').replace(',', '').replace('_', ' ').split())
    
    # DETECTAR si el nombre tiene indicadores de "Parte 02", "Parte 2", etc.
    tiene_parte = 'parte' in nombre_limpio_norm or 'part' in nombre_limpio_norm
    tiene_02 = '02' in nombre_truncado or '2' in nombre_limpio_norm.split()[-1] if nombre_limpio_norm.split() else False
    
    # LÓGICA ESPECIAL: Si el número de hoja es mayor a 10 y hay múltiples cursos similares,
    # probablemente es una "Parte 02" o versión posterior
    es_probablemente_parte2 = numero_hoja is not None and numero_hoja > 10
    
    # Buscar en el JSON por coincidencia
    mejor_match = None
    candidatos = []
    
    for nombre_completo in config_cursos['cursos'].keys():
        nombre_completo_norm = ' '.join(nombre_completo.lower().replace(':', '').replace(',', '').split())
        nombre_completo_tiene_parte = 'parte' in nombre_completo_norm or 'part' in nombre_completo_norm
        
        # Método 1: El nombre limpio es el inicio del nombre completo (exacto)
        if nombre_completo_norm.startswith(nombre_limpio_norm):
            candidatos.append({
                'nombre': nombre_completo,
                'longitud': len(nombre_completo_norm),
                'tiene_parte': nombre_completo_tiene_parte,
                'prioridad': 1000
            })
        
        # Método 2: Coincidencia parcial al inicio (primeras palabras)
        palabras_limpio = nombre_limpio_norm.split()
        palabras_completo = nombre_completo_norm.split()
        
        # Contar cuántas palabras consecutivas coinciden desde el inicio
        coincidencias = 0
        for i, palabra in enumerate(palabras_limpio):
            if i < len(palabras_completo):
                if palabra in palabras_completo[i] or palabras_completo[i] in palabra:
                    coincidencias += 1
                elif any(palabra in pc or pc in palabra for pc in palabras_completo[:len(palabras_limpio)]):
                    coincidencias += 0.5
            else:
                break
        
        if coincidencias >= min(2, len(palabras_limpio) * 0.6):
            candidatos.append({
                'nombre': nombre_completo,
                'longitud': len(nombre_completo_norm),
                'tiene_parte': nombre_completo_tiene_parte,
                'prioridad': coincidencias
            })
    
    # FILTRAR candidatos según si tiene "Parte" o no
    if candidatos:
        if tiene_parte or tiene_02 or es_probablemente_parte2:
            candidatos_con_parte = [c for c in candidatos if c['tiene_parte']]
            if candidatos_con_parte:
                mejor_match = max(candidatos_con_parte, key=lambda x: (x['prioridad'], x['longitud']))['nombre']
            else:
                mejor_match = max(candidatos, key=lambda x: (x['prioridad'], x['longitud']))['nombre']
        else:
            candidatos_sin_parte = [c for c in candidatos if not c['tiene_parte']]
            if candidatos_sin_parte:
                mejor_match = min(candidatos_sin_parte, key=lambda x: x['longitud'])['nombre']
            else:
                mejor_match = min(candidatos, key=lambda x: x['longitud'])['nombre']
    
    if mejor_match:
        return mejor_match
    
    # Método 3: Buscar si el nombre limpio está contenido (menos estricto)
    matches_contenidos = []
    for nombre_completo in config_cursos['cursos'].keys():
        nombre_completo_norm = ' '.join(nombre_completo.lower().replace(':', '').replace(',', '').split())
        nombre_completo_tiene_parte = 'parte' in nombre_completo_norm
        
        if nombre_limpio_norm in nombre_completo_norm:
            matches_contenidos.append({
                'nombre': nombre_completo,
                'longitud': len(nombre_completo),
                'tiene_parte': nombre_completo_tiene_parte
            })
    
    if matches_contenidos:
        if tiene_parte or tiene_02 or es_probablemente_parte2:
            matches_con_parte = [m for m in matches_contenidos if m['tiene_parte']]
            if matches_con_parte:
                return max(matches_con_parte, key=lambda x: x['longitud'])['nombre']
        else:
            matches_sin_parte = [m for m in matches_contenidos if not m['tiene_parte']]
            if matches_sin_parte:
                return min(matches_sin_parte, key=lambda x: x['longitud'])['nombre']
        
        return max(matches_contenidos, key=lambda x: x['longitud'])['nombre']
    
    return nombre_truncado


def buscar_nota_en_maestro(dni, maestro_curso):
    """
    Busca la información de nota de un DNI en el maestro de curso.
    
    Args:
        dni (str): DNI del participante
        maestro_curso (DataFrame): DataFrame con las notas del curso
    
    Returns:
        dict: Información de la nota (NOTA, FECHA DEL EXAMEN, DURACIÓN) o None
    """
    if maestro_curso is None:
        return None
    
    # Detectar columna de DNI en maestro
    possible_dni_cols = ['DNI', 'DOCUMENTO', 'Documento', 'dni', 'documento']
    dni_col_maestro = None
    
    for col in possible_dni_cols:
        if col in maestro_curso.columns:
            dni_col_maestro = col
            break
    
    if not dni_col_maestro:
        return None
    
    # Buscar por DNI (intentar con y sin ceros a la izquierda)
    dni_sin_ceros = str(int(dni)) if dni.isdigit() else dni
    nota_row = maestro_curso[
        (maestro_curso[dni_col_maestro].astype(str) == dni) |
        (maestro_curso[dni_col_maestro].astype(str) == dni_sin_ceros) |
        (maestro_curso[dni_col_maestro].astype(str).str.zfill(8) == dni)
    ]
    
    if not nota_row.empty:
        return nota_row.iloc[0]
    
    return None


def procesar_curso(curso, dnis_procesados, maestro_excel, course_config):
    """
    Procesa un curso individual: extrae datos y genera el DataFrame.
    
    Args:
        curso (str): Nombre del curso (hoja de Excel)
        dnis_procesados (DataFrame): DataFrame con DNIs y datos del personal
        maestro_excel (ExcelFile): Archivo Excel con las notas
        course_config (dict): Configuración específica del curso
    
    Returns:
        tuple: (DataFrame del curso, nombre completo del archivo, prefijo numérico)
    """
    # Cargar la hoja del curso
    try:
        maestro_curso = pd.read_excel(maestro_excel, sheet_name=curso)
    except Exception as e:
        print(f"⚠️ No se pudo cargar datos de {curso}: {e}")
        maestro_curso = None
    
    # Crear DataFrame para este curso
    curso_data = []
    
    for idx, row in dnis_procesados.iterrows():
        dni = str(row['DNI'])
        nota_info = buscar_nota_en_maestro(dni, maestro_curso)
        
        curso_data.append({
            'N°': idx + 1,
            'Apellidos y Nombres': row['Nombre'],
            'DNI': dni,
            'Unidad (Cliente)': row['Unidad'],
            'Nota': nota_info['NOTA'] if nota_info is not None else '',
            'Fecha Examen': nota_info['FECHA DEL EXAMEN'] if nota_info is not None else '',
            'Hora Conexión': nota_info['DURACIÓN'] if nota_info is not None else ''
        })
    
    df_curso = pd.DataFrame(curso_data)
    
    # Obtener nombre completo del curso
    nombre_completo_archivo = course_config['Nombre Curso']
    
    # Extraer el número de hoja original del nombre truncado (ej: "14_IPERC..." -> "14_")
    prefijo_numero = ""
    if curso and curso[0].isdigit():
        for i, char in enumerate(curso):
            if char.isdigit() or char == '_':
                prefijo_numero += char
            else:
                break
    
    return df_curso, nombre_completo_archivo, prefijo_numero


def convertir_excel_a_pdf(excel_data, base_filename, excel_app):
    """
    Convierte un archivo Excel a PDF usando win32com.
    
    Args:
        excel_data (bytes): Datos del archivo Excel
        base_filename (str): Nombre base del archivo (sin extensión)
        excel_app: Instancia de Excel COM (win32com)
    
    Returns:
        bytes: Datos del PDF generado o None si falla
    """
    tmp_excel_path = None
    tmp_pdf_path = None
    wb = None
    pdf_data = None
    
    try:
        # Crear archivo temporal para el Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
            tmp_excel.write(excel_data)
            tmp_excel_path = tmp_excel.name
        
        # Crear archivo temporal para el PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
            tmp_pdf_path = tmp_pdf.name
        
        # Pequeño delay para asegurar que Excel esté listo
        time.sleep(0.5)
        
        # Abrir el workbook y exportar a PDF
        wb = excel_app.Workbooks.Open(os.path.abspath(tmp_excel_path))
        wb.ExportAsFixedFormat(0, os.path.abspath(tmp_pdf_path))  # 0 = xlTypePDF
        wb.Close(False)
        wb = None
        
        # Pequeño delay después de cerrar
        time.sleep(0.3)
        
        # Leer el PDF generado
        if os.path.exists(tmp_pdf_path):
            with open(tmp_pdf_path, 'rb') as pdf_file:
                pdf_data = pdf_file.read()
        
    except Exception as e:
        print(f"⚠️ Error al convertir {base_filename} a PDF: {e}")
        # Cerrar workbook si quedó abierto
        try:
            if wb is not None:
                wb.Close(False)
        except:
            pass
    
    finally:
        # Limpiar archivos temporales
        try:
            if tmp_excel_path and os.path.exists(tmp_excel_path):
                time.sleep(0.2)
                os.unlink(tmp_excel_path)
        except:
            pass
        try:
            if tmp_pdf_path and os.path.exists(tmp_pdf_path):
                time.sleep(0.2)
                os.unlink(tmp_pdf_path)
        except:
            pass
    
    return pdf_data


def generar_zip_formatos(dnis_procesados, selected_courses, maestro_excel, 
                         course_configs, output_format, progress_callback=None):
    """
    Genera un archivo ZIP con los formatos de capacitación en Excel y/o PDF.
    
    Args:
        dnis_procesados (DataFrame): DataFrame con DNIs procesados
        selected_courses (list): Lista de cursos seleccionados
        maestro_excel (ExcelFile): Archivo Excel con notas de todos los cursos
        course_configs (dict): Configuraciones de cada curso
        output_format (str): "Excel (.xlsx)", "PDF", o "Ambos (Excel + PDF)"
        progress_callback (callable): Función para reportar progreso (opcional)
    
    Returns:
        tuple: (BytesIO con el ZIP, str con el nombre sugerido del archivo, list de warnings)
    """
    # Determinar qué formatos generar
    generar_excel = output_format in ["Excel (.xlsx)", "Ambos (Excel + PDF)"]
    generar_pdf = output_format in ["PDF", "Ambos (Excel + PDF)"]
    
    warnings = []
    excel_app = None
    
    # Inicializar Excel COM una sola vez si se necesita PDF
    if generar_pdf:
        try:
            import win32com.client
            import pythoncom
            
            pythoncom.CoInitialize()
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            excel_app.ScreenUpdating = False
        except ImportError:
            warnings.append("⚠️ pywin32 no instalado. Los PDFs se generarán como Excel.")
            generar_pdf = False
            generar_excel = True
        except Exception as e:
            warnings.append(f"⚠️ No se pudo inicializar Excel: {e}")
            generar_pdf = False
            generar_excel = True
    
    zip_buffer = BytesIO()
    
    try:
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for idx, curso in enumerate(selected_courses, 1):
                # Reportar progreso si hay callback
                if progress_callback:
                    progress_callback(idx, len(selected_courses), curso)
                
                # Procesar el curso
                df_curso, nombre_completo_archivo, prefijo_numero = procesar_curso(
                    curso, dnis_procesados, maestro_excel, course_configs[curso]
                )
                
                # Generar Excel
                excel_data = create_formatted_excel(df_curso, course_configs[curso])
                
                if excel_data:
                    # Nombre del archivo
                    unidad = df_curso['Unidad (Cliente)'].iloc[0] if not df_curso.empty else 'Sin_Unidad'
                    
                    if prefijo_numero:
                        base_filename = f"{prefijo_numero}{nombre_completo_archivo} - {unidad}"
                    else:
                        base_filename = f"{nombre_completo_archivo} - {unidad}"
                    
                    # Agregar Excel al ZIP si se solicitó
                    if generar_excel:
                        file_name_excel = f"{base_filename}.xlsx"
                        zip_file.writestr(file_name_excel, excel_data)
                    
                    # Generar y agregar PDF al ZIP si se solicitó
                    if generar_pdf and excel_app is not None:
                        pdf_data = convertir_excel_a_pdf(excel_data, base_filename, excel_app)
                        
                        if pdf_data:
                            file_name_pdf = f"{base_filename}.pdf"
                            zip_file.writestr(file_name_pdf, pdf_data)
                        else:
                            warnings.append(f"⚠️ No se generó el PDF para {base_filename}")
                            # Si solo se pidió PDF y falló, agregar Excel como respaldo
                            if not generar_excel:
                                file_name_excel = f"{base_filename}.xlsx"
                                zip_file.writestr(file_name_excel, excel_data)
        
        zip_buffer.seek(0)
    
    finally:
        # Cerrar Excel y liberar recursos COM
        if excel_app is not None:
            try:
                excel_app.Quit()
                excel_app = None
                time.sleep(1)  # Dar tiempo para que Excel se cierre
            except:
                pass
            
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except:
                pass
    
    # Determinar nombre del archivo ZIP
    if output_format == "Excel (.xlsx)":
        zip_filename = "Formatos_Capacitacion_Excel.zip"
    elif output_format == "PDF":
        zip_filename = "Formatos_Capacitacion_PDF.zip"
    else:  # Ambos
        zip_filename = "Formatos_Capacitacion_Excel_PDF.zip"
    
    return zip_buffer, zip_filename, warnings
