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
from datetime import datetime, timedelta
from formato_excel import create_formatted_excel
import unicodedata
import difflib
import re


def quitar_acentos(texto):
    """
    Quita acentos y caracteres especiales de un texto.
    
    Args:
        texto (str): Texto a normalizar
    
    Returns:
        str: Texto sin acentos
    """
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')


def sanitizar_nombre_archivo(nombre):
    """
    Sanitiza un nombre para que sea válido en nombres de archivo y hojas de Excel.
    
    Args:
        nombre (str): Nombre a sanitizar
    
    Returns:
        str: Nombre sanitizado
    """
    # Caracteres inválidos para nombres de archivo en Windows: < > : " / \ | ? *
    # Caracteres inválidos para nombres de hoja en Excel: : \ / ? * [ ]
    # Reemplazarlos por guión para mantener legibilidad
    nombre_limpio = nombre
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*', '[', ']']
    for char in invalid_chars:
        nombre_limpio = nombre_limpio.replace(char, '-')
    
    # Limpiar espacios múltiples y guiones múltiples
    nombre_limpio = re.sub(r'\s+', ' ', nombre_limpio)
    nombre_limpio = re.sub(r'-+', '-', nombre_limpio)
    
    # Limpiar espacios alrededor de guiones: " - " -> " - " (ya está bien), "- " -> "-", " -" -> "-"
    nombre_limpio = re.sub(r'\s*-\s*', ' - ', nombre_limpio)
    nombre_limpio = re.sub(r'\s+-\s+', ' - ', nombre_limpio)  # Normalizar espacios alrededor de guiones
    
    # Quitar espacios y guiones al inicio y final
    nombre_limpio = nombre_limpio.strip(' -')
    
    return nombre_limpio


def sumar_tiempos(tiempo1, tiempo2):
    """
    Suma dos tiempos en formato HH:MM:SS o H:MM:SS.
    
    Args:
        tiempo1 (str): Primer tiempo (ej: "01:30:00" o "1:30:00")
        tiempo2 (str): Segundo tiempo (ej: "00:45:00" o "45:00")
    
    Returns:
        str: Suma de tiempos en formato HH:MM:SS
    """
    try:
        # Si alguno de los tiempos está vacío, retornar el otro
        if not tiempo1 or pd.isna(tiempo1):
            return str(tiempo2) if tiempo2 and not pd.isna(tiempo2) else "00:00:00"
        if not tiempo2 or pd.isna(tiempo2):
            return str(tiempo1) if tiempo1 and not pd.isna(tiempo1) else "00:00:00"
        
        # Convertir a string por si vienen como otros tipos
        tiempo1 = str(tiempo1).strip()
        tiempo2 = str(tiempo2).strip()
        
        # Parsear tiempo1
        partes1 = tiempo1.split(':')
        if len(partes1) == 3:
            h1, m1, s1 = map(int, partes1)
        elif len(partes1) == 2:
            h1 = 0
            m1, s1 = map(int, partes1)
        else:
            return "00:00:00"
        
        # Parsear tiempo2
        partes2 = tiempo2.split(':')
        if len(partes2) == 3:
            h2, m2, s2 = map(int, partes2)
        elif len(partes2) == 2:
            h2 = 0
            m2, s2 = map(int, partes2)
        else:
            return tiempo1  # Si tiempo2 es inválido, retornar tiempo1
        
        # Crear objetos timedelta y sumar
        delta1 = timedelta(hours=h1, minutes=m1, seconds=s1)
        delta2 = timedelta(hours=h2, minutes=m2, seconds=s2)
        suma = delta1 + delta2
        
        # Convertir de vuelta a formato HH:MM:SS
        total_seconds = int(suma.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    
    except Exception as e:
        print(f"⚠️ Error al sumar tiempos '{tiempo1}' + '{tiempo2}': {e}")
        # En caso de error, retornar el primer tiempo o 00:00:00
        return str(tiempo1) if tiempo1 and not pd.isna(tiempo1) else "00:00:00"


def get_nombre_completo_curso(nombre_truncado, courses_dict):
    """
    Mapea un nombre truncado de hoja de Excel (máx 31 caracteres) 
    al nombre completo del curso en courses_dict con logging detallado.
    
    Args:
        nombre_truncado (str): Nombre truncado de la hoja de Excel
        courses_dict (dict): Diccionario de cursos con nombres completos como keys
    
    Returns:
        str: Nombre completo del curso o el nombre truncado si no se encuentra
    """
    # DICCIONARIO DE ALIAS PARA CASOS ESPECIALES DE TRUNCAMIENTO
    ALIAS_MAP = {
        'eventos indeseables y disturb': 'Eventos indeseables, perturbadores y lugares hostiles',
        'eventos indeseables y disturbar': 'Eventos indeseables, perturbadores y lugares hostiles',
        'eventos indeseables, disturb': 'Eventos indeseables, perturbadores y lugares hostiles',
        'protocolos y procedimientos de': 'Protocolos y procedimientos de agente parking o valet parking',
        'gestion de residuos solidos imp': 'Gestión de residuos sólidos, impactos ambientales y responsabilidad social empresarial',
        'gestion de residuos solidos, imp': 'Gestión de residuos sólidos, impactos ambientales y responsabilidad social empresarial',
        'armas: conocimiento y manipulac': 'Armas: conocimiento y manipulación',
        'armas: conocimiento y manipula': 'Armas: conocimiento y manipulación',
        'armas_ conocimiento y manipu': 'Armas: conocimiento y manipulación',
        'armas_ conocimiento y manipul': 'Armas: conocimiento y manipulación',
    }
    
    print(f"\n[MATCH] === Iniciando matching para: '{nombre_truncado}' ===")
    
    # 1. PREPROCESAMIENTO MÍNIMO
    numero_hoja = None
    if nombre_truncado and nombre_truncado[0].isdigit():
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
    
    # Remover solo números y guiones bajos del inicio
    nombre_limpio = nombre_truncado.lstrip('0123456789_').strip()
    print(f"[MATCH] Nombre limpio (sin prefijo numérico): '{nombre_limpio}'")
    
    # MÉTODO 0: CHECK DE ALIAS PRIMERO (antes de cualquier otro método)
    nombre_limpio_lower = nombre_limpio.lower()
    for alias, nombre_completo in ALIAS_MAP.items():
        if nombre_limpio_lower.startswith(alias) or alias in nombre_limpio_lower:
            if nombre_completo in courses_dict:
                print(f"[MATCH] OK ALIAS encontrado: '{alias}' -> '{nombre_completo}'")
                print(f"[MATCH] === Resultado: '{nombre_completo}' (Metodo: ALIAS) ===\n")
                return nombre_completo
    
    # Detectar variantes
    tiene_parte_02 = 'parte 02' in nombre_limpio.lower() or '(parte 02)' in nombre_limpio.lower()
    tiene_2025 = '2025' in nombre_limpio
    es_hoja_alta = numero_hoja is not None and numero_hoja >= 14  # Cambiado de >15 a >=14
    print(f"[MATCH] Indicadores: Parte02={tiene_parte_02}, 2025={tiene_2025}, Hoja#{numero_hoja}, HojaAlta={es_hoja_alta}")
    
    # 2. MÉTODO 1: MATCH EXACTO POR INICIO (31 caracteres)
    print(f"[MATCH] Método 1: Match exacto por inicio...")
    candidatos_exactos = []
    for nombre_completo in courses_dict.keys():
        # Comparar los primeros N caracteres (longitud del truncado)
        if nombre_completo.startswith(nombre_limpio):
            candidatos_exactos.append(nombre_completo)
            print(f"[MATCH] Candidato exacto: '{nombre_completo}'")
    
    if candidatos_exactos:
        # Si es hoja alta y hay versión "Parte 02", priorizarla
        if es_hoja_alta:
            for candidato in candidatos_exactos:
                if 'parte 02' in candidato.lower() or '(parte 02)' in candidato.lower():
                    print(f"[MATCH] OK Match exacto (Parte 02 por hoja alta): '{candidato}'")
                    print(f"[MATCH] === Resultado: '{candidato}' (Método: EXACTO+HOJA_ALTA) ===\n")
                    return candidato
        
        # Si tiene indicador 2025, priorizar versión 2025
        if tiene_2025:
            for candidato in candidatos_exactos:
                if '2025' in candidato:
                    print(f"[MATCH] OK Match exacto (2025): '{candidato}'")
                    print(f"[MATCH] === Resultado: '{candidato}' (Método: EXACTO+2025) ===\n")
                    return candidato
        
        # Si NO tiene indicadores de variante, priorizar versión base (más corta)
        if not tiene_parte_02 and not tiene_2025 and not es_hoja_alta:
            candidatos_base = [c for c in candidatos_exactos if 'parte 02' not in c.lower() and '2025' not in c]
            if candidatos_base:
                mejor = min(candidatos_base, key=len)
                print(f"[MATCH] OK Match exacto (version base): '{mejor}'")
                print(f"[MATCH] === Resultado: '{mejor}' (Método: EXACTO) ===\n")
                return mejor
        
        # Devolver el primero encontrado
        mejor = candidatos_exactos[0]
        print(f"[MATCH] OK Match exacto encontrado: '{mejor}'")
        print(f"[MATCH] === Resultado: '{mejor}' (Método: EXACTO) ===\n")
        return mejor
    
    # 3. MÉTODO 2: MATCH POR PALABRAS CLAVE ÚNICAS
    print(f"[MATCH] Método 2: Match por palabras clave...")
    palabras_truncado = [p.lower() for p in nombre_limpio.split() if len(p) > 2]
    print(f"[MATCH] Palabras clave extraídas: {palabras_truncado}")
    
    candidatos_palabras_clave = []
    for nombre_completo in courses_dict.keys():
        palabras_completo = [p.lower() for p in nombre_completo.split() if len(p) > 2]
        
        # Contar cuántas palabras del truncado están en el completo EN EL MISMO ORDEN
        palabras_encontradas = 0
        idx_anterior = -1
        for palabra_trunc in palabras_truncado:
            for idx, palabra_comp in enumerate(palabras_completo):
                if idx > idx_anterior and palabra_trunc in palabra_comp:
                    palabras_encontradas += 1
                    idx_anterior = idx
                    break
        
        porcentaje_match = palabras_encontradas / len(palabras_truncado) if palabras_truncado else 0
        
        # Requiere al menos 80% de coincidencia de palabras o mínimo 3 palabras
        if palabras_encontradas >= 3 or porcentaje_match >= 0.8:
            candidatos_palabras_clave.append({
                'nombre': nombre_completo,
                'palabras_match': palabras_encontradas,
                'porcentaje': porcentaje_match,
                'longitud': len(nombre_completo)
            })
    
    if candidatos_palabras_clave:
        print(f"[MATCH] Candidatos por palabras clave: {len(candidatos_palabras_clave)}")
        for c in candidatos_palabras_clave:
            print(f"[MATCH]   - '{c['nombre']}' (match: {c['palabras_match']}/{len(palabras_truncado)}, {c['porcentaje']:.0%})")
    
    # 4. MÉTODO 3: FILTRADO POR VARIANTES
    if candidatos_palabras_clave:
        print(f"[MATCH] Método 3: Filtrado por variantes...")
        
        # Separar candidatos con y sin variantes
        candidatos_con_parte02 = [c for c in candidatos_palabras_clave if 'parte 02' in c['nombre'].lower() or '(parte 02)' in c['nombre'].lower()]
        candidatos_con_2025 = [c for c in candidatos_palabras_clave if '2025' in c['nombre']]
        candidatos_base = [c for c in candidatos_palabras_clave if 'parte 02' not in c['nombre'].lower() and '2025' not in c['nombre']]
        
        # Decidir según indicadores
        if tiene_parte_02 or es_hoja_alta:
            if candidatos_con_parte02:
                mejor = max(candidatos_con_parte02, key=lambda x: (x['palabras_match'], -x['longitud']))
                print(f"[MATCH] OK Seleccionado (con Parte 02): '{mejor['nombre']}'")
                print(f"[MATCH] === Resultado: '{mejor['nombre']}' (Método: PALABRAS_CLAVE+VARIANTE) ===\n")
                return mejor['nombre']
        
        if tiene_2025:
            if candidatos_con_2025:
                mejor = max(candidatos_con_2025, key=lambda x: (x['palabras_match'], -x['longitud']))
                print(f"[MATCH] OK Seleccionado (con 2025): '{mejor['nombre']}'")
                print(f"[MATCH] === Resultado: '{mejor['nombre']}' (Método: PALABRAS_CLAVE+VARIANTE) ===\n")
                return mejor['nombre']
        
        # Si NO tiene indicadores de variante, priorizar versiones base
        if not tiene_parte_02 and not tiene_2025 and not es_hoja_alta:
            if candidatos_base:
                mejor = max(candidatos_base, key=lambda x: (x['palabras_match'], -x['longitud']))
                print(f"[MATCH] OK Seleccionado (version base): '{mejor['nombre']}'")
                print(f"[MATCH] === Resultado: '{mejor['nombre']}' (Método: PALABRAS_CLAVE) ===\n")
                return mejor['nombre']
        
        # Si no hay candidatos base pero sí otros, tomar el mejor
        mejor = max(candidatos_palabras_clave, key=lambda x: (x['palabras_match'], -x['longitud']))
        print(f"[MATCH] OK Seleccionado (mejor disponible): '{mejor['nombre']}'")
        print(f"[MATCH] === Resultado: '{mejor['nombre']}' (Método: PALABRAS_CLAVE) ===\n")
        return mejor['nombre']
    
    # 5. MÉTODO 4: FUZZY MATCHING (ÚLTIMO RECURSO)
    print(f"[MATCH] Método 4: Fuzzy matching (último recurso)...")
    candidatos_fuzzy = []
    for nombre_completo in courses_dict.keys():
        # Comparar el truncado con el inicio del nombre completo
        longitud_comparacion = min(len(nombre_limpio), len(nombre_completo))
        inicio_completo = nombre_completo[:longitud_comparacion]
        
        ratio = difflib.SequenceMatcher(None, nombre_limpio.lower(), inicio_completo.lower()).ratio()
        
        if ratio >= 0.90:  # Umbral MUY alto
            candidatos_fuzzy.append({
                'nombre': nombre_completo,
                'ratio': ratio,
                'longitud': len(nombre_completo)
            })
    
    if candidatos_fuzzy:
        print(f"[MATCH] Candidatos fuzzy (≥90%): {len(candidatos_fuzzy)}")
        for c in candidatos_fuzzy:
            print(f"[MATCH]   - '{c['nombre']}' (similitud: {c['ratio']:.0%})")
        
        mejor = max(candidatos_fuzzy, key=lambda x: (x['ratio'], -x['longitud']))
        print(f"[MATCH] OK Seleccionado (fuzzy): '{mejor['nombre']}'")
        print(f"[MATCH] === Resultado: '{mejor['nombre']}' (Método: FUZZY) ===\n")
        return mejor['nombre']
    
    # 6. NO SE ENCONTRÓ MATCH
    print(f"[MATCH] WARNING: No se encontro match para '{nombre_truncado}'")
    print(f"[MATCH] === Resultado: '{nombre_truncado}' (SIN MATCH) ===\n")
    return nombre_truncado


def test_matching():
    """
    Prueba los casos problemáticos reportados.
    Ejecutar esta función para validar el matching.
    """
    print("\n" + "="*80)
    print("INICIANDO TESTS DE MATCHING")
    print("="*80)
    
    # Simular diccionario de cursos
    courses_dict = {
        "IPERC, mapa de riesgos y procedimientos PETS": {},
        "IPERC, mapa de riesgos y procedimientos PETS (Parte 02)": {},
        "Salud ocupacional y estilo de vida saludable": {},
        "Seguridad y prevención en el puesto de trabajo": {},
        "Prevención de delitos de comercio internacional": {},
        "Personas y vehículos sospechosos": {},
        "Legislación y seguridad privada": {},
        "Normas y procedimientos de seguridad": {},
        "Fundamentos de SGI - 2025": {},
        "Fundamentos del Sistema Integrado de Gestión": {},
        "Eventos indeseables, perturbadores y lugares hostiles": {},
    }
    
    test_cases = [
        ("1_IPERC, mapa de riesgos y pr", "IPERC, mapa de riesgos y procedimientos PETS"),
        ("3_Salud ocupacional y estilo de", "Salud ocupacional y estilo de vida saludable"),
        ("2_Seguridad y prevención en el", "Seguridad y prevención en el puesto de trabajo"),
        ("14_IPERC, mapa de riesgos y pro", "IPERC, mapa de riesgos y procedimientos PETS (Parte 02)"),
        ("Prevención de delitos de comer", "Prevención de delitos de comercio internacional"),
        ("Personas y vehículos sospechoso", "Personas y vehículos sospechosos"),
        ("Fundamentos de SGI - 2025", "Fundamentos de SGI - 2025"),
        ("8_Eventos Indeseables y disturb", "Eventos indeseables, perturbadores y lugares hostiles"),
    ]
    
    resultados = []
    for truncado, esperado in test_cases:
        resultado = get_nombre_completo_curso(truncado, courses_dict)
        es_correcto = resultado == esperado
        resultados.append({
            'truncado': truncado,
            'esperado': esperado,
            'obtenido': resultado,
            'correcto': es_correcto
        })
    
    # Resumen
    print("\n" + "="*80)
    print("RESUMEN DE TESTS")
    print("="*80)
    correctos = sum(1 for r in resultados if r['correcto'])
    total = len(resultados)
    
    for r in resultados:
        status = "[OK] PASS" if r['correcto'] else "[X] FAIL"
        print(f"{status} | Truncado: '{r['truncado']}'")
        print(f"     | Esperado: '{r['esperado']}'")
        print(f"     | Obtenido: '{r['obtenido']}'")
        print()
    
    print(f"Resultado: {correctos}/{total} tests pasados ({correctos/total*100:.0f}%)")
    print("="*80 + "\n")
    
    return correctos == total


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


def procesar_curso(curso, dnis_procesados, maestro_excel, config_cursos):
    """
    Procesa un curso individual: extrae datos y genera el DataFrame.
    
    Args:
        curso (str): Nombre del curso (hoja de Excel)
        dnis_procesados (DataFrame): DataFrame con DNIs y datos del personal
        maestro_excel (ExcelFile): Archivo Excel con las notas
        config_cursos (dict): Configuración de cursos desde JSON
    
    Returns:
        tuple: (DataFrame del curso, nombre completo del archivo, prefijo numérico)
    """
    print(f"\n[PROCESAR] Procesando curso: '{curso}'")
    print(f"[PROCESAR] Cursos disponibles en config_cursos: {list(config_cursos.keys())[:5]}...")
    
    # Obtener la configuración del curso usando el nombre truncado como clave
    course_config = config_cursos.get(curso, {})
    
    print(f"[PROCESAR] Configuración encontrada: {'SI' if course_config else 'NO'}")
    
    # Si no hay configuración, crear una por defecto
    if not course_config:
        print(f"[PROCESAR] ADVERTENCIA: No se encontró configuración para '{curso}'")
        course_config = {
            'Nombre Curso': curso,
            'Duracion': '00:00:00',
            'Tema/Motivo': 'Curso no encontrado',
            'Contenido/ Sub Temas': 'Contenido no disponible',
            'Capacitador/Entrenador': 'Desconocido',
            'Grabacion/ Material': '',
            'Firma': ''
        }
    
    # Obtener el nombre completo del archivo
    nombre_completo_archivo = course_config.get('Nombre Curso', curso)
    print(f"[PROCESAR] Nombre completo para archivo: '{nombre_completo_archivo}'")
    
    # Cargar la hoja del curso
    try:
        print(f"[PROCESAR] Intentando cargar hoja '{curso}' del maestro...")
        maestro_curso = pd.read_excel(maestro_excel, sheet_name=curso)
        print(f"[PROCESAR] Hoja cargada exitosamente: {len(maestro_curso)} registros")
    except Exception as e:
        print(f"[PROCESAR] ERROR: No se pudo cargar datos de '{curso}': {e}")
        print(f"[PROCESAR] Hojas disponibles en maestro: {maestro_excel.sheet_names[:10]}...")
        maestro_curso = None
    
    # Obtener la duración del video desde la configuración del curso
    duracion_video = course_config.get('Duracion', '00:00:00')
    
    # Crear DataFrame para este curso
    curso_data = []
    
    for idx, row in dnis_procesados.iterrows():
        dni = str(row['DNI'])
        
        # Buscar nota específica de este curso en el maestro
        nota_info = buscar_nota_en_maestro(dni, maestro_curso)
        
        # Obtener el tiempo de conexión del maestro
        tiempo_conexion = nota_info['DURACIÓN'] if nota_info is not None else ''
        
        # Sumar el tiempo de conexión con la duración del video
        tiempo_total = sumar_tiempos(tiempo_conexion, duracion_video)
        
        curso_data.append({
            'N°': idx + 1,
            'Apellidos y Nombres': row['Nombre'],
            'DNI': dni,
            'Unidad (Cliente)': row['Unidad'],
            'Nota': nota_info['NOTA'] if nota_info is not None else '',
            'Fecha Examen': nota_info['FECHA DEL EXAMEN'] if nota_info is not None else '',
            'Hora Conexión': tiempo_total
        })
    
    df_curso = pd.DataFrame(curso_data)
    
    # Extraer el número de hoja original del nombre truncado (ej: "14_IPERC..." -> "14_")
    prefijo_numero = ""
    if curso and curso[0].isdigit():
        for i, char in enumerate(curso):
            if char.isdigit() or char == '_':
                prefijo_numero += char
            else:
                break
    
    return df_curso, nombre_completo_archivo, prefijo_numero, course_config


def convertir_excel_a_pdf(excel_data, base_filename, excel_app, max_intentos=3):
    """
    Convierte un archivo Excel a PDF usando win32com con reintentos.
    
    Args:
        excel_data (bytes): Datos del archivo Excel
        base_filename (str): Nombre base del archivo (sin extensión)
        excel_app: Instancia de Excel COM (win32com)
        max_intentos (int): Número máximo de reintentos
    
    Returns:
        bytes: Datos del PDF generado o None si falla
    """
    tmp_excel_path = None
    tmp_pdf_path = None
    wb = None
    pdf_data = None
    
    for intento in range(1, max_intentos + 1):
        try:
            print(f"[PDF] Intento {intento}/{max_intentos} para '{base_filename}'")
            
            # Crear archivo temporal para el Excel
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
                tmp_excel.write(excel_data)
                tmp_excel_path = tmp_excel.name
            
            print(f"[PDF] Excel temporal creado: {tmp_excel_path}")
            
            # Crear archivo temporal para el PDF
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
                tmp_pdf_path = tmp_pdf.name
            
            print(f"[PDF] PDF temporal preparado: {tmp_pdf_path}")
            
            # Delay progresivo según el intento
            delay = 0.5 * intento
            time.sleep(delay)
        
            # Abrir el workbook y configurar para exportar a PDF en A4
            print(f"[PDF] Abriendo workbook...")
            wb = excel_app.Workbooks.Open(os.path.abspath(tmp_excel_path))
            print(f"[PDF] Workbook abierto, configurando hojas...")
            
            # Configurar cada hoja para impresión en A4: columnas en 1 página, filas en múltiples
            for ws in wb.Worksheets:
                ws.PageSetup.PaperSize = 9  # 9 = A4 (210 x 297 mm)
                ws.PageSetup.Orientation = 1  # 1 = Portrait (vertical)
                ws.PageSetup.FitToPagesWide = 1  # Ajustar ANCHO a 1 página (todas las columnas)
                ws.PageSetup.FitToPagesTall = False  # NO ajustar alto (permitir múltiples páginas)
                ws.PageSetup.Zoom = False  # Desactivar zoom manual (usar FitToPages)
                ws.PageSetup.CenterHorizontally = False  # No centrar horizontalmente
                ws.PageSetup.CenterVertically = False  # No centrar verticalmente
                
                # Configurar márgenes pequeños (en pulgadas)
                ws.PageSetup.LeftMargin = excel_app.InchesToPoints(0.25)
                ws.PageSetup.RightMargin = excel_app.InchesToPoints(0.25)
                ws.PageSetup.TopMargin = excel_app.InchesToPoints(0.25)
                ws.PageSetup.BottomMargin = excel_app.InchesToPoints(0.25)
                ws.PageSetup.HeaderMargin = excel_app.InchesToPoints(0.1)
                ws.PageSetup.FooterMargin = excel_app.InchesToPoints(0.1)
            
            print(f"[PDF] Exportando a PDF...")
            # Exportar a PDF con la configuración establecida
            wb.ExportAsFixedFormat(0, os.path.abspath(tmp_pdf_path))  # 0 = xlTypePDF
            print(f"[PDF] Exportación completada")
            
            wb.Close(False)
            wb = None
            
            # Delay más largo después de cerrar para asegurar liberación de archivos
            time.sleep(1.0)
            
            # Verificar que el PDF existe y tiene contenido
            if os.path.exists(tmp_pdf_path):
                file_size = os.path.getsize(tmp_pdf_path)
                print(f"[PDF] Archivo PDF generado: {file_size} bytes")
                
                if file_size > 0:
                    with open(tmp_pdf_path, 'rb') as pdf_file:
                        pdf_data = pdf_file.read()
                    print(f"[PDF] ✓ PDF convertido exitosamente ({len(pdf_data)} bytes)")
                    break  # Salir del bucle de reintentos si fue exitoso
                else:
                    print(f"[PDF] ✗ Archivo PDF vacío en intento {intento}")
            else:
                print(f"[PDF] ✗ Archivo PDF no encontrado en intento {intento}")
            
        except Exception as e:
            print(f"[PDF] ✗ Error en intento {intento}/{max_intentos}: {type(e).__name__}: {e}")
            
            # Cerrar workbook si quedó abierto
            try:
                if wb is not None:
                    wb.Close(False)
                    wb = None
            except:
                pass
            
            # Si no es el último intento, esperar antes de reintentar
            if intento < max_intentos:
                espera = 2 * intento
                print(f"[PDF] Esperando {espera}s antes del siguiente intento...")
                time.sleep(espera)
            else:
                print(f"[PDF] ✗ Fallo definitivo después de {max_intentos} intentos")
    
        finally:
            # Limpiar archivos temporales con reintentos
            for archivo, nombre in [(tmp_excel_path, 'Excel'), (tmp_pdf_path, 'PDF')]:
                if archivo and os.path.exists(archivo):
                    for i in range(3):
                        try:
                            time.sleep(0.5)
                            os.unlink(archivo)
                            print(f"[PDF] Archivo temporal {nombre} eliminado")
                            break
                        except Exception as e:
                            if i == 2:
                                print(f"[PDF] No se pudo eliminar archivo temporal {nombre}: {e}")
    
    if pdf_data is None:
        print(f"[PDF] ✗ NO se pudo generar PDF para '{base_filename}' después de todos los intentos")
    
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
                df_curso, nombre_completo_archivo, prefijo_numero, course_config = procesar_curso(
                    curso, dnis_procesados, maestro_excel, course_configs
                )
                
                # Sanitizar el nombre del curso ANTES de generar el Excel
                nombre_sanitizado = sanitizar_nombre_archivo(nombre_completo_archivo)
                
                # Actualizar la configuración con el nombre sanitizado
                course_config_sanitizado = course_config.copy()
                course_config_sanitizado['Nombre Curso'] = nombre_sanitizado
                
                # Logging detallado
                print(f"[DEBUG] Procesando curso: '{curso}'")
                print(f"[DEBUG] Nombre original: '{nombre_completo_archivo}'")
                print(f"[DEBUG] Nombre sanitizado: '{nombre_sanitizado}'")
                print(f"[DEBUG] Config encontrada: {course_config.get('Nombre Curso', 'NO ENCONTRADO')}")
                
                # Generar Excel con el nombre sanitizado
                excel_data = create_formatted_excel(df_curso, course_config_sanitizado)
                
                if excel_data:
                    # Nombre del archivo
                    unidad = df_curso['Unidad (Cliente)'].iloc[0] if not df_curso.empty else 'Sin_Unidad'
                    
                    # Usar el nombre completo sanitizado
                    base_filename = sanitizar_nombre_archivo(f"{nombre_completo_archivo} - {unidad}")
                    
                    print(f"[DEBUG] Base filename final: '{base_filename}'")
                    
                    # Agregar Excel al ZIP si se solicitó
                    if generar_excel:
                        file_name_excel = f"{base_filename}.xlsx"
                        print(f"[DEBUG] Agregando Excel al ZIP: '{file_name_excel}'")
                        zip_file.writestr(file_name_excel, excel_data)
                    
                    # Generar y agregar PDF al ZIP si se solicitó
                    if generar_pdf and excel_app is not None:
                        print(f"\n{'='*70}")
                        print(f"[DEBUG] === GENERANDO PDF {idx}/{len(selected_courses)} ===")
                        print(f"[DEBUG] Archivo: '{base_filename}'")
                        print(f"{'='*70}")
                        
                        pdf_data = convertir_excel_a_pdf(excel_data, base_filename, excel_app, max_intentos=3)
                        
                        if pdf_data and len(pdf_data) > 0:
                            file_name_pdf = f"{base_filename}.pdf"
                            print(f"[DEBUG] ✓ PDF generado exitosamente: '{file_name_pdf}' ({len(pdf_data)} bytes)")
                            zip_file.writestr(file_name_pdf, pdf_data)
                        else:
                            print(f"[DEBUG] ✗ ERROR: No se pudo generar PDF para '{base_filename}'")
                            warnings.append(f"⚠️ No se generó el PDF para {base_filename}")
                            # Si solo se pidió PDF y falló, agregar Excel como respaldo
                            if not generar_excel:
                                print(f"[DEBUG] Agregando Excel como respaldo...")
                                file_name_excel = f"{base_filename}.xlsx"
                                zip_file.writestr(file_name_excel, excel_data)
                        
                        print(f"{'='*70}\n")
                else:
                    print(f"[DEBUG] ERROR: No se generaron datos de Excel para '{curso}'")
        
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
