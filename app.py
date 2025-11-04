import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile
import json
from formato_excel import create_formatted_excel

# --- Función para mapear nombres truncados de Excel a nombres completos ---
def get_nombre_completo_curso(nombre_truncado, config_cursos):
    """
    Mapea un nombre truncado de hoja de Excel (máx 31 caracteres) 
    al nombre completo del curso en config_cursos.json
    """
    # Remover números y guiones bajos del inicio (ej: "1_IPERC..." -> "IPERC...")
    nombre_limpio = nombre_truncado.lstrip('0123456789_')
    
    # Buscar en el JSON por coincidencia parcial
    for nombre_completo in config_cursos['cursos'].keys():
        # Si el nombre limpio es el inicio del nombre completo
        if nombre_completo.startswith(nombre_limpio):
            return nombre_completo
        # O si el nombre limpio está contenido en el nombre completo
        if nombre_limpio in nombre_completo:
            return nombre_completo
    
    # Si no se encuentra, devolver el nombre truncado original
    return nombre_truncado

# --- Configuración de la Página ---
st.set_page_config(
    page_title="Generador de Formatos de Capacitación",
    page_icon="📋",
    layout="wide"
)

# --- INICIALIZAR SESSION STATE PRIMERO ---
if 'dnis_procesados' not in st.session_state:
    st.session_state.dnis_procesados = None
if 'cursos_disponibles' not in st.session_state:
    st.session_state.cursos_disponibles = []
if 'personal_df' not in st.session_state:
    st.session_state.personal_df = None
if 'maestro_excel' not in st.session_state:
    st.session_state.maestro_excel = None
if 'paso_completado' not in st.session_state:
    st.session_state.paso_completado = {
        'paso1_personal': False,
        'paso1_maestro': False,
        'paso2_dnis': False,
        'paso3_cursos': False
    }
if 'config_cursos' not in st.session_state:
    # Cargar configuración de cursos
    config_path = os.path.join(os.path.dirname(__file__), 'config_cursos.json')
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            st.session_state.config_cursos = json.load(f)
    except Exception as e:
        st.session_state.config_cursos = {'cursos': {}, 'configuracion_default': {}}

# --- BARRA LATERAL CON INFORMACIÓN ---
with st.sidebar:
    st.title("🎯 Guía de Uso")
    st.markdown("""
    ### Pasos a seguir:
    
    **1. Cargar Archivos Base** 📂
    - Personal Asignado (Excel)
    - Maestro de Notas (Excel)
    
    **2. Ingresar DNIs** 🔢
    - Pegar manualmente o subir archivo
    - Procesar y validar datos
    
    **3. Seleccionar Cursos** 📚
    - Elegir de los cursos disponibles
    
    **4. Configurar Detalles** ⚙️
    - Tema, capacitador, duración, etc.
    
    **5. Generar y Descargar** 📥
    - Descargar formatos en ZIP
    """)
    
    st.markdown("---")
    
    # Estado actual
    st.subheader("📊 Estado Actual")
    st.write(f"Personal: {'✅ Cargado' if st.session_state.personal_df is not None else '❌ Pendiente'}")
    st.write(f"Maestro: {'✅ Cargado' if st.session_state.maestro_excel is not None else '❌ Pendiente'}")
    st.write(f"DNIs: {'✅ Procesados' if st.session_state.dnis_procesados is not None else '❌ Pendiente'}")
    st.write(f"Cursos: {'✅ Seleccionados' if st.session_state.paso_completado['paso3_cursos'] else '❌ Pendiente'}")
    
    st.markdown("---")
    
    # Botón de reinicio
    if st.button("🔄 Reiniciar Todo", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

st.title("📋 Generador de Formatos de Capacitación")
st.markdown("---")

# --- BARRA DE PROGRESO ---
pasos_completados = sum(st.session_state.paso_completado.values())
progreso = pasos_completados / 4
st.progress(progreso)
st.caption(f"Progreso: {pasos_completados}/4 pasos completados")

# --- PASO 1: CARGAR ARCHIVOS BASE ---
st.header("📂 Paso 1: Cargar Archivos Base")

# Indicador de estado del paso 1
if st.session_state.paso_completado['paso1_personal'] and st.session_state.paso_completado['paso1_maestro']:
    st.success("✅ Paso 1 completado - Archivos cargados correctamente")
else:
    st.info("ℹ️ Sube ambos archivos para continuar al siguiente paso")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Personal Asignado")
    
    # Mostrar estado actual
    if st.session_state.personal_df is not None:
        st.success(f"✅ Archivo cargado ({len(st.session_state.personal_df)} registros)")
        if st.button("🔄 Recargar archivo de Personal", key="reload_personal"):
            st.session_state.personal_df = None
            st.session_state.paso_completado['paso1_personal'] = False
            st.session_state.dnis_procesados = None
            st.session_state.paso_completado['paso2_dnis'] = False
            st.rerun()
    else:
        personal_file = st.file_uploader(
            "Subir archivo Excel",
            type=["xlsx", "xls"],
            key="personal",
            help="Archivo con la información del personal (DNI, Nombre, Unidad)"
        )

        if personal_file:
            with st.spinner("Cargando archivo..."):
                try:
                    # Leer Excel indicando que los encabezados están en la fila 4 (índice 3)
                    # Primero leer para detectar columnas de DNI
                    df = pd.read_excel(personal_file, header=3)
                    
                    # Detectar columnas de DNI y convertir a string con ceros a la izquierda
                    possible_dni_cols = ['DOCUMENTO', 'DNI', 'Documento', 'dni', 'documento', 'DOC']
                    for col in df.columns:
                        if col in possible_dni_cols or 'DNI' in str(col).upper() or 'DOCUMENTO' in str(col).upper():
                            # Convertir a string preservando ceros a la izquierda
                            df[col] = df[col].apply(lambda x: str(int(x)).zfill(8) if pd.notna(x) and str(x).replace('.','').isdigit() else str(x) if pd.notna(x) else '')

                    # Limpiar filas vacías
                    df = df.dropna(how="all")

                    # Guardar en sesión para reutilizar después
                    st.session_state.personal_df = df
                    st.session_state.paso_completado['paso1_personal'] = True

                    # Mostrar mensaje de éxito
                    st.success(f"✅ Archivo cargado correctamente ({len(df)} registros).")
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Error al leer el archivo: {e}")
    
    # Mostrar vista previa si está cargado
    if st.session_state.personal_df is not None:
        with st.expander("👁️ Ver vista previa"):
            st.dataframe(st.session_state.personal_df.head())
        
        with st.expander("📋 Ver columnas disponibles"):
            st.write(list(st.session_state.personal_df.columns))

with col2:
    st.subheader("📊 Maestro de Notas")
    
    # Mostrar estado actual
    if st.session_state.maestro_excel is not None:
        st.success(f"✅ Maestro cargado ({len(st.session_state.cursos_disponibles)} cursos)")
        if st.button("🔄 Recargar Maestro de Notas", key="reload_maestro"):
            st.session_state.maestro_excel = None
            st.session_state.cursos_disponibles = []
            st.session_state.paso_completado['paso1_maestro'] = False
            st.session_state.paso_completado['paso3_cursos'] = False
            st.rerun()
    else:
        maestro_file = st.file_uploader(
            "Subir archivo Excel con múltiples hojas",
            type=['xlsx', 'xls'],
            key='maestro',
            help="Cada hoja representa un curso con las notas de los participantes"
        )
        
        if maestro_file:
            with st.spinner("⏳ Cargando Maestro de Notas..."):
                try:
                    # Cargar el archivo Excel
                    excel_file = pd.ExcelFile(maestro_file)
                    st.session_state.cursos_disponibles = excel_file.sheet_names
                    st.session_state.maestro_excel = excel_file
                    st.session_state.paso_completado['paso1_maestro'] = True
                    
                    st.success(f"✅ Maestro de Notas cargado: {len(st.session_state.cursos_disponibles)} cursos")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error al cargar Maestro de Notas: {e}")
    
    # Mostrar cursos disponibles si está cargado
    if st.session_state.maestro_excel is not None:
        with st.expander("📚 Ver cursos disponibles"):
            for idx, curso in enumerate(st.session_state.cursos_disponibles, 1):
                st.write(f"{idx}. {curso}")

st.markdown("---")

# --- PASO 2: INGRESAR DNIs ---
st.header("🔢 Paso 2: Ingresar DNIs")

# Verificar si se puede avanzar a este paso
paso1_completo = st.session_state.paso_completado['paso1_personal'] and st.session_state.paso_completado['paso1_maestro']

if not paso1_completo:
    st.warning("⚠️ Completa el Paso 1 antes de continuar")
    st.stop()

# Indicador de estado del paso 2
if st.session_state.paso_completado['paso2_dnis']:
    st.success("✅ Paso 2 completado - DNIs procesados correctamente")
else:
    st.info("ℹ️ Ingresa los DNIs para buscar la información del personal")

dni_input_method = st.radio(
    "Método de ingreso de DNIs:",
    ["Pegar DNIs manualmente", "Subir archivo con DNIs"],
    horizontal=True
)

dnis_list = []

if dni_input_method == "Pegar DNIs manualmente":
    dni_text = st.text_area(
        "Pegar DNIs (uno por línea):",
        height=150,
        placeholder="12345678\n87654321\n01234567"
    )
    if dni_text:
        # Limpiar y convertir a string, preservando ceros a la izquierda (formato 8 dígitos)
        dnis_list = []
        for dni in dni_text.split('\n'):
            if dni.strip():
                dni_clean = dni.strip().replace('.', '').replace(',', '').replace(' ', '')
                if dni_clean.isdigit():
                    # Rellenar con ceros a la izquierda si es necesario (DNI peruano = 8 dígitos)
                    dnis_list.append(dni_clean.zfill(8))

else:  # Subir archivo
    dni_file = st.file_uploader(
        "Subir archivo Excel/CSV con DNIs",
        type=['xlsx', 'xls', 'csv'],
        key='dnis'
    )
    
    if dni_file:
        try:
            if dni_file.name.endswith('.csv'):
                dni_df = pd.read_csv(dni_file)
            else:
                dni_df = pd.read_excel(dni_file)
            
            # Intentar detectar la columna de DNIs
            possible_columns = ['DNI', 'DOCUMENTO', 'Documento', 'dni', 'documento']
            dni_column = None
            
            for col in possible_columns:
                if col in dni_df.columns:
                    dni_column = col
                    break
            
            if dni_column:
                # Limpiar DNIs: convertir a string preservando ceros a la izquierda
                dnis_list = []
                for x in dni_df[dni_column].tolist():
                    if pd.notna(x):
                        dni_str = str(x).replace('.0', '').replace(',', '').strip()
                        if dni_str.isdigit():
                            # Rellenar con ceros a la izquierda (DNI peruano = 8 dígitos)
                            dnis_list.append(dni_str.zfill(8))
            else:
                st.warning("⚠️ No se encontró columna de DNI. Selecciona manualmente:")
                dni_column = st.selectbox("Columna con DNIs:", dni_df.columns)
                if dni_column:
                    dnis_list = []
                    for x in dni_df[dni_column].tolist():
                        if pd.notna(x):
                            dni_str = str(x).replace('.0', '').replace(',', '').strip()
                            if dni_str.isdigit():
                                dnis_list.append(dni_str.zfill(8))
        except Exception as e:
            st.error(f"❌ Error al leer archivo de DNIs: {e}")

if dnis_list:
    st.info(f"📊 Total de DNIs ingresados: {len(dnis_list)}")

# Botón para procesar DNIs
col_btn1, col_btn2 = st.columns([3, 1])
with col_btn1:
    procesar_btn = st.button(
        "🔍 Buscar y Procesar DNIs", 
        type="primary", 
        disabled=not dnis_list,
        use_container_width=True
    )
with col_btn2:
    if st.session_state.dnis_procesados is not None:
        if st.button("🗑️ Limpiar DNIs", use_container_width=True):
            st.session_state.dnis_procesados = None
            st.session_state.paso_completado['paso2_dnis'] = False
            st.rerun()

if procesar_btn:
    if st.session_state.personal_df is None:
        st.error("❌ Primero carga el archivo de Personal Asignado")
    else:
        with st.spinner("Buscando información..."):
            # Detectar columna de DNI en Personal Asignado
            possible_dni_cols = ['DOCUMENTO', 'DNI', 'Documento', 'dni', 'documento', 'DOC']
            dni_col_personal = None
            
            for col in possible_dni_cols:
                if col in st.session_state.personal_df.columns:
                    dni_col_personal = col
                    break
            
            if dni_col_personal is None:
                st.error("❌ No se encontró columna de DNI en Personal Asignado. Columnas disponibles:")
                st.write(list(st.session_state.personal_df.columns))
                st.stop()
            
            # Detectar columna de Nombre
            possible_nombre_cols = ['APELLIDOS Y NOMBRES', 'NOMBRE', 'Nombre', 'nombre', 'NOMBRES Y APELLIDOS']
            nombre_col = None
            
            for col in possible_nombre_cols:
                if col in st.session_state.personal_df.columns:
                    nombre_col = col
                    break
            
            # Detectar columna de Unidad
            possible_unidad_cols = ['UNIDAD', 'Unidad', 'unidad', 'UNID', 'CLIENTE']
            unidad_col = None
            
            for col in possible_unidad_cols:
                if col in st.session_state.personal_df.columns:
                    unidad_col = col
                    break
            
            if nombre_col is None or unidad_col is None:
                st.warning(f"⚠️ Columnas detectadas: DNI={dni_col_personal}, Nombre={nombre_col}, Unidad={unidad_col}")
                st.info("Selecciona manualmente las columnas correctas:")
                
                col1, col2 = st.columns(2)
                with col1:
                    nombre_col = st.selectbox("Columna de Nombres:", st.session_state.personal_df.columns)
                with col2:
                    unidad_col = st.selectbox("Columna de Unidad:", st.session_state.personal_df.columns)
                
                if st.button("Continuar con columnas seleccionadas"):
                    pass
                else:
                    st.stop()
            
            # Procesar cada DNI
            processed_data = []
            
            for dni in dnis_list:
                # Asegurar formato de DNI con ceros a la izquierda
                dni_formatted = str(dni).zfill(8) if str(dni).isdigit() else str(dni)
                
                # Buscar en Personal Asignado (comparar ambos formatos por si acaso)
                person = st.session_state.personal_df[
                    (st.session_state.personal_df[dni_col_personal].astype(str) == dni_formatted) |
                    (st.session_state.personal_df[dni_col_personal].astype(str) == str(int(dni_formatted)))
                ]
                
                if not person.empty:
                    nombre = person.iloc[0][nombre_col]
                    unidad = person.iloc[0][unidad_col]
                else:
                    nombre = None
                    unidad = None
                
                processed_data.append({
                    'DNI': dni_formatted,  # Guardar con formato correcto
                    'Nombre': nombre,
                    'Unidad': unidad
                })
            
            st.session_state.dnis_procesados = pd.DataFrame(processed_data)
            st.session_state.paso_completado['paso2_dnis'] = True
            st.success("✅ DNIs procesados correctamente")
            st.rerun()

# Mostrar datos procesados
if st.session_state.dnis_procesados is not None:
    st.subheader("📋 Datos Procesados")
    
    # Identificar DNIs sin información
    faltantes_count = st.session_state.dnis_procesados['Nombre'].isna().sum()
    
    if faltantes_count > 0:
        st.warning(f"⚠️ {faltantes_count} DNI(s) no encontrados en Personal Asignado - Edita directamente en la tabla")
    else:
        st.success(f"✅ Todos los datos están completos ({len(st.session_state.dnis_procesados)} registros)")
    
    st.info("💡 Puedes editar directamente los campos Nombre y Unidad en la tabla. Los cambios se guardan automáticamente.")
    
    # Usar data_editor para editar directamente
    edited_df = st.data_editor(
        st.session_state.dnis_procesados,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            "DNI": st.column_config.TextColumn("DNI", disabled=True, width="medium"),
            "Nombre": st.column_config.TextColumn("Nombre", required=True, width="large"),
            "Unidad": st.column_config.TextColumn("Unidad", required=True, width="large")
        },
        hide_index=True,
        key="data_editor"
    )
    
    # Actualizar el session state con los datos editados
    if not edited_df.equals(st.session_state.dnis_procesados):
        st.session_state.dnis_procesados = edited_df
        st.success("✅ Cambios guardados automáticamente")

st.markdown("---")

# --- PASO 3: SELECCIONAR CURSOS ---
st.header("📚 Paso 3: Seleccionar Cursos")

# Verificar si se puede avanzar a este paso
if not st.session_state.paso_completado['paso2_dnis']:
    st.warning("⚠️ Completa el Paso 2 antes de continuar")
    st.stop()

# Verificar que no haya datos faltantes
if st.session_state.dnis_procesados is not None:
    faltantes_count = st.session_state.dnis_procesados['Nombre'].isna().sum()
    if faltantes_count > 0:
        st.error(f"❌ Completa los {faltantes_count} datos faltantes en el Paso 2 antes de continuar")
        st.stop()

# Indicador de estado del paso 3
if st.session_state.paso_completado['paso3_cursos']:
    st.success("✅ Paso 3 completado - Cursos seleccionados")
else:
    st.info("ℹ️ Selecciona los cursos para generar los formatos")

if st.session_state.cursos_disponibles:
    selected_courses = st.multiselect(
        "Selecciona los cursos a generar:",
        st.session_state.cursos_disponibles,
        help="Puedes seleccionar múltiples cursos",
        key="selected_courses"
    )
    
    if selected_courses:
        st.session_state.paso_completado['paso3_cursos'] = True
        st.info(f"📌 {len(selected_courses)} curso(s) seleccionado(s)")
        
        # --- PASO 4: CONFIGURAR CADA CURSO ---
        st.markdown("---")
        st.header("⚙️ Paso 4: Configurar Detalles de Cursos")
        st.info("ℹ️ Configura los detalles de cada curso seleccionado")
        
        course_configs = {}
        
        # Botón para editar configuración de cursos
        with st.expander("⚙️ Gestionar configuración de cursos"):
            st.info("💡 Puedes editar el archivo 'config_cursos.json' para configurar los 25 cursos con sus datos específicos")
            
            # Mostrar debug de coincidencias
            cursos_json = list(st.session_state.config_cursos['cursos'].keys())
            st.caption(f"**Cursos en JSON:** {len(cursos_json)}")
            st.caption(f"**Cursos seleccionados:** {len(selected_courses)}")
            
            # Verificar coincidencias con mapeo
            st.markdown("**Mapeo de nombres:**")
            for curso in selected_courses:
                nombre_completo = get_nombre_completo_curso(curso, st.session_state.config_cursos)
                if nombre_completo in cursos_json:
                    st.success(f"✅ '{curso}' → '{nombre_completo}'")
                else:
                    st.error(f"❌ '{curso}' → '{nombre_completo}' (no encontrado)")
                    st.caption(f"Búsqueda: '{curso}'")
            
            if st.button("🔄 Recargar configuración desde archivo"):
                config_path = os.path.join(os.path.dirname(__file__), 'config_cursos.json')
                try:
                    with open(config_path, 'r', encoding='utf-8') as f:
                        st.session_state.config_cursos = json.load(f)
                    st.success("✅ Configuración recargada correctamente")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error al recargar configuración: {e}")
        
        for curso in selected_courses:
            # Mapear nombre truncado a nombre completo
            nombre_completo = get_nombre_completo_curso(curso, st.session_state.config_cursos)
            
            # Obtener configuración del curso desde el archivo JSON usando el nombre completo
            curso_config = st.session_state.config_cursos['cursos'].get(nombre_completo, None)
            
            # Si no se encuentra, usar la configuración default
            if curso_config is None:
                curso_config = st.session_state.config_cursos.get('configuracion_default', {})
                st.warning(f"⚠️ Curso '{curso}' (mapeado a '{nombre_completo}') no encontrado en config_cursos.json. Usando configuración por defecto.")
            
            with st.expander(f"📝 {curso}", expanded=False):
                if nombre_completo != curso:
                    st.caption(f"🔗 Nombre completo: **{nombre_completo}**")
                
                if st.session_state.config_cursos['cursos'].get(nombre_completo, None) is not None:
                    st.caption("✅ Configuración cargada desde config_cursos.json")
                else:
                    st.caption("⚠️ Usando configuración por defecto - Agrega este curso al config_cursos.json")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown(f"**Tema/Motivo:**")
                    st.info(curso_config.get('tema_motivo', 'Capacitación en seguridad'))
                    
                    st.markdown(f"**Capacitador/Entrenador:**")
                    st.info(curso_config.get('capacitador', 'Jose Alvines'))
                    
                    st.markdown(f"**Duración:**")
                    st.info(curso_config.get('duracion', '00:30:00'))
                    
                    st.markdown(f"**Firma:**")
                    st.info(curso_config.get('firma', 'firma_capacitador.png'))
                
                with col2:
                    st.markdown(f"**Contenido/Sub Temas:**")
                    st.info(curso_config.get('contenido_subtemas', '* Tema 1\n* Tema 2\n* Tema 3'))
                    
                    st.markdown(f"**Grabación/Material:**")
                    st.info(curso_config.get('grabacion', 'https://youtu.be/ejemplo'))
                
                st.caption("💡 Para editar esta información, modifica el archivo config_cursos.json")
            
            # Construir configuración directamente desde el JSON usando nombre completo
            course_configs[curso] = {
                'Nombre Curso': nombre_completo,  # Usar nombre completo en el Excel generado
                'Tema/Motivo': curso_config.get('tema_motivo', 'Capacitación en seguridad'),
                'Contenido/ Sub Temas': curso_config.get('contenido_subtemas', '* Tema 1\n* Tema 2\n* Tema 3'),
                'Capacitador/Entrenador': curso_config.get('capacitador', 'Jose Alvines'),
                'Duracion': curso_config.get('duracion', '00:30:00'),
                'Grabacion/ Material': curso_config.get('grabacion', 'https://youtu.be/ejemplo'),
                'Firma': curso_config.get('firma', 'firma_capacitador.png')
            }
        
        st.markdown("---")
        
        # --- PASO 5: GENERAR ARCHIVOS ---
        st.header("📥 Paso 5: Generar y Descargar")
        st.info("ℹ️ Revisa la configuración y genera los formatos")
        
        # Resumen antes de generar
        with st.expander("📋 Resumen de la configuración", expanded=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Personal", len(st.session_state.dnis_procesados))
            with col2:
                st.metric("Cursos seleccionados", len(selected_courses))
            with col3:
                st.metric("Formatos a generar", len(selected_courses))
        
        col1, col2 = st.columns(2)
        with col1:
            output_format = st.radio(
                "Formato de salida:",
                ["Excel (.xlsx)", "PDF"],
                horizontal=True,
                disabled=True,
                help="Por ahora solo está disponible Excel"
            )
        
        generar_btn = st.button(
            "🚀 Generar Formatos", 
            type="primary",
            use_container_width=True,
            help="Click para generar todos los formatos configurados"
        )
        
        if generar_btn:
            if st.session_state.dnis_procesados is None:
                st.error("❌ Primero procesa los DNIs")
            elif st.session_state.dnis_procesados['Nombre'].isna().any():
                st.error("❌ Completa los datos faltantes antes de generar")
            else:
                # Barra de progreso para la generación
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                with st.spinner("Generando formatos..."):
                    zip_buffer = BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for idx, curso in enumerate(selected_courses, 1):
                            # Actualizar progreso
                            progress = idx / len(selected_courses)
                            progress_bar.progress(progress)
                            status_text.text(f"Generando {idx}/{len(selected_courses)}: {curso}")
                            
                            # Cargar la hoja solo cuando se necesita (lazy loading)
                            try:
                                maestro_curso = pd.read_excel(st.session_state.maestro_excel, sheet_name=curso)
                            except Exception as e:
                                st.warning(f"⚠️ No se pudo cargar datos de {curso}: {e}")
                                maestro_curso = None
                            
                            # Crear DataFrame para este curso
                            curso_data = []
                            
                            for idx, row in st.session_state.dnis_procesados.iterrows():
                                dni = str(row['DNI'])
                                
                                # Buscar en maestro de notas
                                nota_info = None
                                if maestro_curso is not None:
                                    # Detectar columna de DNI en maestro
                                    possible_dni_cols = ['DNI', 'DOCUMENTO', 'Documento', 'dni', 'documento']
                                    dni_col_maestro = None
                                    
                                    for col in possible_dni_cols:
                                        if col in maestro_curso.columns:
                                            dni_col_maestro = col
                                            break
                                    
                                    if dni_col_maestro:
                                        # Buscar por DNI (intentar con y sin ceros a la izquierda)
                                        dni_sin_ceros = str(int(dni)) if dni.isdigit() else dni
                                        nota_row = maestro_curso[
                                            (maestro_curso[dni_col_maestro].astype(str) == dni) |
                                            (maestro_curso[dni_col_maestro].astype(str) == dni_sin_ceros) |
                                            (maestro_curso[dni_col_maestro].astype(str).str.zfill(8) == dni)
                                        ]
                                        if not nota_row.empty:
                                            nota_info = nota_row.iloc[0]
                                
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
                            nombre_completo_archivo = course_configs[curso]['Nombre Curso']
                            
                            # Generar Excel
                            excel_data = create_formatted_excel(df_curso, course_configs[curso])
                            
                            if excel_data:
                                # Nombre del archivo: NombreCurso - Unidad (usando nombre completo)
                                unidad = df_curso['Unidad (Cliente)'].iloc[0] if not df_curso.empty else 'Sin_Unidad'
                                file_name = f"{nombre_completo_archivo} - {unidad}.xlsx"
                                
                                zip_file.writestr(file_name, excel_data)
                    
                    zip_buffer.seek(0)
                    
                    # Limpiar barra de progreso
                    progress_bar.empty()
                    status_text.empty()
                    
                    st.success("✅ Formatos generados correctamente")
                    
                    st.download_button(
                        label="📦 Descargar ZIP con todos los formatos",
                        data=zip_buffer.getvalue(),
                        file_name="Formatos_Capacitacion.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
    else:
        st.info("👆 Selecciona al menos un curso para continuar")

else:
    st.warning("⚠️ Carga primero el Maestro de Notas para ver los cursos disponibles")