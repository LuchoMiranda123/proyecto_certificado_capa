import streamlit as st
import pandas as pd
import os
import json
from generador_archivos import get_nombre_completo_curso, generar_zip_formatos

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
                    'DNI': dni_formatted,
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
    
    # --- HERRAMIENTA DE EDICIÓN MASIVA DE UNIDAD ---
    with st.expander("✏️ Cambiar Unidad en Múltiples Registros", expanded=False):
        st.info("💡 Usa esta herramienta para cambiar la Unidad de varios registros a la vez")
        
        col1, col2, col3 = st.columns([2, 2, 1])
        
        with col1:
            # Obtener lista de unidades únicas disponibles
            unidades_disponibles = st.session_state.dnis_procesados['Unidad'].dropna().unique().tolist()
            nueva_unidad = st.text_input(
                "Nueva Unidad:", 
                placeholder="Escribe el nombre de la unidad",
                help="Escribe la unidad que quieres asignar a los registros seleccionados"
            )
            if unidades_disponibles:
                st.caption(f"💡 Unidades existentes: {', '.join(unidades_disponibles[:3])}{'...' if len(unidades_disponibles) > 3 else ''}")
        
        with col2:
            # Opciones de selección
            modo_seleccion = st.radio(
                "Aplicar a:",
                ["Todos los registros", "Registros específicos (por índice)", "Registros con Unidad vacía"],
                help="Elige qué registros quieres actualizar"
            )
        
        with col3:
            st.write("")  # Espaciador
            st.write("")  # Espaciador
            aplicar_cambio = st.button("✅ Aplicar", type="primary", use_container_width=True)
        
        if modo_seleccion == "Registros específicos (por índice)":
            indices_str = st.text_input(
                "Índices (separados por comas):",
                placeholder="Ej: 1,2,3,5-10",
                help="Puedes usar rangos con guión (5-10) o números separados por comas (1,2,3)"
            )
        
        if aplicar_cambio and nueva_unidad:
            try:
                if modo_seleccion == "Todos los registros":
                    st.session_state.dnis_procesados['Unidad'] = nueva_unidad
                    st.success(f"✅ Unidad actualizada a '{nueva_unidad}' en todos los {len(st.session_state.dnis_procesados)} registros")
                    st.rerun()
                
                elif modo_seleccion == "Registros con Unidad vacía":
                    mask = st.session_state.dnis_procesados['Unidad'].isna()
                    count = mask.sum()
                    if count > 0:
                        st.session_state.dnis_procesados.loc[mask, 'Unidad'] = nueva_unidad
                        st.success(f"✅ Unidad actualizada a '{nueva_unidad}' en {count} registros vacíos")
                        st.rerun()
                    else:
                        st.warning("⚠️ No hay registros con Unidad vacía")
                
                elif modo_seleccion == "Registros específicos (por índice)":
                    # Parsear índices
                    indices = []
                    for parte in indices_str.split(','):
                        parte = parte.strip()
                        if '-' in parte:
                            inicio, fin = map(int, parte.split('-'))
                            indices.extend(range(inicio-1, fin))  # -1 porque el usuario ve índices desde 1
                        else:
                            indices.append(int(parte) - 1)
                    
                    # Validar índices
                    indices = [i for i in indices if 0 <= i < len(st.session_state.dnis_procesados)]
                    
                    if indices:
                        st.session_state.dnis_procesados.loc[indices, 'Unidad'] = nueva_unidad
                        st.success(f"✅ Unidad actualizada a '{nueva_unidad}' en {len(indices)} registros")
                        st.rerun()
                    else:
                        st.error("❌ Índices inválidos")
            
            except Exception as e:
                st.error(f"❌ Error al aplicar cambios: {e}")
    
    st.info("💡 También puedes editar directamente en la tabla. Los cambios se guardan automáticamente.")
    
    # Usar data_editor para editar directamente (tabla más pequeña sin scroll)
    edited_df = st.data_editor(
        st.session_state.dnis_procesados,
        use_container_width=True,
        num_rows="fixed",
        height=400,  # Altura fija para evitar scroll excesivo
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
    # Definir categorías de cursos
    CATEGORIAS_CURSOS = {
        'SSOMA': [
            'IPERC, mapa de riesgos y procedimientos PETS',
            'Primeros auxilios y prevención contra incendios',
            'Respuesta ante emergencias, Contingencias y desastres naturales',
            'Respuesta ante emergencias, contingencias y desastres naturales',
            'Salud ocupacional y estilo de vida saludable',
            'Seguridad y prevención en el puesto de trabajo',
            'IPERC, mapa de riesgos y procedimientos PETS (Parte 02)',
            'Gestión de residuos sólidos, impactos ambientales y responsabilidad social empresarial'
        ],
        'TÉCNICO': [
            'Defensa personal y uso de la fuerza',
            'Derechos humanos, principios voluntarios y constitución',
            'Prevención de delitos de comercio internacional',
            'Integridad y ética en la seguridad privada',
            'Armas: conocimiento y manipulación',
            'Armas: Conocimiento y manipulación',
            'Normas y procedimientos de seguridad',
            'Legislación y seguridad privada',
            'Seguridad de instalaciones',
            'Eventos indeseables, perturbadores y lugares hostiles'
        ],
        'ESTRATÉGICO': [
            'Hostigamiento sexual laboral',
            'Fundamentos de SGI - 2025',
            'Fundamentos del Sistema Integrado de Gestión',
            'Prevención de riesgos de soborno',
            'Prevención de delitos relacionados a ciberdelincuencia'
        ],
        'ANTAPACCAY': [
            'Uso de la Fuerza',
            'DDHH y Principios voluntarios - Integridad y ética en la seguridad privada',
            'Falsificacion de documentos'
        ]
    }
    
    # Clasificar cursos disponibles por categoría
    cursos_por_categoria = {cat: [] for cat in CATEGORIAS_CURSOS.keys()}
    cursos_por_categoria['OTROS'] = []
    
    for curso in st.session_state.cursos_disponibles:
        # Mapear nombre truncado a nombre completo
        nombre_completo = get_nombre_completo_curso(curso, st.session_state.config_cursos['cursos'])
        
        asignado = False
        for categoria, lista_cursos in CATEGORIAS_CURSOS.items():
            if nombre_completo in lista_cursos:
                cursos_por_categoria[categoria].append(curso)
                asignado = True
                break
        
        if not asignado:
            cursos_por_categoria['OTROS'].append(curso)
    
    # Mostrar selección por categorías
    st.markdown("### Selecciona cursos por categoría:")
    
    selected_courses = []
    
    # Crear tabs para cada categoría
    tabs = st.tabs(['🛡️ SSOMA', '🔧 TÉCNICO', '📊 ESTRATÉGICO', '📦 OTROS', '⛰️ ANTAPACCAY'])
    
    with tabs[0]:  # SSOMA
        if cursos_por_categoria['SSOMA']:
            st.markdown("**Cursos de Seguridad, Salud Ocupacional y Medio Ambiente:**")
            cursos_ssoma = st.multiselect(
                "Selecciona cursos de SSOMA:",
                cursos_por_categoria['SSOMA'],
                key="ssoma_courses"
            )
            selected_courses.extend(cursos_ssoma)
            st.info(f"📌 {len(cursos_ssoma)} curso(s) de SSOMA seleccionado(s)")
        else:
            st.warning("No hay cursos de SSOMA disponibles")
    
    with tabs[1]:  # TÉCNICO
        if cursos_por_categoria['TÉCNICO']:
            st.markdown("**Cursos Técnicos de Seguridad:**")
            cursos_tecnico = st.multiselect(
                "Selecciona cursos técnicos:",
                cursos_por_categoria['TÉCNICO'],
                key="tecnico_courses"
            )
            selected_courses.extend(cursos_tecnico)
            st.info(f"📌 {len(cursos_tecnico)} curso(s) técnico(s) seleccionado(s)")
        else:
            st.warning("No hay cursos técnicos disponibles")
    
    with tabs[2]:  # ESTRATÉGICO
        if cursos_por_categoria['ESTRATÉGICO']:
            st.markdown("**Cursos Estratégicos y de Gestión:**")
            cursos_estrategico = st.multiselect(
                "Selecciona cursos estratégicos:",
                cursos_por_categoria['ESTRATÉGICO'],
                key="estrategico_courses"
            )
            selected_courses.extend(cursos_estrategico)
            st.info(f"📌 {len(cursos_estrategico)} curso(s) estratégico(s) seleccionado(s)")
        else:
            st.warning("No hay cursos estratégicos disponibles")
    
    with tabs[3]:  # OTROS
        if cursos_por_categoria['OTROS']:
            st.markdown("**Otros Cursos:**")
            cursos_otros = st.multiselect(
                "Selecciona otros cursos:",
                cursos_por_categoria['OTROS'],
                key="otros_courses"
            )
            selected_courses.extend(cursos_otros)
            st.info(f"📌 {len(cursos_otros)} curso(s) adicional(es) seleccionado(s)")
        else:
            st.info("No hay otros cursos disponibles")
    
    with tabs[4]:  # ANTAPACCAY
        if cursos_por_categoria['ANTAPACCAY']:
            st.markdown("**Cursos Antapaccay:**")
            cursos_antapaccay = st.multiselect(
                "Selecciona cursos de Antapaccay:",
                cursos_por_categoria['ANTAPACCAY'],
                key="antapaccay_courses"
            )
            selected_courses.extend(cursos_antapaccay)
            st.info(f"📌 {len(cursos_antapaccay)} curso(s) de Antapaccay seleccionado(s)")
        else:
            st.warning("No hay cursos de Antapaccay disponibles")
    
    # Resumen de selección total
    if selected_courses:
        st.markdown("---")
        st.success(f"✅ **Total: {len(selected_courses)} curso(s) seleccionado(s) en todas las categorías**")
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
                nombre_completo = get_nombre_completo_curso(curso, st.session_state.config_cursos['cursos'])
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
            nombre_completo = get_nombre_completo_curso(curso, st.session_state.config_cursos['cursos'])
            
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
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Personal", len(st.session_state.dnis_procesados))
            with col2:
                st.metric("Total Cursos", len(selected_courses))
            with col3:
                # Contar cursos por categoría seleccionados
                count_ssoma = len([c for c in selected_courses if c in cursos_por_categoria['SSOMA']])
                count_tecnico = len([c for c in selected_courses if c in cursos_por_categoria['TÉCNICO']])
                count_estrategico = len([c for c in selected_courses if c in cursos_por_categoria['ESTRATÉGICO']])
                count_otros = len([c for c in selected_courses if c in cursos_por_categoria['OTROS']])
                st.metric("Por Categoría", f"S:{count_ssoma} T:{count_tecnico} E:{count_estrategico} O:{count_otros}")
            with col4:
                st.metric("Formatos a generar", len(selected_courses))
        
        col1, col2 = st.columns(2)
        with col1:
            output_format = st.radio(
                "Formato de salida:",
                ["Excel (.xlsx)", "PDF", "Ambos (Excel + PDF)"],
                horizontal=True,
                help="Elige el formato de descarga"
            )
        
        # Opciones de descarga por grupo
        st.markdown("### 📦 Opciones de Descarga")
        
        descarga_option = st.radio(
            "¿Cómo deseas descargar los formatos?",
            ["Todo en un solo ZIP", "ZIP separado por categoría"],
            horizontal=True,
            help="Descarga todo junto o separado por categorías SSOMA, TÉCNICO, ESTRATÉGICO"
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
                if descarga_option == "Todo en un solo ZIP":
                    # Generación tradicional: todo en un solo ZIP
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    def actualizar_progreso(idx, total, curso):
                        progress = idx / total
                        progress_bar.progress(progress)
                        status_text.text(f"Generando {idx}/{total}: {curso}")
                    
                    with st.spinner("Generando formatos..."):
                        zip_buffer, zip_filename, warnings = generar_zip_formatos(
                            dnis_procesados=st.session_state.dnis_procesados,
                            selected_courses=selected_courses,
                            maestro_excel=st.session_state.maestro_excel,
                            course_configs=course_configs,
                            output_format=output_format,
                            progress_callback=actualizar_progreso
                        )
                        
                        for warning in warnings:
                            st.warning(warning)
                        
                        progress_bar.empty()
                        status_text.empty()
                        
                        st.success("✅ Formatos generados correctamente")
                        
                        if output_format == "Excel (.xlsx)":
                            label = "📦 Descargar ZIP con archivos Excel"
                        elif output_format == "PDF":
                            label = "📦 Descargar ZIP con archivos PDF"
                        else:
                            label = "📦 Descargar ZIP con archivos Excel y PDF"
                        
                        st.download_button(
                            label=label,
                            data=zip_buffer.getvalue(),
                            file_name=zip_filename,
                            mime="application/zip",
                            use_container_width=True
                        )
                
                else:  # ZIP separado por categoría
                    st.markdown("### 📦 Descargas por Categoría")
                    
                    # Separar cursos por categoría
                    cursos_ssoma_sel = [c for c in selected_courses if c in cursos_por_categoria['SSOMA']]
                    cursos_tecnico_sel = [c for c in selected_courses if c in cursos_por_categoria['TÉCNICO']]
                    cursos_estrategico_sel = [c for c in selected_courses if c in cursos_por_categoria['ESTRATÉGICO']]
                    cursos_otros_sel = [c for c in selected_courses if c in cursos_por_categoria['OTROS']]
                    cursos_antapaccay_sel = [c for c in selected_courses if c in cursos_por_categoria['ANTAPACCAY']]
                    
                    categorias_con_cursos = []
                    if cursos_ssoma_sel:
                        categorias_con_cursos.append(('SSOMA', '🛡️', cursos_ssoma_sel))
                    if cursos_tecnico_sel:
                        categorias_con_cursos.append(('TÉCNICO', '🔧', cursos_tecnico_sel))
                    if cursos_estrategico_sel:
                        categorias_con_cursos.append(('ESTRATÉGICO', '📊', cursos_estrategico_sel))
                    if cursos_otros_sel:
                        categorias_con_cursos.append(('OTROS', '📦', cursos_otros_sel))
                    if cursos_antapaccay_sel:
                        categorias_con_cursos.append(('ANTAPACCAY', '⛰️', cursos_antapaccay_sel))
                    
                    # Generar ZIPs separados
                    for categoria_nombre, icono, cursos_categoria in categorias_con_cursos:
                        with st.expander(f"{icono} {categoria_nombre} ({len(cursos_categoria)} cursos)", expanded=True):
                            st.markdown(f"**Cursos incluidos:**")
                            for curso in cursos_categoria:
                                st.markdown(f"- {curso}")
                            
                            # Generar ZIP para esta categoría
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            def actualizar_progreso_cat(idx, total, curso):
                                progress = idx / total
                                progress_bar.progress(progress)
                                status_text.text(f"[{categoria_nombre}] Generando {idx}/{total}: {curso}")
                            
                            # Filtrar course_configs para esta categoría
                            course_configs_cat = {k: v for k, v in course_configs.items() if k in cursos_categoria}
                            
                            with st.spinner(f"Generando formatos de {categoria_nombre}..."):
                                zip_buffer, zip_filename, warnings = generar_zip_formatos(
                                    dnis_procesados=st.session_state.dnis_procesados,
                                    selected_courses=cursos_categoria,
                                    maestro_excel=st.session_state.maestro_excel,
                                    course_configs=course_configs_cat,
                                    output_format=output_format,
                                    progress_callback=actualizar_progreso_cat
                                )
                                
                                for warning in warnings:
                                    st.warning(warning)
                                
                                progress_bar.empty()
                                status_text.empty()
                                
                                st.success(f"✅ {categoria_nombre} generado correctamente")
                                
                                # Ajustar nombre del archivo ZIP
                                zip_filename_cat = zip_filename.replace('.zip', f'_{categoria_nombre}.zip')
                                
                                if output_format == "Excel (.xlsx)":
                                    label = f"📥 Descargar {categoria_nombre} - Excel"
                                elif output_format == "PDF":
                                    label = f"📥 Descargar {categoria_nombre} - PDF"
                                else:
                                    label = f"📥 Descargar {categoria_nombre} - Excel + PDF"
                                
                                st.download_button(
                                    label=label,
                                    data=zip_buffer.getvalue(),
                                    file_name=zip_filename_cat,
                                    mime="application/zip",
                                    use_container_width=True,
                                    key=f"download_{categoria_nombre}"
                                )
    else:
        st.info("👆 Selecciona al menos un curso para continuar")

else:
    st.warning("⚠️ Carga primero el Maestro de Notas para ver los cursos disponibles")