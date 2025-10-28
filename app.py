import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile
from formato_excel import create_formatted_excel

# --- Configuración de la Página ---
st.set_page_config(
    page_title="Generador de Formatos de Capacitación",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Generador de Formatos de Capacitación")
st.markdown("---")

# --- INICIALIZAR SESSION STATE ---
if 'dnis_procesados' not in st.session_state:
    st.session_state.dnis_procesados = None
if 'cursos_disponibles' not in st.session_state:
    st.session_state.cursos_disponibles = []
if 'personal_df' not in st.session_state:
    st.session_state.personal_df = None
if 'maestro_excel' not in st.session_state:
    st.session_state.maestro_excel = None

# --- PASO 1: CARGAR ARCHIVOS BASE ---
st.header("📂 Paso 1: Cargar Archivos Base")

col1, col2 = st.columns(2)

with col1:
    personal_file = st.file_uploader(
        "Subir archivo: Personal Asignado (Excel)",
        type=["xlsx", "xls"],
        key="personal"
    )

    if personal_file:
        try:
            # Leer Excel indicando que los encabezados están en la fila 4 (índice 3)
            df = pd.read_excel(personal_file, header=3)

            # Guardar en sesión para reutilizar después
            st.session_state.personal_df = df

            # Limpiar filas vacías
            df = df.dropna(how="all")

            # Mostrar mensaje de éxito
            st.success(f"✅ Archivo cargado correctamente ({len(df)} registros).")

            # Mostrar vista previa
            st.dataframe(df.head())

            # Mostrar columnas disponibles
            with st.expander("Ver columnas disponibles"):
                st.write(list(df.columns))

        except Exception as e:
            st.error(f"❌ Error al leer el archivo: {e}")

with col2:
    maestro_file = st.file_uploader(
        "Subir Maestro de Notas (Excel con múltiples hojas)",
        type=['xlsx', 'xls'],
        key='maestro'
    )
    
    if maestro_file and 'maestro_loaded' not in st.session_state:
        with st.spinner("⏳ Cargando Maestro de Notas... Esto puede tomar unos segundos..."):
            try:
                # Cargar el archivo Excel
                excel_file = pd.ExcelFile(maestro_file)
                st.session_state.cursos_disponibles = excel_file.sheet_names
                st.session_state.maestro_excel = excel_file  # Guardar el archivo completo
                st.session_state.maestro_loaded = True
                
                st.success(f"✅ Maestro de Notas cargado: {len(st.session_state.cursos_disponibles)} cursos disponibles")
            except Exception as e:
                st.error(f"❌ Error al cargar Maestro de Notas: {e}")
    
    if 'maestro_loaded' in st.session_state:
        st.success(f"✅ Maestro de Notas cargado: {len(st.session_state.cursos_disponibles)} cursos")
        with st.expander("Ver cursos disponibles"):
            for idx, curso in enumerate(st.session_state.cursos_disponibles, 1):
                st.write(f"{idx}. {curso}")

st.markdown("---")

# --- PASO 2: INGRESAR DNIs ---
st.header("🔢 Paso 2: Ingresar DNIs")

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
        placeholder="12345678\n87654321\n11223344"
    )
    if dni_text:
        # Limpiar y convertir a string, eliminando espacios y puntos
        dnis_list = [str(int(float(dni.strip().replace('.', '').replace(',', '')))) 
                     for dni in dni_text.split('\n') 
                     if dni.strip() and dni.strip().replace('.', '').replace(',', '').isdigit()]

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
                # Limpiar DNIs: convertir a string sin decimales
                dnis_list = [str(int(float(x))) if pd.notna(x) else '' 
                            for x in dni_df[dni_column].tolist()]
                dnis_list = [dni for dni in dnis_list if dni]  # Eliminar vacíos
            else:
                st.warning("⚠️ No se encontró columna de DNI. Selecciona manualmente:")
                dni_column = st.selectbox("Columna con DNIs:", dni_df.columns)
                if dni_column:
                    dnis_list = dni_df[dni_column].astype(str).tolist()
        except Exception as e:
            st.error(f"❌ Error al leer archivo de DNIs: {e}")

if dnis_list:
    st.info(f"📊 Total de DNIs ingresados: {len(dnis_list)}")

# Botón para procesar DNIs
if st.button("🔍 Buscar y Procesar DNIs", type="primary", disabled=not dnis_list):
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
                # Buscar en Personal Asignado
                person = st.session_state.personal_df[
                    st.session_state.personal_df[dni_col_personal].astype(str) == str(dni)
                ]
                
                if not person.empty:
                    nombre = person.iloc[0][nombre_col]
                    unidad = person.iloc[0][unidad_col]
                else:
                    nombre = None
                    unidad = None
                
                processed_data.append({
                    'DNI': dni,
                    'Nombre': nombre,
                    'Unidad': unidad
                })
            
            st.session_state.dnis_procesados = pd.DataFrame(processed_data)
            st.success("✅ DNIs procesados correctamente")

# Mostrar datos procesados
if st.session_state.dnis_procesados is not None:
    st.subheader("📋 Datos Procesados")
    
    # Identificar DNIs sin información
    faltantes = st.session_state.dnis_procesados[
        st.session_state.dnis_procesados['Nombre'].isna()
    ]
    
    if not faltantes.empty:
        st.warning(f"⚠️ {len(faltantes)} DNI(s) no encontrados en Personal Asignado")
        
        with st.expander("Completar datos manualmente"):
            for idx, row in faltantes.iterrows():
                st.write(f"**DNI: {row['DNI']}**")
                col1, col2 = st.columns(2)
                with col1:
                    nombre = st.text_input(f"Nombre para {row['DNI']}", key=f"nombre_{row['DNI']}")
                with col2:
                    unidad = st.text_input(f"Unidad para {row['DNI']}", key=f"unidad_{row['DNI']}")
                
                if nombre and unidad:
                    st.session_state.dnis_procesados.loc[idx, 'Nombre'] = nombre
                    st.session_state.dnis_procesados.loc[idx, 'Unidad'] = unidad
    
    st.dataframe(st.session_state.dnis_procesados, use_container_width=True)

st.markdown("---")

# --- PASO 3: SELECCIONAR CURSOS ---
st.header("📚 Paso 3: Seleccionar Cursos")

if st.session_state.cursos_disponibles:
    selected_courses = st.multiselect(
        "Selecciona los cursos a generar:",
        st.session_state.cursos_disponibles,
        help="Puedes seleccionar múltiples cursos"
    )
    
    if selected_courses:
        st.info(f"📌 {len(selected_courses)} curso(s) seleccionado(s)")
        
        # --- PASO 4: CONFIGURAR CADA CURSO ---
        st.markdown("---")
        st.header("⚙️ Paso 4: Configurar Detalles de Cursos")
        
        course_configs = {}
        
        for curso in selected_courses:
            with st.expander(f"📝 Configurar: {curso}", expanded=False):
                col1, col2 = st.columns(2)
                
                with col1:
                    tema = st.text_input(
                        "Tema/Motivo:",
                        key=f"tema_{curso}",
                        value="Capacitación en seguridad"
                    )
                    capacitador = st.text_input(
                        "Capacitador/Entrenador:",
                        key=f"capacitador_{curso}",
                        value="Jose Alvines"
                    )
                    duracion = st.text_input(
                        "Duración (HH:MM:SS):",
                        key=f"duracion_{curso}",
                        value="00:30:00"
                    )
                
                with col2:
                    contenido = st.text_area(
                        "Contenido/Sub Temas:",
                        key=f"contenido_{curso}",
                        height=100,
                        value="* Tema 1\n* Tema 2\n* Tema 3"
                    )
                    grabacion = st.text_input(
                        "Grabación/Material (URL):",
                        key=f"grabacion_{curso}",
                        value="https://youtu.be/ejemplo"
                    )
                
                course_configs[curso] = {
                    'Nombre Curso': curso,
                    'Tema/Motivo': tema,
                    'Contenido/ Sub Temas': contenido,
                    'Capacitador/Entrenador': capacitador,
                    'Duracion': duracion,
                    'Grabacion/ Material': grabacion
                }
        
        st.markdown("---")
        
        # --- PASO 5: GENERAR ARCHIVOS ---
        st.header("📥 Paso 5: Generar y Descargar")
        
        col1, col2 = st.columns(2)
        with col1:
            output_format = st.radio(
                "Formato de salida:",
                ["Excel (.xlsx)", "PDF"],
                horizontal=True
            )
        
        if st.button("🚀 Generar Formatos", type="primary"):
            if st.session_state.dnis_procesados is None:
                st.error("❌ Primero procesa los DNIs")
            elif st.session_state.dnis_procesados['Nombre'].isna().any():
                st.error("❌ Completa los datos faltantes antes de generar")
            else:
                with st.spinner("Generando formatos..."):
                    zip_buffer = BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for idx, curso in enumerate(selected_courses, 1):
                            st.write(f"Generando {idx}/{len(selected_courses)}: {curso}")
                            
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
                                    nota_row = maestro_curso[
                                        maestro_curso['DNI'].astype(str) == dni
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
                            
                            # Generar Excel
                            excel_data = create_formatted_excel(df_curso, course_configs[curso])
                            
                            if excel_data:
                                # Nombre del archivo: NombreCurso - Unidad
                                unidad = df_curso['Unidad (Cliente)'].iloc[0] if not df_curso.empty else 'Sin_Unidad'
                                file_name = f"{curso} - {unidad}.xlsx"
                                
                                zip_file.writestr(file_name, excel_data)
                    
                    zip_buffer.seek(0)
                    
                    st.success("✅ Formatos generados correctamente")
                    
                    st.download_button(
                        label="📦 Descargar ZIP con todos los formatos",
                        data=zip_buffer.getvalue(),
                        file_name="Formatos_Capacitacion.zip",
                        mime="application/zip"
                    )

else:
    st.warning("⚠️ Carga primero el Maestro de Notas para ver los cursos disponibles")