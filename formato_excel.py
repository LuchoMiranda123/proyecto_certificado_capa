import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as OpenpyxlImage
from io import BytesIO
import datetime
import os

# --- Rutas de Archivos Estáticos (Logo, Firma) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "plantillas", "logo_liderman.png")
FIRMA_PATH = os.path.join(BASE_DIR, "plantillas", "firma_capacitador.png")

def apply_border_to_range(ws, start_cell, end_cell, border):
    """
    Aplica bordes a todas las celdas en un rango, incluso si están combinadas.
    """
    from openpyxl.utils import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(f"{start_cell}:{end_cell}")
    
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row=row, column=col).border = border

def create_formatted_excel(df_course, course_details):
    """
    Toma un DataFrame y detalles de curso, y devuelve 
    un archivo Excel (en bytes) con el formato exacto.
    """
    try:
        output = BytesIO()
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = course_details['Nombre Curso'][:31]

        # --- Definir Estilos ---
        font_bold_14 = Font(bold=True, size=14)
        font_bold_10 = Font(bold=True, size=10)
        font_normal = Font(size=10)

        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
        top_left_align = Alignment(horizontal='left', vertical='top', wrap_text=True)
        top_center_align = Alignment(horizontal='center', vertical='top', wrap_text=True)

        thin_border = Border(left=Side(style='thin'), 
                             right=Side(style='thin'), 
                             top=Side(style='thin'), 
                             bottom=Side(style='thin'))
        
        grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        
        # --- Ajustar Anchos de Columna ---
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 52
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 57
        ws.column_dimensions['E'].width = 19
        ws.column_dimensions['F'].width = 19
        ws.column_dimensions['G'].width = 19
        ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 12 

        # --- Fila 1-4: Logo y Encabezado ---
        if os.path.exists(LOGO_PATH):
            try:
                img = OpenpyxlImage(LOGO_PATH)
                img.width = 180
                img.height = 60
                ws.merge_cells('A1:B4')
                ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
                ws.add_image(img, 'A1')
            except Exception as e:
                print(f"Advertencia: No se pudo cargar el logo: {e}")
        else:
            print(f"Advertencia: No se encontró el archivo de logo en {LOGO_PATH}")

        ws.merge_cells('C1:G4')
        cell = ws['C1']
        cell.value = "FORMATO\n\nLISTA DE ASISTENCIA VIRTUAL"
        cell.font = font_bold_14
        cell.alignment = center_align
        apply_border_to_range(ws, 'C1', 'G4', thin_border)

        ws.cell(row=1, column=8, value="Código:").font = font_bold_10
        ws.cell(row=1, column=8).border = thin_border
        ws.cell(row=1, column=9, value="JV-GTH-F-111").font = font_normal
        ws.cell(row=1, column=9).border = thin_border

        ws.cell(row=2, column=8, value="Versión:").font = font_bold_10
        ws.cell(row=2, column=8).border = thin_border
        ws.cell(row=2, column=9, value=4).font = font_normal
        ws.cell(row=2, column=9).border = thin_border

        ws.cell(row=3, column=8, value="Fecha:").font = font_bold_10
        ws.cell(row=3, column=8).border = thin_border
        ws.cell(row=3, column=9, value=datetime.date.today().strftime('%d/%m/%Y')).font = font_normal
        ws.cell(row=3, column=9).border = thin_border

        ws.cell(row=4, column=8, value="Página 01").font = font_normal
        ws.cell(row=4, column=8).border = thin_border
        ws.cell(row=4, column=9, value="").border = thin_border

        # --- Fila 5: Datos del empleador ---
        ws.merge_cells('A5:I5')
        cell = ws['A5']
        cell.value = "Datos del empleador:"
        cell.font = font_bold_10
        cell.alignment = left_align
        apply_border_to_range(ws, 'A5', 'I5', thin_border)

        # --- FILA 6: Cabeceras de la tabla superior ---
        ws.cell(row=6, column=1, value="MARCAR").font = font_bold_10
        ws.cell(row=6, column=1).alignment = center_align
        ws.cell(row=6, column=1).border = thin_border
        
        ws.cell(row=6, column=2, value="RAZÓN SOCIAL").font = font_bold_10
        ws.cell(row=6, column=2).alignment = center_align
        ws.cell(row=6, column=2).border = thin_border
        
        ws.cell(row=6, column=3, value="RUC").font = font_bold_10
        ws.cell(row=6, column=3).alignment = center_align
        ws.cell(row=6, column=3).border = thin_border
        
        ws.merge_cells('D6:G6')
        ws.cell(row=6, column=4, value="DOMICILIO").font = font_bold_10
        ws.cell(row=6, column=4).alignment = center_align
        apply_border_to_range(ws, 'D6', 'G6', thin_border)
        
        ws.merge_cells('H6:I6')
        ws.cell(row=6, column=8, value="ACTIVIDAD ECONOMICA").font = font_bold_10
        ws.cell(row=6, column=8).alignment = center_align
        apply_border_to_range(ws, 'H6', 'I6', thin_border)

        # --- FILAS 7-10: Datos de las empresas ---
        # Fila 7
        ws.cell(row=7, column=1, value="X").font = font_normal
        ws.cell(row=7, column=1).alignment = center_align
        ws.cell(row=7, column=1).border = thin_border
        
        ws.cell(row=7, column=2, value="J&V RESGUARDO SAC").font = font_normal
        ws.cell(row=7, column=2).alignment = center_align
        ws.cell(row=7, column=2).border = thin_border
        
        ws.cell(row=7, column=3, value="20100901481").font = font_normal
        ws.cell(row=7, column=3).alignment = center_align
        ws.cell(row=7, column=3).border = thin_border
        
        ws.merge_cells('D7:G7')
        ws.cell(row=7, column=4, value="AV. DEFENSORES DEL MORRO N°1620 - CHORRILLOS").font = font_normal
        ws.cell(row=7, column=4).alignment = center_align
        apply_border_to_range(ws, 'D7', 'G7', thin_border)
        
        ws.merge_cells('H7:I7')
        ws.cell(row=7, column=8, value="VIGILANCIA PRIVADA").font = font_normal
        ws.cell(row=7, column=8).alignment = center_align
        apply_border_to_range(ws, 'H7', 'I7', thin_border)
        
        # Fila 8
        ws.cell(row=8, column=1).border = thin_border
        ws.cell(row=8, column=2, value="J&V RESGUARDO SELVA SAC").font = font_normal
        ws.cell(row=8, column=2).alignment = center_align
        ws.cell(row=8, column=2).border = thin_border
        
        ws.cell(row=8, column=3, value="20493762789").font = font_normal
        ws.cell(row=8, column=3).alignment = center_align
        ws.cell(row=8, column=3).border = thin_border
        
        ws.merge_cells('D8:G8')
        ws.cell(row=8, column=4, value="JR. NAUTA 269 - IQUITOS").font = font_normal
        ws.cell(row=8, column=4).alignment = center_align
        apply_border_to_range(ws, 'D8', 'G8', thin_border)
        
        ws.merge_cells('H8:I8')
        ws.cell(row=8, column=8, value="VIGILANCIA PRIVADA").font = font_normal
        ws.cell(row=8, column=8).alignment = center_align
        apply_border_to_range(ws, 'H8', 'I8', thin_border)

        # Fila 9
        ws.cell(row=9, column=1).border = thin_border
        ws.cell(row=9, column=2, value="LIDERMAN SERVICIOS SAC").font = font_normal
        ws.cell(row=9, column=2).alignment = center_align
        ws.cell(row=9, column=2).border = thin_border
        
        ws.cell(row=9, column=3, value="20601355761").font = font_normal
        ws.cell(row=9, column=3).alignment = center_align
        ws.cell(row=9, column=3).border = thin_border
        
        ws.merge_cells('D9:G9')
        ws.cell(row=9, column=4, value="AV. DEFENSORES DEL MORRO N°1620 - CHORRILLOS").font = font_normal
        ws.cell(row=9, column=4).alignment = center_align
        apply_border_to_range(ws, 'D9', 'G9', thin_border)
        
        ws.merge_cells('H9:I9')
        ws.cell(row=9, column=8, value="ACTIVIDADES DE TRANSPORTE").font = font_normal
        ws.cell(row=9, column=8).alignment = center_align
        apply_border_to_range(ws, 'H9', 'I9', thin_border)
        
        # Fila 10
        ws.cell(row=10, column=1).border = thin_border
        ws.cell(row=10, column=2, value="J&V ALARMAS S.A.C.").font = font_normal
        ws.cell(row=10, column=2).alignment = center_align
        ws.cell(row=10, column=2).border = thin_border
        
        ws.cell(row=10, column=3, value="20303166573").font = font_normal
        ws.cell(row=10, column=3).alignment = center_align
        ws.cell(row=10, column=3).border = thin_border
        
        ws.merge_cells('D10:G10')
        ws.cell(row=10, column=4, value="AV. DEFENSORES DEL MORRO N°1620 - CHORRILLOS").font = font_normal
        ws.cell(row=10, column=4).alignment = center_align
        apply_border_to_range(ws, 'D10', 'G10', thin_border)
        
        ws.merge_cells('H10:I10')
        ws.cell(row=10, column=8, value="ACTIVIDAD DE INVESTIGACIÓN").font = font_normal
        ws.cell(row=10, column=8).alignment = center_align
        apply_border_to_range(ws, 'H10', 'I10', thin_border)

        # --- FILA 11: Tema/Motivo ---
        ws.row_dimensions[11].height = 39.60
        
        ws.merge_cells('A11:B11')
        ws['A11'].value = "Tema/Motivo:"
        ws['A11'].font = font_bold_10
        ws['A11'].alignment = center_align
        apply_border_to_range(ws, 'A11', 'B11', thin_border)
        
        ws.merge_cells('C11:E11')
        ws['C11'].value = course_details['Tema/Motivo']
        ws['C11'].font = font_normal
        ws['C11'].alignment = center_align
        apply_border_to_range(ws, 'C11', 'E11', thin_border)
        
        ws.merge_cells('F11:G11')
        ws['F11'].value = "Grabación/ Material:"
        ws['F11'].font = font_bold_10
        ws['F11'].alignment = center_align
        apply_border_to_range(ws, 'F11', 'G11', thin_border)
        
        ws.merge_cells('H11:I11')
        ws['H11'].value = course_details['Grabacion/ Material']
        ws['H11'].font = font_normal
        ws['H11'].alignment = center_align
        apply_border_to_range(ws, 'H11', 'I11', thin_border)

        # --- FILA 12: Contenido ---
        ws.row_dimensions[12].height = 285
        
        ws.merge_cells('A12:B12')
        ws['A12'].value = "Contenido/ Sub Temas:"
        ws['A12'].font = font_bold_10
        ws['A12'].alignment = top_center_align
        apply_border_to_range(ws, 'A12', 'B12', thin_border)
        
        ws.merge_cells('C12:I12')
        ws['C12'].value = course_details['Contenido/ Sub Temas']
        ws['C12'].font = font_normal
        ws['C12'].alignment = top_left_align
        apply_border_to_range(ws, 'C12', 'I12', thin_border)

        # --- FILA 13: Capacitador/Firma ---
        ws.row_dimensions[13].height = 50
        
        ws.merge_cells('A13:B13')
        ws['A13'].value = "Capacitador/Entrenador:"
        ws['A13'].font = font_bold_10
        ws['A13'].alignment = center_align
        apply_border_to_range(ws, 'A13', 'B13', thin_border)
        
        ws.merge_cells('C13:D13')
        ws['C13'].value = course_details['Capacitador/Entrenador']
        ws['C13'].font = font_normal
        ws['C13'].alignment = center_align
        apply_border_to_range(ws, 'C13', 'D13', thin_border)
        
        ws['E13'].value = "Firma:"
        ws['E13'].font = font_bold_10
        ws['E13'].alignment = center_align
        ws['E13'].border = thin_border
        
        # Firma (Imagen) - Ajustada al tamaño de la celda
        ws['F13'].border = thin_border
        if os.path.exists(FIRMA_PATH):
            try:
                img_firma_cap = OpenpyxlImage(FIRMA_PATH)
                img_firma_cap.width = 120  # Píxeles
                img_firma_cap.height = 45   # Píxeles
                ws['F13'].alignment = Alignment(horizontal='center', vertical='center')
                ws.add_image(img_firma_cap, 'F13')
            except Exception as e:
                print(f"Advertencia: No se pudo cargar la firma del capacitador: {e}")
        
        ws['G13'].value = "Duración:"
        ws['G13'].font = font_bold_10
        ws['G13'].alignment = center_align
        ws['G13'].border = thin_border
        
        ws.merge_cells('H13:I13')
        ws['H13'].value = course_details['Duracion']
        ws['H13'].font = font_normal
        ws['H13'].alignment = center_align
        apply_border_to_range(ws, 'H13', 'I13', thin_border)

        # --- FILA 14: Motivo ---
        ws.merge_cells('A14:I14')
        ws['A14'].value = "Motivo:"
        ws['A14'].font = font_bold_10
        ws['A14'].alignment = left_align
        apply_border_to_range(ws, 'A14', 'I14', thin_border)

        # --- FILA 15: Tipo de actividad y N° de Trabajadores ---
        ws.merge_cells('A15:F15')
        ws['A15'].value = "Inducción ( ) Capacitación (X) Entrenamiento ( ) Charla de 5 minutos ( )"
        ws['A15'].font = font_normal
        ws['A15'].alignment = left_align
        apply_border_to_range(ws, 'A15', 'F15', thin_border)
        
        ws['G15'].value = "N° de Trabajadores:"
        ws['G15'].font = font_bold_10
        ws['G15'].alignment = center_align
        ws['G15'].border = thin_border
        
        ws.merge_cells('H15:I15')
        ws['H15'].value = len(df_course)
        ws['H15'].font = font_normal
        ws['H15'].alignment = center_align
        apply_border_to_range(ws, 'H15', 'I15', thin_border)

        # --- FILA 16: Cabeceras de la tabla de datos ---
        data_header_row = 16
        headers = ["N°", "Apellidos y Nombres", "DNI", "Unidad (Cliente)", "Nota", "Fecha Examen", "Hora Conexión", "Observación"]
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=data_header_row, column=col_idx, value=header)
            cell.font = font_bold_10
            cell.alignment = center_align
            cell.border = thin_border
            cell.fill = grey_fill

         # Agregar columna "Observación" que abarca H16:I16
        ws.merge_cells('H16:I16')
        ws['H16'].value = "Observación"
        ws['H16'].font = font_bold_10
        ws['H16'].alignment = center_align
        ws['H16'].fill = grey_fill
        apply_border_to_range(ws, 'H16', 'I16', thin_border)

        # --- Pegar los datos del DataFrame ---
        rows = df_course.values.tolist()
        for row_idx, row_data in enumerate(rows, start=data_header_row + 1):
            for col_idx, cell_value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                cell.font = font_normal
                cell.alignment = left_align if col_idx == 2 else center_align
                cell.border = thin_border

        # --- Pie de página: Responsable del Registro ---
        footer_start_row = data_header_row + len(df_course) + 3

        ws.merge_cells(f'A{footer_start_row}:C{footer_start_row}')
        ws[f'A{footer_start_row}'].value = "Responsable del Registro:"
        ws[f'A{footer_start_row}'].font = font_bold_10

        ws.merge_cells(f'A{footer_start_row+1}:C{footer_start_row+1}')
        ws[f'A{footer_start_row+1}'].value = "Apellidos y Nombres: Ciudad Olano, Karina Alejandra"
        ws[f'A{footer_start_row+1}'].font = font_normal
        ws[f'A{footer_start_row+1}'].alignment = left_align

        ws.merge_cells(f'A{footer_start_row+2}:B{footer_start_row+2}')
        ws[f'A{footer_start_row+2}'].value = "Cargo: Coordinadora de Capacitación y Desarrollo"
        ws[f'A{footer_start_row+2}'].font = font_normal
        ws[f'A{footer_start_row+2}'].alignment = left_align

        ws.merge_cells(f'D{footer_start_row+1}:E{footer_start_row+1}')
        ws[f'D{footer_start_row+1}'].value = "Firma:"
        ws[f'D{footer_start_row+1}'].font = font_bold_10
        ws[f'D{footer_start_row+1}'].alignment = left_align

        if os.path.exists(FIRMA_PATH):
            try:
                img_firma = OpenpyxlImage(FIRMA_PATH)
                img_firma.width = 100
                img_firma.height = 50
                ws.add_image(img_firma, f'F{footer_start_row + 1}')
            except Exception as e:
                print(f"Advertencia: No se pudo cargar la firma: {e}")

        ws.merge_cells(f'D{footer_start_row+2}:E{footer_start_row+2}')
        ws[f'D{footer_start_row+2}'].value = "Fecha:"
        ws[f'D{footer_start_row+2}'].font = font_bold_10
        ws[f'D{footer_start_row+2}'].alignment = left_align

        ws.merge_cells(f'F{footer_start_row+2}:G{footer_start_row+2}')
        ws[f'F{footer_start_row+2}'].value = datetime.date.today().strftime('%d/%m/%Y')
        ws[f'F{footer_start_row+2}'].font = font_normal
        ws[f'F{footer_start_row+2}'].alignment = left_align

        # Guardar el libro de trabajo en el buffer de memoria
        wb.save(output)
        return output.getvalue()

    except Exception as e:
        print(f"Error fatal al crear Excel: {e}")
        return None