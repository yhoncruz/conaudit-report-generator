import win32com.client
import os
import time

def main():
    base_dir = os.path.abspath(os.path.dirname(__file__))
    
    # LISTA DE ARCHIVOS ESCALABLE
    # Puedes agregar más archivos aquí en el futuro
    excel_files = [
        'Reporte_Cierre_Paquete.xlsx',
        'Reporte_Cierre_Paquete_2.xlsx'
    ]
    
    pptx_path = os.path.join(base_dir, 'Presentación Reunión de Seguimiento.pptx')
    out_pptx_path = os.path.join(base_dir, 'Presentacion_Actualizada_Fase3.pptx')

    if os.path.exists(out_pptx_path):
        try: os.remove(out_pptx_path)
        except: pass

    print("Iniciando aplicaciones Office...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    
    workbooks = []
    try:
        pres = ppt.Presentations.Open(pptx_path, WithWindow=False)
        
        # Eliminar tabla de muestra
        base_slide = pres.Slides(1)
        shapes_to_delete = []
        for i in range(1, base_slide.Shapes.Count + 1):
            if base_slide.Shapes(i).Type == 13 and base_slide.Shapes(i).Width > 200:
                shapes_to_delete.append(base_slide.Shapes(i))
        for shape in shapes_to_delete:
            shape.Delete()

        # Datos de mapeo
        # master_razones: conjunto de todas las razones sociales únicas
        # file_info: info por cada archivo (ws, header_row, col_razon, last_col)
        # data_map: diccionario {razon_social: {file_index: row_number}}
        master_razones = set()
        file_info = []
        data_map = {}

        print("Analizando archivos Excel...")
        for file_idx, filename in enumerate(excel_files):
            file_path = os.path.join(base_dir, filename)
            if not os.path.exists(file_path):
                print(f"Advertencia: No se encontró {filename}")
                continue
                
            # Preguntar al usuario por el título de este archivo
            titulo = input(f"Ingrese el título para la tabla del archivo {filename}: ")
                
            wb = excel.Workbooks.Open(file_path)
            workbooks.append(wb)
            
            try:
                ws = wb.Sheets("Resumen por Razón Social")
            except:
                print(f"Advertencia: No hay hoja 'Resumen por Razón Social' en {filename}")
                continue
                
            # Encontrar encabezados
            header_row = 1
            col_razon = 1
            last_col = 7
            
            for r in range(1, 10):
                for c in range(1, 10):
                    val = ws.Cells(r, c).Value
                    if val and "Razon_Social_Reclamante" in str(val):
                        header_row = r
                        col_razon = c
                        break
                if header_row != 1 or col_razon != 1:
                    if "Razon_Social_Reclamante" in str(ws.Cells(r, col_razon).Value):
                         break
            
            c = col_razon
            last_col = c
            while ws.Cells(header_row, c).Value is not None and str(ws.Cells(header_row, c).Value).strip() != "":
                val_str = str(ws.Cells(header_row, c).Value).strip()
                if "TIPIFICACION" in val_str:
                    last_col = c - 1
                    break
                if "Porcentaje_Glosado" in val_str:
                    last_col = c
                    break
                c += 1
                last_col = c - 1
            
            file_info.append({
                'ws': ws,
                'header_row': header_row,
                'col_razon': col_razon,
                'last_col': last_col,
                'filename': filename,
                'titulo': titulo
            })

            # Mapear filas
            current_row = header_row + 1
            while True:
                val = ws.Cells(current_row, col_razon).Value
                if val is None or str(val).strip() == "":
                    current_row += 1
                    if current_row > header_row + 1000: break
                    continue
                    
                razon = str(val).strip()
                if razon == "TOTAL GENERAL":
                    break
                    
                master_razones.add(razon)
                if razon not in data_map:
                    data_map[razon] = {}
                data_map[razon][file_idx] = current_row
                current_row += 1

        print(f"Total de razones sociales únicas encontradas: {len(master_razones)}")
        
        # Iterar sobre las razones sociales únicas (ordenadas alfabéticamente)
        slide_index = 1
        sorted_razones = sorted(list(master_razones))
        
        # Usaremos el primer workbook para crear la hoja temporal
        wb_main = workbooks[0]
        
        for razon in sorted_razones:
            print(f"Procesando: {razon}")
            
            ws_temp = wb_main.Sheets.Add(After=wb_main.Sheets(wb_main.Sheets.Count))
            ws_temp.Name = "TempCopy"
            
            temp_current_row = 1
            
            # Recorrer cada archivo y si la razon existe en él, copiar
            for file_idx, info in enumerate(file_info):
                if file_idx in data_map[razon]:
                    source_row = data_map[razon][file_idx]
                    ws_source = info['ws']
                    h_row = info['header_row']
                    c_razon = info['col_razon']
                    l_col = info['last_col']
                    titulo_archivo = info['titulo']
                    
                    # Pintar título
                    if titulo_archivo.strip():
                        ws_temp.Cells(temp_current_row, 1).Value = titulo_archivo
                        ws_temp.Cells(temp_current_row, 1).Font.Bold = True
                    temp_current_row += 1
                    
                    # Copiar encabezados
                    ws_source.Range(ws_source.Cells(h_row, c_razon), ws_source.Cells(h_row, l_col)).Copy()
                    ws_temp.Cells(temp_current_row, 1).PasteSpecial(Paste=-4104) # xlPasteAll
                    ws_temp.Cells(temp_current_row, 1).PasteSpecial(Paste=8)    # xlPasteColumnWidths
                    
                    # Copiar fila de datos
                    ws_source.Range(ws_source.Cells(source_row, c_razon), ws_source.Cells(source_row, l_col)).Copy()
                    ws_temp.Cells(temp_current_row + 1, 1).PasteSpecial(Paste=-4104)
                    
                    temp_current_row += 3 # Deja una fila en blanco
            
            # Copiar imagen compuesta
            copy_range = ws_temp.Range(ws_temp.Cells(1, 1), ws_temp.Cells(temp_current_row - 2, info['last_col'] - info['col_razon'] + 1))
            
            success = False
            for attempt in range(5):
                try:
                    copy_range.CopyPicture(Appearance=1, Format=2)
                    success = True
                    break
                except:
                    time.sleep(1)
            
            if not success:
                print(f"Fallo al copiar imagen para {razon}")
                
            # PowerPoint: Siempre duplicamos la plantilla limpia (Slide 1)
            slide_range = pres.Slides(1).Duplicate()
            slide = slide_range(1)
            slide.MoveTo(pres.Slides.Count)
                
            # Título
            title_shape = None
            for shape in slide.Shapes:
                if shape.HasTextFrame and "Resumen Ejecutivo" in shape.TextFrame.TextRange.Text:
                    shape.TextFrame.TextRange.Text = f"Resumen Ejecutivo de Gestión de\n{razon}"
                    title_shape = shape
                    break
            
            if not title_shape and slide.Shapes.HasTitle:
                slide.Shapes.Title.TextFrame.TextRange.Text = f"Resumen Ejecutivo de Gestión de\n{razon}"
                title_shape = slide.Shapes.Title
                
            # Pegar
            paste_success = False
            for attempt in range(5):
                try:
                    slide.Shapes.Paste()
                    paste_success = True
                    break
                except:
                    time.sleep(1)
                    
            if paste_success:
                pic = slide.Shapes(slide.Shapes.Count)
                if title_shape:
                    pic.Top = title_shape.Top + title_shape.Height + 30
                else:
                    pic.Top = 150
                    
                max_width = pres.PageSetup.SlideWidth * 0.95
                if pic.Width > max_width:
                    pic.LockAspectRatio = True
                    pic.Width = max_width
                    
                pic.Left = (pres.PageSetup.SlideWidth - pic.Width) / 2
            else:
                print(f"Fallo al pegar imagen para {razon}")
                
            # Borrar hoja temporal
            excel.DisplayAlerts = False
            ws_temp.Delete()
            excel.DisplayAlerts = True
            
            slide_index += 1

        # Eliminar la plantilla limpia al final
        pres.Slides(1).Delete()

        print("Guardando presentación...")
        pres.SaveAs(out_pptx_path)
        print(f"Guardado como {out_pptx_path}")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        try: pres.Close()
        except: pass
        for wb in workbooks:
            try: wb.Close(SaveChanges=False)
            except: pass
        ppt.Quit()
        excel.Quit()

if __name__ == "__main__":
    main()
