import win32com.client
import os
import time
import pandas as pd

def rgb_to_int(r, g, b):
    return r + (g << 8) + (b << 16)

def procesar_y_crear_hoja_glosas(excel, base_dir, reporte_filename):
    glosas_filename = reporte_filename.replace("Reporte_Cierre_Paquete", "GLOSAS_RECLAMACIONES")
    glosas_path = os.path.join(base_dir, glosas_filename)
    
    if not os.path.exists(glosas_path):
        print(f"Advertencia: No se encontró el archivo de glosas {glosas_filename}")
        return None
        
    print(f"Procesando glosas para {glosas_filename}...")
    
    try:
        df = pd.read_excel(glosas_path, sheet_name="GLOSAS_RECLAMACIONES")
    except Exception as e:
        print(f"Error al leer hoja GLOSAS_RECLAMACIONES de {glosas_filename}: {e}")
        return None
        
    # Limpieza de datos
    df = df.dropna(subset=["IPS", "reclamacionID", "Codigo_glosa"])
    df["IPS"] = df["IPS"].astype(str).str.strip()
    df["Codigo_glosa"] = df["Codigo_glosa"].astype(str).str.strip()
    df["Descripcion_glosa"] = df["Descripcion_glosa"].fillna("").astype(str).str.strip()
    df["reclamacionID"] = df["reclamacionID"].astype(int)
    
    # Calcular estadísticas agrupando por IPS, código de glosa y descripción
    agg = df.groupby(["IPS", "Codigo_glosa", "Descripcion_glosa"]).agg(
        Q_Reclamaciones=("reclamacionID", "nunique"),
        Q_items=("reclamacionID", "count")
    ).reset_index()
    
    ips_totals = agg.groupby("IPS")["Q_items"].transform("sum")
    agg["% participacion"] = agg["Q_items"] / ips_totals
    
    # Ordenar por IPS (alfabético) y luego por % participacion de mayor a menor (descendente)
    agg = agg.sort_values(by=["IPS", "% participacion"], ascending=[True, False])
    
    wb = excel.Workbooks.Open(glosas_path)
    
    try:
        sheet_name = "TD_GLOSAS_RECLAMACIONES"
        excel.DisplayAlerts = False
        try:
            ws = wb.Sheets(sheet_name)
            ws.Delete()
        except:
            pass
        excel.DisplayAlerts = True
            
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheet_name
        
        ws.Activate()
        excel.ActiveWindow.DisplayGridlines = True
        
        # Constantes de diseño
        header_color = rgb_to_int(217, 226, 243) # azul/gris claro (RGB 217, 226, 243)
        border_color = rgb_to_int(192, 192, 192) # gris claro (RGB 192, 192, 192)
        
        unique_ips = sorted(agg["IPS"].unique())
        ips_ranges = {}
        
        current_row = 1
        for ips in unique_ips:
            ips_data = agg[agg["IPS"] == ips]
            if len(ips_data) == 0:
                continue
                
            start_table_row = current_row
            
            # Fila 1: Filtro IPS
            ws.Cells(current_row, 1).Value = "IPS"
            ws.Cells(current_row, 1).Font.Bold = True
            ws.Cells(current_row, 2).Value = ips
            ws.Cells(current_row, 2).Font.Bold = True
            
            # Bordes para el filtro
            filter_range = ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 2))
            for b_id in [7, 8, 9, 10]:
                filter_range.Borders(b_id).LineStyle = 1
                filter_range.Borders(b_id).Weight = 2
                filter_range.Borders(b_id).Color = border_color
            
            current_row += 2 # Fila vacía, luego fila de cabecera
            header_row_idx = current_row
            
            # Cabeceras de la tabla
            headers = ["Codigo_glosa", "Descripcion_glosa", "Q Reclamaciones", "Q items", "% participacion"]
            for col_idx, h in enumerate(headers, 1):
                ws.Cells(current_row, col_idx).Value = h
            
            # Formatear cabecera
            header_range = ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 5))
            header_range.Font.Bold = True
            header_range.Interior.Color = header_color
            header_range.HorizontalAlignment = -4108 # xlCenter
            
            # Escribir las filas de datos ordenadas
            for _, row_data in ips_data.iterrows():
                current_row += 1
                
                # Codigo_glosa
                cell = ws.Cells(current_row, 1)
                cell.Value = str(row_data["Codigo_glosa"])
                cell.HorizontalAlignment = -4108
                
                # Descripcion_glosa
                cell = ws.Cells(current_row, 2)
                cell.Value = str(row_data["Descripcion_glosa"])
                cell.HorizontalAlignment = -4131
                
                # Q Reclamaciones
                cell = ws.Cells(current_row, 3)
                cell.Value = int(row_data["Q_Reclamaciones"])
                cell.NumberFormatLocal = "#.##0"
                cell.HorizontalAlignment = -4152
                
                # Q items
                cell = ws.Cells(current_row, 4)
                cell.Value = int(row_data["Q_items"])
                cell.NumberFormatLocal = "#.##0"
                cell.HorizontalAlignment = -4152
                
                # % participacion
                cell = ws.Cells(current_row, 5)
                cell.Value = float(row_data["% participacion"])
                cell.NumberFormatLocal = "0,00%"
                cell.HorizontalAlignment = -4152
            
            # Fila de Total general
            current_row += 1
            
            ws.Cells(current_row, 1).Value = "Total general"
            ws.Cells(current_row, 2).Value = ""
            
            ips_rows = df[df["IPS"] == ips]
            total_rec = ips_rows["reclamacionID"].nunique()
            cell_total_rec = ws.Cells(current_row, 3)
            cell_total_rec.Value = int(total_rec)
            cell_total_rec.NumberFormatLocal = "#.##0"
            cell_total_rec.HorizontalAlignment = -4152
            
            total_items = ips_data["Q_items"].sum()
            cell_total_items = ws.Cells(current_row, 4)
            cell_total_items.Value = int(total_items)
            cell_total_items.NumberFormatLocal = "#.##0"
            cell_total_items.HorizontalAlignment = -4152
            
            cell_total_pct = ws.Cells(current_row, 5)
            cell_total_pct.Value = 1.0
            cell_total_pct.NumberFormatLocal = "0,00%"
            cell_total_pct.HorizontalAlignment = -4152
            
            total_range = ws.Range(ws.Cells(current_row, 1), ws.Cells(current_row, 5))
            total_range.Font.Bold = True
            total_range.Interior.Color = header_color
            
            end_table_row = current_row
            
            # Aplicar bordes a toda la tabla
            table_range = ws.Range(ws.Cells(header_row_idx, 1), ws.Cells(end_table_row, 5))
            for b_id in range(7, 13):
                try:
                    table_range.Borders(b_id).LineStyle = 1
                    table_range.Borders(b_id).Weight = 2 # xlThin
                    table_range.Borders(b_id).Color = border_color
                except:
                    pass
            
            ips_ranges[ips] = (header_row_idx, end_table_row)
            
            current_row += 3 # Espaciado
            
        # Formatear anchos de columnas
        ws.Columns(1).ColumnWidth = 12
        ws.Columns(2).ColumnWidth = 85
        ws.Columns(2).WrapText = True
        ws.Columns(3).ColumnWidth = 15
        ws.Columns(4).ColumnWidth = 10
        ws.Columns(5).ColumnWidth = 15
        
        wb.Save()
        return wb, ws, ips_ranges
        
    except Exception as e:
        print(f"Error al estructurar hoja en {glosas_filename}: {e}")
        try: wb.Close(SaveChanges=False)
        except: pass
        return None

def main():
    base_dir = os.path.abspath(os.path.dirname(__file__))
    
    # LISTA DE ARCHIVOS ESCALABLE
    # Puedes agregar más archivos aquí en el futuro
    excel_files = [
        'Reporte_Cierre_Paquete.xlsx',
        'Reporte_Cierre_Paquete_2.xlsx'
    ]
    
    pptx_path = os.path.join(base_dir, 'Presentación Reunión de Seguimiento.pptx')
    out_pptx_path = os.path.join(base_dir, 'Presentacion_Actualizada_Fase4.pptx')

    if os.path.exists(out_pptx_path):
        try: os.remove(out_pptx_path)
        except: pass

    print("Iniciando aplicaciones Office...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    
    workbooks = []
    glosas_info = {}
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
            
            # Procesar el archivo de glosas correspondiente de forma dinámica
            info_glosas = procesar_y_crear_hoja_glosas(excel, base_dir, filename)
            if info_glosas:
                glosas_info[file_idx] = info_glosas
            
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
            
            # --- 2. DIAPOSITIVA DE MOTIVOS DE GLOSA (Nueva) ---
            for file_idx, info in enumerate(file_info):
                if file_idx in data_map[razon] and file_idx in glosas_info:
                    wb_g, ws_g, ips_ranges = glosas_info[file_idx]
                    
                    if razon in ips_ranges:
                        start_row, end_row = ips_ranges[razon]
                        print(f"Copiando tabla de glosas para {razon} desde {excel_files[file_idx]} (filas {start_row}-{end_row})")
                        
                        copy_range_glosas = ws_g.Range(ws_g.Cells(start_row, 1), ws_g.Cells(end_row, 5))
                        
                        success_g = False
                        for attempt in range(5):
                            try:
                                copy_range_glosas.CopyPicture(Appearance=1, Format=2)
                                success_g = True
                                break
                            except:
                                time.sleep(1)
                                
                        if success_g:
                            slide_range_g = pres.Slides(1).Duplicate()
                            slide_g = slide_range_g(1)
                            slide_g.MoveTo(pres.Slides.Count)
                            
                            title_text_g = f"Reporte de Gestión de motivos de glosa\n{razon}"
                            title_shape_g = None
                            
                            for shape in slide_g.Shapes:
                                if shape.HasTextFrame and "Resumen Ejecutivo" in shape.TextFrame.TextRange.Text:
                                    shape.TextFrame.TextRange.Text = title_text_g
                                    title_shape_g = shape
                                    break
                            
                            if not title_shape_g and slide_g.Shapes.HasTitle:
                                slide_g.Shapes.Title.TextFrame.TextRange.Text = title_text_g
                                title_shape_g = slide_g.Shapes.Title
                                
                            paste_success_g = False
                            for attempt in range(5):
                                try:
                                    slide_g.Shapes.Paste()
                                    paste_success_g = True
                                    break
                                except:
                                    time.sleep(1)
                                    
                            if paste_success_g:
                                pic_g = slide_g.Shapes(slide_g.Shapes.Count)
                                if title_shape_g:
                                    pic_g.Top = title_shape_g.Top + title_shape_g.Height + 30
                                else:
                                    pic_g.Top = 150
                                    
                                # Max width to avoid green bar on the right side
                                max_width = 760
                                if pic_g.Width > max_width:
                                    pic_g.LockAspectRatio = True
                                    pic_g.Width = max_width
                                    
                                # Center the table image within the white area (from 40 to 800 pixels)
                                pic_g.Left = 40 + (760 - pic_g.Width) / 2
                            else:
                                print(f"Fallo al pegar imagen de glosas para {razon}")
                        else:
                            print(f"Fallo al copiar imagen de glosas para {razon}")
            
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
        for file_idx, g_info in glosas_info.items():
            try: g_info[0].Close(SaveChanges=True)
            except: pass
        try: ppt.Quit()
        except: pass
        try: excel.Quit()
        except: pass

if __name__ == "__main__":
    main()
