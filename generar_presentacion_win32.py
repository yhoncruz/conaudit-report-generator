import win32com.client
import os

def main():
    base_dir = os.path.abspath(os.path.dirname(__file__))
    excel_path = os.path.join(base_dir, 'Reporte_Cierre_Paquete.xlsx')
    pptx_path = os.path.join(base_dir, 'Presentación Reunión de Seguimiento.pptx')
    out_pptx_path = os.path.join(base_dir, 'Presentacion_Actualizada_v3.pptx')

    # Si ya existe el archivo de salida, intentar borrarlo
    if os.path.exists(out_pptx_path):
        try:
            os.remove(out_pptx_path)
        except:
            pass

    print("Iniciando aplicaciones Office...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    # PowerPoint requires Visible=True to manipulate some shapes sometimes, but let's try without it or with WithWindow=False
    
    try:
        wb = excel.Workbooks.Open(excel_path)
        ws_source = wb.Sheets("Resumen por Razón Social")
        
        pres = ppt.Presentations.Open(pptx_path, WithWindow=False)
        
        # Eliminar la tabla de muestra de la primera diapositiva (la plantilla) para que no se duplique
        # Sabemos que es una imagen (Type 13) y que es muy ancha (Width > 200) a diferencia del logo
        base_slide = pres.Slides(1)
        shapes_to_delete = []
        for i in range(1, base_slide.Shapes.Count + 1):
            shape = base_slide.Shapes(i)
            # Type 13 = msoPicture
            if shape.Type == 13 and shape.Width > 200:
                shapes_to_delete.append(shape)
                
        for shape in shapes_to_delete:
            shape.Delete()
            print("Tabla de muestra eliminada de la plantilla.")
        
        # Encontrar la fila de encabezados y columnas
        # Asumimos que los encabezados están en la fila 1 o 2. Buscaremos "Razon_Social_Reclamante"
        header_row = 1
        col_razon = 1
        last_col = 7
        
        for r in range(1, 10):
            for c in range(1, 10):
                val = ws_source.Cells(r, c).Value
                if val and "Razon_Social_Reclamante" in str(val):
                    header_row = r
                    col_razon = c
                    break
            if header_row != 1 or col_razon != 1:
                if "Razon_Social_Reclamante" in str(ws_source.Cells(r, col_razon).Value):
                     break
                     
        # Determinar last_col
        c = col_razon
        while ws_source.Cells(header_row, c).Value is not None and str(ws_source.Cells(header_row, c).Value).strip() != "":
            c += 1
        last_col = c - 1

        print(f"Encabezados encontrados en fila {header_row}, columnas {col_razon} a {last_col}")

        # Iterar sobre las filas de datos
        current_row = header_row + 1
        slide_index = 1 # Índice de la diapositiva en PowerPoint (1-based)
        
        while True:
            razon_social = ws_source.Cells(current_row, col_razon).Value
            
            if razon_social is None or str(razon_social).strip() == "":
                current_row += 1
                if current_row > header_row + 1000: # safety limit
                    break
                continue
                
            razon_str = str(razon_social).strip()
            if razon_str == "TOTAL GENERAL":
                print("Se encontró TOTAL GENERAL. Terminando.")
                break
                
            print(f"Procesando: {razon_str}")
            
            # Crear hoja temporal para copiar
            ws_temp = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
            ws_temp.Name = "TempCopy"
            
            # Copiar encabezados
            ws_source.Range(ws_source.Cells(header_row, col_razon), ws_source.Cells(header_row, last_col)).Copy()
            ws_temp.Cells(1, 1).PasteSpecial(Paste=-4104) # xlPasteAll
            ws_temp.Cells(1, 1).PasteSpecial(Paste=8)    # xlPasteColumnWidths
            
            # Copiar fila de datos
            ws_source.Range(ws_source.Cells(current_row, col_razon), ws_source.Cells(current_row, last_col)).Copy()
            ws_temp.Cells(2, 1).PasteSpecial(Paste=-4104)
            
            # Copiar el rango temporal como imagen
            import time
            copy_range = ws_temp.Range(ws_temp.Cells(1, 1), ws_temp.Cells(2, last_col - col_razon + 1))
            
            success = False
            for attempt in range(5):
                try:
                    copy_range.CopyPicture(Appearance=1, Format=2) # xlScreen, xlPicture
                    success = True
                    break
                except Exception as e:
                    time.sleep(1)
            
            if not success:
                print(f"Fallo al copiar imagen para {razon_str}")
            
            # Manipular PowerPoint
            if slide_index == 1:
                slide = pres.Slides(1)
            else:
                slide_range = pres.Slides(1).Duplicate()
                slide = slide_range(1)
                slide.MoveTo(pres.Slides.Count)
                
            # Modificar título
            title_shape = None
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    text = shape.TextFrame.TextRange.Text
                    if "Resumen Ejecutivo" in text:
                        shape.TextFrame.TextRange.Text = f"Resumen Ejecutivo de Gestión de\n{razon_str}"
                        title_shape = shape
                        break
            
            if title_shape is None and slide.Shapes.HasTitle:
                slide.Shapes.Title.TextFrame.TextRange.Text = f"Resumen Ejecutivo de Gestión de\n{razon_str}"
                title_shape = slide.Shapes.Title
                
            # Pegar imagen
            paste_success = False
            for attempt in range(5):
                try:
                    slide.Shapes.Paste()
                    paste_success = True
                    break
                except:
                    time.sleep(1)
                    
            if not paste_success:
                print(f"Fallo al pegar imagen para {razon_str}")
                continue
                
            # La imagen pegada es la última añadida
            pic = slide.Shapes(slide.Shapes.Count)
            
            # Posicionar 3 enters abajo del título (aprox)
            if title_shape:
                pic.Top = title_shape.Top + title_shape.Height + 30 # 30 points ~ 0.5 inches
            else:
                pic.Top = 150
                
            # Centrar horizontalmente y ajustar si es muy ancha
            max_width = pres.PageSetup.SlideWidth * 0.95
            if pic.Width > max_width:
                pic.LockAspectRatio = True
                pic.Width = max_width
                
            pic.Left = (pres.PageSetup.SlideWidth - pic.Width) / 2
            
            # Eliminar hoja temporal
            wb.Application.DisplayAlerts = False
            ws_temp.Delete()
            wb.Application.DisplayAlerts = True
            
            current_row += 1
            slide_index += 1
            
        print("Guardando presentación...")
        pres.SaveAs(out_pptx_path)
        print(f"Guardado como {out_pptx_path}")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        try:
            pres.Close()
        except:
            pass
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        ppt.Quit()
        excel.Quit()

if __name__ == "__main__":
    main()
