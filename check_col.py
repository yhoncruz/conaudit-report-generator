import win32com.client
excel = win32com.client.Dispatch('Excel.Application')
wb = excel.Workbooks.Open(r'g:\Mi unidad\Amorsh_toxic\Conaudit\Reporte_Cierre_Paquete.xlsx')
ws = wb.Sheets('Resumen por Razón Social')

print(f"Header 1: {ws.Cells(1,1).Value}")
header_row = 1
col_razon = 2
c = col_razon
while ws.Cells(header_row, c).Value is not None and str(ws.Cells(header_row, c).Value).strip() != '':
    print(f"Col {c}: {ws.Cells(header_row, c).Value}")
    c += 1
print('last_col File 1:', c-1)

wb2 = excel.Workbooks.Open(r'g:\Mi unidad\Amorsh_toxic\Conaudit\Reporte_Cierre_Paquete_2.xlsx')
ws2 = wb2.Sheets('Resumen por Razón Social')
c = col_razon
while ws2.Cells(header_row, c).Value is not None and str(ws2.Cells(header_row, c).Value).strip() != '':
    print(f"Col {c}: {ws2.Cells(header_row, c).Value}")
    c += 1
print('last_col File 2:', c-1)

wb.Close()
wb2.Close()
excel.Quit()
