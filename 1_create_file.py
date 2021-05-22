from openpyxl import Workbook
wb = Workbook() 
ws = wb.active # 활성화된 시트
ws.title = "NadoSheet"
wb.save("sample.xlsx")
wb.close() 
# make new sheet and file whose name is made by me. 
