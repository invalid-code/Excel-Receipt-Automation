from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws["A1"] = "Name of Company" 
wb.save("C:\\Users\\JessG\\Documents\\Mom's Work\\test.xlsx")