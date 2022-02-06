import openpyxl as xl
wb = xl.Workbook()
ws = wb.active
print(wb,ws)
file_location = "C:\\Users\\JessG\\Documents\\Mom's Work\\test.xlsx"

with open(file_location, "w") as f:
	f.write("")
help(xl)