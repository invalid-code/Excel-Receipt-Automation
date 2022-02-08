from openpyxl import Workbook
from openpyxl.utils import get_column_letter
wb = Workbook()
ws = wb.active
count = 1
ws["A1"] = "Name of Company" 
ws.append(["Product Name", "Prev. Inv.", "Pull-Out Stocks", "Total Stocks", "SOH", "Stocks Sold", "Unit", "Amount"])
while ws[get_column_letter(count) + "2"].value:
	count += 1
ws.merge_cells("A1:" + get_column_letter(count) + "1")
wb.save("C:\\Users\\JessG\\Documents\\Mom's Work\\test.xlsx")