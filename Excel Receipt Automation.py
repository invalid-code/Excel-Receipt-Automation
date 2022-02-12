import openpyxl.styles as styles
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side
wb = Workbook()
ws = wb.active
count = 1
global_border = Border(top=Side(border_style="thin"), bottom=Side(border_style="thin"), right=Side(border_style="thin"), left=Side(border_style="thin"))
ws["A1"] = "Name of Company"
ws["A1"].border = global_border 
ws.append(["Product Name", "Prev. Inv.", "Pull-Out Stocks", "Total Stocks", "SOH", "Stocks Sold", "Unit", "Amount"])
while ws[get_column_letter(count) + "2"].value:
	ws[get_column_letter(count) + "2"].border = global_border
	count += 1
ws.merge_cells("A1:" + get_column_letter(count-1) + "1")
wb.save("C:\\Users\\JessG\\Documents\\Mom's Work\\test.xlsx")