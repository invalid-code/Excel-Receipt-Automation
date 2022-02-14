from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Alignment
wb = Workbook()
ws = wb.active
def styling():
	count = 1
	global_border = Border(top=Side(border_style="thin"), bottom=Side(border_style="thin"), right=Side(border_style="thin"), left=Side(border_style="thin"))
	center_alignment = Alignment(horizontal="center")
	ws["A1"].border = global_border
	ws["A1"].alignment = center_alignment
	ws["A2"].border = global_border
	ws["E2"].border = global_border
	ws["E2"].alignment = Alignment(horizontal="right")
	while ws[get_column_letter(count) + "3"].value:
		ws[get_column_letter(count) + "3"].border = global_border
		count += 1
	ws.merge_cells("A1:" + get_column_letter(count-1) + "1")
	ws.merge_cells("A2:" + get_column_letter((count-1) // 2)  + "2")
	ws.merge_cells(get_column_letter((count+1) // 2) + "2:" + get_column_letter(count-1) + "2")
	for row in range(4, count+3):
		for col in range(1,count):
			ws[get_column_letter(col) + str(row)].border = global_border
def main():
	section = ["Product Name", "Prev. Inv.", "Pull-Out Stocks", "Total Stocks", "SOH", "Stocks Sold", "Unit", "Amount"]
	ws["A1"] = "Name of Company"
	ws["A2"] = "Supplier: " 
	ws["E2"] = "Date: "
	ws.append(section)
	styling()
	wb.save("C:\\Users\\JessG\\Desktop\\Mom's Work\\test.xlsx")
if __name__ == '__main__':
	main()