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
	while ws[get_column_letter(count) + "3"].value:
		ws[get_column_letter(count) + "3"].border = global_border
		count += 1
	ws.merge_cells("A1:" + get_column_letter(count-1) + "1")
	ws.merge_cells("A2:" + get_column_letter(count-1) + "2")
def main():
	section = ["Product Name", "Prev. Inv.", "Pull-Out Stocks", "Total Stocks", "SOH", "Stocks Sold", "Unit", "Amount"]
	ws["A1"] = "Name of Company"
	ws["A2"] = "nothing" 
	ws.append(section)
	styling()
	wb.save("C:\\Users\\JessG\\Desktop\\Mom's Work\\test.xlsx")
if __name__ == '__main__':
	main()