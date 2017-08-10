from openpyxl import *



def read_spreadsheet(filename):
	wb = load_workbook(filename)
	sheet_name = wb.get_sheet_by_name(input('Enter sheet name: '))
	need_cell = sheet_name[input('Enter cell name: ')].value
	print(need_cell)

if __name__ == '__main__':
	read_spreadsheet(input('Enter filename: '))
