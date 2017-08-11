from openpyxl import *


def read_spreadsheet(filename):
	wb = load_workbook(filename)
	sheet = wb.active
	cells = sheet['A1' : 'B7']
	for c1, c2 in cells:
		print("{0:8} {1:8}".format(c1.value, c2.value))

if __name__ == '__main__':
	read_spreadsheet(input('Enter filename: '))
