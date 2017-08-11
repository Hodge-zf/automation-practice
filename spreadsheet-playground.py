from openpyxl import *


def read_spreadsheet(filename):
	wb = load_workbook(filename)
	sheet = wb.active
	cells = sheet[input('First cell in range: ') : input('Last cell in range: ')]
	cell1 = []
	cell2 = []
	for c1, c2 in cells:
		cell1.append(c1.value)
		cell2.append(c2.value)
		print("{0:8} {1:8}".format(c1.value, c2.value))

	return cell1, cell2


def calculate_mean(numbers):
	s = sum(numbers)
	n = len(numbers)
	mean = s/n

	return mean


if __name__ == '__main__':
	cell1, cell2 = read_spreadsheet(input('Enter filename: '))
	mean1 = calculate_mean(cell1)
	print('The mean of column 1 is: '+str(mean1))
	mean2 = calculate_mean(cell2)
	print('The mean of column 2 is: '+str(mean2))
