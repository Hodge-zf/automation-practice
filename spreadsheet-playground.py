from openpyxl import *
from collections import Counter

def read_spreadsheet(filename):
	wb = load_workbook(filename)
	sheet = wb.active
	columns = sheet[input('First column in range: ') : input('Last column in range: ')]
	column1 = []
	column2 = []
	for c1, c2 in columns:
		column1.append(c1.value)
		column2.append(c2.value)
		print("{0:8} {1:8}".format(c1.value, c2.value))

	return column1, column2


def calculate_mean(numbers):
	s = sum(numbers)
	n = len(numbers)
	mean = s/n

	return mean


def calculate_median(numbers):
	n = len(numbers)
	numbers.sort()

	if n % 2 == 0:
		m1 = n/2
		m2 = (n/2) + 1

		m1 = int(m1) - 1
		m2 = int(m2) - 1
		median = (numbers[m1] + numbers[m2]) / 2

	else:
		m = (n + 1) / 2
		m = int(m) - 1
		median = numbers[m]

	return median


def calculate_mode(numbers):
	c = Counter(numbers)
	numbers_freq = c.most_common()
	max_count = numbers_freq[0][1]

	modes = []
	for num in numbers_freq:
		if num[1] == max_count:
			modes.append(num[0])
	return modes


if __name__ == '__main__':
	column1, column2 = read_spreadsheet(input('Enter filename: '))

	mean1 = calculate_mean(column1)
	print('The mean of column 1 is: '+str(mean1))

	median1 = calculate_median(column1)
	print('The median of column 1 is: '+str(median1))

	mode1 = calculate_mode(column1)
	print('The mode(s) of the list are: ')
	for mode in mode1:
		print(mode)

	mean2 = calculate_mean(column2)
	print('The mean of column 2 is: '+str(mean2))

	median2 = calculate_median(column2)
	print('The median of column 2 is: '+str(median2))

	mode2 = calculate_mode(column2)
	print('The mode(s) of the list are: ')
	for mode in mode2:
		print(mode)
