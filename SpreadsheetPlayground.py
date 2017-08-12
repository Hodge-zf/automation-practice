from openpyxl import *
from collections import Counter
import matplotlib.pyplot as plt

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


def find_differences(numbers):
	mean = calculate_mean(numbers)
	diff = []
	for num in numbers:
		diff.append(num-mean)

	return diff


def calculate_variance(numbers):
	diff = find_differences(numbers)
	squared_diff = []
	for d in diff:
		squared_diff.append(d**2)
	sum_squared_diff = sum(squared_diff)
	variance = sum_squared_diff/len(numbers)
	return variance


def plot_points(col1, col2):
	plt.plot(col1, col2, marker='o')
	plt.xlabel('Column 1')
	plt.ylabel('Column 2')
	plt.title('Visualization of data sets')
	plt.show()


if __name__ == '__main__':
	column1, column2 = read_spreadsheet(input('Enter filename: '))

	mean1 = calculate_mean(column1)
	print('The mean of column 1 is {0}'.format(str(mean1)))

	median1 = calculate_median(column1)
	print('The median of column 1 is {0}'.format(str(median1)))

	mode1 = calculate_mode(column1)
	print('The mode(s) of the column 1 are: ')
	for mode in mode1:
		print(mode)

	variance1 = calculate_variance(column1)
	std1 = variance1**0.5
	print('The variance of column 1 is {0} and the standard deviation is {1}'.format(variance1, std1))

	mean2 = calculate_mean(column2)
	print('The mean of column 2 is {0}'.format(str(mean2)))

	median2 = calculate_median(column2)
	print('The median of column 2 is {0}'.format(str(median2)))

	mode2 = calculate_mode(column2)
	print('The mode(s) of the column 2 are: ')
	for mode in mode2:
		print(mode)

	variance2 = calculate_variance(column2)
	std2 = variance2**0.5
	print('The variance of column 2 is {0} and the standard deviation is {1}'.format(variance2, std2))

	plot_points(column1, column2)

