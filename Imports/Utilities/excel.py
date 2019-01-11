from .const import *
from ..Classes.Component import *

from openpyxl import load_workbook # TODO look into this 
import os

def generateTaskList():
	taskList = []

def generateProductDictionary():
	# TODO allow user to select filename
	# TODO assert existence of files, allow backout


	taskWorkBook = load_workbook(os.path.join(TASK_LISTS_DIR_PATH, "gtProductAssembly010819.xlsx"), read_only = True)
	# TODO assert existence of sheet
	componentSheet = taskWorkBook['components']
	# TODO assert worksheet dimensions from ws.calculate_dimension() sheet.get_highest_column()
	# TODO assert column headers are as expected

	rowMax = componentSheet.max_row
	# use regex to assert expected formatting of rows
	productDictionary = {}
	for i in range(2, componentSheet.max_row):
		product = str(componentSheet.cell(row = i, column = 1).value)
		part = str(componentSheet.cell(row = i, column = 2).value)
		quantity = componentSheet.cell(row = i, column = 3).value

		if product not in productDictionary.keys():
			productDictionary[product] = []

		productDictionary[product].append(Component(part, quantity))

	return productDictionary
