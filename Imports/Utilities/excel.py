from .const import *
from .fileManagement import *
from ..Classes.Component import *
from ..Classes.Task import *

from openpyxl import load_workbook

import os

def loadReadOnlyWorkbook():
	workBook = load_workbook(selectTaskFile(), read_only = True)
	assertExpectedWorkSheets(TASKS, workBook)
	return workBook


def assertExpectedWorkSheets(fileType, workBook):
	for sheet in FILE_SHEETS[fileType]:
		if not sheet in workBook.sheetnames:
			quit()


def generateTaskList():
	taskList = []

	taskWorkBook = loadReadOnlyWorkbook()
	taskSheet = taskWorkBook['tasks']

	productDictionary = generateProductDictionary(taskWorkBook)

	for i in range(2, taskSheet.max_row):
		product = str(taskSheet.cell(row = i, column = 1).value)
		if product == "None":
			break
		quantity = taskSheet.cell(row = i, column = 2).value
		components = productDictionary[product]
		taskList.append(Task(product, quantity, components))

def generateProductDictionary(taskWorkBook):
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
