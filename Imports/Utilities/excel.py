from .const import *
from .fileManagement import *
from ..Classes.Component import *
from ..Classes.Task import *

from openpyxl import load_workbook

import os

def loadWorkbook(fileType, filePath, readOnly = False):
	workBook = load_workbook(filePath, readOnly)
	assertExpectedWorkSheets(fileType, workBook)
	return workBook


def assertExpectedWorkSheets(fileType, workBook):
	for sheet in FILE_SHEETS[fileType]:
		if not sheet in workBook.sheetnames:
			print('Selected task file is missing ' + sheet + ' worksheet.')
			quit()


def generateTaskList():
	taskList = []

	workBook = loadWorkbook(TASKS, selectTaskFile(), True)
	taskSheet = workBook['tasks']

	productDictionary = generateProductDictionary(workBook)

	for i in range(2, taskSheet.max_row):
		product = str(taskSheet.cell(row = i, column = 1).value)
		if product == "None":
			break
		quantity = taskSheet.cell(row = i, column = 2).value
		components = productDictionary[product]
		taskList.append(Task(product, quantity, components))

	return taskList


def generateProductDictionary(workBook):
	componentSheet = workBook['components']
	# TODO assert worksheet dimensions from ws.calculate_dimension() sheet.get_highest_column()
	# TODO assert column headers are as expected
	# TODO use regex to assert expected formatting of rows
	productDictionary = {}
	for i in range(2, componentSheet.max_row):
		product = str(componentSheet.cell(row = i, column = 1).value)
		part = str(componentSheet.cell(row = i, column = 2).value)
		quantity = componentSheet.cell(row = i, column = 3).value

		if product not in productDictionary.keys():
			productDictionary[product] = []

		productDictionary[product].append(Component(part, quantity))

	return productDictionary


def populateAssemblyOrderFile(taskList):
	workBook = loadWorkbook(TEMPLATE, ASSEMBLY_ORDER_FILE_NAME_PATH)
	for i in range (0, len(taskList)):
		pass