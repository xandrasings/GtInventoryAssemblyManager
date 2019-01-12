from .const import *
from .fileManagement import *
from .sys import *
from ..Classes.Component import *
from ..Classes.Task import *

from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles.borders import Border, Side

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
	taskSheet = workBook[TASK_FILE_TASK_SHEET]

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
	componentSheet = workBook[TASK_FILE_COMPONENTS_SHEET]
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

	assertComponentMax(productDictionary)
	return productDictionary


def assertComponentMax(productDictionary):
	for product in productDictionary:
		if len(productDictionary[product]) > PRODUCT_COMPONENT_MAX:
			print('Product ' + product + ' has over the max (' + str(PRODUCT_COMPONENT_MAX) + ') of assembly components.')
			quit()


def generateAssemblyOrderFile(taskList):
	createAssemblyOrderFile()
	try:
		populateAssemblyOrderFile(taskList)
		renameAssemblyOrderFile()
	except:
		print('Something went wrong! Cancelling process.')
		deleteAssemblyOrderFile()
		quit()


def populateAssemblyOrderFile(taskList):
	workBook = loadWorkbook(TEMPLATE, ASSEMBLY_ORDER_FILE_PATH)
	templateSheet = workBook[TEMPLATE_FILE_TEMPLATE_SHEET]
	taskCount = len(taskList)

	for i in range (0, taskCount):
		product = taskList[i].getProduct()
		newSheet = workBook.copy_worksheet(templateSheet)
		newSheet.title = generateSheetName(i, product)

		populateCell(newSheet, FRACTION, generateFraction(i, taskCount))
		populateCell(newSheet, PRODUCT, product)
		populateCell(newSheet, QUANTITY, taskList[i].getQuantity())

		currentRow = 8
		for component in taskList[i].getComponents():
			populateCell(newSheet, PART, component.getPart(), currentRow)
			populateCell(newSheet, PART_QUANTITY, component.getQuantity(), currentRow)
			currentRow = currentRow + 1
		formatSheet(newSheet)

	workBook.remove(templateSheet)
	workBook.save(ASSEMBLY_ORDER_FILE_PATH)
	

def generateSheetName(counter, product):
	return str(counter + 1) + " " + product[:27]


def generateFraction(counter, total):
	return str(counter + 1) + "/" + str(total)


def populateCell(sheet, dataElementType, value, currentRow = None):
	sheet.cell(row = calculateDataElementRow(dataElementType, currentRow), column = getDataElementColumn[dataElementType]).value = value;


def calculateDataElementRow(dataElementType, currentRow):
	if dataElementType not in [PART, PART_QUANTITY]:
		currentRow = getDataElementRow[dataElementType]

	if currentRow == None:
		print('Encountered error calculating target row for ' + dataElementType)
		quit()

	return currentRow

def formatSheet(sheet):
	addImage(sheet, GLACIER_TEK_IMAGE_FILE_PATH, 'A1')
	addSingleRowBoxBorder(sheet, 7, 1, 8)


def addSingleRowBoxBorder(sheet, targetRow, start, stop):
	border = Border(top = Side(style='thin', color='00000000'), bottom = Side(style='thin', color='00000000'))
	leftEndBorder = Border(top = Side(style='thin', color='00000000'), bottom = Side(style='thin', color='00000000'), left = Side(style='thin', color='00000000'))
	rightEndBorder = Border(top = Side(style='thin', color='00000000'), bottom = Side(style='thin', color='00000000'), right = Side(style='thin', color='00000000'))
	
	for targetColumn in range(start + 1, stop - 1):
		sheet.cell(row = targetRow, column = targetColumn).border = border

	sheet.cell(row = targetRow, column = start).border = leftEndBorder
	sheet.cell(row = targetRow, column = stop - 1).border = rightEndBorder

def addImage(sheet, imagePath, targetCell):
	image = Image(imagePath)
	# image.anchor(sheet.cell(row = targetRow, column = targetColumn))
	sheet.add_image(image, targetCell)
