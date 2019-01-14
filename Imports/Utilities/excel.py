from .const import *
from .fileManagement import *
from .sys import *
from ..Classes.Component import *
from ..Classes.Task import *

from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.worksheet.page import PageMargins

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

	for i in range(2, taskSheet.max_row + 1):
		product = str(taskSheet.cell(row = i, column = TASK_LIST_PRODUCT_COLUMN).value)
		if product == "None":
			break
		quantity = taskSheet.cell(row = i, column = TASK_LIST_QUANTITY_COLUMN).value
		components = productDictionary[product]
		taskList.append(Task(product, quantity, components))

	return taskList


def generateProductDictionary(workBook):
	componentSheet = workBook[TASK_FILE_COMPONENTS_SHEET]
	verifyComponentSheetFormatting(componentSheet)
	productDictionary = {}
	for i in range(2, componentSheet.max_row + 1):
		product = str(componentSheet.cell(row = i, column = COMPONENT_LIST_PRODUCT_COLUMN).value)
		part = str(componentSheet.cell(row = i, column = COMPONENT_LIST_PART_COLUMN).value)
		quantity = componentSheet.cell(row = i, column = COMPONENT_LIST_QUANTITY_COLUMN).value

		if product not in productDictionary.keys():
			productDictionary[product] = []

		productDictionary[product].append(Component(part, quantity))

	assertComponentMax(productDictionary)
	return productDictionary


def verifyComponentSheetFormatting(sheet):
	# TODO assert worksheet dimensions from ws.calculate_dimension() sheet.get_highest_column()
	# TODO assert column headers are as expected
	# TODO use regex to assert expected formatting of rows
	pass


def assertComponentMax(productDictionary):
	for product in productDictionary:
		if len(productDictionary[product]) > PRODUCT_COMPONENT_MAX:
			print('Product ' + product + ' has over the max (' + str(PRODUCT_COMPONENT_MAX) + ') of assembly components.')
			quit()


def generateAssemblyOrderFile(taskList):
	createAssemblyOrderFile()
	try:
		workBook = loadWorkbook(TEMPLATE, ASSEMBLY_ORDER_FILE_PATH)
		populateAssemblyOrderFile(workBook, taskList)
		fileName = saveAndRenameWorkbook(workBook)
		print('Successfully generated assembly order file: ' + fileName)
	except Exception as e:
		print('Something went wrong! Cancelling process.')
		print('\n\n\nerror detail:')
		print(e)
		print('\n\n\n')
		deleteAssemblyOrderFile()
		quit()


def saveAndRenameWorkbook(workBook):
		workBook.save(ASSEMBLY_ORDER_FILE_PATH)
		return renameAssemblyOrderFile()


def populateAssemblyOrderFile(workBook, taskList):
	populateCopyrightPage(workBook)
	populateAssemblyOrderPages(workBook, taskList)


def populateCopyrightPage(workBook):
	copyrightSheet = workBook[TEMPLATE_FILE_COPYRIGHT_SHEET]
	addImage(copyrightSheet, VERTEX_IMAGE_FILE_PATH, 'B3')


def populateAssemblyOrderPages(workBook, taskList):
	templateSheet = workBook[TEMPLATE_FILE_TEMPLATE_SHEET]
	taskCount = len(taskList)

	for i in range (0, taskCount):
		product = taskList[i].getProduct()
		newSheet = workBook.copy_worksheet(templateSheet)
		newSheet.title = generateSheetName(i, product)
		pageMargins = PageMargins(top = margin[TOP], bottom = margin[BOTTOM], left = margin[LEFT], right = margin[RIGHT], header = margin[HEADER], footer = margin[FOOTER])

		populateCell(newSheet, FRACTION, generateFraction(i, taskCount))
		populateCell(newSheet, PRODUCT, product)
		populateCell(newSheet, QUANTITY, taskList[i].getQuantity())

		currentRow = PRODUCT_COMPONENT_START_ROW
		for component in taskList[i].getComponents():
			populateCell(newSheet, PART, component.getPart(), currentRow)
			populateCell(newSheet, PART_QUANTITY, component.getQuantity(), currentRow)
			currentRow = currentRow + 1
		formatAssemblyOrderSheet(newSheet, pageMargins)
	workBook.remove(templateSheet)
	

def generateSheetName(counter, product):
	return str(counter + 1) + " " + product[:27]


def generateFraction(counter, total):
	return str(counter + 1) + "/" + str(total)


def populateCell(sheet, dataElementType, value, currentRow = None):
	sheet.cell(row = calculateDataElementRow(dataElementType, currentRow), column = dataElementColumn[dataElementType]).value = value;


def calculateDataElementRow(dataElementType, currentRow):
	if dataElementType not in [PART, PART_QUANTITY]:
		currentRow = dataElementRow[dataElementType]

	if currentRow == None:
		print('Encountered error calculating target row for ' + dataElementType)
		quit()

	return currentRow


def addImage(sheet, imagePath, targetCell):
	image = Image(imagePath)
	sheet.add_image(image, targetCell)


def addProductAndQuantityUnderline(sheet):
	for dataElementType in [PRODUCT, QUANTITY]:
		addUnderline(sheet, dataElementRow[dataElementType] + dataElementRowBorderOffset[dataElementType], dataElementColumn[dataElementType] + dataElementColumnBorderOffset[dataElementType])


def addUnderline(sheet, targetRow, targetColumn):
	border = Border(bottom = Side(style='thin', color='00000000'))
	sheet.cell(row = targetRow, column = targetColumn).border = border


def addSingleRowBoxBorder(sheet, targetRow, start, stop):
	border = Border(top = Side(style='thin', color='00000000'), bottom = Side(style='thin', color='00000000'))
	leftEndBorder = Border(top = Side(style='thin', color='00000000'), bottom = Side(style='thin', color='00000000'), left = Side(style='thin', color='00000000'))
	rightEndBorder = Border(top = Side(style='thin', color='00000000'), bottom = Side(style='thin', color='00000000'), right = Side(style='thin', color='00000000'))
	
	for targetColumn in range(start + 1, stop - 1):
		sheet.cell(row = targetRow, column = targetColumn).border = border

	sheet.cell(row = targetRow, column = start).border = leftEndBorder
	sheet.cell(row = targetRow, column = stop - 1).border = rightEndBorder


def leftAlignAssemblyOrderNumbers(sheet):
	for targetRow in range (PRODUCT_COMPONENT_START_ROW, PRODUCT_COMPONENT_START_ROW + PRODUCT_COMPONENT_MAX):
		sheet.cell(row = targetRow, column = dataElementColumn[PART_QUANTITY]).alignment = Alignment(horizontal = 'left')
		sheet.cell(row = targetRow, column = dataElementColumn[ORDER_QUANTITY]).alignment = Alignment(horizontal = 'left')


def assignPageMargins(sheet, pageMargins):
	sheet.page_margins = pageMargins
	sheet.print_options.horizontalCentered = True


def assignPrintArea(sheet):
	sheet.print_area = PRINT_AREA


def formatAssemblyOrderSheet(sheet, pageMargins):
	addImage(sheet, GLACIER_TEK_IMAGE_FILE_PATH, 'A1')
	addProductAndQuantityUnderline(sheet)
	addSingleRowBoxBorder(sheet, PRODUCT_COMPONENT_START_ROW - 1, PRODUCT_COMPONENT_LIST_START_COLUMN, PRODUCT_COMPONENT_LIST_END_COLUMN + 1)
	leftAlignAssemblyOrderNumbers(sheet)
	assignPageMargins(sheet, pageMargins)
	assignPrintArea(sheet)
