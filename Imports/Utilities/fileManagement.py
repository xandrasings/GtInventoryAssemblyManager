from .const import *
from .dateTime import *
from .sys import *
from .userInteraction import *

from os import listdir

import shutil
import os

def extendPath(filePath, extension):
	return os.path.join(filePath, extension)


def selectTaskFile():
	return extendPath(TASK_LISTS_DIR_PATH, promptFileSelection())


def promptFileSelection():
	validFiles = getValidTaskFiles()
	print('Which task file would you like to use to generate assembly orders?')
	optionIndex = solicitPathOptionIndex(validFiles)
	optionName = validFiles[optionIndex]
	return optionName


def getValidTaskFiles():
	validFiles = []

	directoryContent = listdir(TASK_LISTS_DIR_PATH)

	for item in list(directoryContent):
		if isValidExcelFile(item):
			validFiles.append(item)

	return validFiles

def isValidExcelFile(item):
	return (
		(
			item.endswith('.xlsx') or
			item.endswith('.xls')
		) and (	
			not item.startswith('~')
		)
	)


def solicitPathOptionIndex(validOptions):
	printFileOptions(validOptions)
	optionIndex = getPathIndex(validOptions)
	return optionIndex


def printFileOptions(pathContent):
	i = 1
	for item in list(pathContent):
		printOption('- (' + str(i) + ') ' + item)
		i = i + 1
	printOption('- (Q)uit')


def getPathIndex(validOptions):
	fileIndex = -1
	maxInput = len(validOptions)

	while fileIndex < 0:
		inputIndex = prompt()
		try:
			fileIndex = int(inputIndex) - 1
			if fileIndex < 0 or fileIndex >= maxInput:
				print('Selection should be between 1 and ' + str(maxInput))
				fileIndex = -1
		except:
			if inputIndex == 'Q':
				quit()
			print('Selection should be an integer')

	return fileIndex


def generateAssemblyOrderFileName():
	return ASSEMBLY_ORDER_FILE_NAME_PREFIX + getDateTimeAsFilePathSegment() + EXCEL_FILE_NAME_SUFFIX


def generateAssemblyOrderFilePath(fileName):
	return extendPath(ASSEMBLY_ORDERS_DIR_PATH, fileName)


def createAssemblyOrderFile():
	if not os.path.exists(TEMPLATE_FILE_NAME_PATH):
		print("Template file does not exist at " + TEMPLATE_FILE_NAME_PATH)
		quit()
	shutil.copy2(TEMPLATE_FILE_NAME_PATH, ASSEMBLY_ORDERS_DIR_PATH)


def renameAssemblyOrderFile():
	fileName = generateAssemblyOrderFileName()
	shutil.move(ASSEMBLY_ORDER_FILE_PATH, generateAssemblyOrderFilePath(fileName))
	return fileName


def deleteAssemblyOrderFile():
	os.remove(ASSEMBLY_ORDER_FILE_PATH)

