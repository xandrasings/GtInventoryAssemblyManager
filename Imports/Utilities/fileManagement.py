from .const import *
from .dateTime import *

import shutil
import os

def generateAssemblyOrderFileName():
	return ASSEMBLY_ORDER_FILE_NAME_PREFIX + getDateTimeAsFilePathSegment() + EXCEL_FILE_NAME_SUFFIX

def generateAssemblyOrderFilePath():
	return os.path.join(ASSEMBLY_ORDERS_DIR_PATH, generateAssemblyOrderFileName())

def createAssemblyOrderFile():
	shutil.copy2(TEMPLATE_FILE_NAME_PATH, ASSEMBLY_ORDERS_DIR_PATH)

def renameAssemblyOrderFile():
	shutil.move(ASSEMBLY_ORDER_FILE_NAME_PATH, generateAssemblyOrderFilePath())

def deleteAssemblyOrderFile():
	os.remove(ASSEMBLY_ORDER_FILE_NAME_PATH)
