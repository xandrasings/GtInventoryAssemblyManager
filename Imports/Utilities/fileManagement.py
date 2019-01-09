from .const import *

import shutil
import os

def createAssemblyOrderFile():
	shutil.copy2(TEMPLATE_FILE_NAME_PATH, ASSEMBLY_ORDERS_DIR_PATH)

def renameAssemblyOrderFile():
	destinationFilePath = os.path.join(ASSEMBLY_ORDERS_DIR_PATH, "sillyName.xlsx")
	shutil.move(TEMPLATE_FILE_NAME_PATH, destinationFilePath)

def deleteAssemblyOrderFile():
	os.remove(ASSEMBLY_ORDER_FILE_NAME_PATH)
