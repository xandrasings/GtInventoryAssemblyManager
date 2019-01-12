import os

CURRENT_DIR = os.curdir
EXCEL_FILES_DIR = "ExcelFiles"
EXCEL_FILES_DIR_PATH = os.path.join(CURRENT_DIR, EXCEL_FILES_DIR)

TASK_LISTS_DIR = "TaskLists"
TASK_LISTS_DIR_PATH = os.path.join(EXCEL_FILES_DIR_PATH, TASK_LISTS_DIR)

TEMPLATES_DIR = "Templates"
TEMPLATES_DIR_PATH = os.path.join(EXCEL_FILES_DIR_PATH, TEMPLATES_DIR)
TEMPLATE_FILE_NAME = "assemblyOrderTemplate.xlsx"
TEMPLATE_FILE_NAME_PATH = os.path.join(TEMPLATES_DIR_PATH, TEMPLATE_FILE_NAME)

ASSEMBLY_ORDERS_DIR = "AssemblyOrders"
ASSEMBLY_ORDERS_DIR_PATH = os.path.join(EXCEL_FILES_DIR_PATH, ASSEMBLY_ORDERS_DIR)
ASSEMBLY_ORDER_FILE_NAME_PATH = os.path.join(ASSEMBLY_ORDERS_DIR_PATH, TEMPLATE_FILE_NAME)
ASSEMBLY_ORDER_FILE_NAME_PREFIX = "assembly_order_"
EXCEL_FILE_NAME_SUFFIX = ".xlsx"

TASKS = 'TASKS'
TASK_FILE_TASK_SHEET = 'tasks'
TASK_FILE_COMPONENTS_SHEET = 'components'
TASK_FILE_SHEETS = [TASK_FILE_TASK_SHEET, TASK_FILE_COMPONENTS_SHEET]

TEMPLATE = 'TEMPLATE'
TEMPLATE_FILE_TEMPLATE_SHEET = 'template'
TEMPLATE_FILE_SHEETS = [TEMPLATE_FILE_TEMPLATE_SHEET]

FILE_SHEETS = {
	TASKS : TASK_FILE_SHEETS,
	TEMPLATE : TEMPLATE_FILE_SHEETS
}

PRODUCT_COMPONENT_MAX = 30

PRODUCT = 'product'
FRACTION = 'fraction'
QUANTITY = 'quantity'
PART = 'part'
PART_QUANTITY = 'part_quantity'

getDataElementRow = {
	FRACTION : 1,
	PRODUCT : 2,
	QUANTITY : 4
}

getDataElementColumn = {
	FRACTION : 7,
	PRODUCT : 4,
	QUANTITY : 4,
	PART : 1,
	PART_QUANTITY : 4
}
