import os

CURRENT_DIR = os.curdir
REFERENCE_FILES_DIR = "ReferenceFiles"
REFERENCE_FILES_DIR_PATH = os.path.join(CURRENT_DIR, REFERENCE_FILES_DIR)

TASK_LISTS_DIR = "TaskLists"
TASK_LISTS_DIR_PATH = os.path.join(REFERENCE_FILES_DIR_PATH, TASK_LISTS_DIR)

TEMPLATES_DIR = "Templates"
TEMPLATES_DIR_PATH = os.path.join(REFERENCE_FILES_DIR_PATH, TEMPLATES_DIR)
TEMPLATE_FILE_NAME = "assemblyOrderTemplate.xlsx"
TEMPLATE_FILE_NAME_PATH = os.path.join(TEMPLATES_DIR_PATH, TEMPLATE_FILE_NAME)

ASSEMBLY_ORDERS_DIR = "AssemblyOrders"
ASSEMBLY_ORDERS_DIR_PATH = os.path.join(REFERENCE_FILES_DIR_PATH, ASSEMBLY_ORDERS_DIR)
ASSEMBLY_ORDER_FILE_PATH = os.path.join(ASSEMBLY_ORDERS_DIR_PATH, TEMPLATE_FILE_NAME)
ASSEMBLY_ORDER_FILE_NAME_PREFIX = "assembly_order_"
EXCEL_FILE_NAME_SUFFIX = ".xlsx"

IMAGES_DIR = "Images"
IMAGES_DIR_PATH = os.path.join(REFERENCE_FILES_DIR_PATH, IMAGES_DIR)
GLACIER_TEK_IMAGE_FILE_NAME = 'GlacierTek.png'
GLACIER_TEK_IMAGE_FILE_PATH = os.path.join(IMAGES_DIR_PATH, GLACIER_TEK_IMAGE_FILE_NAME)
VERTEX_IMAGE_FILE_NAME = 'Vertex42.png'
VERTEX_IMAGE_FILE_PATH = os.path.join(IMAGES_DIR_PATH, VERTEX_IMAGE_FILE_NAME)

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

PRODUCT_COMPONENT_MAX = 26

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
