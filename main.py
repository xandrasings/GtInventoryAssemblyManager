from Imports.Utilities.excel import *
from Imports.Utilities.fileManagement import *
from Imports.Utilities.sys import *
from Imports.Utilities.userInteraction import *

def main():
	clearTerminal()
	taskList = generateTaskList()
	createAssemblyOrderFile()
	try:
		populateAssemblyOrderFile(taskList)
		renameAssemblyOrderFile()
	except:
		print('Something went wrong! Cancelling process.')
		deleteAssemblyOrderFile()
		quit()


main()