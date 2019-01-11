from Imports.Utilities.excel import *
from Imports.Utilities.fileManagement import *
from Imports.Utilities.userInteraction import *

def main():

	clearTerminal()
	taskList = generateTaskList()

	# TODO check existence of template file
	createAssemblyOrderFile()
	# in case of unexpected failure
	# deleteAssemblyOrderFile()
	renameAssemblyOrderFile()


main()