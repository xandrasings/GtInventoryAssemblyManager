from Imports.Classes.Product import *
from Imports.Utilities.excel import *
from Imports.Utilities.fileManagement import *

def main():

	productDictionary = generateProductDictionary()

	# TODO check existence of template file
	createAssemblyOrderFile()
	# in case of unexpected failure
	# deleteAssemblyOrderFile()
	renameAssemblyOrderFile()


main()