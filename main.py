from Imports.Classes.Product import *
from Imports.Utilities.fileManagement import *

def main():

	productDictionary = {}
	p = Product("name")
	p.addComponent(Component("thingy",4))
	p.addComponent(Component("thungy",5))
	productDictionary["test"] = p
	print(productDictionary["test"].summarize())

	# TODO check existence of template file
	createAssemblyOrderFile()
	renameAssemblyOrderFile()

	# deleteAssemblyOrderFile()

main()