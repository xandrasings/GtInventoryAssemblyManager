from Imports.Utilities.excel import *

def main():
	clearTerminal()
	taskList = generateTaskList()
	generateAssemblyOrderFile(taskList)

main()
