from Imports.Utilities.excel import *

import os

def main():
	os.system('cls' if os.name == 'nt' else 'clear')
	taskList = generateTaskList()
	generateAssemblyOrderFile(taskList)

main()
