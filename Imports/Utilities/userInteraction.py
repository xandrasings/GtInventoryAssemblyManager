import os


def clearTerminal():
	os.system('cls' if os.name == 'nt' else 'clear')


def printOption(text):
	leftIndex = text.index('(')
	rightIndex = text.index(')') + 1
	print(text[0:leftIndex] + text[leftIndex:rightIndex] + text[rightIndex:])


def prompt(text = ''):
	return input(text + ' > ').upper()