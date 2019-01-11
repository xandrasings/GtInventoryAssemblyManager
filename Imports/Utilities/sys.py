import sys

def quit():
	print('Exiting Assembly Manager.')
	sys.exit()

def fatalQuit(exception):
	print(exception)
	quit()