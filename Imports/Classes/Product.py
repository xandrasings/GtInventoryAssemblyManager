from .Component import *

class Product:
	def __init__(self, name):
		self.name = name
		self.components = []

	def getName(self):
		return self.name

	def getComponents(self):
		return self.components

	def addComponent(self, component):
		self.components.append(component)

	def summarize(self):
		summary = "{name: " + self.name + ", components: ["
		summary = summary + ",".join([component.summarize() for component in self.components])
		summary = summary + "]}"
		return summary