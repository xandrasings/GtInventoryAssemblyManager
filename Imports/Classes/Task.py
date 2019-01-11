from .Component import *

class Task:
	def __init__(self, product, quantity, components):
		self.product = product
		self.quantity = quantity
		self.components = components

	def getProduct(self):
		return self.product

	def getQuantity(self):
		return self.quantity

	def getComponents(self):
		return self.components

	def addComponent(self, component):
		self.components.append(component)

	def summarize(self):
		summary = "{product: " + self.product + ", quantity: " + str(self.quantity) + ", components: ["
		summary = summary + ",".join([component.summarize() for component in self.components])
		summary = summary + "]}"
		return summary