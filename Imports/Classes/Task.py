from .Component import *
from .Helpers.TaskHelper import *

class Task:
	def __init__(self, index, product, quantity):
		self.index = index
		self.product = product
		self.quantity = quantity
		self.timeEstimate = calculateTimeEstimate(product.getTimeEstimate(), quantity)

	def getIndex(self):
		return self.index

	def getProduct(self):
		return self.product

	def getComponents(self):
		return self.product.getComponents()

	def getQuantity(self):
		return self.quantity

	def getTimeEstimate(self):
		return self.timeEstimate

	def summarize(self):
		return "{index: " + str(self.index) + ", product: " + self.product.summarize() + ", quantity: " + str(self.quantity) + ", time estimate: " + self.timeEstimate + "]}"
		return summarize