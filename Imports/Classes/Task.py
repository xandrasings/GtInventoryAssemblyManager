from .Component import *
from .Helpers.TaskHelper import *
from ..Utilities.constants import *

class Task:
	def __init__(self, index, product, quantity):
		self.index = index
		self.product = product
		self.quantity = quantity
		self.timeEstimate = calculateTimeEstimate(product.getTimeEstimate(), quantity)

	def __eq__(self, other):
		return self.index == other.index and self.isBelowThreshold() == other.isBelowThreshold()

	def __lt__(self, other):
		if self.isBelowThreshold() == other.isBelowThreshold():
			return self.getIndex() < other.getIndex()
		else:
			return self.isBelowThreshold()

	def isBelowThreshold(self):
		return (self.timeEstimate < TIME_THRESHOLD)

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

	def getReadableTimeEstimate(self):
		return summarizeTimeEstimate(self.timeEstimate)

	def summarize(self):
		return "{index: " + str(self.index) + ", product: " + self.product.summarize() + ", quantity: " + str(self.quantity) + ", time estimate: " + str(self.timeEstimate) + " or " + self.getReadableTimeEstimate() + "}"
		return summarize