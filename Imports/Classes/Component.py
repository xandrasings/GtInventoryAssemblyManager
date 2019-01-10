class Component:
	def __init__(self, part, quantity):
		self.part = part
		self.quantity = quantity

	def getPart(self):
		return self.part

	def getQuantity(self):
		return self.quantity

	def summarize(self):
		return "{part: " + self.part + ", quantity: " + str(self.quantity) + "}"