class Product:
	def __init__(self, name, timeEstimate):
		self.name = name
		self.timeEstimate = timeEstimate
		self.components = []

	def getName(self):
		return self.name

	def getTimeEstimate(self):
		return self.timeEstimate

	def getComponents(self):
		return self.components

	def addComponent(self, component):
		self.components.append(component)

	def summarize(self):
		return "{name: " + self.name + ", time estimate: " + str(self.timeEstimate) + ", components: [" + ",".join([component.summarize() for component in self.components]) + "]}"