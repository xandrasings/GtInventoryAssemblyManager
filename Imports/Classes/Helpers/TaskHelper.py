def calculateTimeEstimate(timeEstimate, quantity):
	product = (timeEstimate * quantity) / 60
	if product < .8:
		return "< 1 hour"
	elif product < 1.5:
		return "~ 1 hour"
	else:
		return "~ " + str(round(product)) + " hours"