def calculateTimeEstimate(timeEstimate, quantity):
	product = (timeEstimate * quantity) / 60
	if product < .8:
		return 0
	else:
		return round(product)

def summarizeTimeEstimate(timeEstimate):
    conversionDictionary = {
        0: '< 1 hour',
        1: '~ 1 hour'
    }
    return conversionDictionary.get(timeEstimate, '~ ' + str(timeEstimate) + ' hours')